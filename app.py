import streamlit as st
import pandas as pd
from datetime import datetime
from engine import (
    analyze_call, analyze_chat,
    build_excel_bytes,
    IHLAL_TYPES_CALL, IHLAL_TYPES_CHAT,
    fmt
)

st.set_page_config(
    page_title="Vardiya Uyumsuzluk Analizi",
    page_icon="📋",
    layout="wide"
)

# ── Stil ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
.main-title { font-size: 28px; font-weight: 700; color: #2E5FA3; margin-bottom: 4px; }
.sub-title  { font-size: 14px; color: #666; margin-bottom: 24px; }
.metric-box {
    background: #F0F4FF; border-radius: 10px; padding: 16px 20px;
    border-left: 4px solid #2E5FA3; margin-bottom: 8px;
}
.metric-label { font-size: 12px; color: #666; margin: 0; }
.metric-value { font-size: 22px; font-weight: 700; color: #2E5FA3; margin: 0; }
.ihlal-row { display: flex; justify-content: space-between; padding: 6px 0;
             border-bottom: 1px solid #EEE; }
</style>
""", unsafe_allow_html=True)

# ── Başlık ────────────────────────────────────────────────────────────────────
st.markdown('<p class="main-title">📋 Vardiya Uyumsuzluk Analizi</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Vardiya planı ve statü raporunu yükleyerek ihlal analizi yapın.</p>', unsafe_allow_html=True)

# ── Ekip konfigürasyonları ────────────────────────────────────────────────────
EKIP_CONFIG = {
    "OA Satış (Mplus)": {
        "tip": "call",
        "agent_col": "Temsilci Adı",
        "fix_year": False,
        "break_limit": 1800,
        "lunch_limit": 1800,
        "call_cross": False,
        "dosyalar": ["vardiya", "statu"],
    },
    "OA Call (Mplus)": {
        "tip": "call",
        "agent_col": "Temsilci",
        "fix_year": True,
        "break_limit": 1800,
        "lunch_limit": 1800,
        "call_cross": True,   # Chat Invisible ile çapraz
        "name_map": {"esma çiftci": "Esma Çiftçi"},
        "dosyalar": ["vardiya", "statu", "chat_statu"],
    },
    "OA Chat (Mplus)": {
        "tip": "chat",
        "agent_col": "Temsilci Adı",
        "fix_year": False,
        "break_limit": 1800,
        "lunch_limit": 1800,
        "name_map": {"esma çiftci": "Esma Çiftçi"},
        "dosyalar": ["vardiya", "statu", "call_statu"],
    },
    "OU Call (Mplus)": {
        "tip": "call",
        "agent_col": "Temsilci",
        "fix_year": True,
        "break_limit": 1800,
        "lunch_limit": 1800,
        "call_cross": False,
        "dosyalar": ["vardiya", "statu"],
    },
    "OU Chat (Mplus)": {
        "tip": "chat",
        "agent_col": "Temsilci",
        "fix_year": True,
        "break_limit": 1800,
        "lunch_limit": 1800,
        "dosyalar": ["vardiya", "statu", "call_statu"],
        "vardiya_sheet": "Sayfa1",
    },
    "OU Chat (Concentrix)": {
        "tip": "chat",
        "agent_col": "Temsilci",
        "fix_year": False,
        "break_limit": 2700,
        "lunch_limit": 2700,
        "dosyalar": ["vardiya", "statu", "call_statu"],
    },
}

# ── Ekip seçimi ───────────────────────────────────────────────────────────────
st.markdown("### 1️⃣ Ekip Seç")
ekip = st.selectbox("Ekip", list(EKIP_CONFIG.keys()), label_visibility="collapsed")
cfg  = EKIP_CONFIG[ekip]

st.markdown("---")

# ── Dosya yükleme ─────────────────────────────────────────────────────────────
st.markdown("### 2️⃣ Dosyaları Yükle")

col1, col2, col3 = st.columns(3)

with col1:
    vardiya_file = st.file_uploader(
        "📅 Vardiya Planı (.xlsx)",
        type=["xlsx","xlsm"],
        key="vardiya"
    )
with col2:
    statu_file = st.file_uploader(
        "📊 Statü Raporu (.xlsx / .csv)",
        type=["xlsx","xlsm","csv"],
        key="statu"
    )
with col3:
    call_statu_file = None
    if "call_statu" in cfg.get("dosyalar",[]) or "chat_statu" in cfg.get("dosyalar",[]):
        label = "📞 Call Statü Raporu (.xlsx / .csv)" if cfg["tip"]=="chat" else "💬 Chat Statü Raporu (.xlsx)"
        call_statu_file = st.file_uploader(
            label,
            type=["xlsx","xlsm","csv"],
            key="call_statu"
        )

st.markdown("---")

# ── Analiz ────────────────────────────────────────────────────────────────────
st.markdown("### 3️⃣ Analizi Başlat")

if st.button("🔍 Rapor Oluştur", type="primary", use_container_width=True):

    # Dosya kontrolü
    required = [vardiya_file, statu_file]
    if ("call_statu" in cfg.get("dosyalar",[]) or
        "chat_statu" in cfg.get("dosyalar",[])):
        required.append(call_statu_file)

    if any(f is None for f in required):
        st.error("⚠️ Lütfen gerekli tüm dosyaları yükleyin.")
        st.stop()

    with st.spinner("Analiz yapılıyor..."):

        # Vardiya oku
        sheet = cfg.get("vardiya_sheet", 0)
        df_vd = pd.read_excel(vardiya_file, sheet_name=sheet)
        df_vd['agent'] = df_vd[cfg["agent_col"]].str.strip()
        date_cols = [c for c in df_vd.columns if isinstance(c, datetime)]
        if cfg["fix_year"]:
            dcf = {col: col.replace(year=2026) for col in date_cols}
        else:
            dcf = {col: col for col in date_cols}

        # Statü oku
        if statu_file.name.endswith(".csv"):
            df_st = pd.read_csv(statu_file, sep=';')
        else:
            df_st = pd.read_excel(statu_file)

        # Call/Chat çapraz statü
        df_cross = None
        if call_statu_file is not None:
            if call_statu_file.name.endswith(".csv"):
                df_cross = pd.read_csv(call_statu_file, sep=';')
            else:
                df_cross = pd.read_excel(call_statu_file, engine='openpyxl')
            if 'Musteri_Temsilcisi' in df_cross.columns:
                df_cross['agent'] = df_cross['Musteri_Temsilcisi'].str.strip()

        name_map = cfg.get("name_map", None)

        # ── CALL analizi ──────────────────────────────────────────────────────
        if cfg["tip"] == "call":
            df_st['Tarih'] = pd.to_datetime(df_st['Tarih'], errors='coerce')
            chat_st = df_cross if cfg.get("call_cross") else None
            df_v, ew, login_data, ss = analyze_call(
                df_st, df_vd, cfg["agent_col"], date_cols, dcf,
                df_chat_status=chat_st, name_map=name_map
            )
            all_agents = sorted(set(k[0] for k in ss.keys()))
            excel_bytes = build_excel_bytes(
                df_v, ew, login_data, IHLAL_TYPES_CALL,
                sistem_kes=None, all_agents=all_agents
            )
            ihlal_types = IHLAL_TYPES_CALL

        # ── CHAT analizi ──────────────────────────────────────────────────────
        else:
            df_v, ew, login_data, sistem_kes, ss = analyze_chat(
                df_st, df_vd, cfg["agent_col"], date_cols, dcf,
                df_call=df_cross,
                break_limit=cfg["break_limit"],
                lunch_limit=cfg["lunch_limit"],
                name_map=name_map
            )
            all_agents = sorted(set(k[0] for k in ss.keys()))
            excel_bytes = build_excel_bytes(
                df_v, ew, login_data, IHLAL_TYPES_CHAT,
                sistem_kes=sistem_kes, all_agents=all_agents
            )
            ihlal_types = IHLAL_TYPES_CHAT

    # ── Sonuçlar ──────────────────────────────────────────────────────────────
    st.success("✅ Analiz tamamlandı!")
    st.markdown("---")
    st.markdown("### 📊 Özet")

    total_ihlal = len(df_v)
    total_sure  = df_v['İhlal Süresi (sn)'].sum() if len(df_v) > 0 else 0
    total_tamam = sum(ew.values())
    total_login = sum(login_data.values())

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("Toplam İhlal", f"{total_ihlal:,}")
    with c2:
        st.metric("Toplam İhlal Süresi", fmt(total_sure))
    with c3:
        st.metric("Tamamlatılan Süre", fmt(total_tamam))
    with c4:
        st.metric("Login Süresi", fmt(total_login))

    # İhlal türü dağılımı
    if len(df_v) > 0:
        st.markdown("**İhlal Türü Dağılımı**")
        counts = df_v['İhlal Türü'].value_counts()
        cols = st.columns(len(ihlal_types))
        for i, ihlal in enumerate(ihlal_types):
            with cols[i]:
                adet = counts.get(ihlal, 0)
                if adet > 0:
                    st.metric(ihlal, adet)

    # Chat için sistem kesintisi
    if cfg["tip"] == "chat" and sistem_kes:
        total_kes = sum(sistem_kes.values())
        if total_kes > 0:
            st.info(f"⚡ Sistem kesintisi nedeniyle sayılmayan Invisible toplamı: **{fmt(total_kes)}**")

    st.markdown("---")

    # İndir butonu
    st.download_button(
        label="⬇️ Excel Raporu İndir",
        data=excel_bytes,
        file_name=f"{ekip.replace(' ','_').replace('(','').replace(')','')}_Vardiya_Uyumsuzluk.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary"
    )

    # Önizleme
    if len(df_v) > 0:
        with st.expander("📋 İlk 50 satır önizleme"):
            st.dataframe(
                df_v.head(50)[['Temsilci','Tarih','Vardiya','İhlal Türü','Detay']],
                use_container_width=True,
                height=400
            )
