import streamlit as st
import pandas as pd
import io, urllib.parse, hashlib, sqlite3
from datetime import datetime, date, timedelta
import plotly.graph_objects as go
from fpdf import FPDF
from docx import Document
from pptx import Presentation
from streamlit_js_eval import get_geolocation

# --- 1. CONFIGURAÇÃO DE UI/UX SUPREME ---
st.set_page_config(page_title="SGI Apex V76", layout="wide", page_icon="🛡️")

st.markdown("""
    <style>
    header {visibility: hidden;}
    .stApp { background-color: #05070A; color: #E0E6ED; }
    .stButton>button { width: 100%; border-radius: 12px; height: 4em; font-weight: bold; background: linear-gradient(135deg, #00D1FF, #0052D4); color: white; border: none; font-size: 16px; box-shadow: 0px 4px 15px rgba(0, 209, 255, 0.3); }
    div[data-testid="stMetric"] { background-color: rgba(28, 37, 51, 0.9); padding: 25px; border-radius: 15px; border-bottom: 6px solid #00D1FF; }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] { background-color: #1C2533; border-radius: 10px; padding: 12px 25px; color: white; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. MOTOR DE INTELIGÊNCIA JURÍDICO-EXECUTIVA ---
def motor_apex_omega(dados):
    v = {"Baixa": 45000, "Média": 250000, "Alta": 950000, "Crítica": 6500000}
    tcr = v.get(dados['grav'], 50000) * 4.8
    leis = {
        "OFFSHORE / PLATAFORMA": "NR-37 + NORMAM-204/DPC + NR-13 + Lei 5.811/72.",
        "AEROPORTO / AVIAÇÃO": "RBAC 107 (AVSEC) + RBAC 153/155 + RBAC 117 (Fadiga) + ANAC.",
        "PORTO / MARÍTIMO": "ISPS Code + NORMAM-01 + Antaq + NR-29 (Portuário).",
        "LOGÍSTICA / GALPÃO": "NR-11 + NR-12 + PMOC + IT-17 Bombeiros + ISO 22301.",
        "HOSPITALAR / SAÚDE": "RDC 222 ANVISA + NR-32 + ISO 9001:2026 + ANVISA.",
        "SEGURANÇA PRIVADA": "Lei 14.967/2024 (Estatuto) + Portaria 18.045/23 PF + Art. 482 CLT."
    }
    lei_ref = leis.get(dados['setor'], "Legislação Federal e SGI Master")
    texto = f"Prezado Cliente, informamos a resolução técnica no posto {dados['local']} relativa ao alvo {dados['alvo']}.\n\n1. PARECER: O desvio identificado infringe a {lei_ref}.\n2. TRATAMENTO: Aplicamos a {dados['medida']} imediata.\n3. IMPACTO: Evitamos um TCR de R$ {tcr:,.2f}.\n\nHash SHA-512: {hashlib.sha512(dados['relato'].encode()).hexdigest()[:16].upper()}"
    return texto, tcr

# --- 3. INTERFACE TITAN ---
st.title("🛡️ SGI Alpha Apex V76")
tabs = st.tabs(["🧐 Auditoria & Devolutiva", "⚙️ Ativos/LTO", "👥 Equipe/CIPA", "📊 Dashboard BI", "🚀 Boutique Executive"])

with tabs[0]:
    with st.form("audit_v76"):
        c1, c2, c3 = st.columns(3)
        setor = c1.selectbox("Setor", ["OFFSHORE / PLATAFORMA", "AEROPORTO / AVIAÇÃO", "PORTO / MARÍTIMO", "LOGÍSTICA / GALPÃO", "HOSPITALAR / SAÚDE", "SEGURANÇA PRIVADA"])
        local = c1.text_input("Unidade / Posto")
        alvo = c2.text_input("Alvo (Pessoa ou Ativo)")
        grav = c2.select_slider("Criticidade", ["Baixa", "Média", "Alta", "Crítica"])
        medida = c3.selectbox("Ação:", ["Advertência", "Suspensão", "Substituição", "Manutenção", "Justa Causa", "Plano CIPA"])
        foto = c3.camera_input("📸 Foto da Solução")
        relato = st.text_area("Descrição da Situação:")
        
        if st.form_submit_button("⚖️ EXECUTAR INTELIGÊNCIA SOBERANA"):
            txt, tcr = motor_apex_omega({'setor': setor, 'grav': grav, 'alvo': alvo, 'local': local, 'medida': medida, 'relato': relato})
            st.session_state['v76'] = {'txt': txt, 'tcr': tcr, 'local': local, 'alvo': alvo}
            st.success("✅ Processado com Sucesso!")

with tabs[3]:
    if 'v76' in st.session_state:
        st.metric("ROI de Gestão (Lucro Protegido)", f"R$ {st.session_state['v76']['tcr']:,.2f}")
    st.info("💡 Gráficos e indicadores de performance para apresentação executiva.")

with tabs[4]:
    if 'v76' in st.session_state:
        d = st.session_state['v76']; u = urllib.parse.quote(d['txt'])
        st.subheader("📲 Canais de Despacho (WhatsApp)")
        st.markdown(f'<a href="https://wa.me/?text={u}" target="_blank"><button style="background:#25D366; width:100%; height:50px; color:white; border:none; border-radius:10px; font-weight:bold;">📲 WHATSAPP CLIENTE</button></a>', unsafe_allow_html=True)
        st.divider()
        st.subheader("💾 Exportação de Arquivos")
        st.download_button("📄 PDF DOSSIÊ", d['txt'], f"SGI_{d['local']}.pdf")
        st.text_area("Devolutiva Gerada:", d['txt'], height=300)
    else: st.warning("Faça uma auditoria primeiro.")
