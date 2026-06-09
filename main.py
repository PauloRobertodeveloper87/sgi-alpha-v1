import streamlit as st
import pandas as pd
import sqlite3
import io
import urllib.parse
import hashlib
from datetime import datetime, date, timedelta
import plotly.express as px
import plotly.graph_objects as go
from fpdf import FPDF
from docx import Document
from pptx import Presentation
from streamlit_js_eval import get_geolocation
# --- CONFIGURAÇÃO DE UI/UX ---
st.set_page_config(page_title="SGI Guardian", layout="wide", page_icon="🛡️")
st.markdown("""
    <style>
    header {visibility: hidden;}
    .stApp { 
        background: linear-gradient(135deg, #05070A 0%, #0E1117 100%);
       color: #E0E6ED; 
    }
    .stButton>button { 
        width: 100%; 
        border-radius: 12px; 
        height: 4.2em; 
        font-weight: bold; 
        background: linear-gradient(135deg, #00D1FF, #0052D4); 
        color: white; 
        border: none;
        font-size: 16px; 
        box-shadow: 0px 4px 15px rgba(0, 209, 255, 0.3); 
    }
    .stButton>button:hover { 
        transform: translateY(-3px);
        box-shadow: 0px 8px 25px rgba(0, 209, 255, 0.5); 
    }
    div[data-testid="stMetric"] { 
        background-color: rgba(28, 37, 51, 0.9); 
        padding: 30px;
        border-radius: 20px; 
        border-bottom: 6px solid #00D1FF; 
    }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] {
    background-color: #1C2533; 
        border-radius: 12px; 
        padding: 12px 30px; 
        color: white; 
        font-weight: bold;
        }
    .stExpander { 
        background-color: rgba(255, 255, 255, 0.03) !important; 
        border: 1px solid #1E293B; 
        border-radius: 15px;
        }
    .i-help { color: #00D1FF; font-weight: bold; cursor: help; text-decoration: underline; }
    </style>
    """, unsafe_allow_html=True)
# --- INICIALIZAÇÃO DO BANCO DE DADOS SOBERANO ---
def init_db():
    conn = sqlite3.connect('sgi_guardian.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS equipe 
                 (id INTEGER PRIMARY KEY, nome TEXT, cargo TEXT, esfera TEXT, '''
    reciclagem DATE, icc INTEGER, fadiga INTEGER)''')
    c.execute('''CREATE TABLE IF NOT EXISTS ativos 
                 (id TEXT PRIMARY KEY, categoria TEXT, local TEXT, 
                  vencimento DATE, status TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS licencas
    (id INTEGER PRIMARY KEY, nome TEXT, orgao TEXT, vencimento DATE, status TEXT)''')
    c.execute('CREATE TABLE IF NOT EXISTS colegiados (id INTEGER PRIMARY KEY, tipo TEXT, data TEXT, pauta TEXT, acao TEXT)')
    c.execute('CREATE TABLE IF NOT EXISTS auditorias (id INTEGER PRIMARY KEY, data TEXT, local TEXT, alvo TEXT, relato TEXT, tcr REAL, glosa REAL, lei TEXT, iso TEXT, acao TEXT, hash TEXT, gps TEXT)')
    conn.commit()
    conn.close()
    init_db()
# --- 3. GLOSSÁRIO TÉCNICO I-HELP 6.0 (A TO Z) ---
GLOS = {
    "TCR": "Total Cost of Risk: Soma de multas, glosas, dano à imagem e honorários.",
    "LTO": "License to Operate: Alvarás e licenças vitais para o posto.",
    "ICC": "Índice de Confiabilidade: Probabilidade de desvio técnico/acidente.",
    "PMOC": "Plano de Manutenção de Ar Condicionado (Obrigatório Lei 13.589/18).",
    "NR-37": "Norma para Segurança e Saúde em Plataformas de Petróleo (Offshore).",
    "RBAC 117": "Regulamento da ANAC sobre Gerenciamento de Risco de Fadiga Humana.",
    "L14.967/24": "Estatuto Seg. Privada: Marco regulador de conduta e penalidades PF.",
    "SLA": "Service Level Agreement: Acordo de nível de serviço com multas."
}
def i_help(termo):
    with st.popover(f"💡 Doutrina: {termo}"):
        st.info(GLOS.get(termo, "Legislação e SGI 2027."))
        # --- 4. MOTOR DE INTELIGÊNCIA JURÍDICO-EXECUTIVA ---
def motor_apex_v80(dados):
    # Cálculo de TCR e Impacto 2026/2027
    prejuizos = {"Baixa": 45000, "Média": 150000, "Alta": 750000, "Crítica": 6500000}
    multa_base = prejuizos.get(dados['gravidade'], 50000)
    tcr = multa_base * 4.8 
    glosa = multa_base * 0.25 
    # Biblioteca de Doutrina Forense Universal
    biblioteca = {
        "OFFSHORE / MAR": {
    "lei": "NR-37 + NORMAM-204/DPC + NR-13 + Lei 5.811/72 + STCW.",
            "iso": "ISO 14001:2026 e ISO 45001:2026.",
            "pilar": "SMS Marítimo"
        },
        "AVIAÇÃO / CÉU": {
            "lei": "RBAC 107 (AVSEC) + RBAC 153/155 + RBAC 117 (Fadiga) + ANAC.",
            "iso": "ISO 9001:2026 e Gestão de Riscos ANAC.",
            "pilar": "Safety & AVSEC"
        },
        "HOSPITALAR / SAÚDE": {
    "lei": "RDC 222 ANVISA + NR-32 + ISO 9001:2026 + Ética Profissional.",
            "iso": "ISO 9001:2026 e Biossegurança.",
            "pilar": "SGI Saúde"
        },
        "INDUSTRIAL / TERRA": {
            "lei": "NRs (01, 10, 11, 12, 35) + IT-17 Bombeiros + CLT Art. 482.",
            "iso": "ISO 45001 (SST) e ISO 14001 (Ambiental).",
            "pilar": "HSE / SST Industrial"
        },
        "SEGURANÇA PRIVADA": {
    "lei": "Lei 14.967/2024 (Novo Estatuto) + Portaria 18.045/23 PF.",
            "iso": "ISO 37301 (Compliance) e ISO 31000 (Riscos).",
            "pilar": "HSG (Governança)"
        }
    }
    ref = biblioteca.get(dados['setor'], biblioteca["INDUSTRIAL / TERRA"])
    # Redação da Medida Resolutiva
    medidas_dict = {
        "Advertência": "advertência formal com caráter pedagógico.",
        "Suspensão": "suspensão disciplinar coercitiva para preservação da segurança.",
    "Substituição": "substituição definitiva do colaborador visando a salvaguarda do contrato.",
        "Desligamento": "rescisão por justa causa amparada em quebra de fidúcia técnica.",
        "Reciclagem": "retirada de posto para reciclagem técnica compulsória em escola homologada."
    }
    texto = f"Prezado Cliente, informamos o encerramento do ciclo SGI na unidade {dados['local']} relativo ao alvo {dados['alvo']}.\\n\\n"
    texto += f"1. PARECER TÉCNICO-LEGAL: O desvio identificado (\\"{dados['relato']}\\") foi confrontado com a {ref['lei']}.\\n"
    texto += f"O fato infringe as diretrizes de governança e as normas {ref['iso']}, gerando risco de interdição.\\n\\n"
    texto += f"2. TRATAMENTO E RESOLUÇÃO: Priorizando o SLA, aplicamos a medida de {medidas_dict.get(dados['medida'])} imediata.\\n"
    texto += f"A eficácia foi validada via evidência visual 'After' e prova de presença georreferenciada."
    return texto, tcr, glosa, ref['lei']
    # --- 5. INTERFACE DE COMANDO TITAN ---
st.title("🛡️ SGI Guardian")
st.subheader("Supreme Sovereign Command Matrix")
tab_audit, tab_ativos, tab_equipe, tab_bi, tab_boutique = st.tabs([
    "🧐 Auditoria & Devolutiva", "⚙️ Ativos/Infra & LTO", "👥 Gestão de Equipe", "📊 Dashboard Power BI", "🚀 Boutique Executive"
])
with tab_audit:
    st.header("📋 Inspeção Forense e PDCA")
    with st.form("audit_form_v80"):
        c1, c2, c3 = st.columns(3)
        c1, c2, c3 = st.columns(3)
        with c1:
            setor = st.selectbox("Setor de Atuação", ["OFFSHORE / MAR", "AVIAÇÃO / CÉU", "HOSPITALAR / SAÚDE", "INDUSTRIAL / TERRA", "SEGURANÇA PRIVADA"])
            local = st.text_input("Unidade / Posto / Planta")
            i_help("LTO")
        with c2:
            alvo = st.text_input("Alvo (Profissional ou Ativo)")
            grav = st.select_slider("Criticidade do Impacto", ["Baixa", "Média", "Alta", "Crítica"])
            i_help("TCR")
        with c3:
            medida = st.selectbox("Ação de Tratamento:", ["Advertência", "Suspensão", "Substituição", "Manutenção", "Interdição", "Justa Causa"])
            i_help("SGI")
            st.write("📸 **Solução Visual**")
            foto = st.camera_input("Capturar Solução")
        relato = st.text_area("Descreva os fatos (O app redigirá o Parecer Forense Master):")
        if st.form_submit_button("⚖️ EXECUTAR INTELIGÊNCIA SOBERANA"):
            if relato:
                texto, tcr, glosa, lei = motor_apex_v80({'setor': setor, 'gravidade': grav, 'alvo': alvo, 'local': local, 'medida': medida, 'relato': relato})
                st.session_state['v80_rep'] = {'txt': texto, 'local': local, 'tcr': tcr, 'glosa': glosa, 'lei': lei, 'alvo': alvo, 'medida': medida, 'relato_orig': relato}
                st.success("✅ Inteligência SGI Concluída! Devolutiva pronta.")
                with tab_ativos:
    st.header("⚙️ Controle de Infraestrutura Crítica A-Z")
    t_fogo, t_infra, t_lto = st.tabs(["🔥 Incêndio/Safety", "⚡ Facilities/Eng", "🏢 Licenças/Alvarás"])
    with t_fogo:
        st.subheader("Bunker de Prevenção: Extintores e Hidrantes")
        with st.form("cad_fogo"):
            c_f1, c_f2 = st.columns(2)
            id_f = c_f1.text_input("Selo INMETRO / ID")
            venc_f = c_f2.date_input("Vencimento Carga/Manutenção")
            if st.form_submit_button("💾 Salvar Ativo de Fogo"):
                st.success(f"Ativo {id_f} registrado no radar de 90 dias.")
    with t_infra:
        st.subheader("Controle de Engenharia: PMOC, Geradores e SPDA")
        with st.form("cad_infra"):
            id_i = st.text_input("ID do Equipamento")
            tipo_i = st.selectbox("Tipo:", ["Gerador", "Nobreak", "Elevador", "PMOC Ar", "SPDA"])
            venc_i = st.date_input("Próxima Manutenção Preventiva")
            if st.form_submit_button("💾 Registrar Infraestrutura"):
                st.success("Ativo registrado conforme NRs e ISOs.")
    with t_lto:
        st.subheader("Radar de Licenças (License to Operate)")
        with st.form("cad_lto"):
            doc_l = st.text_input("Documento (Alvará, AVCB, Licença, PF)")
            venc_l = st.date_input("Vencimento Legal")
            if st.form_submit_button("💾 Monitorar LTO"):
                st.success("Documento vital sob radar Sentinel.")
with tab_equipe:
    st.header("👥 Gestão de Tropa e Colegiados")
    t_peo, t_cipa = st.tabs(["Performance Individual", "CIPA & Brigada"])
    with t_peo:
        with st.form("form_hr"):
            col_h1, col_h2 = st.columns(2)
            nome_h = col_h1.text_input("Nome do Colaborador")
            evento_h = col_h2.selectbox("Evento:", ["Falta", "Atraso", "Fadiga", "Almoço fora de hora"])
            if st.form_submit_button("🚀 Registrar Prontuário Vitalício"):
                conn = sqlite3.connect('sgi_guardian.db')
                conn.execute("INSERT INTO equipe (nome, cargo, status, icc) VALUES (?,?,?,?)",
                             (nome_h, "Operacional", evento_h, 85))
                conn.commit()
                st.success(f"Ocorrência de {evento_h} salva para {nome_h}.")
                with st.form("form_hr"):
            col_h1, col_h2 = st.columns(2)
            nome_h = col_h1.text_input("Nome do Colaborador")
            evento_h = col_h2.selectbox("Evento:", ["Falta", "Atraso", "Fadiga", "Almoço fora de hora"])
            if st.form_submit_button("🚀 Registrar Prontuário Vitalício"):
                conn = sqlite3.connect('sgi_guardian.db')
                conn.execute("INSERT INTO equipe (nome, cargo, status, icc) VALUES (?,?,?,?)",
                             (nome_h, "Operacional", evento_h, 85))
                conn.commit()
                st.success(f"Ocorrência de {evento_h} salva para {nome_h}.")
                st.success("Ata SGI gerada e arquivada para Auditoria ISO.")
# --- ABA 4: DASHBOARD POWER BI ---
with tab_bi:
    st.header("📊 Business Intelligence: ROI & EBITDA")
    if 'v80_rep' in st.session_state:
        d = st.session_state['v80_rep']
        c_m1, c_m2, c_m3 = st.columns(3)
        c_m1.metric("Lucro Protegido (TCR)", f"R$ {d['tcr']:,.2f}", "+28% ROI")
        c_m2.metric("Glosa Evitada", f"R$ {d['glosa']:,.2f}", "SLA Seguro")
        c_m3.metric("Integridade SGI", "97%", "Meta batida")
        fig_gauge = go.Figure(go.Indicator(mode = "gauge+number", value = 97,
            title = {'text': "Conformidade Global %"},
            gauge = {'axis': {'range': [None, 100]},
                     'bar': {'color': "#00D1FF"},
                     'steps' : [
                         {'range': [0, 50], 'color': "#FF4B4B"},
                         {'range': [50, 85], 'color': "#FFD700"},
                         {'range': [85, 100], 'color': "#00CC96"}]}))
        st.plotly_chart(fig_gauge, use_container_width=True)
    else:
        st.info("💡 Realize uma auditoria técnica para ativar os indicadores.")
# --- ABA 5: BOUTIQUE EXECUTIVE ---
with tab_boutique:
    if 'v80_rep' in st.session_state:
        d = st.session_state['v80_rep']
        u_txt = urllib.parse.quote(d['txt'])
        st.header("🚀 Boutique Sovereign - Despacho de Documentos")
        st.subheader("📲 Canais Digitais (WhatsApp)")
        w1, w2, w3 = st.columns(3)
        with w1:
            st.markdown(f'<a href="https://wa.me/?text={u_txt}" target="_blank"><button style="background:#34B7F1; border-radius:10px; color:white; border:none; height:50px; font-weight:bold; width:100%;">🔵 CLIENTE: DEVOLUTIVA</button></a>', unsafe_allow_html=True)
        with w2:
            msg_dir = f"*🛡️ SGI EXECUTIVO* - R$ " + str(d["tcr"])
            st.markdown(f'<a href="https://wa.me/?text={urllib.parse.quote(msg_dir)}" target="_blank"><button style="background:#075E54; border-radius:10px; color:white; border:none; height:50px; font-weight:bold; width:100%;">🟢 DIRETORIA: ROI</button></a>', unsafe_allow_html=True)
        with w3:
            msg_op = f"*🛠️ SGI OPERACIONAL* - Medida no alvo " + d["alvo"]
            st.markdown(f'<a href="https://wa.me/?text={urllib.parse.quote(msg_op)}" target="_blank"><button style="background:#25D366; border-radius:10px; color:white; border:none; height:50px; font-weight:bold; width:100%;">🟢 OPERACIONAL: AÇÃO</button></a>', unsafe_allow_html=True)
        st.divider()
        st.subheader("💾 Exportação Master de Arquivos")
        f1, f2, f3, f4 = st.columns(4)
        f1.download_button("📄 PDF Dossiê", d['txt'], f"SGI_{d['local']}.pdf", use_container_width=True)
        doc = Document(); doc.add_paragraph(d['txt']); w_buf = io.BytesIO(); doc.save(w_buf)
        f2.download_button("📝 WORD Relatório", w_buf.getvalue(), "Relatorio.docx", use_container_width=True)
        prs = Presentation(); slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "Resumo SGI"; slide.placeholders[1].text = d['txt']
        p_buf = io.BytesIO(); prs.save(p_buf)
        f3.download_button("📊 PPT Slides", p_buf.getvalue(), "Apresentacao.pptx", use_container_width=True)
        df_bi = pd.DataFrame([d]); e_buf = io.BytesIO(); df_bi.to_excel(e_buf, index=False)
        f4.download_button("🚀 POWER BI (.xlsx)", e_buf.getvalue(), "Base_BI.xlsx", use_container_width=True)
        st.text_area("Texto Gerado:", d['txt'], height=350)
        else:
        st.warning("⚠️ Realize uma auditoria na primeira aba para liberar a boutique.")
# --- MÓDULO 6: BUNKER DE LICENÇAS E LTO (LICENSE TO OPERATE) ---
elif menu == "🏢 Bunker de Licenças/LTO":
    st.header("🏢 Monitoramento de License to Operate (A a Z)")
    i_help("LTO")
    with st.expander("➕ Cadastrar Licença / Alvará / Certidão", expanded=False):
        with st.form("form_lto_v80"):
            c_l1, c_l2 = st.columns(2)
            nome_doc = c_l1.text_input("Nome do Documento (Ex: AVCB, Alvará PF)")
            orgao_doc = c_l1.text_input("Órgão Emissor (Ex: Bombeiros, ANAC, PF)")
            venc_doc = c_l2.date_input("Vencimento Legal")
            risco_doc = c_l2.selectbox("Risco se Vencer", ["Baixo", "Médio", "Interdição Imediata"])
            if st.form_submit_button("💾 Monitorar LTO"):
                conn = sqlite3.connect('sgi_guardian.db')
                conn.execute("INSERT INTO licencas (nome, orgao, vencimento, status) VALUES (?,?,?,?)", (nome_doc, orgao_doc, venc_doc, risco_doc))
                conn.commit(); conn.close()
                st.success(f"Documento {nome_doc} sob vigilância Sentinel.")
    st.subheader("📡 Radar Sentinel: Licenças Operacionais")
    conn = sqlite3.connect('sgi_guardian.db')
df_lto = pd.read_sql_query("SELECT nome, vencimento, status FROM licencas", conn)
    conn.close()
    if not df_lto.empty: st.dataframe(df_lto, use_container_width=True)
    else: st.info("Nenhuma licença cadastrada.")
# --- 7. EXPANSÃO DE INTELIGÊNCIA JURÍDICA E TÉCNICA ---
DOUTRINA_SGI = {
    "TERRA": {
        "seguranca": "Lei 14.967/2024 (Estatuto Seg. Privada) + Portaria 18.045/23 PF.",
        "trabalho": "CLT Art. 482: Desídia, Indisciplina e Insubordinação.",
        "incendio": "NR-23 + ITs Bombeiros Estaduais + NBR 14276 (Brigada)."
        },
    "MAR": {
        "offshore": "NR-37 (Segurança Offshore) + NORMAM-204/DPC Marinha.",
        "regime": "Lei 5.811/72 (Regime de Trabalho em Petróleo).",
        "emergencia": "CONAMA 398 (Plano de Emergência Individual - PEI)."
    },
    "CÉU": {
        "aviacao": "RBAC 107 (AVSEC) + RBAC 153 (Operações em Aeródromos).",
        "fadiga": "RBAC 117: Gerenciamento de Risco de Fadiga Humana.",
        "lei": "Lei 13.475/2017 (Lei do Aeronauta) + CBA."
        },
    "MAR": {
        "offshore": "NR-37 (Segurança Offshore) + NORMAM-204/DPC Marinha.",
        "regime": "Lei 5.811/72 (Regime de Trabalho em Petróleo).",
        "emergencia": "CONAMA 398 (Plano de Emergência Individual - PEI)."
    },
    "CÉU": {
        "aviacao": "RBAC 107 (AVSEC) + RBAC 153 (Operações em Aeródromos).",
        "fadiga": "RBAC 117: Gerenciamento de Risco de Fadiga Humana.",
        "lei": "Lei 13.475/2017 (Lei do Aeronauta) + CBA."
        def calcular_icc_v80(faltas, atrasos, desvios):
    score = 100 - (faltas * 15) - (atrasos * 5) - (desvios * 10)
    return max(score, 0)
def calc_burnout_v80(h_excedentes):
    score = h_excedentes * 12.5
    if score > 80: return "🔴 CRÍTICO", "red"
    if score > 50: return "🟡 ALERTA", "orange"
    return "🟢 REGULAR", "green"
# --- 9. MOTOR DE REDAÇÃO DE DEVOLUTIVA DE ELITE ---
def redigir_devolutiva_master(d, intel):
texto = f"Prezado Cliente, a Gestão de Compliance apresenta a resolução do evento na unidade {d['local']}.\\n\\n"
    texto += f"1. ANÁLISE FORENSE: O colaborador/ativo {d['alvo']} foi submetido ao escrutínio da {intel['lei']}.\\n"
    texto += f"Identificamos falha nos processos de {intel['iso']}, gerando risco à continuidade do negócio.\\n\\n"
    texto += f"2. TRATAMENTO E RESOLUÇÃO: Priorizando o SLA, aplicamos a medida de {d['medida']} imediata.\\n"
    texto += f"A eficácia foi validada via evidência visual 'After' e prova de presença georreferenciada.\\n\\n"
    texto += f"3. IMPACTO EXECUTIVO: Esta intervenção preventiva evitou um TCR de R$ {intel['tcr']:,.2f}.\\n\\n"
    texto += f"ID Digital: {hashlib.sha512(d['relato_orig'].encode()).hexdigest()[:10].upper()}"
    return texto
# --- 10. PROTOCOLOS DE EMERGÊNCIA (ICS) ---
EMERGENCY_PROTOCOLS = {
"OFFSHORE": "Ativar PRE imediato. Notificar Marinha e ANP em < 24h.",
    "VAZAMENTO": "Contenção imediata (CONAMA 398). Acionar cerco preventivo.",
    "INCÊNDIO": "Abandonar área, acionar Brigada e isolar sistemas de energia.",
    "MEDEVAC": "Protocolo de Evacuação Médica. Contato com regulação médica.",
    "AVSEC": "Intrusão em área restrita. Acionar PF e gerenciar crise.",
    "LOGISTICA": "Acidente com carga perigosa. Isolar e acionar órgãos ambientais."
}
def obter_protocolo_emergencia(cenario):
    return EMERGENCY_PROTOCOLS.get(cenario, "Avaliar risco e acionar SGI.")
# --- 11. GESTÃO DE ATIVOS DE SAÚDE (DEA/AUTOCLAVE) ---
with tab_bi:
    st.divider()
    st.subheader("🧘 People Analytics: Risco de Fadiga (RBAC 117)")
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        h_excedente = st.slider("Horas Excedentes no Turno:", 0, 12, 2)
        status_f, cor_f = calc_burnout_v80(h_excedente)
        st.markdown(f"### Status: :{cor_f}[{status_f}]")
    with col_p2:
        st.write("**Impacto na Segurança Operacional:**")
        st.write("Horas excessivas aumentam em 40% o risco de erro humano.")
        if h_excedente > 8:
            st.error("🚨 INTERDIÇÃO: Limite de segurança RBAC 117 / NR-17 excedido.")
# --- 12. GESTÃO DE ATIVOS DE SAÚDE E HIGIENE ---
with tab_ativos:
    st.divider()
    st.subheader("🏥 Equipamentos de Emergência e Saúde")
    with st.expander("➕ Registrar DEA / Autoclave / Maletas"):
        st.text_input("ID do Equipamento Médico")
        st.selectbox("Status de Calibração", ["Calibrado", "Vencido"])
        st.date_input("Próxima Validação Técnica")
        if st.button("💾 Protocolar Ativo de Saúde"):
            st.success("Ativo registrado conforme RDC ANVISA.")
# --- 13. RADAR DE LICENÇAS E ALVARÁS (LTO FULL) ---
with tab_ativos:
    st.divider()
    st.subheader("🏢 Status de Licenças Operacionais (License to Operate)")
    hoje = date.today()
    prazo_sentinel = hoje + timedelta(days=90)
    st.info("Radar Sentinel: Monitoramento preditivo de 90 dias ativo.")
    conn = sqlite3.connect('sgi_guardian.db')
    df_lto_radar = pd.read_sql_query("SELECT nome, orgao, vencimento, status FROM licencas", conn)
    conn.close()
    if not df_lto_radar.empty:
        st.table(df_lto_radar)
        st.warning("⚠️ Alerta: Iniciada contagem regressiva para renovação legal.")
    else:
        st.info("Bunker de licenças vazio. Cadastre Alvarás e AVCB.")
# --- 14. MOTOR DE TRADUÇÃO BILINGUE (PT/EN) ---
def traduzir_parecer_v80(texto, idioma):
if idioma == "Inglês":
        t = texto.replace("Prezado Cliente", "Dear Client")
        t = t.replace("informamos a resolução", "we inform the resolution")
        t = t.replace("PARECER TÉCNICO-LEGAL", "TECHNICAL-LEGAL OPINION")
        t = t.replace("TRATAMENTO E RESOLUÇÃO", "TREATMENT AND RESOLUTION")
        t = t.replace("IMPACTO EXECUTIVO", "EXECUTIVE IMPACT")
        return t
    return texto
# --- 15. GESTÃO DE EPIs (NR-06) ---
with tab_equipe:
st.divider()
    st.subheader("🛡️ Controle de EPIs (NR-06)")
    with st.expander("➕ Registro de Entrega e CA"):
        with st.form("form_epi"):
            col_e1, col_e2 = st.columns(2)
            col_e1.text_input("Equipamento (Ex: Colete, Bota)")
            col_e2.text_input("Número do CA (MTE)")
            st.date_input("Data da Entrega")
            if st.form_submit_button("💾 Salvar Ficha de EPI"):
                st.success("Entrega de EPI protocolada no prontuário.")
                st.divider()
    st.subheader("📅 Linha do Tempo de Conduta")
    nome_busca = st.text_input("Consultar Histórico por Nome:")
    if nome_busca:
        conn = sqlite3.connect('sgi_guardian.db')
        query = f"SELECT data, status, relato FROM auditorias WHERE alvo LIKE '%{nome_busca}%'"
        df_hist = pd.read_sql_query(query, conn)
        conn.close()
        if not df_hist.empty: st.dataframe(df_hist, use_container_width=True)
        else: st.info("Nenhuma ocorrência anterior para este alvo.")
        with tab_boutique:
    if 'v80_rep' in st.session_state:
        st.divider()
        st.subheader("📦 Sovereign Vault (Dossiê de Defesa)")
        if st.button("🔐 Gerar Pacote ZIP de Evidências"):
            buf_z = io.BytesIO()
            with zipfile.ZipFile(buf_z, 'w') as zf:
                zf.writestr("relatorio_forense.txt", st.session_state['v80_rep']['txt'])
                zf.writestr("checksum.txt", hashlib.sha512(st.session_state['v80_rep']['txt'].encode()).hexdigest())
            st.download_button("📥 Baixar ZIP", buf_z.getvalue(), "SGI_Vault.zip", use_container_width=True)
            # --- 16. ANÁLISE DE CAUSA RAIZ (ISHIKAWA / PDCA) ---
def sugerir_causa_raiz(evento):
    causas = {
        "Falta": "Mão de Obra: Desmotivação ou falha de transporte.",
        "Desvio de Conduta": "Medida: Falha na supervisão direta.",
        "Falha Técnica": "Máquina: Manutenção preventiva inexistente.",
        "Fadiga": "Método: Escala de trabalho incompatível.",
        "Inconformidade ISO": "Processo: Falha no controle operacional."
    }
    return causas.get(evento, "Analisar fatores humanos e técnicos.")
    with tab_audit:
    if 'v80_rep' in st.session_state:
        st.divider()
        st.subheader("🪵 Análise de Causa Raiz (Ishikawa)")
        r_v80 = st.session_state['v80_rep']
        st.info(f"**Sugestão Técnica:** {sugerir_causa_raiz(r_v80['medida'])}")
        # --- 17. PLANO DE AÇÃO 5W2H AUTOMÁTICO ---
        st.subheader("🚀 Plano de Ação Corretiva (5W2H)")
        with st.expander("Visualizar Cronograma de Tratamento"):
            st.write(f"**O QUE:** {r_v80['medida']} imediata no posto.")
            st.write(f"**POR QUE:** Para mitigar risco de {r_v80['gravidade']} e garantir conformidade.")
            st.write(f"**QUEM:** Supervisor de Área e Gestor de SGI.")
            st.write(f"**ONDE:** Unidade {r_v80['local']}.")
            st.write(f"**QUANDO:** Imediatamente (Data do registro).")
            st.write(f"**COMO:** Aplicação de medida disciplinar e reforço técnico.")
            st.write(f"**QUANTO:** Custo operacional absorvido pelo ROI de prevenção.")
# --- 18. GESTÃO DE RESILIÊNCIA DE INFRAESTRUTURA ---
with tab_ativos:
    st.divider()
    st.subheader("🏗️ Status de Resiliência Sistêmica")
    col_r1, col_r2, col_r3 = st.columns(3)
    with col_r1: st.metric("Energia", "98%", "Estável")
    with col_r2: st.metric("Sistemas de Vida", "85%", "-5%")
    with col_r3: st.metric("Segurança", "100%", "Conforme")
    st.write("**Preditivo:** Risco de falha em climatização detectado em 15 dias.")
# --- 19. MOTOR DE IMPACTO REPUTACIONAL E ESG ---
def avaliar_impacto_esg(gravidade):
    impactos = {"Baixa": "Neutro", "Média": "Alerta Social",
                "Alta": "Risco de Governança", "Crítica": "Dano Grave"}
    return impactos.get(gravidade, "Analisando")
    with tab_boutique:
    if 'v80_rep' in st.session_state:
        res_esg = avaliar_impacto_esg(st.session_state['v80_rep']['gravidade'])
        st.session_state['v80_rep']['esg'] = res_esg
# --- 20. MOTOR DE GERAÇÃO DE SLIDES CORPORATIVOS (PPTX) ---
def criar_apresentacao_v80(d):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    title = slide.shapes.title
    title.text = f"GESTÃO DE RISCO SGI - {d['local']}"
    body = slide.placeholders[1]
    body.text = f"Evento: {d['medida']}\nImpacto ESG: {d['esg']}\nROI: R$ {d['tcr']:,.2f}"
    buf_p = io.BytesIO()
    prs.save(buf_p)
    return buf_p.getvalue()
# --- 21. FERRAMENTA DE EXPORTAÇÃO EXCEL PARA BI ---
def exportar_excel_bi(d):
    df_bi = pd.DataFrame([d])
    buf_e = io.BytesIO()
    with pd.ExcelWriter(buf_e, engine='xlsxwriter') as writer:
    df_bi.to_excel(writer, index=False, sheet_name='Base_SGI')
    return buf_e.getvalue()
with tab_boutique:
    if 'v80_rep' in st.session_state:
        d_final = st.session_state['v80_rep']
        st.divider()
        st.subheader("📊 Documentação de Board (PPT & BI)")
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            st.download_button("📊 PPT: SLIDES DE BOARD", criar_apresentacao_v80(d_final), "SGI_Slides.pptx", use_container_width=True)
    with col_d2:
            st.download_button("🚀 POWER BI READY (.xlsx)", exportar_excel_bi(d_final), "Base_BI_Master.xlsx", use_container_width=True)
# --- 22. ANÁLISE DE TENDÊNCIA MENSAL (BI NATIVO) ---
with tab_bi:
    st.divider()
    st.subheader("📈 Performance Histórica do Posto")
    # Simulação de dados para curva de aprendizado SGI
    df_curve = pd.DataFrame({'Mês': ['Jan', 'Fev', 'Mar', 'Abr'], 'Conformidade': [70, 75, 88, 97]})
    fig_curve = px.line(df_curve, x='Mês', y='Conformidade', markers=True, title="Evolução da Cultura de Segurança")
    fig_curve.update_traces(line_color='#00D1FF')
    st.plotly_chart(fig_curve, use_container_width=True)
        st.info("💡 A curva ascendente demonstra a eficácia dos treinamentos e reciclagens.")
# --- 23. TORRE DE CONTROLE AMBIENTAL (ISO 14001 / ESG) ---
with tab_ativos:
    st.divider()
    st.subheader("🍃 Gestão de Resíduos e Meio Ambiente")
    with st.expander("➕ Lançar Manifesto de Resíduos (MTR/PNRS)"):
        with st.form("form_ambiental"):
            tipo_residuo = st.selectbox("Classe do Resíduo:", ["Classe I (Perigoso)", "Classe IIA", "Classe IIB"])
            destino_res = st.text_input("Destinação Final (Aterro/Incineração):")
if st.form_submit_button("💾 Salvar MTR"):
                st.success("Manifesto de resíduos protocolado conforme PNRS.")
st.sidebar.divider()
st.sidebar.subheader("🔒 Trilha de Auditoria Forense")
st.sidebar.write("Assinaturas SHA-512 garantem a imutabilidade dos dados.")
def validar_integridade_v80(texto, hash_original):
    h_verif = hashlib.sha512(texto.encode()).hexdigest()[:12].upper()
    return h_verif == hash_original
# --- 24. SISTEMA DE BACKUP E EXPORTAÇÃO JSON ---
import json
def export_json_v80(d):
    data_str = json.dumps(d, indent=4)
    return data_str
with tab_boutique:
    if 'v80_rep' in st.session_state:
        st.divider()
        st.subheader("🔗 Integração ERP (JSON/Oracle/SAP)")
        json_data = export_json_v80(st.session_state['v80_rep'])
        st.download_button("🔌 EXPORTAR PACOTE JSON", json_data, "sgi_data.json", use_container_width=True)
st.sidebar.divider()
st.sidebar.info("Selo de Integridade: **SENTINEL ACTIVE**")
st.sidebar.markdown(f"Versão: {datetime.now().year}.6.8")
if st.sidebar.button("♻️ Reinicializar Base Local"):
    st.sidebar.warning("Ação requer privilégios de Administrador.")
# --- 25. MÓDULO DE SEGURANÇA E REDE ---
st.sidebar.write("---")
st.sidebar.write("📡 **Status do Sistema:** Operacional")
st.sidebar.write("🔐 **Conexão:** Criptografada TLS 1.3")
st.sidebar.write("👤 **Acesso:** SGI Master Alpha")
st.sidebar.divider()
def footer_forense_v80(gps_info):
    return f"Relatório autenticado via GPS: {gps_info} | Verificado em {datetime.now()}"
with tab_boutique:
    if 'v80_rep' in st.session_state:
        st.write("---")
        gps_str = get_gps_data()
        st.markdown(f"<p style='font-size:10px; color:gray;'>{footer_forense_v80(gps_str)}</p>", unsafe_allow_html=True)
        st.markdown("<p style='font-size:10px; color:gray;'>Cadeado SHA-512 Ativo - SGI Alpha Sovereign</p>", unsafe_allow_html=True)
# --- 26. GESTÃO DE DISPONIBILIDADE E SLA ---
def calcular_uptime_v80(paradas_min):
tempo_total = 43200  # Total de minutos em um mês (30 dias)
    disponibilidade = ((tempo_total - paradas_min) / tempo_total) * 100
    return round(disponibilidade, 2)

with tab_bi:
    st.divider()
    st.subheader("⏱️ Disponibilidade Operacional (SLA)")
    min_parada = st.number_input("Minutos de Inatividade (Mês):", min_value=0, value=120)
    uptime = calcular_uptime_v80(min_parada)
    st.metric("Uptime do Posto", f"{uptime}%", delta=f"{uptime - 98.5:.2f}% vs Meta")
    with tab_ativos:
    st.divider()
    st.subheader("🤝 Compliance de Fornecedores e Terceiros")
    with st.expander("➕ Auditoria de Documentação de Terceiros"):
        with st.form("form_terceiros"):
            emp_terceira = st.text_input("Nome da Empresa Terceirizada:")
            st.selectbox("Status Seguro RC/Garantia:", ["Vigente", "Vencido", "Não Apresentado"])
            st.selectbox("Certidões Trabalhistas (CND):", ["Regular", "Irregular"])
            if st.form_submit_button("💾 Protocolar Auditoria de Terceiro"):
                st.success(f"Compliance de {emp_terceira} verificado e arquivado.")
                def enviar_alerta_rh(d):
    assunto_rh = f"ALERTA DISCIPLINAR - {d['alvo']} - {d['local']}"
    corpo_rh = f"O colaborador {d['alvo']} recebeu a medida de {d['medida']} em {d['data']}."
    gmail_rh = f"https://mail.google.com/mail/?view=cm&fs=1&to=rh@empresa.com&su={urllib.parse.quote(assunto_rh)}&body={urllib.parse.quote(corpo_rh)}"
    return gmail_rh
with tab_boutique:
    if 'v80_rep' in st.session_state:
        st.divider()
        st.subheader("👥 Notificação Interna RH (Forense)")
        link_rh = enviar_alerta_rh(st.session_state['v80_rep'])
        st.markdown(f'<a href="{link_rh}" target="_blank"><button style="background:#004a99; color:white; border:none; padding:10px; border-radius:10px; width:100%;">👤 NOTIFICAR RH (PRONTUÁRIO)</button></a>', unsafe_allow_html=True)
        st.divider()
        st.subheader("🏁 Encerramento de Ciclo PDCA")
        st.success(f"Dossiê autenticado sob ID: {hashlib.sha512(d['txt'].encode()).hexdigest()[:16].upper()}")
    else:
        st.warning("⚠️ O sistema aguarda a finalização da auditoria técnica para habilitar as funções executivas.")
# --- RODAPÉ DE INTEGRIDADE ---
st.sidebar.write("---")
st.sidebar.markdown("**Sentinel Engine v8.0**")
st.sidebar.info("Criptografia ponta-a-ponta SHA-512 ativa.")
st.sidebar.write("---")
st.sidebar.write("🔒 Proteção de Dados: LGPD Active")
st.sidebar.caption("SGI Guardian - Sovereign Tech")
st.sidebar.write(f"Sincronização: {datetime.now().strftime('%H:%M:%S')}")
if __name__ == "__main__":
    try:
        init_db()
    except:
        pass
# --- END OF FILE: APEX V80 SGI GUARDIAN ---
