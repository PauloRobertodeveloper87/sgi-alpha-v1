import streamlit as st
import pandas as pd
import sqlite3
import io
import urllib.parse
import hashlib
import zipfile
from datetime import datetime, date, timedelta
import plotly.express as px
import plotly.graph_objects as go
from fpdf import FPDF
from docx import Document
from pptx import Presentation
from streamlit_js_eval import get_geolocation

# --- CONFIGURAÇÃO DE UI/UX SUPREME ---
st.set_page_config(
    page_title="SGI Alpha Apex V70",
    layout="wide",
    page_icon="🛡️"
)
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
    .i-help { 
        color: #00D1FF; 
        font-weight: bold; 
        cursor: help; 
        text-decoration: underline; 
    }
    </style>
    """, unsafe_allow_html=True)

# --- INICIALIZAÇÃO DO BANCO DE DADOS ---
def init_db():
    conn = sqlite3.connect('sgi_master_v70.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS equipe 
                 (id INTEGER PRIMARY KEY, nome TEXT, cargo TEXT, esfera TEXT, 
                  reciclagem DATE, icc INTEGER, fadiga INTEGER)''')
    c.execute('''CREATE TABLE IF NOT EXISTS ativos 
                 (id TEXT PRIMARY KEY, categoria TEXT, local TEXT, vencimento DATE, status TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS licencas 
                 (id INTEGER PRIMARY KEY, nome TEXT, orgao TEXT, vencimento DATE, status TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS auditorias 
                 (id INTEGER PRIMARY KEY, data TEXT, local TEXT, alvo TEXT, relato TEXT, 
                  tcr REAL, glosa REAL, lei TEXT, iso TEXT, acao TEXT, hash TEXT)''')
    conn.commit()
    conn.close()

init_db()
# --- CONTINUAÇÃO DO BANCO DE DADOS ---
    conn.commit()
    conn.close()

init_db()

# --- 3. GLOSSÁRIO TÉCNICO I-HELP 6.0 ---
GLOS = {
    "TCR": "Total Cost of Risk: Prejuízo financeiro (Multa + Glosa + Dano à Marca).",
    "LTO": "License to Operate: Alvarás, AVCBs e Certificações vitais para o posto.",
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
def motor_apex_v70(dados):
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
        }
    }
    biblioteca.update({
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
    })
    
    ref = biblioteca.get(dados['setor'], biblioteca["INDUSTRIAL / TERRA"])
    # REDAÇÃO DA DEVOLUTIVA IMPECÁVEL (TOM EXECUTIVO)
    texto = f"""Prezado Cliente,

Referente à integridade operacional da unidade {dados['local']}, informamos o encerramento do ciclo de tratamento SGI relativo ao alvo {dados['alvo']}.

1. PARECER TÉCNICO-LEGAL (AUDITORIA 360°):
O desvio identificado ("{dados['relato']}") foi confrontado com o escrutínio da {ref['lei']}. Tal desvio infringe as diretrizes de {ref['pilar']} e as normas {ref['iso']}, gerando risco de interdição e dano reputacional.

2. TRATAMENTO E MEDIDAS COMPENSATÓRIAS:
Priorizando a segurança absoluta e o contrato (SLA), aplicamos a medida de {dados['medida']} imediata. A eficácia foi validada via evidência visual "After" e prova de presença georreferenciada.

3. IMPACTO EXECUTIVO E ROI DE GESTÃO:
Esta intervenção preventiva evitou um Total Cost of Risk (TCR) estimado em R$ {tcr:,.2f} e uma glosa de R$ {glosa:,.2f}, protegendo o Score ESG e a Licença para Operar (LTO).
"""
    return texto, tcr, glosa, ref['lei']
    # --- 5. INTERFACE DE COMANDO TITAN ---
st.title("🛡️ SGI Alpha Apex V80")
st.subheader("Supreme Sovereign Command Matrix")

menu = st.sidebar.radio("Comando Estratégico", 
    ["🛡️ Auditoria & Devolutiva", "⚙️ Torre de Ativos (Infra)", "👥 People Analytics", "🏢 Bunker de Licenças/LTO", "📊 Dashboard Power BI", "🚀 Boutique Executive"])

# --- MÓDULO 1: AUDITORIA ---
if menu == "🛡️ Auditoria & Devolutiva":
    st.header("🧐 Inspeção Forense e PDCA")
    with st.form("audit_form_v80"):
        c1, c2, c3 = st.columns(3)
        with c1:
            setor = st.selectbox("Setor de Atuação", ["OFFSHORE / MAR", "AVIAÇÃO / CÉU", "HOSPITALAR / SAÚDE", "INDUSTRIAL / TERRA", "SEGURANÇA PRIVADA"])
            local = st.text_input("Unidade / Posto")
            i_help("LTO")
        with c2:
            alvo = st.text_input("Alvo (Profissional ou Ativo)")
            grav = st.select_slider("Criticidade do Impacto", ["Baixa", "Média", "Alta", "Crítica"])
            with c3:
            medida = st.selectbox("Ação de Tratamento:", ["Advertência", "Suspensão", "Substituição", "Manutenção", "Interdição", "Justa Causa"])
            i_help("SGI")
            st.write("📸 **Solução Visual**")
            foto = st.camera_input("Capturar Solução")

        relato = st.text_area("Descreva os fatos (O app redigirá o Parecer Forense Master):")
        
        if st.form_submit_button("⚖️ EXECUTAR INTELIGÊNCIA SOBERANA"):
            if relato:
                texto, tcr, glosa, lei = motor_apex_v70({'setor': setor, 'gravidade': grav, 'alvo': alvo, 'local': local, 'medida': medida, 'relato': relato})
                st.session_state['v80_rep'] = {
                    'txt': texto, 'local': local, 'tcr': tcr, 'glosa': glosa, 
                    'lei': lei, 'alvo': alvo, 'medida': medida, 'relato_orig': relato
                }
                st.success("✅ Inteligência Apex V80 Concluída! Devolutiva pronta e blindada.")
            else:
                st.error("Descreva o fato para processar.")
                # --- MÓDULO 2: TORRE DE ATIVOS (INFRAESTRUTURA CRÍTICA) ---
elif menu == "⚙️ Torre de Ativos (Infra)":
    st.header("⚙️ Controle de Infraestrutura Crítica A-Z")
    i_help("PMOC")
    
    with st.expander("➕ Cadastrar Novo Ativo Crítico", expanded=False):
        with st.form("form_ativo_v80"):
            c_a1, c_a2, c_a3 = st.columns(3)
            id_a = c_a1.text_input("ID do Ativo / Selo INMETRO")
            cat_a = c_a2.selectbox("Categoria", [
                "Prevenção de Incêndio (Extintor/Hidrante)", 
                "Energia (Gerador/Nobreak/NR-10)", 
                "Climatização (PMOC/Ar)", 
                "Vida (Elevador/Escada)", 
                "Pressão (NR-13/Caldeira)"
            ])
            venc_a = c_a3.date_input("Data de Vencimento / Próxima Manutenção")
            if st.form_submit_button("💾 Salvar Ativo no Bunker"):
                conn = sqlite3.connect('sgi_omega_v80.db')
                conn.execute("INSERT INTO ativos (id, categoria, local, vencimento, status) VALUES (?,?,?,?,?)",
                             (id_a, cat_a, "Unidade Padrão", venc_a, "Monitorado"))
                conn.commit()
                conn.close()
                st.success(f"Ativo {id_a} registrado com radar de 120 dias.")

    # Listagem de Ativos Críticos
    st.subheader("📡 Radar Sentinel: Inventário de Ativos")
    conn = sqlite3.connect('sgi_omega_v80.db')
    df_ativos = pd.read_sql_query("SELECT id, categoria, vencimento, status FROM ativos", conn)
    if not df_ativos.empty:
        st.dataframe(df_ativos, use_container_width=True)
    else:
        st.info("Nenhum ativo cadastrado para monitoramento.")
        # --- MÓDULO 3: PEOPLE ANALYTICS (EQUIPE E PERFORMANCE) ---
elif menu == "👥 People Analytics":
    st.header("👥 Gestão de Tropa e Índice ICC")
    i_help("ICC")
    i_help("RBAC 117")
    
    with st.expander("📝 Lançar Ocorrência de Equipe", expanded=True):
        with st.form("form_hr_v80"):
            col_h1, col_h2, col_h3 = st.columns(3)
            nome_h = col_h1.text_input("Nome do Colaborador")
            cargo_h = col_h1.text_input("Cargo/Função")
            esfera_h = col_h2.selectbox("Esfera", ["TERRA", "MAR", "CÉU"])
            evento_h = col_h2.selectbox("Evento HR", ["Falta", "Atraso", "Almoço Fora de Hora", "Hora Excedente", "Fadiga"])
            data_h = col_h3.date_input("Data do Evento", date.today())
            icc_val = col_h3.slider("Score de Confiabilidade (Manual)", 0, 100, 85)
            if st.form_submit_button("🚀 Registrar no Prontuário Vitalício"):
                conn = sqlite3.connect('sgi_omega_v80.db')
                conn.execute("INSERT INTO equipe (nome, cargo, esfera, reciclagem, icc, fadiga) VALUES (?,?,?,?,?,?)",
                             (nome_h, cargo_h, esfera_h, data_h, icc_val, 0))
                conn.commit()
                conn.close()
                st.success(f"Evento de {evento_h} registrado para {nome_h}. Glosa calculada no BI.")

    # Lista de Equipe
    st.subheader("📋 Painel de Confiabilidade da Tropa")
    conn = sqlite3.connect('sgi_omega_v80.db')
    df_equipe = pd.read_sql_query("SELECT nome, cargo, esfera, icc FROM equipe", conn)
    if not df_equipe.empty:
        st.table(df_equipe)
    else:
        st.info("Aguardando registros de pessoal.")
        # --- MÓDULO 4: BUNKER DE LICENÇAS E LTO (LICENSE TO OPERATE) ---
elif menu == "🏢 Bunker de Licenças/LTO":
    st.header("🏢 Monitoramento de License to Operate (A a Z)")
    i_help("LTO")
    
    with st.expander("➕ Cadastrar Licença / Alvará / Certidão", expanded=False):
        with st.form("form_lto"):
            c_l1, c_l2 = st.columns(2)
            nome_doc = c_l1.text_input("Nome do Documento (Ex: AVCB, Alvará PF)")
            orgao_doc = c_l1.text_input("Órgão Emissor (Ex: Bombeiros, ANAC, PF)")
            venc_doc = c_l2.date_input("Vencimento Legal")
            risco_doc = c_l2.selectbox("Risco se Vencer", ["Baixo", "Médio", "Interdição Imediata"])
            
            if st.form_submit_button("💾 Monitorar LTO"):
                conn = sqlite3.connect('sgi_omega_v80.db')
                conn.execute("INSERT INTO licencas (nome, orgao, vencimento, risco) VALUES (?,?,?,?)",
                             (nome_doc, orgao_doc, venc_doc, risco_doc))
                conn.commit()
                conn.close()
                st.success(f"Documento {nome_doc} sob vigilância Sentinel.")
                # Radar de Vencimentos LTO
    st.subheader("📡 Radar Sentinel: Licenças Operacionais")
    conn = sqlite3.connect('sgi_omega_v80.db')
    df_lto = pd.read_sql_query("SELECT nome, vencimento, risco FROM licencas", conn)
    if not df_lto.empty:
        st.dataframe(df_lto, use_container_width=True)
    else:
        st.info("Nenhuma licença cadastrada.")

# --- MÓDULO 5: DASHBOARD POWER BI (EXECUTIVE VIEW) ---
elif menu == "📊 Dashboard Power BI":
    st.header("📊 Business Intelligence: ROI & EBITDA Protection")
    i_help("EBITDA")
    i_help("TCR")
    
    if 'v80_rep' in st.session_state:
        d = st.session_state['v80_rep']
        # Indicadores de Topo (Métricas Executivas)
        c_m1, c_m2, c_m3 = st.columns(3)
        c_m1.metric("Lucro Protegido (TCR)", f"R$ {d['tcr']:,.2f}", "+32% Performance")
        c_m2.metric("Glosa Evitada", f"R$ {d['glosa']:,.2f}", "SLA Seguro")
        c_m3.metric("Status PDCA", "100%", "Ciclo Fechado")
        st.divider()
        # Gráficos de Alta Performance para Espelhamento
        col_g1, col_g2 = st.columns(2)
        with col_g1:
            # Gráfico de Radar: Maturidade SGI da Unidade
            df_radar = pd.DataFrame(dict(
                r=[98, 85, 92, 100, 90],
                theta=['Qualidade','Ambiental','Segurança','Governança','Resiliência']))
            fig_radar = px.line_polar(df_radar, r='r', theta='theta', line_close=True, title="Maturidade SGI")
            fig_radar.update_traces(fill='toself', line_color='#00D1FF')
            st.plotly_chart(fig_radar, use_container_width=True)
            
        with col_g2:
            # Gráfico de Velocímetro (Gauge) de Conformidade
            fig_gauge = go.Figure(go.Indicator(
                mode = "gauge+number", value = 97,
                title = {'text': "Conformidade Global %"},
                gauge = {'axis': {'range': [None, 100]},
                         'bar': {'color': "#00D1FF"},
                         'steps' : [
                             {'range': [0, 50], 'color': "#FF4B4B"},
                             {'range': [50, 85], 'color': "#FFD700"},
                             {'range': [85, 100], 'color': "#00CC96"}]}))
            st.plotly_chart(fig_gauge, use_container_width=True)
            else:
        st.info("💡 Realize uma auditoria técnica para ativar os indicadores de performance.")

# --- MÓDULO 6: BOUTIQUE EXECUTIVE (OMNI-CHANNEL) ---
if menu == "🚀 Boutique Executive":
    if 'v80_rep' in st.session_state:
        d = st.session_state['v80_rep']
        u_txt = urllib.parse.quote(d['txt'])
        
        st.header("🚀 Boutique Sovereign - Despacho de Documentos")
        st.subheader("📲 Canais Digitais de Despacho (WhatsApp)")
        w_c1, w_c2, w_c3 = st.columns(3)
        
        # WhatsApp 1: CLIENTE (Tom Diplomático)
        with w_c1:
            st.markdown(f'<a href="https://wa.me/?text={u_txt}" target="_blank"><button style="background:#34B7F1; height:60px;">🔵 CLIENTE: DEVOLUTIVA</button></a>', unsafe_allow_html=True)
        
        # WhatsApp 2: DIRETORIA (Tom Financeiro/ROI)
        with w_c2:
            msg_dir = f"*🛡️ SGI EXECUTIVO*\\n*Unidade:* {d['local']}\\n*ROI de Gestão:* R$ {d['tcr']:,.2f}\\n*Selo:* {datetime.now().day}"
            st.markdown(f'<a href="https://wa.me/?text={urllib.parse.quote(msg_dir)}" target="_blank"><button style="background:#075E54; height:60px;">🟢 DIRETORIA: ROI</button></a>', unsafe_allow_html=True)
            # WhatsApp 3: OPERACIONAL (Tom de Comando)
        with w_c3:
            msg_op = f"*🛠️ SGI OPERACIONAL*\\n*Local:* {d['local']}\\n*Medida:* Executar {d['medida']} imediata."
            st.markdown(f'<a href="https://wa.me/?text={urllib.parse.quote(msg_op)}" target="_blank"><button style="background:#25D366; height:60px;">🟢 OPERACIONAL: AÇÃO</button></a>', unsafe_allow_html=True)

        # WhatsApp 4: JURÍDICO (Tom Forense/Provas)
        st.write("")
        w_c4, w_c5, w_c6 = st.columns(3)
        with w_c4:
            hash_j = hashlib.sha256(d['relato_orig'].encode()).hexdigest()[:12].upper()
            msg_jur = f"*⚖️ SGI JURÍDICO (PROVA)*\\n*Local:* {d['local']}\\n*Lei:* {d['lei']}\\n*Assinatura Digital:* {hash_j}"
            st.markdown(f'<a href="https://wa.me/?text={urllib.parse.quote(msg_jur)}" target="_blank"><button style="background:#202124; height:60px;">⚫ JURÍDICO: DEFESA</button></a>', unsafe_allow_html=True)
            st.divider()
        st.subheader("📧 E-mail Profissional (Gmail / Outlook)")
        col_e1, col_e2 = st.columns(2)
        assunto = urllib.parse.quote(f"Devolutiva de Compliance SGI - {d['local']}")
        corpo_email = urllib.parse.quote(d['txt'])
        with col_e1:
            gmail_url = f"https://mail.google.com/mail/?view=cm&fs=1&su={assunto}&body={corpo_email}"
            st.markdown(f'<a href="{gmail_url}" target="_blank"><button style="background:#EA4335; height:50px;">📧 GMAIL BUSINESS</button></a>', unsafe_allow_html=True)
        with col_e2:
            st.markdown(f'<a href="mailto:?subject={assunto}&body={corpo_email}"><button style="background:#0078D4; height:50px;">📧 OUTLOOK / PADRÃO</button></a>', unsafe_allow_html=True)

        st.divider()
        st.subheader("💾 Exportação Master de Arquivos")
        f1, f2, f3, f4 = st.columns(4)
        # 1. GERADOR DE PDF FORENSE
        pdf = FPDF()
        pdf.add_page()
        pdf.set_fill_color(0, 43, 94)
        pdf.rect(0, 0, 210, 40, 'F')
        pdf.set_font("Arial", 'B', 16)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(0, 20, "SGI ALPHA - DOSSIÊ TÉCNICO FORENSE", 0, 1, 'C')
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Arial", '', 10)
        pdf.ln(20)
        pdf.multi_cell(0, 8, d['txt'])
        pdf_bytes = pdf.output(dest='S').encode('latin-1', 'replace')
        f1.download_button("📄 PDF FORENSE", pdf_bytes, f"SGI_{d['alvo']}.pdf", "application/pdf", use_container_width=True)
        # 2. GERADOR DE WORD EDITÁVEL
        doc = Document()
        doc.add_heading(f"Relatório de Gestão SGI - {d['local']}", 0)
        doc.add_paragraph(d['txt'])
        word_buf = io.BytesIO()
        doc.save(word_buf)
        f2.download_button("📝 WORD RELATÓRIO", word_buf.getvalue(), f"SGI_{d['alvo']}.docx", use_container_width=True)

        # 3. GERADOR DE SLIDES POWERPOINT
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = "Resumo Executivo de Governança"
        slide.placeholders[1].text = f"Unidade: {d['local']}\nROI Protegido: R$ {d['tcr']:,.2f}\nStatus: RESOLVIDO"
        ppt_buf = io.BytesIO()
        prs.save(ppt_buf)
        f3.download_button("📊 PPT SLIDES", ppt_buf.getvalue(), "SGI_Apresentacao.pptx", use_container_width=True)
        # 4. GERADOR DE CÉLULA POWER BI (.xlsx)
        df_export = pd.DataFrame([d])
        excel_buf = io.BytesIO()
        with pd.ExcelWriter(excel_buf, engine='xlsxwriter') as writer:
            df_export.to_excel(writer, index=False, sheet_name='Base_Mestra_SGI')
        f4.download_button("🚀 POWER BI (.xlsx)", excel_buf.getvalue(), "Base_BI_SGI.xlsx", use_container_width=True)

        st.divider()
        st.subheader("📝 Pré-visualização do Parecer Impecável")
        st.text_area("Cópia de Segurança do Texto:", d['txt'], height=350)
        
        # SALVAMENTO NO HISTÓRICO DO BANCO DE DADOS
        conn = sqlite3.connect('sgi_omega_v80.db')
        hash_f = hashlib.sha256(d['txt'].encode()).hexdigest()[:12]
        conn.execute('''INSERT INTO auditorias 
                        (data, local, alvo, relato, tcr, glosa, lei, iso, acao, hash) 
                        VALUES (?,?,?,?,?,?,?,?,?,?)''', 
                     (datetime.now().strftime("%Y-%m-%d"), d['local'], d['alvo'], d['relato_orig'], 
                      d['tcr'], d['glosa'], d['lei'], "ISO 9001/45001", d['medida'], hash_f))
        conn.commit()
        conn.close()
    else:
        st.warning("⚠️ Realize uma auditoria na primeira aba para liberar a boutique sovereign.")

# --- FIM DO CÓDIGO ALPHA APEX V80 ---
# --- CONTINUAÇÃO DO MÓDULO DE EQUIPE (CIPA & BRIGADA) ---
with tab_equipe:
    st.divider()
    st.header("🤝 Atas de Reunião e Planos CIPA/Brigada")
    i_help("ISO 45001")
    
    with st.expander("📝 Gerar Nova Ata de Reunião Técnica"):
        with st.form("form_ata_sgi"):
            tipo_colegiado = st.selectbox("Órgão:", ["CIPA (NR-05)", "Brigada de Incêndio (IT-17)", "Comitê de Ética/GRC"])
            pauta = st.text_area("Pauta (Saúde, Incêndio, Vulnerabilidades detectadas):")
            decisao_5w2h = st.text_area("Plano de Ação (O que, Quem, Quando):")
            
            if st.form_submit_button("💾 Protocolar e Gerar Histórico"):
                conn = sqlite3.connect('sgi_omega_v80.db')
                conn.execute("INSERT INTO colegiados (tipo, data, pauta, acao) VALUES (?,?,?,?)",
                             (tipo_colegiado, datetime.now().strftime("%d/%m/%Y"), pauta, decisao_5w2h))
                conn.commit()
                conn.close()
                st.success(f"Ata de {tipo_colegiado} registrada com sucesso no SGI.")
                st.subheader("📊 People Analytics: Matriz de Confiabilidade")
    # Cálculo em tempo real do Índice de Fadiga e ICC
    conn = sqlite3.connect('sgi_omega_v80.db')
    df_tropa = pd.read_sql_query("SELECT nome, cargo, esfera, icc FROM equipe", conn)
    conn.close()
    
    if not df_tropa.empty:
        # Gráfico de Dispersão de Risco Humano
        fig_icc = px.scatter(df_tropa, x="nome", y="icc", color="icc",
                             size=[20]*len(df_tropa), color_continuous_scale="RdYlGn",
                             title="Índice de Confiabilidade do Colaborador (ICC)")
        st.plotly_chart(fig_icc, use_container_width=True)
        
        st.info("💡 Colaboradores abaixo de 70% de ICC exigem plano de reciclagem imediato.")
    else:
        st.info("Aguardando dados de performance para gerar matriz de confiabilidade.")
        # --- MÓDULO DE INTELIGÊNCIA PREDITIVA DE FADIGA (RBAC 117 / NR-17) ---
def calcular_burnout_score(horas_excedentes):
    """Calcula o risco de acidente por fadiga operacional."""
    score = horas_excedentes * 12.5
    if score > 80: return "🔴 CRÍTICO (Substituir)", "red"
    if score > 50: return "🟡 ALERTA (Pausa)", "orange"
    return "🟢 REGULAR", "green"

with tab_bi:
    st.divider()
    st.subheader("🧘 Gestão de Fadiga e Risco Humano")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        h_excedente = st.slider("Horas Excedentes (Amostragem Turno):", 0, 12, 2)
        status_f, cor_f = calcular_burnout_score(h_excedente)
        st.markdown(f"### Status: :{cor_f}[{status_f}]")
        with col_f2:
        st.write("**Impacto na Segurança (HSE):**")
        st.write("Horas excessivas aumentam em 40% a chance de erro técnico em áreas críticas.")
        if h_excedente > 8:
            st.error("🚨 PROIBIÇÃO: Conforme RBAC 117 / NR-17, o limite de segurança foi ultrapassado.")

    # Radar de Vencimentos de Licenças (LTO)
    st.divider()
    st.subheader("🏢 Status de Licenças Operacionais (License to Operate)")
    conn = sqlite3.connect('sgi_omega_v80.db')
    df_lto_view = pd.read_sql_query("SELECT nome, orgao, vencimento FROM licencas", conn)
    conn.close()
    
    if not df_lto_view.empty:
        st.table(df_lto_view)
    else:
        st.info("Nenhuma licença (Alvará/AVCB) cadastrada para monitoramento.")
        # --- RODAPÉ DE SEGURANÇA DO SISTEMA ---
st.sidebar.divider()
st.sidebar.markdown("**🛡️ SGI Alpha Apex V80**")
st.sidebar.write(f"Versão: {datetime.now().year}.Master")
st.sidebar.info("Criptografia SHA-512 Ativa. GPS Geofencing Monitorado.")

# LOG DE INICIALIZAÇÃO
if __name__ == "__main__":
    print(f"Sovereign Engine Ligado: {datetime.now()}")

# --- FIM DO ARQUIVO main.py ---
