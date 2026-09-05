import io
import re
import time
import urllib.parse
from datetime import date, datetime
import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(
    page_title="Controle de Vencimentos & Conformidade SST",
    page_icon="🛡️",
    layout="wide"
)

# Estilização CSS Clean Power BI + Responsividade Mobile
st.markdown("""
<style>
    .stApp {
        background-color: #F4F6F9;
        color: #1E293B;
    }
    .header-card {
        background-color: #FFFFFF;
        padding: 16px 20px;
        border-radius: 8px;
        margin-bottom: 16px;
        border: 1px solid #E2E8F0;
        border-left: 6px solid #0F766E;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.07);
    }
    .header-title {
        color: #0F172A !important;
        font-size: 22px;
        font-weight: 800;
        margin: 0;
        text-transform: uppercase;
    }
    .header-subtitle {
        color: #64748B !important;
        font-size: 12px;
        margin: 4px 0 0 0;
        font-weight: 500;
    }
    .kpi-card {
        background-color: #FFFFFF;
        border-radius: 8px;
        padding: 12px 6px;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.07);
        text-align: center;
        border: 1px solid #E2E8F0;
        margin-bottom: 8px;
    }
    .kpi-title {
        color: #64748B;
        font-size: 11px;
        font-weight: 700;
        margin-bottom: 4px;
        text-transform: uppercase;
    }
    .kpi-value {
        font-size: 24px;
        font-weight: 800;
    }
    .val-red { color: #DC2626 !important; }
    .val-yellow { color: #D97706 !important; }
    .val-green { color: #16A34A !important; }
    .val-blue { color: #0284C7 !important; }
    
    @media (max-width: 768px) {
        .header-title { font-size: 16px; }
        .kpi-value { font-size: 18px; }
        .kpi-title { font-size: 10px; }
    }
</style>
""", unsafe_allow_html=True)

# 1. Gerador de Planilha Modelo
@st.cache_data
def gerar_planilha_modelo_vencimentos():
    hoje = datetime.now()
    df_base = pd.DataFrame({
        "Categoria": [
            "ASO (NR-07)", "ASO (NR-07)",
            "CA de EPI (NR-06)", "CA de EPI (NR-06)",
            "Treinamento Normativo", "Treinamento Normativo",
            "Inspeção NR-13", "Inspeção NR-13",
            "Calibração de Instrumento", "Calibração de Instrumento"
        ],
        "Item_Colaborador_Equipamento": [
            "Carlos Eduardo Silva", "Mariana Costa",
            "Luva Nitrílica (CA 38290)", "Respirador PFF2 (CA 41550)",
            "NR-35 Trabalho em Altura - Roberto Dias", "NR-10 Segurança em Eletricidade - Ana Paula",
            "Caldeira Principal - B1", "Vaso de Pressão - Compressor Ar 02",
            "Dosímetro de Ruído - DOS-01", "Luxímetro Digital - LUX-03"
        ],
        "Setor": [
            "Operações", "Manutenção",
            "Almoxarifado Geral", "Pintura Industrial",
            "Montagem", "Elétrica",
            "Utilidades", "Usinagem",
            "SST / Higiene", "SST / Higiene"
        ],
        "Data_Validade": [
            (hoje - pd.Timedelta(days=15)).strftime("%d/%m/%Y"),
            (hoje + pd.Timedelta(days=18)).strftime("%d/%m/%Y"),
            (hoje - pd.Timedelta(days=5)).strftime("%d/%m/%Y"),
            (hoje + pd.Timedelta(days=25)).strftime("%d/%m/%Y"),
            (hoje + pd.Timedelta(days=120)).strftime("%d/%m/%Y"),
            (hoje - pd.Timedelta(days=2)).strftime("%d/%m/%Y"),
            (hoje + pd.Timedelta(days=10)).strftime("%d/%m/%Y"),
            (hoje + pd.Timedelta(days=240)).strftime("%d/%m/%Y"),
            (hoje - pd.Timedelta(days=30)).strftime("%d/%m/%Y"),
            (hoje + pd.Timedelta(days=90)).strftime("%d/%m/%Y")
        ],
        "Responsavel": [
            "Dr. Médico do Trabalho", "Dr. Médico do Trabalho",
            "Comprador SST", "Comprador SST",
            "Instrutor Interno", "Consultoria Externa",
            "Eng. Mecânico PH", "Eng. Mecânico PH",
            "Técnico SST", "Técnico SST"
        ],
        "Telefone_Responsavel": [
            "5581999990001", "5581999990001",
            "5581999990002", "5581999990002",
            "5581999990003", "5581999990004",
            "5581999990005", "5581999990005",
            "5581999990006", "5581999990006"
        ]
    })
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_base.to_excel(writer, sheet_name="Vencimentos", index=False)
    buffer.seek(0)
    return buffer

# 2. Barra Lateral: Configurações e Fonte de Dados
with st.sidebar:
    st.header("⚙️ Configurações")
    dias_aviso = st.slider("Avisar itens a vencer em até:", min_value=7, max_value=90, value=30, step=1)
    
    st.divider()
    st.subheader("📥 Planilha Modelo")
    st.download_button(
        label="⬇️ Baixar Planilha Padrão",
        data=gerar_planilha_modelo_vencimentos(),
        file_name="modelo_vencimentos_sst.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.divider()
    st.subheader("📊 Conectar seus Dados")
    
    tipo_fonte = st.radio(
        "Como deseja carregar os dados?",
        ["Planilha Online (Google Sheets)", "Subir Arquivo (.xlsx)"],
        index=0
    )
    
    url_sheets = ""
    upload_arquivo = None
    
    if tipo_fonte == "Planilha Online (Google Sheets)":
        url_sheets = st.text_input(
            "Cole o link do Google Sheets:",
            placeholder="https://docs.google.com/spreadsheets/d/...",
            help="Certifique-se de que a planilha esteja com acesso 'Qualquer pessoa com o link pode ler'."
        )
        if st.button("🔄 Atualizar Dados Agora", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
    else:
        upload_arquivo = st.file_uploader("Suba sua planilha (.xlsx)", type=["xlsx"])

# 3. Leitura e Processamento dos Dados
df = None

def carregar_google_sheets(url):
    id_match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    if not id_match:
        return None
    sheet_id = id_match.group(1)
    
    # Captura a aba específica (gid) para evitar ler a aba errada ou desatualizada
    gid_match = re.search(r"[#&?]gid=([0-9]+)", url)
    gid_str = f"&gid={gid_match.group(1)}" if gid_match else ""
    
    # Header anti-cache com timestamp
    timestamp = int(time.time())
    csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv{gid_str}&nocache={timestamp}"
    return pd.read_csv(csv_url)

if tipo_fonte == "Planilha Online (Google Sheets)" and url_sheets.strip() != "":
    try:
        df = carregar_google_sheets(url_sheets)
        if df is None or df.empty:
            st.sidebar.error("Não foi possível ler o link. Verifique as permissões da planilha.")
    except Exception as e:
        st.sidebar.error(f"Erro ao ler Google Sheets: {e}")

if df is None and upload_arquivo is not None:
    try:
        df = pd.read_excel(upload_arquivo, sheet_name="Vencimentos")
    except Exception:
        df = pd.read_excel(upload_arquivo)

if df is None:
    buffer = gerar_planilha_modelo_vencimentos()
    df = pd.read_excel(buffer, sheet_name="Vencimentos")

if "Telefone_Responsavel" not in df.columns:
    df["Telefone_Responsavel"] = ""

# Parser robusto para datas
def tratar_data_universal(valor):
    if pd.isna(valor) or str(valor).strip().lower() in ["nan", "nat", "", "none"]:
        return pd.NaT
    if isinstance(valor, (datetime, date)):
        return pd.to_datetime(valor)
    try:
        val_float = float(valor)
        return pd.to_datetime(val_float, unit='D', origin='1899-12-30')
    except (ValueError, TypeError):
        pass
    return pd.to_datetime(str(valor).strip(), dayfirst=True, errors="coerce")

df["Data_Validade"] = df["Data_Validade"].apply(tratar_data_universal)

hoje_ts = pd.to_datetime(date.today())
df["Dias_Restantes"] = (df["Data_Validade"] - hoje_ts).dt.days

def classificar_status(dias):
    if pd.isna(dias):
        return "Sem Data"
    elif dias < 0:
        return "🔴 Vencido"
    elif dias <= dias_aviso:
        return "🟡 A Vencer"
    else:
        return "🟢 Em Dia"

df["Status"] = df["Dias_Restantes"].apply(classificar_status)

# Filtros na Barra Lateral
with st.sidebar:
    st.divider()
    st.subheader("🔍 Filtros")
    categorias_lista = ["Todas"] + sorted(df["Categoria"].dropna().unique().tolist())
    categoria_sel = st.selectbox("Categoria:", categorias_lista)
    
    status_lista = ["Todos", "🔴 Vencido", "🟡 A Vencer", "🟢 Em Dia"]
    status_sel = st.selectbox("Status:", status_lista)

df_filtrado = df.copy()
if categoria_sel != "Todas":
    df_filtrado = df_filtrado[df_filtrado["Categoria"] == categoria_sel]
if status_sel != "Todos":
    df_filtrado = df_filtrado[df_filtrado["Status"] == status_sel]

# 4. Banner Superior
st.markdown("""
<div class="header-card">
    <div class="header-title">🛡️ CONTROLE DE VENCIMENTOS & CONFORMIDADE LEGAL SST</div>
    <div class="header-subtitle">Gestão Preventiva de ASOs • CAs de EPIs • Treinamentos • Inspeções NR-13 • Calibração de Equipamentos</div>
</div>
""", unsafe_allow_html=True)

# 5. Métricas de Topo
total_itens = len(df)
total_vencidos = len(df[df["Status"] == "🔴 Vencido"])
total_a_vencer = len(df[df["Status"] == "🟡 A Vencer"])
total_em_dia = len(df[df["Status"] == "🟢 Em Dia"])
taxa_conformidade = (total_em_dia / total_itens * 100) if total_itens > 0 else 100

k1, k2, k3, k4, k5 = st.columns(5)
with k1:
    st.markdown(f'<div class="kpi-card"><div class="kpi-title">Total Cadastrado</div><div class="kpi-value val-blue">{total_itens}</div></div>', unsafe_allow_html=True)
with k2:
    st.markdown(f'<div class="kpi-card"><div class="kpi-title">🔴 Vencidos (Crítico)</div><div class="kpi-value val-red">{total_vencidos}</div></div>', unsafe_allow_html=True)
with k3:
    st.markdown(f'<div class="kpi-card"><div class="kpi-title">🟡 A Vencer (≤ {dias_aviso} dias)</div><div class="kpi-value val-yellow">{total_a_vencer}</div></div>', unsafe_allow_html=True)
with k4:
    st.markdown(f'<div class="kpi-card"><div class="kpi-title">🟢 Em Dia / Conforme</div><div class="kpi-value val-green">{total_em_dia}</div></div>', unsafe_allow_html=True)
with k5:
    st.markdown(f'<div class="kpi-card"><div class="kpi-title">Conformidade Legal</div><div class="kpi-value">{taxa_conformidade:.1f}%</div></div>', unsafe_allow_html=True)

st.write("")

# 6. Gráficos Analíticos
c_graf1, c_graf2 = st.columns(2)
config_limpo = {"displayModeBar": False}

with c_graf1:
    dados_cat = df.groupby(["Categoria", "Status"]).size().reset_index(name="Qtd")
    fig_cat = px.bar(
        dados_cat, x="Categoria", y="Qtd", color="Status",
        barmode="stack",
        title="<b>Status de Conformidade por Categoria</b>",
        color_discrete_map={"🔴 Vencido": "#DC2626", "🟡 A Vencer": "#D97706", "🟢 Em Dia": "#16A34A"}
    )
    fig_cat.update_layout(
        plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
        margin=dict(l=15, r=15, t=40, b=80),
        xaxis=dict(tickfont=dict(color="#0F172A", size=10), showgrid=False, title=None),
        yaxis=dict(tickfont=dict(color="#0F172A", size=10), gridcolor="#E2E8F0", title="Quantidade"),
        legend=dict(orientation="h", yanchor="bottom", y=-0.50, xanchor="center", x=0.5, title=None)
    )
    st.plotly_chart(fig_cat, use_container_width=True, config=config_limpo)

with c_graf2:
    dados_status = df["Status"].value_counts().reset_index()
    dados_status.columns = ["Status", "Qtd"]
    fig_status = px.pie(
        dados_status, names="Status", values="Qtd",
        hole=0.55,
        title="<b>Distribuição Geral do Inventário</b>",
        color="Status",
        color_discrete_map={"🔴 Vencido": "#DC2626", "🟡 A Vencer": "#D97706", "🟢 Em Dia": "#16A34A"}
    )
    fig_status.update_layout(
        paper_bgcolor="#FFFFFF",
        margin=dict(l=15, r=15, t=40, b=20),
        legend=dict(orientation="v", yanchor="middle", y=0.5)
    )
    st.plotly_chart(fig_status, use_container_width=True, config=config_limpo)

st.write("")

# 7. Resumo Geral de WhatsApp
itens_alerta = df[df["Status"].isin(["🔴 Vencido", "🟡 A Vencer"])]
texto_alerta_zap = f"🚨 *ALERTA SST - CONTROLE DE VENCIMENTOS* 🚨%0A%0A"
texto_alerta_zap += f"Data: {date.today().strftime('%d/%m/%Y')}%0A"
texto_alerta_zap += f"Total Vencidos: {total_vencidos} | A Vencer (≤ {dias_aviso} dias): {total_a_vencer}%0A%0A"
for _, row in itens_alerta.head(8).iterrows():
    data_formatada = row['Data_Validade'].strftime('%d/%m/%Y') if pd.notna(row['Data_Validade']) else "Sem data"
    texto_alerta_zap += f"- [{row['Status']}] {row['Categoria']}: {row['Item_Colaborador_Equipamento']} (Prazo: {data_formatada})%0A"

link_zap_geral = f"https://api.whatsapp.com/send?text={texto_alerta_zap}"

st.markdown(f"""
    <div style="display: flex; justify-content: flex-end; margin-bottom: 12px;">
        <a href="{link_zap_geral}" target="_blank" style="text-decoration: none;">
            <div style="background-color: #25D366; color: white; padding: 10px 18px; border-radius: 6px; font-weight: 700; font-size: 14px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                📲 Compartilhar Resumo Geral no WhatsApp
            </div>
        </a>
    </div>
""", unsafe_allow_html=True)

# 8. Tabela de Gestão com Notificação Individual Direta
st.markdown("### 📋 Itens em Monitoramento")

def gerar_link_zap_individual(row):
    tel = str(row["Telefone_Responsavel"]).replace(".0", "").strip()
    if not tel or tel.lower() in ["nan", "none", ""]:
        return '<span style="color: #94A3B8;">Sem contato</span>'
    
    data_formatada = row['Data_Validade'].strftime('%d/%m/%Y') if pd.notna(row['Data_Validade']) else "Sem data"
    msg = f"Olá {row['Responsavel']}, atenção para a seguinte pendência de SST sob sua responsabilidade: {row['Item_Colaborador_Equipamento']} está com status {row['Status']} (Validade: {data_formatada})."
    msg_cod = urllib.parse.quote(msg)
    url = f"https://api.whatsapp.com/send?phone={tel}&text={msg_cod}"
    primeiro_nome = str(row["Responsavel"]).split()[0] if pd.notna(row["Responsavel"]) else "Responsável"
    return f'<a href="{url}" target="_blank" style="text-decoration: none; font-weight: bold; color: #0F766E;">📲 Notificar {primeiro_nome}</a>'

df_exibir = df_filtrado.copy()
df_exibir["Aviso WhatsApp"] = df_exibir.apply(gerar_link_zap_individual, axis=1)
df_exibir["Data_Validade"] = df_exibir["Data_Validade"].apply(
    lambda d: d.strftime("%d/%m/%Y") if pd.notna(d) else "Sem Data"
)

st.write(
    df_exibir[["Status", "Categoria", "Item_Colaborador_Equipamento", "Setor", "Data_Validade", "Dias_Restantes", "Responsavel", "Aviso WhatsApp"]].to_html(escape=False, index=False),
    unsafe_allow_html=True
)
