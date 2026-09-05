import io
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

# Estilização CSS Clean Power BI + Tema Claro
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
</style>
""", unsafe_allow_html=True)

# 1. Gerador de Planilha Modelo com Abas Específicas
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
            (hoje - pd.Timedelta(days=15)).strftime("%Y-%m-%d"),  # Vencido
            (hoje + pd.Timedelta(days=18)).strftime("%Y-%m-%d"),  # Vencendo
            (hoje - pd.Timedelta(days=5)).strftime("%Y-%m-%d"),   # Vencido
            (hoje + pd.Timedelta(days=25)).strftime("%Y-%m-%d"),  # Vencendo
            (hoje + pd.Timedelta(days=120)).strftime("%Y-%m-%d"), # Em dia
            (hoje - pd.Timedelta(days=2)).strftime("%Y-%m-%d"),   # Vencido
            (hoje + pd.Timedelta(days=10)).strftime("%Y-%m-%d"),  # Vencendo
            (hoje + pd.Timedelta(days=240)).strftime("%Y-%m-%d"), # Em dia
            (hoje - pd.Timedelta(days=30)).strftime("%Y-%m-%d"),  # Vencido
            (hoje + pd.Timedelta(days=90)).strftime("%Y-%m-%d")   # Em dia
        ],
        "Responsavel": [
            "Dr. Médico do Trabalho", "Dr. Médico do Trabalho",
            "Comprador SST", "Comprador SST",
            "Instrutor Interno", "Consultoria Externa",
            "Eng. Mecânico PH", "Eng. Mecânico PH",
            "Técnico SST", "Técnico SST"
        ]
    })

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_base.to_excel(writer, sheet_name="Vencimentos", index=False)
    buffer.seek(0)
    return buffer

# 2. Barra Lateral: Configurações, Download e Upload
with st.sidebar:
    st.header("⚙️ Configuração de Alertas")
    dias_aviso = st.slider("Avisar itens a vencer em até:", min_value=7, max_value=90, value=30, step=1, help="Dias de antecedência para status de atenção")
    
    st.divider()
    st.subheader("📥 Planilha Modelo")
    st.download_button(
        label="⬇️ Baixar Planilha Padrão",
        data=gerar_planilha_modelo_vencimentos(),
        file_name="modelo_vencimentos_sst.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.divider()
    st.subheader("📁 Upload de Dados")
    upload_arquivo = st.file_uploader("Suba sua planilha (.xlsx)", type=["xlsx"])

# 3. Leitura e Processamento dos Dados
if upload_arquivo:
    try:
        df = pd.read_excel(upload_arquivo, sheet_name="Vencimentos")
    except Exception:
        df = pd.read_excel(upload_arquivo)
else:
    buffer = gerar_planilha_modelo_vencimentos()
    df = pd.read_excel(buffer, sheet_name="Vencimentos")

# Normalização e Cálculo de Prazos
df["Data_Validade"] = pd.to_datetime(df["Data_Validade"], errors="coerce")
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

# Aplicar filtros
df_filtrado = df.copy()
if categoria_sel != "Todas":
    df_filtrado = df_filtrado[df_filtrado["Categoria"] == categoria_sel]
if status_sel != "Todos":
    df_filtrado = df_filtrado[df_filtrado["Status"] == status_sel]

# 4. Cabeçalho Principal
st.markdown("""
<div class="header-card">
    <div class="header-title">🛡️ CONTROLE DE VENCIMENTOS & CONFORMIDADE LEGAL SST</div>
    <div class="header-subtitle">Gestão Preventiva de ASOs • CAs de EPIs • Treinamentos • Inspeções NR-13 • Calibração de Equipamentos</div>
</div>
""", unsafe_allow_html=True)

# 5. Métricas de Topo (Cards Resumo)
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
        margin=dict(l=15, r=15, t=40, b=20),
        xaxis=dict(tickfont=dict(color="#0F172A", size=10), showgrid=False),
        yaxis=dict(tickfont=dict(color="#0F172A", size=10), gridcolor="#E2E8F0"),
        legend=dict(orientation="h", yanchor="bottom", y=-0.35, xanchor="center", x=0.5)
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

# 7. Tabela de Gestão e Ações de Alerta
st.markdown("### 📋 Itens em Monitoramento")

# Geração de Mensagem para WhatsApp API
itens_alerta = df[df["Status"].isin(["🔴 Vencido", "🟡 A Vencer"])]
texto_alerta_zap = f"🚨 *ALERTA SST - CONTROLE DE VENCIMENTOS* 🚨%0A%0A"
texto_alerta_zap += f"Data do Relatório: {date.today().strftime('%d/%m/%Y')}%0A"
texto_alerta_zap += f"Total Vencidos: {total_vencidos}%0A"
texto_alerta_zap += f"Total A Vencer (≤ {dias_aviso} dias): {total_a_vencer}%0A%0A"
texto_alerta_zap += "*Atenção para as principais pendências:*%0A"

for _, row in itens_alerta.head(8).iterrows():
    texto_alerta_zap += f"- [{row['Status']}] {row['Categoria']}: {row['Item_Colaborador_Equipamento']} (Prazo: {row['Data_Validade'].strftime('%d/%m/%Y')})%0A"

link_zap = f"https://api.whatsapp.com/send?text={texto_alerta_zap}"

col_tab1, col_tab2 = st.columns([3, 1])
with col_tab2:
    st.markdown(f"""
        <a href="{link_zap}" target="_blank" style="text-decoration: none;">
            <div style="background-color: #25D366; color: white; padding: 10px 14px; border-radius: 6px; text-align: center; font-weight: 700; font-size: 13px; margin-bottom: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
                📲 Enviar Alerta via WhatsApp
            </div>
        </a>
    """, unsafe_allow_html=True)

df_exibir = df_filtrado.copy()
df_exibir["Data_Validade"] = df_exibir["Data_Validade"].dt.strftime("%d/%m/%Y")
st.dataframe(
    df_exibir[["Status", "Categoria", "Item_Colaborador_Equipamento", "Setor", "Data_Validade", "Dias_Restantes", "Responsavel"]],
    hide_index=True,
    use_container_width=True
)
