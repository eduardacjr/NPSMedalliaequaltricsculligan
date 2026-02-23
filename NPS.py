import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
from io import BytesIO 

# 1. Configuração da Página
st.set_page_config(
    page_title="Dashboard NPS - Diretoria",
    page_icon="📊",
    layout="wide"
)

# --- PALETAS DE CORES ---
CORES_NPS_PASTEL = {
    'Promotor': '#a8e6cf', 
    'Neutro': '#fdffab',   
    'Detrator': '#ffaaa5'  
}

CORES_BLUES = [
    '#08306b', '#08519c', '#2171b5', '#4292c6', 
    '#6baed6', '#9ecae1', '#c6dbef', '#deebf7'
]

# Cores Pastéis para os KPIs
KPI_BG_NPS = "#e3f2fd"   
KPI_BG_VOL = "#f5f5f5"   
KPI_BG_5ST = "#e8f5e9"   

# Constantes de nomes de arquivo
ARQUIVO_GERAL = "NPS Geral.xlsx"
ARQUIVO_CLASSIFICADO = "NPS Classificado.xlsx"
ARQUIVO_DATA = "data_atualizacao.txt"

# Mapa Global de Meses para formatação
MAPA_MESES_GLOBAL = {
    1: 'Jan', 2: 'Fev', 3: 'Mar', 4: 'Abr', 5: 'Mai', 6: 'Jun', 
    7: 'Jul', 8: 'Ago', 9: 'Set', 10: 'Out', 11: 'Nov', 12: 'Dez'
}

# --- FUNÇÕES AUXILIARES GLOBAIS ---
def classificar_nps(nota):
    if pd.isna(nota): return "Sem Nota"
    if nota >= 9: return "Promotor"
    elif nota >= 7: return "Neutro"
    else: return "Detrator"

def calcular_nps_score(df_input):
    if df_input.empty: return 0
    counts = df_input['Classificacao'].value_counts()
    total = len(df_input)
    if total == 0: return 0
    promotores = counts.get('Promotor', 0)
    detratores = counts.get('Detrator', 0)
    return ((promotores - detratores) / total) * 100

def fmt_milhar(valor):
    if pd.isna(valor): return "-"
    return f"{int(valor):,}".replace(",", ".")

def ler_data_atualizacao():
    if os.path.exists(ARQUIVO_DATA):
        with open(ARQUIVO_DATA, "r", encoding="utf-8") as f:
            return f.read().strip()
    return "Data n/d"

# Função para converter DataFrame em Excel Bytes para download
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Dados')
    processed_data = output.getvalue()
    return processed_data

# 2. Funções de Carregamento
@st.cache_data
def load_data_geral(file_path):
    if not os.path.exists(file_path): return None
    try:
        df = pd.read_excel(file_path)
        cols_data = ['Data de criação local', 'Data da resposta local', 'Data']
        for col in cols_data:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce', dayfirst=True)
        
        if 'NPS Purificador BTP' in df.columns:
            df['NPS Purificador BTP'] = pd.to_numeric(df['NPS Purificador BTP'], errors='coerce')
        if 'Avaliação do Técnico' in df.columns:
            df['Avaliação do Técnico'] = pd.to_numeric(df['Avaliação do Técnico'], errors='coerce')
        if 'Num OS' in df.columns:
            df['Num OS'] = df['Num OS'].astype(str).str.replace('.0', '', regex=False)

        df['Classificacao'] = df['NPS Purificador BTP'].apply(classificar_nps)
        df['Mes_Ano_Sort'] = df['Data da resposta local'].dt.to_period('M').astype(str)
        df['Ano'] = df['Data da resposta local'].dt.year
        
        df['Mes_Num'] = df['Data da resposta local'].dt.month
        df['Mes_Nome'] = df['Mes_Num'].map(MAPA_MESES_GLOBAL)
        
        # --- MAPEAMENTO DA FORMA JURÍDICA (NOVO) ---
        if 'Forma Jurídica' in df.columns:
            def map_segmento(val):
                if pd.isna(val): return 'Não Informado'
                v_str = str(val).strip()
                if v_str == 'Não atribuído': return 'PF'
                if v_str == 'P1': return 'PME'
                if v_str == 'C1': return 'Corporativo'
                return v_str
            df['Segmento'] = df['Forma Jurídica'].apply(map_segmento)
        else:
            df['Segmento'] = 'Não Informado'
            
        return df
    except Exception as e:
        st.error(f"Erro ao ler {file_path}: {e}")
        return None

@st.cache_data
def load_data_classificado(file_path):
    if not os.path.exists(file_path): return None
    try:
        df = pd.read_excel(file_path)
        if 'Data' in df.columns:
            df['Data'] = pd.to_datetime(df['Data'], errors='coerce', dayfirst=True)
            df['Mes_Ano_Sort'] = df['Data'].dt.to_period('M').astype(str)
            df['Ano'] = df['Data'].dt.year
            df['Mes_Num'] = df['Data'].dt.month
            df['Mes_Nome'] = df['Mes_Num'].map(MAPA_MESES_GLOBAL)
        
        if 'Num OS' in df.columns:
            df['Num OS'] = df['Num OS'].astype(str).str.replace('.0', '', regex=False)
            
        if 'NPS Purificador BTP' in df.columns:
             df['NPS Purificador BTP'] = pd.to_numeric(df['NPS Purificador BTP'], errors='coerce')
             df['Classificacao'] = df['NPS Purificador BTP'].apply(classificar_nps)
        else:
             df['Classificacao'] = None 
             
        return df
    except Exception as e:
        st.error(f"Erro ao ler {file_path}: {e}")
        return None

def filtrar_por_programa(df, coluna_programa, selecao):
    if selecao == "Geral": return df
    elif selecao == "Pós OS":
        if coluna_programa in df.columns: return df[df[coluna_programa] == "Pós OS"]
    else: 
        if coluna_programa in df.columns: return df[df[coluna_programa].str.contains("Instala", na=False, case=False)]
    return pd.DataFrame()

# --- KPI CARD (Cores Pastéis) ---
def criar_card_kpi(titulo, valor, cor_bg="#ffffff"):
    html_card = f"""
    <div style="
        background-color: {cor_bg};
        padding: 10px 5px;
        border-radius: 8px;
        box-shadow: 1px 1px 4px rgba(0,0,0,0.05);
        text-align: center;
        margin-bottom: 10px;
        height: 90px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
    ">
        <p style="
            font-size: 13px; 
            color: #555; 
            margin: 0; 
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        ">{titulo}</p>
        <p style="
            font-size: 22px; 
            font-weight: 800; 
            color: #333; 
            margin: 4px 0 0 0;
        ">{valor}</p>
    </div>
    """
    return st.markdown(html_card, unsafe_allow_html=True)

# --- FUNÇÃO DE AJUDA PARA TEXTO ---
def gerar_texto_ofensores(df_target):
    if df_target.empty: return "Sem dados."
    
    if 'Classificacao' not in df_target.columns: return "Sem classificação."
    df_probs = df_target[df_target['Classificacao'].isin(['Detrator', 'Neutro'])]
    
    if df_probs.empty: return "Não houveram Detratores/Neutros no período classificado."
    
    txt_saida = ""
    if 'Categorização Primária' in df_probs.columns:
        top_cats = df_probs['Categorização Primária'].value_counts().head(3)
        for cat, qtd in top_cats.items():
            txt_saida += f"- Macro: {cat} ({qtd} casos)\n"
            if 'Subcategorização Primária' in df_probs.columns:
                df_sub = df_probs[df_probs['Categorização Primária'] == cat]
                top_sub = df_sub['Subcategorização Primária'].value_counts().head(2)
                for sub, qtd_sub in top_sub.items():
                    txt_saida += f"   * Detalhe: {sub} ({qtd_sub})\n"
    return txt_saida

def gerar_texto_franquias(df_target):
    if df_target.empty or 'Franquia' not in df_target.columns: return "Sem dados."
    
    df_agg = df_target.groupby('Franquia').apply(
        lambda x: pd.Series({'NPS': calcular_nps_score(x), 'Vol': len(x)})
    ).reset_index()
    df_agg = df_agg[df_agg['Vol'] >= 3]
    
    if df_agg.empty: return "Volume insuficiente por franquia."
    
    melhores = df_agg.nlargest(3, 'NPS')
    piores = df_agg.nsmallest(3, 'NPS')
    
    txt = "TOP 3 (Melhores):\n" + "\n".join([f"- {r['Franquia']}: NPS {r['NPS']:.1f}" for _, r in melhores.iterrows()])
    txt += "\n\nBOTTOM 3 (Atenção):\n" + "\n".join([f"- {r['Franquia']}: NPS {r['NPS']:.1f}" for _, r in piores.iterrows()])
    return txt

# --- Interface Principal ---
st.sidebar.title("📊 Dashboard de Indicadores NPS")

# --- DATA DE ATUALIZAÇÃO ---
data_atualizacao = ler_data_atualizacao()
st.sidebar.markdown(f"**Atualizado em:** {data_atualizacao}")
st.sidebar.markdown("---")

df_geral = load_data_geral(ARQUIVO_GERAL)
df_classificado = load_data_classificado(ARQUIVO_CLASSIFICADO)

if df_geral is not None and df_classificado is not None:
    
    # GARANTIR QUE SEGMENTO ESTEJA NO CLASSIFICADO PARA OS FILTROS
    if 'Num OS' in df_geral.columns and 'Num OS' in df_classificado.columns:
        if 'Segmento' in df_geral.columns and 'Segmento' not in df_classificado.columns:
            temp_seg = df_geral[['Num OS', 'Segmento']].drop_duplicates('Num OS')
            df_classificado = df_classificado.merge(temp_seg, on='Num OS', how='left')
            df_classificado['Segmento'] = df_classificado['Segmento'].fillna('Não Informado')

    # --- FILTROS GLOBAIS ---
    st.sidebar.header("Filtros Globais")
    
    anos_disponiveis = sorted(df_geral['Ano'].dropna().unique().astype(int))
    opcoes_anos = ['Todos'] + [str(a) for a in anos_disponiveis]
    ano_selecionado = st.sidebar.selectbox("Selecione o Ano:", opcoes_anos)
    
    meses_ordem = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
    opcoes_meses = ['Todos'] + meses_ordem
    meses_selecionados = st.sidebar.multiselect("Selecione o(s) Mês(es):", options=opcoes_meses, default=['Todos'])

    # NOVO FILTRO: Segmento
    st.sidebar.markdown("---")
    if 'Segmento' in df_geral.columns:
        segmentos_disp = sorted([str(s) for s in df_geral['Segmento'].dropna().unique()])
        segmentos_selecionados = st.sidebar.multiselect("Selecione o Segmento:", options=['Todos'] + segmentos_disp, default=['Todos'])
    else:
        segmentos_selecionados = ['Todos']

    st.sidebar.markdown("---")

    # FILTRO: Franquias
    franquias_geral = set(df_geral['Franquia'].dropna().unique())
    franquias_class = set(df_classificado['Franquia'].dropna().unique()) if 'Franquia' in df_classificado.columns else set()
    todas_franquias = sorted(list(franquias_geral.union(franquias_class)))
    
    usar_todas_franquias = st.sidebar.checkbox("Selecionar Todas as Franquias", value=True)
    if usar_todas_franquias:
        franquias_selecionadas = todas_franquias
        st.sidebar.info(f"Todas as {len(todas_franquias)} franquias selecionadas.")
    else:
        franquias_selecionadas = st.sidebar.multiselect("Selecione as Franquias:", options=todas_franquias, default=[])
    st.sidebar.markdown("---")

    # --- APLICAÇÃO DOS FILTROS ---
    df_geral_filt = df_geral.copy()
    df_class_filt = df_classificado.copy()

    if ano_selecionado != 'Todos':
        df_geral_filt = df_geral_filt[df_geral_filt['Ano'] == int(ano_selecionado)]
        if 'Ano' in df_class_filt.columns: df_class_filt = df_class_filt[df_class_filt['Ano'] == int(ano_selecionado)]

    if "Todos" not in meses_selecionados:
        df_geral_filt = df_geral_filt[df_geral_filt['Mes_Nome'].isin(meses_selecionados)]
        if 'Mes_Nome' in df_class_filt.columns: df_class_filt = df_class_filt[df_class_filt['Mes_Nome'].isin(meses_selecionados)]

    if "Todos" not in segmentos_selecionados:
        df_geral_filt = df_geral_filt[df_geral_filt['Segmento'].isin(segmentos_selecionados)]
        if 'Segmento' in df_class_filt.columns: df_class_filt = df_class_filt[df_class_filt['Segmento'].isin(segmentos_selecionados)]

    if franquias_selecionadas:
        df_geral_filt = df_geral_filt[df_geral_filt['Franquia'].isin(franquias_selecionadas)]
        if 'Franquia' in df_class_filt.columns: df_class_filt = df_class_filt[df_class_filt['Franquia'].isin(franquias_selecionadas)]

    # --- KPIS ---
    df_pos_global = df_geral_filt[df_geral_filt['Programa de Pesquisa'] == 'Pós OS']
    df_inst_global = df_geral_filt[df_geral_filt['Programa de Pesquisa'].str.contains("Instala", na=False, case=False)]

    st.markdown("### Indicadores de Performance NPS")
    k1, k2, k3, k4, k5, k6, k7 = st.columns(7)
    
    with k1: criar_card_kpi("NPS", f"{calcular_nps_score(df_geral_filt):.1f}".replace('.', ','), KPI_BG_NPS)
    with k2: criar_card_kpi("Respostas", fmt_milhar(len(df_geral_filt)), KPI_BG_VOL)
    with k3: criar_card_kpi("NPS Pós OS", f"{calcular_nps_score(df_pos_global):.1f}".replace('.', ','), KPI_BG_NPS)
    with k4: criar_card_kpi("Resp. Pós OS", fmt_milhar(len(df_pos_global)), KPI_BG_VOL)
    with k5: criar_card_kpi("NPS Instalação", f"{calcular_nps_score(df_inst_global):.1f}".replace('.', ','), KPI_BG_NPS)
    with k6: criar_card_kpi("Resp. Instalação", fmt_milhar(len(df_inst_global)), KPI_BG_VOL)
    val_5s = df_geral_filt['Avaliação do Técnico'].mean()
    with k7: criar_card_kpi("5Star", f"{val_5s:.2f}".replace('.', ',') if pd.notnull(val_5s) else "-", KPI_BG_5ST)
    st.markdown("---")

    # --- ABAS ---
    tabs = st.tabs(["Visão Geral", "Análise Consolidada", "NPS Franquias Detratores e Neutros", "Classificação NPS", "5Star", "Detalhes", "🧠 Análises Avançadas"])
    (tab_visao, tab_consolidada, tab_franquia, tab_kpis, tab_tecnico, tab_detalhes, tab_analises) = tabs

    # 1. Visão Geral
    with tab_visao:
        def gerar_analise_nps_visual(dataframe, titulo_secao):
            if titulo_secao:
                st.subheader(f"{titulo_secao}")
                
            if dataframe.empty:
                st.info(f"Não há dados para {titulo_secao}.")
                return
            pivot_contagem = dataframe.pivot_table(index='Classificacao', columns='Mes_Ano_Sort', values='NPS Purificador BTP', aggfunc='count', fill_value=0).sort_index(axis=1)
            pivot_pct = pivot_contagem.div(pivot_contagem.sum(axis=0), axis=1) * 100
            
            s_prom = pivot_pct.loc['Promotor'] if 'Promotor' in pivot_pct.index else pd.Series(0, index=pivot_pct.columns)
            s_detr = pivot_pct.loc['Detrator'] if 'Detrator' in pivot_pct.index else pd.Series(0, index=pivot_pct.columns)
            nps_raw = (s_prom - s_detr).fillna(0)
            
            nps_text = [f"{x:.1f}".replace('.', ',') for x in nps_raw.values]

            df_chart = pivot_contagem.reset_index().melt(id_vars='Classificacao', var_name='Mes', value_name='Quantidade')
            
            titulo_grafico = "Evolução Mensal (NPS)" if titulo_secao == "NPS Geral" or titulo_secao == "" else f"Evolução Mensal - {titulo_secao}"
            
            fig = px.bar(df_chart, x='Mes', y='Quantidade', color='Classificacao', title=titulo_grafico, color_discrete_map=CORES_NPS_PASTEL, text_auto=True)
            
            fig.add_trace(go.Scatter(x=nps_raw.index.astype(str), y=[pivot_contagem.sum().max() * 1.1] * len(nps_raw), text=nps_text, mode='text+markers', textposition='top center', name='NPS Score', textfont=dict(size=14, color='black'), marker=dict(size=1, color='rgba(0,0,0,0)')))
            
            vals_x = df_chart['Mes'].unique()
            text_x = []
            for v in vals_x:
                try:
                    mes_num = int(v.split('-')[1])
                    text_x.append(MAPA_MESES_GLOBAL.get(mes_num, v))
                except:
                    text_x.append(v)

            fig.update_layout(
                barmode='stack', yaxis_title="Quantidade", xaxis_title="Mês", 
                yaxis_range=[0, pivot_contagem.sum().max() * 1.25], 
                margin=dict(t=40),
                title_font=dict(size=14),
                xaxis=dict(
                    tickmode='array',
                    tickvals=vals_x,
                    ticktext=text_x
                )
            )
            st.plotly_chart(fig, use_container_width=True)

            col_a, col_b, col_c = st.columns(3)
            def criar_tabela_cat_simples(nome, icone, cor):
                if nome in pivot_contagem.index:
                    st.markdown(f"##### {icone} <span style='color:{cor}'>{nome}s</span>", unsafe_allow_html=True)
                    s_qtd = pivot_contagem.loc[nome]
                    s_pct = pivot_pct.loc[nome]
                    df_d = pd.DataFrame([s_qtd.apply(lambda x: f"{int(x)}").tolist() + [f"{int(s_qtd.sum())}", f"{s_qtd.mean():.1f}".replace('.', ',')],
                                         s_pct.apply(lambda x: f"{x:.1f}%".replace('.', ',')).tolist() + ["-", "-"]], columns=list(s_qtd.index.astype(str)) + ['Soma', 'Média'], index=["Qtd", "%"])
                    st.dataframe(df_d, use_container_width=True)
                    
                    excel_data = convert_df_to_excel(df_d)
                    st.download_button(
                        label=f"📥 Baixar Tabela (.xlsx)",
                        data=excel_data,
                        file_name=f"Tabela_{nome}_{titulo_secao}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"btn_down_{nome}_{titulo_secao}"
                    )

            with col_a: criar_tabela_cat_simples("Promotor", "🟢", "green")
            with col_b: criar_tabela_cat_simples("Neutro", "🟡", "#b58900")
            with col_c: criar_tabela_cat_simples("Detrator", "🔴", "red")
            st.markdown("---")

        gerar_analise_nps_visual(df_geral_filt, "")
        
        if 'Programa de Pesquisa' in df_geral_filt.columns:
            gerar_analise_nps_visual(df_geral_filt[df_geral_filt['Programa de Pesquisa'] == 'Pós OS'], "Pós OS")
            gerar_analise_nps_visual(df_geral_filt[df_geral_filt['Programa de Pesquisa'].str.contains("Instala", na=False, case=False)], "Instalação")

    # 2. Análise Consolidada
    with tab_consolidada:
        st.subheader("📑 Visão Executiva Consolidada")
        st.info("""
        **Como ler o Mapa de Calor:** Este gráfico exibe a **concentração de casos** por categoria e mês.  
        As cores mais escuras (azul forte) indicam **maior volume de ocorrências**, facilitando a identificação imediata dos principais ofensores e padrões de sazonalidade.
        """)
        tipo_pes = st.radio("Programa:", ["Geral", "Pós OS", "Instalação"], horizontal=True, key="rd_cons")
        df_cons = filtrar_por_programa(df_class_filt, 'Programa de Pesquisa', tipo_pes)
        if not df_cons.empty and 'Categorização Primária' in df_cons.columns:
            meses_p = sorted(df_cons['Mes_Num'].unique())
            mapa_i = {1:'Jan', 2:'Fev', 3:'Mar', 4:'Abr', 5:'Mai', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Set', 10:'Out', 11:'Nov', 12:'Dez'}
            cols_ord = [mapa_i[m] for m in meses_p if m in mapa_i]
            piv_res = df_cons.pivot_table(index='Categorização Primária', columns='Mes_Nome', values='ID', aggfunc='count', fill_value=0).reindex(columns=cols_ord, fill_value=0)
            piv_res['Total'] = piv_res.sum(axis=1)
            piv_heat = piv_res.sort_values('Total', ascending=False).drop(columns=['Total'])
            
            fig = px.imshow(piv_heat, labels=dict(x="Mês", y="Categoria", color="Qtd"), color_continuous_scale='Blues', text_auto=True, aspect="auto", title="Mapa de Calor: Categoria x Mês")
            fig.update_layout(title_font=dict(size=14))
            st.plotly_chart(fig, use_container_width=True)
            st.markdown("---")
            st.markdown("### 📋 Top 5 Ofensores")
            for cat in piv_heat.index:
                with st.expander(f"📂 {cat}", expanded=False):
                    df_c = df_cons[df_cons['Categorização Primária'] == cat]
                    piv_s = df_c.pivot_table(index='Subcategorização Primária', columns='Mes_Nome', values='ID', aggfunc='count', fill_value=0).reindex(columns=cols_ord, fill_value=0)
                    piv_s['Total'] = piv_s.sum(axis=1)
                    
                    df_top5 = piv_s.nlargest(5, 'Total').drop(columns=['Total'])
                    st.dataframe(df_top5.style.background_gradient(cmap='Blues', axis=None).format("{:.0f}"), use_container_width=True)
                    
                    excel_data = convert_df_to_excel(df_top5)
                    st.download_button(
                        label=f"📥 Baixar {cat} (.xlsx)",
                        data=excel_data,
                        file_name=f"Top5_{cat}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"btn_down_{cat}"
                    )
        else: st.warning("Sem dados.")

    # 3. Franquias
    with tab_franquia:
        st.subheader("Análise Detalhada por Franquia")
        tp_frq = st.radio("Programa:", ["Geral", "Pós OS", "Instalação"], horizontal=True, key="rd_frq")
        df_fr = filtrar_por_programa(df_class_filt, 'Programa de Pesquisa', tp_frq)
        if not df_fr.empty and 'Franquia' in df_fr.columns:
            frqs = sorted(df_fr['Franquia'].dropna().unique())
            sel_fr = st.multiselect("Franquias (Visão Local):", ['Todas']+frqs, default=['Todas'])
            df_fin = df_fr if "Todas" in sel_fr else df_fr[df_fr['Franquia'].isin(sel_fr)]
            
            if not df_fin.empty and 'Subcategorização Primária' in df_fin.columns:
                c1, c2 = st.columns([3, 4])
                df_t = df_fin['Subcategorização Primária'].value_counts().reset_index()
                df_t.columns = ['Motivo', 'Quantidade']
                df_t['%'] = (df_t['Quantidade']/df_t['Quantidade'].sum()*100).map('{:.1f}%'.format)
                
                with c1:
                    st.dataframe(df_t, use_container_width=True, hide_index=True)
                    excel_data_t = convert_df_to_excel(df_t)
                    st.download_button(
                        label="📥 Baixar Resumo (.xlsx)",
                        data=excel_data_t,
                        file_name="Resumo_Franquias.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_down_frq_resumo"
                    )
                
                with c2:
                    df_g = df_t.copy()
                    if len(df_g)>5: df_g = pd.concat([df_g.iloc[:5], pd.DataFrame({'Motivo':['Outros'], 'Quantidade':[df_g.iloc[5:]['Quantidade'].sum()]})])
                    
                    fig = go.Figure(data=[go.Pie(
                        labels=df_g['Motivo'], 
                        values=df_g['Quantidade'], 
                        hole=0.5, 
                        marker=dict(colors=CORES_BLUES), 
                        showlegend=False,
                        textinfo='label+percent',
                        textposition='outside'
                    )])
                    fig.update_layout(margin=dict(t=40, b=20, l=120, r=120), title_text="Distribuição por Motivo", title_font=dict(size=14))
                    st.plotly_chart(fig, use_container_width=True)
                
                # MERGE PARA TRAZER FR, SEGMENTO, TECNICO E COMENTARIOS
                if 'Num OS' in df_fin.columns and 'Num OS' in df_geral.columns:
                    cols_source = ['Num OS', 'Comentário NPS Ecohouse', 'Franquia', 'Nome do Técnico', 'Segmento']
                    cols_source = [c for c in cols_source if c in df_geral.columns]
                    
                    temp_geral_info = df_geral[cols_source].drop_duplicates('Num OS')
                    df_fin = df_fin.merge(temp_geral_info, on='Num OS', how='left', suffixes=('', '_geral'))
                    
                    if 'Franquia_geral' in df_fin.columns: df_fin['Franquia'] = df_fin['Franquia'].fillna(df_fin['Franquia_geral'])
                    if 'Nome do Técnico_geral' in df_fin.columns: df_fin['Nome do Técnico'] = df_fin['Nome do Técnico_geral']
                    if 'Segmento_geral' in df_fin.columns: df_fin['Segmento'] = df_fin['Segmento'].fillna(df_fin['Segmento_geral'])
                    if 'Comentário NPS Ecohouse_geral' in df_fin.columns: df_fin['Comentário NPS Ecohouse'] = df_fin['Comentário NPS Ecohouse'].fillna(df_fin['Comentário NPS Ecohouse_geral'])

                cols_ver = ['Data', 'Num OS', 'Franquia', 'Segmento', 'Nome do Técnico', 'Categorização Primária', 'Subcategorização Primária', 'Comentário NPS Ecohouse']
                cols_fin = [c for c in cols_ver if c in df_fin.columns]
                
                df_extract = df_fin[cols_fin].sort_values('Data', ascending=False)
                st.dataframe(df_extract, use_container_width=True, hide_index=True)
                
                excel_data_ext = convert_df_to_excel(df_extract)
                st.download_button(
                    label="📥 Baixar Extrato Detalhado (.xlsx)",
                    data=excel_data_ext,
                    file_name="Extrato_Franquias.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_down_frq_extract"
                )

            else: st.warning("Sem dados.")
        else: st.warning("Sem dados.")

    # 4. Classificação NPS
    with tab_kpis:
        st.subheader("Classificação NPS")
        tp_kp = st.radio("Visão:", ["Geral", "Pós OS", "Instalação"], horizontal=True, key="rd_kp")
        df_kp = filtrar_por_programa(df_class_filt, 'Programa de Pesquisa', tp_kp)
        
        if not df_kp.empty:
            c1, c2 = st.columns([1, 1])
            
            with c1:
                st.markdown("**Resumo por Categoria**")
                res = df_kp.groupby(['Categorização Primária', 'Subcategorização Primária']).size().reset_index(name='Qtd').sort_values('Qtd', ascending=False)
                res['%'] = (res['Qtd']/res['Qtd'].sum()*100).map('{:.1f}%'.format)
                st.dataframe(res, use_container_width=True, hide_index=True)
                
                excel_data_res = convert_df_to_excel(res)
                st.download_button(
                    label="📥 Baixar Resumo (.xlsx)",
                    data=excel_data_res,
                    file_name="Resumo_Classificacao.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_down_class_resumo"
                )
            
            with c2:
                if 'Categorização Primária' in df_kp.columns:
                    piz = df_kp['Categorização Primária'].value_counts().reset_index()
                    piz.columns = ['Cat', 'Qtd']
                    if len(piz)>6: piz = pd.concat([piz.iloc[:6], pd.DataFrame({'Cat':['Outros'], 'Qtd':[piz.iloc[6:]['Qtd'].sum()]})])
                    
                    fig = go.Figure(data=[go.Pie(
                        labels=piz['Cat'], 
                        values=piz['Qtd'], 
                        hole=.5, 
                        marker=dict(colors=CORES_BLUES), 
                        showlegend=False,
                        textinfo='label+percent',
                        textposition='outside'
                    )])
                    fig.update_layout(
                        title_text="Distribuição Macro", 
                        margin=dict(t=40, b=20, l=60, r=60), 
                        height=350,
                        title_font=dict(size=14)
                    )
                    st.plotly_chart(fig, use_container_width=True)

            st.divider()
            st.subheader("🔎 Filtro de Detalhamento e Extrato")
            
            if 'Categorização Primária' in df_kp.columns:
                opcoes_cat = sorted(df_kp['Categorização Primária'].astype(str).unique())
                sel_cat_prim = st.multiselect("Selecione a Categoria Primária:", opcoes_cat)
                
                if sel_cat_prim:
                    df_filtered = df_kp[df_kp['Categorização Primária'].isin(sel_cat_prim)]
                else:
                    df_filtered = df_kp 
                
                if 'Subcategorização Primária' in df_filtered.columns:
                    opcoes_sub = sorted(df_filtered['Subcategorização Primária'].astype(str).unique())
                    sel_cat_sec = st.multiselect("Selecione a Subcategoria:", opcoes_sub)
                    
                    if sel_cat_sec:
                        df_final_extrato = df_filtered[df_filtered['Subcategorização Primária'].isin(sel_cat_sec)]
                    else:
                        df_final_extrato = df_filtered
                else:
                    df_final_extrato = df_filtered

                st.markdown("#### 📄 Extrato da Seleção")
                
                # MERGE PARA TRAZER FR, SEGMENTO, TECNICO E COMENTARIOS
                if 'Num OS' in df_final_extrato.columns and 'Num OS' in df_geral.columns:
                    cols_extra = ['Num OS', 'Comentário NPS Ecohouse', 'Franquia', 'Nome do Técnico', 'Segmento']
                    cols_extra = [c for c in cols_extra if c in df_geral.columns]
                    
                    temp_data = df_geral[cols_extra].drop_duplicates('Num OS')
                    df_final_extrato = df_final_extrato.merge(temp_data, on='Num OS', how='left', suffixes=('', '_geral'))
                    
                    if 'Franquia_geral' in df_final_extrato.columns: df_final_extrato['Franquia'] = df_final_extrato['Franquia'].fillna(df_final_extrato['Franquia_geral'])
                    if 'Nome do Técnico_geral' in df_final_extrato.columns: df_final_extrato['Nome do Técnico'] = df_final_extrato['Nome do Técnico_geral']
                    if 'Segmento_geral' in df_final_extrato.columns: df_final_extrato['Segmento'] = df_final_extrato['Segmento'].fillna(df_final_extrato['Segmento_geral'])
                    if 'Comentário NPS Ecohouse_geral' in df_final_extrato.columns: df_final_extrato['Comentário NPS Ecohouse'] = df_final_extrato['Comentário NPS Ecohouse'].fillna(df_final_extrato['Comentário NPS Ecohouse_geral'])

                cols_display = ['Data', 'Num OS', 'Franquia', 'Segmento', 'Nome do Técnico', 'Categorização Primária', 'Subcategorização Primária', 'Comentário NPS Ecohouse']
                cols_present = [c for c in cols_display if c in df_final_extrato.columns]
                
                df_display = df_final_extrato[cols_present].sort_values('Data', ascending=False)
                st.dataframe(df_display, use_container_width=True, hide_index=True)
                
                excel_data_display = convert_df_to_excel(df_display)
                st.download_button(
                    label="📥 Baixar Extrato Filtrado (.xlsx)",
                    data=excel_data_display,
                    file_name="Extrato_Classificacao_Filtrado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_down_class_filtered"
                )

        else: st.warning("Sem dados classificados para os filtros globais.")

    # 5. 5Star
    with tab_tecnico:
        st.subheader("⭐ Programa 5Star")
        cr, cf, cm = st.columns([1.2, 1, 1]) 
        
        tp_tec = cr.radio("Programa:", ["Geral", "Pós OS", "Instalação"], horizontal=True, key="rd_tec")
        ops = ['Todas'] + sorted(df_geral_filt['Franquia'].unique())
        sel_loc = cf.multiselect("Franquias:", ops, default=['Todas'])
        
        df_tc = filtrar_por_programa(df_geral_filt, 'Programa de Pesquisa', tp_tec)
        if "Todas" not in sel_loc: df_tc = df_tc[df_tc['Franquia'].isin(sel_loc)]
        
        if not df_tc.empty and 'Avaliação do Técnico' in df_tc.columns:
            media_val = df_tc['Avaliação do Técnico'].mean()
            with cm:
                criar_card_kpi("Média Geral", f"{media_val:.2f}", KPI_BG_5ST)
            
            df_evol = df_tc.groupby('Mes_Ano_Sort')['Avaliação do Técnico'].mean().reset_index()
            fig = px.bar(df_evol, x='Mes_Ano_Sort', y='Avaliação do Técnico', title="Evolução Mensal da Nota", text_auto='.2f', color_discrete_sequence=['#08306b'])
            
            vals_x_tc = df_evol['Mes_Ano_Sort'].unique()
            text_x_tc = []
            for v in vals_x_tc:
                try:
                    mes_num = int(v.split('-')[1])
                    text_x_tc.append(MAPA_MESES_GLOBAL.get(mes_num, v))
                except:
                    text_x_tc.append(v)
            
            fig.update_yaxes(range=[0, 5.5])
            fig.update_layout(
                title_font=dict(size=14), 
                margin=dict(t=40),
                xaxis=dict(
                    tickmode='array',
                    tickvals=vals_x_tc,
                    ticktext=text_x_tc
                )
            )
            st.plotly_chart(fig, use_container_width=True)
            
            st.markdown("---")
            sel_t = st.selectbox("Técnico:", ['Todos'] + sorted(df_tc['Nome do Técnico'].unique()))
            df_tf = df_tc if sel_t == 'Todos' else df_tc[df_tc['Nome do Técnico'] == sel_t]
            
            ind = df_tf.groupby(['Nome do Técnico', 'Franquia']).agg(Media=('Avaliação do Técnico', 'mean'), Qtd=('Avaliação do Técnico', 'count')).reset_index()
            
            df_rank = ind[ind['Qtd']>0].sort_values('Media', ascending=False)
            st.dataframe(df_rank.style.format({'Media':'{:.2f}'}).background_gradient(subset=['Media'], cmap='RdYlGn', vmin=1, vmax=5), use_container_width=True)
            
            excel_data_rank = convert_df_to_excel(df_rank)
            st.download_button(
                label="📥 Baixar Ranking (.xlsx)",
                data=excel_data_rank,
                file_name="Ranking_Tecnicos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_down_rank"
            )

            df_det_tec = df_tf[['Data da resposta local', 'Nome do Técnico', 'Avaliação do Técnico', 'Num OS', 'Franquia']].sort_values('Data da resposta local', ascending=False)
            st.dataframe(df_det_tec, use_container_width=True, hide_index=True)
            
            excel_data_det_tec = convert_df_to_excel(df_det_tec)
            st.download_button(
                label="📥 Baixar Extrato de Notas (.xlsx)",
                data=excel_data_det_tec,
                file_name="Extrato_Notas_Tecnicos.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_down_det_tec"
            )

        else: st.warning("Sem dados.")

    # 6. Detalhes
    with tab_detalhes:
        st.subheader("🔍 Detalhamento OS")
        os_in = st.text_input("Número da OS:")
        if os_in:
            clean = os_in.strip().replace('.0', '')
            rg = df_geral[df_geral['Num OS'] == clean]
            rc = df_classificado[df_classificado['Num OS'] == clean] if 'Num OS' in df_classificado.columns else pd.DataFrame()
            if rg.empty and rc.empty: st.error("Não encontrado.")
            else:
                if not rg.empty:
                    st.dataframe(rg, use_container_width=True, hide_index=True)
                    excel_data_rg = convert_df_to_excel(rg)
                    st.download_button(
                        label="📥 Baixar Dados Gerais (.xlsx)",
                        data=excel_data_rg,
                        file_name=f"OS_{clean}_Geral.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_down_os_geral"
                    )

                    if 'Comentário NPS Ecohouse' in rg.columns: st.info(f"Comentário: {rg.iloc[0]['Comentário NPS Ecohouse']}")
                if not rc.empty:
                    st.markdown(f"**Classificação:** `{rc.iloc[0].get('Categorização Primária','-')}` > `{rc.iloc[0].get('Subcategorização Primária','-')}`")
                    st.dataframe(rc, use_container_width=True, hide_index=True)
                    excel_data_rc = convert_df_to_excel(rc)
                    st.download_button(
                        label="📥 Baixar Classificação (.xlsx)",
                        data=excel_data_rc,
                        file_name=f"OS_{clean}_Classificado.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_down_os_class"
                    )

    # 7. Análises Avançadas
    with tab_analises:
        st.subheader("🧠 Command Center Estratégico & Insights")
        senha = st.text_input("🔒 Senha de Acesso:", type="password")
        
        if senha == "1010":
            modo_analise = st.radio(
                "Selecione o Modo de Análise:",
                ["📊 Fechamento do Período (Atual)", "⚔️ Comparativo Estratégico (A vs B)"],
                horizontal=True
            )
            st.markdown("---")

            if modo_analise == "📊 Fechamento do Período (Atual)":
                st.markdown("Gera um relatório completo sobre o período selecionado nos filtros laterais.")
                
                if 'Classificacao' not in df_class_filt.columns or df_class_filt['Classificacao'].isnull().all():
                     if 'Num OS' in df_class_filt.columns and 'Num OS' in df_geral_filt.columns:
                         temp_class = df_geral_filt[['Num OS', 'Classificacao']].drop_duplicates('Num OS')
                         df_class_filt = df_class_filt.merge(temp_class, on='Num OS', how='left', suffixes=('', '_y'))
                         if 'Classificacao_y' in df_class_filt.columns:
                             df_class_filt['Classificacao'] = df_class_filt['Classificacao'].fillna(df_class_filt['Classificacao_y'])

                df_pos = df_geral_filt[df_geral_filt['Programa de Pesquisa'] == 'Pós OS']
                df_inst = df_geral_filt[df_geral_filt['Programa de Pesquisa'].str.contains("Instala", na=False, case=False)]
                
                df_class_pos = filtrar_por_programa(df_class_filt, 'Programa de Pesquisa', "Pós OS")
                df_class_inst = filtrar_por_programa(df_class_filt, 'Programa de Pesquisa', "Instalação")

                txt_comentarios = "Sem comentários disponíveis."
                if 'Comentário NPS Ecohouse' in df_geral_filt.columns:
                    df_coments_det = df_geral_filt[
                        (df_geral_filt['Classificacao'] == 'Detrator') & 
                        (df_geral_filt['Comentário NPS Ecohouse'].notna()) & 
                        (df_geral_filt['Comentário NPS Ecohouse'] != '-')
                    ]
                    df_coments_prom = df_geral_filt[
                        (df_geral_filt['Classificacao'] == 'Promotor') & 
                        (df_geral_filt['Comentário NPS Ecohouse'].notna())
                    ]
                    
                    if not df_coments_det.empty:
                        sample_det = df_coments_det['Comentário NPS Ecohouse'].sample(min(len(df_coments_det), 3)).tolist()
                    else: sample_det = []
                    
                    if not df_coments_prom.empty:
                        sample_prom = df_coments_prom['Comentário NPS Ecohouse'].sample(min(len(df_coments_prom), 3)).tolist()
                    else: sample_prom = []
                    
                    if sample_det or sample_prom:
                        txt_comentarios = "💬 O que os Detratores estão dizendo (Amostra):\n" + "\n".join([f'- "{c}"' for c in sample_det])
                        txt_comentarios += "\n\n💬 O que os Promotores elogiam (Amostra):\n" + "\n".join([f'- "{c}"' for c in sample_prom])

                txt_pareto = "Dados insuficientes."
                if 'Avaliação do Técnico' in df_geral_filt.columns:
                    df_bad_tec = df_geral_filt[df_geral_filt['Avaliação do Técnico'] < 4]
                    total_tecnicos_ativos = df_geral_filt['Nome do Técnico'].nunique()
                    tecnicos_com_erro = df_bad_tec['Nome do Técnico'].nunique()
                    
                    if total_tecnicos_ativos > 0:
                        pct_impacto = (tecnicos_com_erro / total_tecnicos_ativos) * 100
                        txt_pareto = f"De {total_tecnicos_ativos} técnicos ativos no período, {tecnicos_com_erro} ({pct_impacto:.1f}%) receberam pelo menos uma avaliação negativa (<4)."
                        if pct_impacto < 20:
                            txt_pareto += " -> Problema CONCENTRADO em poucos indivíduos (Ação: Treinamento/Reciclagem pontual)."
                        else:
                            txt_pareto += " -> Problema SISTÊMICO espalhado na equipe (Ação: Revisão de Processo Global)."

                prompt_text = f"""
Atue como Head de Customer Experience. Analise os dados do dashboard (Filtros: {anos_disponiveis if ano_selecionado == 'Todos' else ano_selecionado} - {meses_selecionados} - Seg: {segmentos_selecionados}).

1. CONTEXTO GERAL
- Volume Total: {len(df_geral_filt)}
- NPS Global: {calcular_nps_score(df_geral_filt):.1f}

2. ANÁLISE QUALITATIVA (VOZ DO CLIENTE)
{txt_comentarios}

3. ANÁLISE DE EQUIPE (PARETO)
{txt_pareto}

4. DETALHAMENTO POR PROGRAMA
🅰️ PÓS OS (Reparo/Manutenção)
- NPS: {calcular_nps_score(df_pos):.1f} (Vol: {len(df_pos)})
- Principais Ofensores (Categorias > Subcategorias):
{gerar_texto_ofensores(df_class_pos)}
- Performance de Franquias (Pós OS):
{gerar_texto_franquias(df_pos)}

🅱️ INSTALAÇÃO (Novos Clientes)
- NPS: {calcular_nps_score(df_inst):.1f} (Vol: {len(df_inst)})
- Principais Ofensores:
{gerar_texto_ofensores(df_class_inst)}
- Performance de Franquias (Instalação):
{gerar_texto_franquias(df_inst)}

5. PROGRAMA TÉCNICO (5STAR)
- Nota Média Geral: {df_geral_filt['Avaliação do Técnico'].mean():.2f}/5.0
- Técnicos em Alerta (Nota < 4.5 e Vol > 3): 
{", ".join([f"{row['Nome do Técnico']} ({row['Avaliação do Técnico']:.1f})" for i, row in df_geral_filt[df_geral_filt['Avaliação do Técnico'] < 4.5].groupby('Nome do Técnico')['Avaliação do Técnico'].mean().reset_index().head(5).iterrows()])}

---
TAREFA:
Crie um relatório estratégico contendo:
1. **Análise de Sentimento:** Baseado nos comentários, qual é o tom emocional do cliente?
2. **Diagnóstico Operacional:** O problema é gente (Pareto) ou processo (Ofensores)?
3. **Plano de Ação:** 3 ações mandatórias para o próximo mês.
"""
                st.info("👇 Copie para IA (Gera Fechamento):")
                st.code(prompt_text, language="text")

            else:
                st.subheader("Selecione os Períodos para Comparação")
                c_a, c_b = st.columns(2)
                if 'Mes_Ano_Sort' in df_geral.columns:
                    periodos_disp = sorted(df_geral['Mes_Ano_Sort'].unique())
                else: periodos_disp = []

                with c_a: 
                    per_a = st.selectbox("📅 Período A (Base):", periodos_disp, index=len(periodos_disp)-2 if len(periodos_disp)>1 else 0)
                with c_b:
                    per_b = st.selectbox("📅 Período B (Atual/Comp):", periodos_disp, index=len(periodos_disp)-1 if len(periodos_disp)>0 else 0)

                if st.button("Gerar Comparativo Estratégico"):
                    df_a_geral = df_geral[df_geral['Mes_Ano_Sort'] == per_a]
                    df_b_geral = df_geral[df_geral['Mes_Ano_Sort'] == per_b]
                    
                    df_class_total = load_data_classificado(ARQUIVO_CLASSIFICADO)
                    if 'Num OS' in df_class_total.columns and 'Num OS' in df_geral.columns:
                         temp_cls = df_geral[['Num OS', 'Classificacao']].drop_duplicates('Num OS')
                         df_class_total = df_class_total.merge(temp_cls, on='Num OS', how='left', suffixes=('', '_y'))
                         if 'Classificacao_y' in df_class_total.columns:
                             df_class_total['Classificacao'] = df_class_total['Classificacao'].fillna(df_class_total['Classificacao_y'])
                    
                    df_a_class = df_class_total[df_class_total['Mes_Ano_Sort'] == per_a]
                    df_b_class = df_class_total[df_class_total['Mes_Ano_Sort'] == per_b]

                    nps_a, nps_b = calcular_nps_score(df_a_geral), calcular_nps_score(df_b_geral)
                    vol_a, vol_b = len(df_a_geral), len(df_b_geral)
                    tec_a, tec_b = df_a_geral['Avaliação do Técnico'].mean(), df_b_geral['Avaliação do Técnico'].mean()

                    delta_nps = nps_b - nps_a
                    delta_vol = vol_b - vol_a
                    
                    prompt_comp = f"""
Atue como Head de Estratégia CX. Realize uma análise comparativa (Year-over-Year ou Month-over-Month) entre dois períodos.

PERÍODO A ({per_a})  vs  PERÍODO B ({per_b})

1. KPIs MACRO
- NPS: {nps_a:.1f}  ➡️  {nps_b:.1f} (Delta: {delta_nps:+.1f})
- Volume: {vol_a}  ➡️  {vol_b} (Delta: {delta_vol:+})
- Nota Técnica: {tec_a:.2f} ➡️ {tec_b:.2f}

2. EVOLUÇÃO DOS OFENSORES (O problema mudou?)
[Período A - Detalhes]
{gerar_texto_ofensores(df_a_class)}

[Período B - Detalhes]
{gerar_texto_ofensores(df_b_class)}

3. DETALHE POR PROGRAMA (NPS A -> NPS B)
- Pós OS: {calcular_nps_score(df_a_geral[df_a_geral['Programa de Pesquisa']=='Pós OS']):.1f} -> {calcular_nps_score(df_b_geral[df_b_geral['Programa de Pesquisa']=='Pós OS']):.1f}
- Instalação: {calcular_nps_score(df_a_geral[df_a_geral['Programa de Pesquisa'].str.contains("Instala",na=False)]):.1f} -> {calcular_nps_score(df_b_geral[df_b_geral['Programa de Pesquisa'].str.contains("Instala",na=False)]):.1f}

---
TAREFA ANALÍTICA:
1. **Veredito da Evolução:** O NPS subiu ou caiu? Foi impulsionado por Pós OS ou Instalação?
2. **Análise de Causa Raiz:** O principal ofensor do Período A foi resolvido no B? Surgiu um novo ofensor crítico?
3. **Recomendação Tática:** O que a diretoria deve fazer para manter a tendência de alta ou reverter a queda no próximo ciclo?
"""
                    st.info("👇 Copie para IA (Gera Comparativo):")
                    st.code(prompt_comp, language="text")

        elif senha:
            st.error("Senha Incorreta.")

else:
    st.error(f"Arquivos {ARQUIVO_GERAL} e {ARQUIVO_CLASSIFICADO} não encontrados.")
