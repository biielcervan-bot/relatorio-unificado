import streamlit as st
import pandas as pd
import plotly.express as px
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import io

# -----------------------------------------------------------------------------
# CONFIGURAÇÃO DA PÁGINA
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Processador Unificado - Leitura & Entrega",
    page_icon="⚡",
    layout="wide"
)

st.title("⚡ Processador Operacional Unificado (Leitura & Entrega)")
st.markdown("Consolidação automática, padronização de campos e geração de relatórios de alta performance.")
st.markdown("---")

# -----------------------------------------------------------------------------
# FUNÇÃO CACHEADA DE PROCESSAMENTO E UNIFICAÇÃO
# -----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def processar_arquivos_unificados(uploaded_files):
    dfs_processados = []

    for file in uploaded_files:
        file_bytes = file.read()
        df_temp = None

        # Tenta decodificar o arquivo (.csv ou .xlsx) com múltiplos encodings e separadores
        if file.name.lower().endswith('.csv'):
            for enc in ['latin1', 'utf-8', 'iso-8859-1']:
                for sep in [';', ',']:
                    try:
                        temp = pd.read_csv(
                            io.BytesIO(file_bytes),
                            sep=sep,
                            encoding=enc,
                            low_memory=False,
                            on_bad_lines='skip'
                        )
                        if temp is not None and len(temp.columns) > 1:
                            df_temp = temp
                            break
                    except Exception:
                        continue
                if df_temp is not None:
                    break
        else:
            try:
                df_temp = pd.read_excel(io.BytesIO(file_bytes))
            except Exception:
                continue

        if df_temp is None or df_temp.empty:
            continue

        # Padroniza os nomes de colunas em maiúsculo e sem espaços nas pontas
        df_temp.columns = df_temp.columns.astype(str).str.strip().str.upper()

        # IDENTIFICAÇÃO DO TIPO DE PLANILHA (Entrega vs Leitura)
        is_entrega = ('DATA_HORA_APROXIMADA' in df_temp.columns or 'DAT_PREVISTA_ENTREGA' in df_temp.columns)

        df_padrao = pd.DataFrame()

        if is_entrega:
            # =========================================================
            # Mapeamento para Planilha de ENTREGA
            # =========================================================
            col_dt = 'DATA_HORA_APROXIMADA' if 'DATA_HORA_APROXIMADA' in df_temp.columns else 'DAT_PREVISTA_ENTREGA'
            col_base = 'NOM_BASE_OPERACIONAL' if 'NOM_BASE_OPERACIONAL' in df_temp.columns else 'BASE_OPERACIONAL'
            col_mun = 'NOM_MUNICIPIO' if 'NOM_MUNICIPIO' in df_temp.columns else 'MUNICIPIO'
            col_unidade = 'NOM_UNIDADE_LEITURA' if 'NOM_UNIDADE_LEITURA' in df_temp.columns else 'UNIDADE_LEITURA'
            col_agente = 'COD_AGENTE_COMERCIAL' if 'COD_AGENTE_COMERCIAL' in df_temp.columns else 'COD_AGENTE'
            col_tarefa = 'SEQ_TAREFA' if 'SEQ_TAREFA' in df_temp.columns else df_temp.index

            dt_series = pd.to_datetime(df_temp[col_dt], dayfirst=True, errors='coerce') if col_dt in df_temp.columns else pd.Series(pd.NaT, index=df_temp.index)

            df_padrao['DATA_HORA_DT'] = dt_series
            df_padrao['DATA_REAL'] = dt_series.dt.strftime('%d/%m/%Y').fillna('Sem Data')
            df_padrao['HORA'] = dt_series.dt.strftime('%H:%M').fillna('N/A')

            df_padrao['BASE_STD'] = df_temp[col_base].fillna('N/A').astype(str).str.strip() if col_base in df_temp.columns else 'N/A'
            df_padrao['MUNICIPIO_STD'] = df_temp[col_mun].fillna('N/A').astype(str).str.strip() if col_mun in df_temp.columns else 'N/A'
            df_padrao['UNIDADE_STD'] = df_temp[col_unidade].fillna('N/A').astype(str).str.strip() if col_unidade in df_temp.columns else 'N/A'
            df_padrao['COD_AGENTE_STD'] = df_temp[col_agente].fillna('N/A').astype(str).str.strip().str.replace(r'\.0$', '', regex=True) if col_agente in df_temp.columns else 'N/A'
            df_padrao['NOM_AGENTE_STD'] = ""
            df_padrao['LOTE_STD'] = ""
            df_padrao['LOCALIZACAO_STD'] = ""
            df_padrao['TIPO_ATIVIDADE_STD'] = "E"
            df_padrao['TAREFA_STD'] = df_temp[col_tarefa] if col_tarefa in df_temp.columns else df_temp.index

            # Premissas Fixas de Entrega
            df_padrao['IMP_GRUPO_1'] = 0
            df_padrao['IMP_GRUPO_2'] = 0
            df_padrao['QTD_FOTO_NUM'] = 0
            df_padrao['LEITURA_LIMPA'] = 1
            df_padrao['ORIGEM_DADO'] = "Entrega"

        else:
            # =========================================================
            # Mapeamento para Planilha de LEITURA
            # =========================================================
            col_dt = 'DT_INI_ACAO' if 'DT_INI_ACAO' in df_temp.columns else 'DAT_PREVISTA'
            col_base = 'NOM_BASE_OPERACIONAL' if 'NOM_BASE_OPERACIONAL' in df_temp.columns else 'BASE_OPERACIONAL'
            col_mun = 'NOM_MUNICIPIO' if 'NOM_MUNICIPIO' in df_temp.columns else 'MUNICIPIO'
            col_lote = 'LOTE' if 'LOTE' in df_temp.columns else 'NUM_LOTE'
            col_unidade = 'NOM_UNIDADE_LEITURA' if 'NOM_UNIDADE_LEITURA' in df_temp.columns else 'UNIDADE_LEITURA'
            col_cod_agente = 'COD_AGENTE' if 'COD_AGENTE' in df_temp.columns else 'CODIGO_AGENTE'
            col_nom_agente = 'AGENTE' if 'AGENTE' in df_temp.columns else 'NOM_AGENTE'
            col_loc = 'LOCALIZACAO' if 'LOCALIZACAO' in df_temp.columns else 'ZONA'
            col_tipo = 'TIPO_ATIVIDADE' if 'TIPO_ATIVIDADE' in df_temp.columns and df_temp['TIPO_ATIVIDADE'].notna().any() else 'IND_TIPO'
            col_status = 'IND_STATUS_VISITA' if 'IND_STATUS_VISITA' in df_temp.columns else 'STATUS'
            col_foto = 'QTD_FOTO' if 'QTD_FOTO' in df_temp.columns else 'FOTO'

            dt_series = pd.to_datetime(df_temp[col_dt], dayfirst=True, errors='coerce') if col_dt in df_temp.columns else pd.Series(pd.NaT, index=df_temp.index)

            df_padrao['DATA_HORA_DT'] = dt_series
            df_padrao['DATA_REAL'] = dt_series.dt.strftime('%d/%m/%Y').fillna('Sem Data')
            df_padrao['HORA'] = dt_series.dt.strftime('%H:%M').fillna('N/A')

            df_padrao['BASE_STD'] = df_temp[col_base].fillna('N/A').astype(str).str.strip() if col_base in df_temp.columns else 'N/A'
            df_padrao['MUNICIPIO_STD'] = df_temp[col_mun].fillna('N/A').astype(str).str.strip() if col_mun in df_temp.columns else 'N/A'
            df_padrao['LOTE_STD'] = df_temp[col_lote].fillna('').astype(str).str.strip() if col_lote in df_temp.columns else ''
            df_padrao['UNIDADE_STD'] = df_temp[col_unidade].fillna('N/A').astype(str).str.strip() if col_unidade in df_temp.columns else 'N/A'
            df_padrao['COD_AGENTE_STD'] = df_temp[col_cod_agente].fillna('N/A').astype(str).str.strip().str.replace(r'\.0$', '', regex=True) if col_cod_agente in df_temp.columns else 'N/A'
            df_padrao['NOM_AGENTE_STD'] = df_temp[col_nom_agente].fillna('').astype(str).str.strip() if col_nom_agente in df_temp.columns else ''
            df_padrao['LOCALIZACAO_STD'] = df_temp[col_loc].fillna('').astype(str).str.strip() if col_loc in df_temp.columns else ''
            
            if col_tipo in df_temp.columns:
                df_padrao['TIPO_ATIVIDADE_STD'] = df_temp[col_tipo].fillna('Leitura').astype(str).str.strip()
            else:
                df_padrao['TIPO_ATIVIDADE_STD'] = 'Leitura'

            df_padrao['TAREFA_STD'] = df_temp.index

            # Cálculo de Impedimentos G1 e G2
            status_series = df_temp[col_status].astype(str).str.upper() if col_status in df_temp.columns else pd.Series('', index=df_temp.index)
            df_padrao['IMP_GRUPO_1'] = status_series.apply(lambda x: 1 if 'G1' in x or 'GRUPO 1' in x else 0)
            df_padrao['IMP_GRUPO_2'] = status_series.apply(lambda x: 1 if 'G2' in x or 'GRUPO 2' in x else 0)
            df_padrao['LEITURA_LIMPA'] = ((df_padrao['IMP_GRUPO_1'] == 0) & (df_padrao['IMP_GRUPO_2'] == 0)).astype(int)

            # Cálculo do total de fotos
            df_padrao['QTD_FOTO_NUM'] = pd.to_numeric(df_temp[col_foto], errors='coerce').fillna(0).astype(int) if col_foto in df_temp.columns else 0
            df_padrao['ORIGEM_DADO'] = "Leitura"

        # Concatenação formatada do Código + Nome do Agente
        df_padrao['AGENTE_COMPLETO'] = df_padrao['COD_AGENTE_STD'] + df_padrao['NOM_AGENTE_STD'].apply(lambda x: f" - {x}" if x else "")
        dfs_processados.append(df_padrao)

    if not dfs_processados:
        return None

    return pd.concat(dfs_processados, ignore_index=True)

# -----------------------------------------------------------------------------
# INTERFACE PRINCIPAL DO STREAMLIT
# -----------------------------------------------------------------------------
uploaded_files = st.file_uploader(
    "📁 Envie suas planilhas de Leitura e/ou Entrega (.csv ou .xlsx):",
    type=["csv", "xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    with st.spinner("🚀 Processando e padronizando os arquivos..."):
        df = processar_arquivos_unificados(uploaded_files)

    if df is not None and not df.empty:
        # BARRA LATERAL: FILTROS DINÂMICOS
        st.sidebar.header("🎯 Filtros Unificados")

        def criar_multiselect(label, col_name):
            opcoes = sorted([str(x) for x in df[col_name].unique() if str(x) not in ['nan', 'N/A', 'Sem Data', '']])
            if not opcoes:
                opcoes = sorted([str(x) for x in df[col_name].unique() if str(x) != 'nan'])
            return st.sidebar.multiselect(label, options=opcoes, default=opcoes)

        f_origem = criar_multiselect("Origem do Dado", 'ORIGEM_DADO')
        f_base = criar_multiselect("Base Operacional", 'BASE_STD')
        f_mun = criar_multiselect("Município", 'MUNICIPIO_STD')
        f_unidade = criar_multiselect("Unidade de Leitura", 'UNIDADE_STD')
        f_loc = criar_multiselect("Localização", 'LOCALIZACAO_STD')
        f_tipo = criar_multiselect("Tipo / Atividade", 'TIPO_ATIVIDADE_STD')
        f_agente = criar_multiselect("Agente Comercial", 'AGENTE_COMPLETO')
        f_data_real = criar_multiselect("Data da Ação", 'DATA_REAL')

        # Aplicação Rápida dos Filtros
        df_filtrado = df[
            (df['ORIGEM_DADO'].astype(str).isin(f_origem)) &
            (df['BASE_STD'].astype(str).isin(f_base)) &
            (df['MUNICIPIO_STD'].astype(str).isin(f_mun)) &
            (df['UNIDADE_STD'].astype(str).isin(f_unidade)) &
            (df['LOCALIZACAO_STD'].astype(str).isin(f_loc)) &
            (df['TIPO_ATIVIDADE_STD'].astype(str).isin(f_tipo)) &
            (df['AGENTE_COMPLETO'].astype(str).isin(f_agente)) &
            (df['DATA_REAL'].astype(str).isin(f_data_real))
        ]

        if df_filtrado.empty:
            st.warning("⚠️ Nenhum registro encontrado para a combinação de filtros selecionada.")
        else:
            # 1. METRIC CARDS / INDICADORES GLOBAIS
            col1, col2, col3, col4, col5 = st.columns(5)
            col1.metric("Total Registros", f"{len(df_filtrado):,}".replace(",", "."))
            col2.metric("Registros Limpos", f"{df_filtrado['LEITURA_LIMPA'].sum():,}".replace(",", "."))
            col3.metric("Impedimentos G1", f"{df_filtrado['IMP_GRUPO_1'].sum():,}".replace(",", "."))
            col4.metric("Impedimentos G2", f"{df_filtrado['IMP_GRUPO_2'].sum():,}".replace(",", "."))
            col5.metric("Total de Fotos", f"{df_filtrado['QTD_FOTO_NUM'].sum():,}".replace(",", "."))

            st.markdown("---")

            # Helper para hora inicial e final de ação
            def hora_min_max(s, tipo):
                v = s.dropna()
                if v.empty: return "N/A"
                res = v.min() if tipo == 'min' else v.max()
                return res.strftime('%H:%M') if pd.notna(res) else "N/A"

            # 2. AGRUPAMENTO RESUMO CONSOLIDADO
            df_resumo = df_filtrado.groupby([
                'DATA_REAL', 'BASE_STD', 'MUNICIPIO_STD', 'LOTE_STD', 'UNIDADE_STD',
                'LOCALIZACAO_STD', 'TIPO_ATIVIDADE_STD', 'COD_AGENTE_STD', 'NOM_AGENTE_STD'
            ], as_index=False).agg(
                TOTAL_REGISTROS=('TAREFA_STD', 'count'),
                LEITURAS_LIMPAS=('LEITURA_LIMPA', 'sum'),
                IMP_G1=('IMP_GRUPO_1', 'sum'),
                IMP_G2=('IMP_GRUPO_2', 'sum'),
                TOTAL_FOTOS=('QTD_FOTO_NUM', 'sum'),
                HORA_INI=('DATA_HORA_DT', lambda x: hora_min_max(x, 'min')),
                HORA_FIM=('DATA_HORA_DT', lambda x: hora_min_max(x, 'max'))
            )

            df_resumo.columns = [
                'Data Realização', 'Base Operacional', 'Município', 'Lote',
                'Unidade de Leitura', 'Localização', 'Tipo / Atividade',
                'Código Agente', 'Nome Agente',
                'Total Registros', 'Limpos / Sucesso',
                'Impedimentos G1', 'Impedimentos G2',
                'Total Fotos', '1ª Ação', 'Última Ação'
            ]

            # 3. PAINEL DE GRÁFICOS
            col_graf1, col_graf2 = st.columns(2)
            with col_graf1:
                st.subheader("🏙️ Volume por Município")
                df_cidade = df_filtrado.groupby('MUNICIPIO_STD').size().reset_index(name='Qtd Registros')
                fig_cidade = px.bar(
                    df_cidade, x='MUNICIPIO_STD', y='Qtd Registros', text_auto=True,
                    labels={'MUNICIPIO_STD': 'Município', 'Qtd Registros': 'Volume'},
                    color_discrete_sequence=['#1f77b4']
                )
                fig_cidade.update_layout(xaxis_title="", yaxis_title="")
                st.plotly_chart(fig_cidade, use_container_width=True)

            with col_graf2:
                st.subheader("📊 Produção por Origem (Leitura vs Entrega)")
                df_origem_graf = df_filtrado.groupby('ORIGEM_DADO').size().reset_index(name='Qtd Registros')
                fig_origem = px.pie(
                    df_origem_graf, names='ORIGEM_DADO', values='Qtd Registros', hole=0.45,
                    color_discrete_sequence=['#2ca02c', '#ff7f0e']
                )
                st.plotly_chart(fig_origem, use_container_width=True)

            st.markdown("---")
            st.subheader("📋 Resumo Consolidado Operacional")
            st.dataframe(df_resumo, use_container_width=True)

            # 4. EXPORTAÇÃO EXCEL EM ALTA VELOCIDADE (OPENPYXL)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                # Aba 1: Resumo Consolidado
                df_resumo.to_excel(writer, sheet_name="Resumo Consolidado", index=False)

                # Aba 2: Base Filtrada Detalhada
                df_detalhado_export = df_filtrado[[
                    'ORIGEM_DADO', 'DATA_REAL', 'BASE_STD', 'MUNICIPIO_STD', 'LOTE_STD', 'UNIDADE_STD',
                    'LOCALIZACAO_STD', 'TIPO_ATIVIDADE_STD', 'COD_AGENTE_STD', 'NOM_AGENTE_STD', 'HORA',
                    'LEITURA_LIMPA', 'IMP_GRUPO_1', 'IMP_GRUPO_2', 'QTD_FOTO_NUM'
                ]].copy()

                df_detalhado_export.columns = [
                    'Origem', 'Data Realização', 'Base Operacional', 'Município', 'Lote', 'Unidade de Leitura',
                    'Localização', 'Tipo / Atividade', 'Código Agente', 'Nome Agente', 'Hora',
                    'Registro Limpo (1/0)', 'Impedimento G1', 'Impedimento G2', 'Qtd Fotos'
                ]
                df_detalhado_export.to_excel(writer, sheet_name="Base Filtrada Detalhada", index=False)

                # Formatação de Estilo Corporativo nas Abas
                workbook = writer.book
                header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

                for ws in workbook.worksheets:
                    for cell in ws[1]:
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = Alignment(horizontal="center", vertical="center")

                    for col in ws.columns:
                        col_letter = get_column_letter(col[0].column)
                        ws.column_dimensions[col_letter].width = 18

            st.download_button(
                label="📥 Baixar Relatório Unificado (Excel)",
                data=buffer.getvalue(),
                file_name="relatorio_unificado_operacional.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    else:
        st.error("Não foi possível extrair dados válidos dos arquivos fornecidos.")
else:
    st.info("👆 Selecione um ou mais arquivos de Leitura e/ou Entrega acima para iniciar o processamento.")
