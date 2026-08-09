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
    page_title="Processador Operacional - Leituras e Entregas",
    page_icon="⚡",
    layout="wide"
)

st.title("⚡ Processador Operacional (Leituras Produtivas & Entregas)")
st.markdown("Filtragem estrita de IND_TIPO (P, R, C) com rastreamento de impedimentos via COD_NOTA_VISITA (Família 1 inicia com 1, Família 2 inicia com 2).")
st.markdown("---")

# -----------------------------------------------------------------------------
# FUNÇÃO DE PROCESSAMENTO
# -----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def processar_arquivos_unificados(uploaded_files):
    dfs_processados = []

    for file in uploaded_files:
        file.seek(0)
        file_bytes = file.read()
        df_temp = None

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

        # Normalização dos nomes das colunas
        df_temp.columns = df_temp.columns.astype(str).str.strip().str.upper()

        # Higienização de texto
        for col in df_temp.select_dtypes(include='object').columns:
            df_temp[col] = (
                df_temp[col]
                .astype(str)
                .str.replace(r'[\r\n]+', ' ', regex=True)
                .str.replace('"', '')
                .str.strip()
            )

        is_entrega = ('DATA_HORA_APROXIMADA' in df_temp.columns or 'DAT_PREVISTA_ENTREGA' in df_temp.columns)
        df_padrao = pd.DataFrame()

        if is_entrega:
            # =========================================================
            # REGISTROS DE ENTREGA (Sem Impedimentos)
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

            df_padrao['BASE_STD'] = df_temp[col_base].fillna('Não Informado').replace({'': 'Não Informado', 'nan': 'Não Informado'}).astype(str).str.strip() if col_base in df_temp.columns else 'Não Informado'
            df_padrao['MUNICIPIO_STD'] = df_temp[col_mun].fillna('Não Informado').replace({'': 'Não Informado', 'nan': 'Não Informado'}).astype(str).str.strip() if col_mun in df_temp.columns else 'Não Informado'
            df_padrao['UNIDADE_STD'] = df_temp[col_unidade].fillna('Não Informado').replace({'': 'Não Informado', 'nan': 'Não Informado'}).astype(str).str.strip() if col_unidade in df_temp.columns else 'Não Informado'
            
            cod_ag = df_temp[col_agente].fillna('Sem Código').astype(str).str.strip().str.replace(r'\.0$', '', regex=True) if col_agente in df_temp.columns else pd.Series('Sem Código', index=df_temp.index)
            df_padrao['COD_AGENTE_STD'] = cod_ag.replace({'nan': 'Sem Código', '': 'Sem Código'})
            df_padrao['NOM_AGENTE_STD'] = ""
            
            df_padrao['LOTE_STD'] = "(Sem Lote)"
            df_padrao['LOCALIZACAO_STD'] = "(N/A - Entrega)"
            
            df_padrao['IND_TIPO_STD'] = "E"
            df_padrao['TIPO_ATIVIDADE_STD'] = "Entrega"
            df_padrao['TAREFA_STD'] = df_temp[col_tarefa] if col_tarefa in df_temp.columns else df_temp.index

            # Entregas não possuem impedimentos operacionais
            df_padrao['IMP_GRUPO_1'] = 0
            df_padrao['IMP_GRUPO_2'] = 0
            df_padrao['TOTAL_IMP'] = 0
            df_padrao['LEITURA_LIMPA'] = 1
            df_padrao['ORIGEM_DADO'] = "Entrega"

            dfs_processados.append(df_padrao)

        else:
            # =========================================================
            # REGISTROS DE LEITURA (Filtro Estrito P, R, C + COD_NOTA_VISITA)
            # =========================================================
            col_dt = 'DT_INI_ACAO' if 'DT_INI_ACAO' in df_temp.columns else 'DAT_PREVISTA'
            col_base = 'NOM_BASE_OPERACIONAL' if 'NOM_BASE_OPERACIONAL' in df_temp.columns else 'BASE_OPERACIONAL'
            col_mun = 'NOM_MUNICIPIO' if 'NOM_MUNICIPIO' in df_temp.columns else 'MUNICIPIO'
            col_lote = 'LOTE' if 'LOTE' in df_temp.columns else 'NUM_LOTE'
            col_unidade = 'NOM_UNIDADE_LEITURA' if 'NOM_UNIDADE_LEITURA' in df_temp.columns else 'UNIDADE_LEITURA'
            col_cod_agente = 'COD_AGENTE' if 'COD_AGENTE' in df_temp.columns else 'CODIGO_AGENTE'
            col_nom_agente = 'AGENTE' if 'AGENTE' in df_temp.columns else 'NOM_AGENTE'
            col_loc = 'LOCALIZACAO' if 'LOCALIZACAO' in df_temp.columns else 'ZONA'
            col_ind_tipo = 'IND_TIPO' if 'IND_TIPO' in df_temp.columns else 'TIPO'
            col_tipo_atv = 'TIPO_ATIVIDADE' if 'TIPO_ATIVIDADE' in df_temp.columns else 'TIPO_SERVICO'
            col_nota = 'COD_NOTA_VISITA' if 'COD_NOTA_VISITA' in df_temp.columns else None

            # Filtro estrito: Apenas linhas onde IND_TIPO é P, R ou C
            if col_ind_tipo in df_temp.columns:
                df_temp = df_temp[df_temp[col_ind_tipo].astype(str).str.strip().isin(['P', 'R', 'C'])].copy()

            if not df_temp.empty:
                dt_series = pd.to_datetime(df_temp[col_dt], dayfirst=True, errors='coerce') if col_dt in df_temp.columns else pd.Series(pd.NaT, index=df_temp.index)

                df_padrao['DATA_HORA_DT'] = dt_series
                df_padrao['DATA_REAL'] = dt_series.dt.strftime('%d/%m/%Y').fillna('Sem Data')
                df_padrao['HORA'] = dt_series.dt.strftime('%H:%M').fillna('N/A')

                df_padrao['BASE_STD'] = df_temp[col_base].fillna('Não Informado').replace({'': 'Não Informado', 'nan': 'Não Informado'}).astype(str).str.strip() if col_base in df_temp.columns else 'Não Informado'
                df_padrao['MUNICIPIO_STD'] = df_temp[col_mun].fillna('Não Informado').replace({'': 'Não Informado', 'nan': 'Não Informado'}).astype(str).str.strip() if col_mun in df_temp.columns else 'Não Informado'
                
                lote_s = df_temp[col_lote].fillna('').astype(str).str.strip() if col_lote in df_temp.columns else pd.Series('', index=df_temp.index)
                df_padrao['LOTE_STD'] = lote_s.replace({'': '(Sem Lote)', 'nan': '(Sem Lote)'})
                
                df_padrao['UNIDADE_STD'] = df_temp[col_unidade].fillna('Não Informado').replace({'': 'Não Informado', 'nan': 'Não Informado'}).astype(str).str.strip() if col_unidade in df_temp.columns else 'Não Informado'
                
                cod_ag = df_temp[col_cod_agente].fillna('Sem Código').astype(str).str.strip().str.replace(r'\.0$', '', regex=True) if col_cod_agente in df_temp.columns else pd.Series('Sem Código', index=df_temp.index)
                df_padrao['COD_AGENTE_STD'] = cod_ag.replace({'nan': 'Sem Código', '': 'Sem Código'})
                
                nom_ag = df_temp[col_nom_agente].fillna('').astype(str).str.strip() if col_nom_agente in df_temp.columns else pd.Series('', index=df_temp.index)
                df_padrao['NOM_AGENTE_STD'] = nom_ag.replace({'nan': ''})
                
                loc_s = df_temp[col_loc].fillna('').astype(str).str.strip() if col_loc in df_temp.columns else pd.Series('', index=df_temp.index)
                df_padrao['LOCALIZACAO_STD'] = loc_s.replace({'': 'Não Informado', 'nan': 'Não Informado'})
                
                df_padrao['IND_TIPO_STD'] = df_temp[col_ind_tipo].astype(str).str.strip()
                
                atv_s = df_temp[col_tipo_atv].fillna('Leitura').astype(str).str.strip() if col_tipo_atv in df_temp.columns else pd.Series('Leitura', index=df_temp.index)
                df_padrao['TIPO_ATIVIDADE_STD'] = atv_s.replace({'': 'Leitura', 'nan': 'Leitura'})

                df_padrao['TAREFA_STD'] = df_temp.index

                # Tratamento estrito de COD_NOTA_VISITA para impedimentos
                if col_nota and col_nota in df_temp.columns:
                    nota_series = df_temp[col_nota].fillna('').astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                else:
                    nota_series = pd.Series('', index=df_temp.index)

                # Família 1: Começa com '1'
                df_padrao['IMP_GRUPO_1'] = nota_series.str.startswith('1').astype(int)
                # Família 2: Começa com '2'
                df_padrao['IMP_GRUPO_2'] = nota_series.str.startswith('2').astype(int)
                # Total de Impedimentos (G1 + G2)
                df_padrao['TOTAL_IMP'] = df_padrao['IMP_GRUPO_1'] + df_padrao['IMP_GRUPO_2']
                # Leitura Limpa: Sem impedimentos das famílias 1 ou 2
                df_padrao['LEITURA_LIMPA'] = (df_padrao['TOTAL_IMP'] == 0).astype(int)

                df_padrao['ORIGEM_DADO'] = "Leitura"

                dfs_processados.append(df_padrao)

    if not dfs_processados:
        return None

    df_concat = pd.concat(dfs_processados, ignore_index=True)

    # Preenchimento cruzado de agentes sem nome
    agentes_map = df_concat[df_concat['NOM_AGENTE_STD'] != ''].groupby('COD_AGENTE_STD')['NOM_AGENTE_STD'].first().to_dict()
    
    def atualizar_agente_completo(row):
        cod = row['COD_AGENTE_STD']
        nom = row['NOM_AGENTE_STD']
        if not nom and cod in agentes_map:
            nom = agentes_map[cod]
        return f"{cod} - {nom}" if nom else cod

    df_concat['AGENTE_COMPLETO'] = df_concat.apply(atualizar_agente_completo, axis=1)

    return df_concat

# -----------------------------------------------------------------------------
# INTERFACE PRINCIPAL
# -----------------------------------------------------------------------------
uploaded_files = st.file_uploader(
    "📁 Envie suas planilhas de Leitura e/ou Entrega (.csv ou .xlsx):",
    type=["csv", "xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    with st.spinner("🚀 Processando e calculando impedimentos via COD_NOTA_VISITA..."):
        df = processar_arquivos_unificados(uploaded_files)

    if df is not None and not df.empty:
        st.sidebar.header("🎯 Filtros Unificados")

        def criar_multiselect(label, col_name):
            opcoes = sorted([str(x) for x in df[col_name].unique() if str(x) not in ['nan', 'nan - nan']])
            return st.sidebar.multiselect(label, options=opcoes, default=opcoes)

        f_origem = criar_multiselect("Origem do Dado", 'ORIGEM_DADO')
        f_base = criar_multiselect("Base Operacional", 'BASE_STD')
        f_mun = criar_multiselect("Município", 'MUNICIPIO_STD')
        f_unidade = criar_multiselect("Unidade de Leitura", 'UNIDADE_STD')
        f_lote = criar_multiselect("Lote", 'LOTE_STD')
        f_loc = criar_multiselect("Localização", 'LOCALIZACAO_STD')
        f_ind_tipo = criar_multiselect("IND_TIPO (E, P, R, C)", 'IND_TIPO_STD')
        f_tipo_atv = criar_multiselect("Tipo de Atividade", 'TIPO_ATIVIDADE_STD')
        f_agente = criar_multiselect("Agente Comercial", 'AGENTE_COMPLETO')
        f_data_real = criar_multiselect("Data da Ação", 'DATA_REAL')

        # Aplicação dos Filtros
        df_filtrado = df[
            (df['ORIGEM_DADO'].astype(str).isin(f_origem)) &
            (df['BASE_STD'].astype(str).isin(f_base)) &
            (df['MUNICIPIO_STD'].astype(str).isin(f_mun)) &
            (df['UNIDADE_STD'].astype(str).isin(f_unidade)) &
            (df['LOTE_STD'].astype(str).isin(f_lote)) &
            (df['LOCALIZACAO_STD'].astype(str).isin(f_loc)) &
            (df['IND_TIPO_STD'].astype(str).isin(f_ind_tipo)) &
            (df['TIPO_ATIVIDADE_STD'].astype(str).isin(f_tipo_atv)) &
            (df['AGENTE_COMPLETO'].astype(str).isin(f_agente)) &
            (df['DATA_REAL'].astype(str).isin(f_data_real))
        ]

        if df_filtrado.empty:
            st.warning("⚠️ Nenhum registro encontrado para a combinação de filtros selecionada.")
        else:
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total Registros", f"{len(df_filtrado):,}".replace(",", "."))
            col2.metric("Registros Limpos", f"{df_filtrado['LEITURA_LIMPA'].sum():,}".replace(",", "."))
            col3.metric("Impedimentos G1 (Inicia com 1)", f"{df_filtrado['IMP_GRUPO_1'].sum():,}".replace(",", "."))
            col4.metric("Impedimentos G2 (Inicia com 2)", f"{df_filtrado['IMP_GRUPO_2'].sum():,}".replace(",", "."))
            
            # Métrica extra destacada para o Total de Impedimentos
            total_geral_imp = df_filtrado['TOTAL_IMP'].sum()
            st.info(f"🚨 **Total Geral de Impedimentos Operacionais:** {total_geral_imp:,}".replace(",", "."))

            st.markdown("---")

            def hora_min_max(s, tipo):
                v = s.dropna()
                if v.empty: return "N/A"
                res = v.min() if tipo == 'min' else v.max()
                return res.strftime('%H:%M') if pd.notna(res) else "N/A"

            df_resumo = df_filtrado.groupby([
                'DATA_REAL', 'BASE_STD', 'MUNICIPIO_STD', 'LOTE_STD', 'UNIDADE_STD',
                'LOCALIZACAO_STD', 'IND_TIPO_STD', 'TIPO_ATIVIDADE_STD', 'AGENTE_COMPLETO'
            ], as_index=False).agg(
                TOTAL_REGISTROS=('TAREFA_STD', 'count'),
                LEITURAS_LIMPAS=('LEITURA_LIMPA', 'sum'),
                IMP_G1=('IMP_GRUPO_1', 'sum'),
                IMP_G2=('IMP_GRUPO_2', 'sum'),
                TOTAL_IMP=('TOTAL_IMP', 'sum'),
                HORA_INI=('DATA_HORA_DT', lambda x: hora_min_max(x, 'min')),
                HORA_FIM=('DATA_HORA_DT', lambda x: hora_min_max(x, 'max'))
            )

            df_resumo.columns = [
                'Data Realização', 'Base Operacional', 'Município', 'Lote',
                'Unidade de Leitura', 'Localização', 'IND_TIPO', 'Tipo Atividade',
                'Agente Comercial', 'Total Registros', 'Limpos / Sucesso',
                'Impedimentos G1', 'Impedimentos G2', 'Total Impedimentos', '1ª Ação', 'Última Ação'
            ]

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
                st.subheader("📊 Produção por IND_TIPO")
                df_ind_graf = df_filtrado.groupby('IND_TIPO_STD').size().reset_index(name='Qtd Registros')
                fig_ind = px.pie(
                    df_ind_graf, names='IND_TIPO_STD', values='Qtd Registros', hole=0.45,
                    color_discrete_sequence=px.colors.qualitative.Set2
                )
                st.plotly_chart(fig_ind, use_container_width=True)

            st.markdown("---")
            st.subheader("📋 Resumo Consolidado por Atividade e Agente")
            st.dataframe(df_resumo, use_container_width=True)

            # Exportação Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_resumo.to_excel(writer, sheet_name="Resumo Consolidado", index=False)

                df_detalhado_export = df_filtrado[[
                    'ORIGEM_DADO', 'DATA_REAL', 'BASE_STD', 'MUNICIPIO_STD', 'LOTE_STD', 'UNIDADE_STD',
                    'LOCALIZACAO_STD', 'IND_TIPO_STD', 'TIPO_ATIVIDADE_STD', 'AGENTE_COMPLETO', 'HORA',
                    'LEITURA_LIMPA', 'IMP_GRUPO_1', 'IMP_GRUPO_2', 'TOTAL_IMP'
                ]].copy()

                df_detalhado_export.columns = [
                    'Origem', 'Data Realização', 'Base Operacional', 'Município', 'Lote', 'Unidade de Leitura',
                    'Localização', 'IND_TIPO', 'Tipo Atividade', 'Agente Comercial', 'Hora',
                    'Registro Limpo (1/0)', 'Impedimento G1', 'Impedimento G2', 'Total Impedimentos'
                ]
                df_detalhado_export.to_excel(writer, sheet_name="Base Filtrada Detalhada", index=False)

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
                        ws.column_dimensions[col_letter].width = 20

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
