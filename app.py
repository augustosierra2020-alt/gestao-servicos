import io
import re
import pandas as pd
import streamlit as st
from openpyxl.styles import PatternFill
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Configuração da página do Streamlit
st.set_page_config(
    page_title="Gestão de Serviços & OS - Hyper Tork", page_icon="📊", layout="wide"
)

st.title("📊 Gestão de Serviços & Emissão de OS")
st.write(
    "Filtragem, cálculo de valores, remoção de duplicadas por Matrícula e geração automática de Ordens de Serviço (Word)."
)

# Criando abas para organizar o fluxo do sistema de forma limpa e profissional
aba1, aba2 = st.tabs(["📋 Processamento da Planilha", "📄 Gerar Ordem de Serviço"])

# Inicializando variáveis na sessão do Streamlit para compartilhar dados entre as abas
if "df_filtrado" not in st.session_state:
    st.session_state.df_filtrado = None
if "df_excel_final" not in st.session_state:
    st.session_state.df_excel_final = None
if "colunas_finais" not in st.session_state:
    st.session_state.colunas_finais = []

# --- FUNÇÃO DE CÁLCULO DE VALOR ---
def calcular_valor_inicial(linha):
    descricao = str(linha.get("Nome arquivo", "")).upper().strip()
    veiculo = str(linha.get("Fabricante", "")).upper().strip()

    descricao = re.sub(r"\s+", " ", descricao)
    veiculo = re.sub(r"\s+", " ", veiculo)

    termos_stg2 = ["STG2", "STG 2", "STAG2", "STAG 2"]
    termos_mod_off = ["MOD", "OFF"]

    fabricantes_especiais = [
        "NEW HOLLAND", "VALTRA", "CASE IH", "CASE", "MASSEY FERGUSSON", 
        "MASSEY", "CLAAS", "JHON DEERE", "JOHN DEERE", "DEERE", 
        "FENDT", "JACTO", "DOPPSTADT", "JAN", "VOLVO CONSTRUCTION EQUIPMENT", 
        "VOLVO CONSTRUCTION", "VOLVO CE"
    ]

    eh_especial = any(fab in veiculo for fab in fabricantes_especiais)
    if "VOLVO TRUCK" in veiculo:
        eh_especial = False

    if any(termo in descricao for termo in termos_stg2):
        return 1400 if eh_especial else 650
    elif any(termo in descricao for termo in termos_mod_off):
        return 700 if eh_especial else 350

    return None

# --- FUNÇÃO PARA FILTRAR APENAS OS TERMOS SOLICITADOS NA DESCRIÇÃO DA OS ---
def limpar_descricao_os(desc_original):
    desc_upper = str(desc_original).upper().strip()
    if "STAG 2" in desc_upper or "STAG2" in desc_upper:
        return "STAG 2"
    elif "STG 2" in desc_upper or "STG2" in desc_upper:
        return "STG 2"
    elif "MOD" in desc_upper:
        return "MOD"
    elif "OFF" in desc_upper:
        return "OFF"
    return desc_original

# --- FUNÇÃO AUXILIAR PARA COR DE FUNDO NAS CÉLULAS DO WORD ---
def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)
    tcPr.append(shd)

# --- FUNÇÃO PARA CRIAR E REDIGIR O ARQUIVO WORD (.DOCX) ---
def gerar_docx_os(flash_point, cliente_nome, cidade, contato, linhas_tabela, total_valor):
    doc = Document()
    
    # Ajustando margens da página do relatório
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # Estilo base Arial
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(11)

    # Título da Marca Cabeçalho
    p_header = doc.add_paragraph()
    p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_header.add_run("HYPER TORK PERFORMANCE\n")
    run_title.bold = True
    run_title.size = Pt(18)
    run_title.font.color.rgb = RGBColor(30, 41, 59) # Tom de Azul Escuro Corporativo
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = p_sub.add_run("Serviços Realizados Remap")
    run_sub.bold = True
    run_sub.size = Pt(13)
    run_sub.font.color.rgb = RGBColor(100, 116, 139)

    doc.add_paragraph() # Espaçamento

    # Tabela com as Informações Cadastrais solicitadas
    table_cli = doc.add_table(rows=3, cols=2)
    table_cli.style = 'Table Grid'
    
    labels_cli = [
        f"Cliente: {cliente_nome} - {flash_point}",
        f"Cidade: {cidade}",
        f"Contato: {contato}"
    ]
    
    for i, texto in enumerate(labels_cli):
        row = table_cli.rows[i]
        cell_lbl = row.cells[0]
        cell_lbl.text = texto
        row.cells[0].merge(row.cells[1]) # Ocupa toda a extensão da tabela
        for paragraph in cell_lbl.paragraphs:
            for run in paragraph.runs:
                run.font.name = 'Arial'
                run.font.size = Pt(11)

    doc.add_paragraph() 

    # Estrutura de Cabeçalho da Tabela de Serviços Reestruturada
    headers = ["Nº MAPA", "Data", "Veículo", "Placa", "Descrição", "Valor"]
    table_serv = doc.add_table(rows=1, cols=6)
    table_serv.style = 'Table Grid'
    
    hdr_cells = table_serv.rows[0].cells
    for i, header_text in enumerate(headers):
        hdr_cells[i].text = header_text
        set_cell_background(hdr_cells[i], "1E293B") # Header elegante escuro
        for p in hdr_cells[i].paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.bold = True
                r.font.color.rgb = RGBColor(255, 255, 255)
                r.font.name = 'Arial'
                r.font.size = Pt(10)

    # Inserindo os itens extraídos da planilha do cliente correspondente
    for linha in linhas_tabela:
        row_cells = table_serv.add_row().cells
        dados_linha = [
            str(linha.get("Nº Mapa", "")),
            str(linha.get("Data", "")),
            str(linha.get("Veículo", "")),
            str(linha.get("Placa", "")),
            limpar_descricao_os(linha.get("Descrição", "")),
            f"R$ {linha.get('Valor', '')}" if linha.get('Valor') is not None else ""
        ]
        for idx, valor_celula in enumerate(dados_linha):
            row_cells[idx].text = valor_celula
            for p in row_cells[idx].paragraphs:
                if idx in [0, 1, 3, 5]: # Alinha dados curtos de forma simétrica ao centro
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in p.runs:
                    r.font.name = 'Arial'
                    r.font.size = Pt(10)

    doc.add_paragraph()

    # Seção do Valor Total Formatado
    p_total = doc.add_paragraph()
    p_total.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_tot_lbl = p_total.add_run("TOTAL: ")
    r_tot_lbl.bold = True
    r_tot_lbl.size = Pt(12)
    r_tot_val = p_total.add_run(f"R$ {total_valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    r_tot_val.bold = True
    r_tot_val.size = Pt(12)
    r_tot_val.font.color.rgb = RGBColor(234, 88, 12) # Cor Laranja Escura Profissional para dar destaque

    doc.add_paragraph()

    # Bloco Fixo de Dados de Pagamentos da Empresa (Cópia fiel do rodapé do modelo original)
    p_pag = doc.add_paragraph()
    r_pag_title = p_pag.add_run("Formas de Pagamento\n")
    r_pag_title.bold = True
    r_pag_title.size = Pt(11)
    
    r_pag_det = p_pag.add_run(
        "DEPOSITO BANCARIO = AG: 0737 | C/C: 91538-4 | SICREDI\n"
        "PIX = Hyper tork Performance | CNPJ: 61.430.678.0001-07"
    )
    r_pag_det.size = Pt(10)
    r_pag_det.italic = True

    # Criação do buffer de saída
    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target


# ==============================================================================
# --- ABA 1: LÓGICA PRINCIPAL DE TRATAMENTO DA PLANILHA ---
# ==============================================================================
with aba1:
    arquivo_carregado = st.file_uploader(
        "Arraste ou selecione a planilha FPF_List para iniciar:",
        type=["xlsx", "xls", "csv"],
        key="uploader_planilha"
    )

    if arquivo_carregado is not None:
        try:
            conteudo = arquivo_carregado.read()
            try:
                excel_file = pd.ExcelFile(io.BytesIO(conteudo))
                abas = excel_file.sheet_names
                if len(abas) > 1:
                    aba_selecionada = st.selectbox("Selecione a aba com os dados:", abas, key="selecao_abas_app")
                else:
                    aba_selecionada = abas[0]
                df = pd.read_excel(io.BytesIO(conteudo), sheet_name=aba_selecionada)
            except Exception:
                try:
                    df = pd.read_csv(io.BytesIO(conteudo), sep=";", encoding="utf-8")
                    if df.shape[1] <= 1:
                        df = pd.read_csv(io.BytesIO(conteudo), sep=",", encoding="utf-8")
                except Exception:
                    df = pd.read_csv(io.BytesIO(conteudo), sep=";", encoding="iso-8859-1")

            if df is None or df.empty or len(df.columns) == 0:
                st.error("Erro: Não foi possível processar a estrutura de dados deste arquivo.")
            else:
                df.columns = df.columns.str.strip()

                # 1. Filtro Mandatório da coluna T contendo apenas "MOD"
                if "T" in df.columns:
                    df["T"] = df["T"].astype(str).str.strip()
                    df_filtrado = df[df["T"] == "MOD"].copy()
                else:
                    st.warning("Aviso: A coluna 'T' não foi encontrada.")
                    df_filtrado = df.copy()

                colunas_originais = ["Arquivo ID", "Fabricante", "Matrícula", "FlashPoint", "Cliente", "Nome arquivo", "Dada"]
                colunas_existentes = [col for col in colunas_originais if col in df_filtrado.columns]
                df_filtrado = df_filtrado[colunas_existentes].copy()

                # Geração de preço automatizada por regras
                df_filtrado["Valor"] = df_filtrado.apply(calcular_valor_inicial, axis=1)

                # Renomeação estrutural de colunas exigida
                dicionario_renomear = {
                    "Arquivo ID": "Nº Mapa",
                    "Fabricante": "Veículo",
                    "Matrícula": "Placa",
                    "Nome arquivo": "Descrição",
                    "Dada": "Data",
                    "FlashPoint": "Flash Point",
                }
                df_filtrado = df_filtrado.rename(columns=dicionario_renomear)

                # Ordenação exata de colunas da esquerda para a direita
                ordem_solicitada = ["Nº Mapa", "Data", "Veículo", "Placa", "Flash Point", "Descrição", "Valor"]
                colunas_finais = [col for col in ordem_solicitada if col in df_filtrado.columns]
                df_filtrado = df_filtrado[colunas_finais].copy()

                if "Flash Point" in df_filtrado.columns:
                    df_filtrado = df_filtrado.sort_values(by=["Flash Point", "Nº Mapa"] if "Nº Mapa" in df_filtrado.columns else ["Flash Point"], ascending=True)

                # Disponibilizando os dados limpos globalmente no sistema
                st.session_state.df_filtrado = df_filtrado.copy()

                # Construção dos agrupamentos visuais em blocos para a planilha Excel final
                if not df_filtrado.empty and "Flash Point" in df_filtrado.columns:
                    lista_linhas = []
                    linhas_amarelas = []
                    linhas_laranjas = []
                    contador_linha_excel = 2

                    for fp, bloco in df_filtrado.groupby("Flash Point", sort=False):
                        placas_vistas = set()

                        for idx, linha in bloco.iterrows():
                            linha_dict = linha.to_dict()
                            placa_atual = str(linha.get("Placa", "")).strip()

                            # Validação crítica contra placas repetidas no mesmo cliente
                            if placa_atual in placas_vistas and placa_atual != "":
                                linha_dict["Valor"] = None
                                linhas_amarelas.append(contador_linha_excel)
                            else:
                                if placa_atual != "":
                                    placas_vistas.add(placa_atual)

                            lista_linhas.append(linha_dict)
                            contador_linha_excel += 1

                        # Cálculo dinâmico da soma das linhas válidas do bloco
                        df_bloco_temp = pd.DataFrame(lista_linhas[-len(bloco) :])
                        soma_bloco = pd.to_numeric(df_bloco_temp["Valor"], errors="coerce").sum()

                        linha_total = {col: "" for col in colunas_finais}
                        linha_total["Flash Point"] = fp
                        linha_total["Descrição"] = "VALOR TOTAL:"
                        linha_total["Valor"] = float(soma_bloco) if soma_bloco > 0 else ""

                        lista_linhas.append(linha_total)
                        linhas_laranjas.append(contador_linha_excel)
                        contador_linha_excel += 1

                        linha_espacamento = {col: "" for col in colunas_finais}
                        lista_linhas.append(linha_espacamento)
                        contador_linha_excel += 1

                    if lista_linhas:
                        lista_linhas.pop()

                    df_excel_final = pd.DataFrame(lista_linhas, columns=colunas_finais)
                    st.session_state.df_excel_final = df_excel_final
                    st.session_state.colunas_finais = colunas_finais

                    st.subheader("📋 Visualização Prévia dos Dados Processados")
                    st.dataframe(df_filtrado)

                    # Exportando o binário do Excel aplicando as máscaras de preenchimento de cores
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                        df_excel_final.to_excel(writer, index=False, sheet_name="FPF Realizados")

                        workbook = writer.book
                        worksheet = writer.sheets["FPF Realizados"]

                        amarelo_claro = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
                        laranja_claro = PatternFill(start_color="FFE6CC", end_color="FFE6CC", fill_type="solid")

                        for num_linha in linhas_amarelas:
                            for col_idx in range(1, len(colunas_finais) + 1):
                                worksheet.cell(row=num_linha, column=col_idx).fill = amarelo_claro

                        for num_linha in linhas_laranjas:
                            for col_idx in range(1, len(colunas_finais) + 1):
                                worksheet.cell(row=num_linha, column=col_idx).fill = laranja_claro

                    st.success("Planilha processada com sucesso na memória! Vá para a aba ao lado se quiser gerar Ordens de Serviço.")
                    st.download_button(
                        label="📥 Baixar Planilha Processada (Excel)",
                        data=buffer.getvalue(),
                        file_name="FPF_Relatorio_Final.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
        except Exception as e:
            st.error(f"Erro crítico no processamento: {e}")

# ==============================================================================
# --- ABA 2: TELA INTEGRADA PARA EMISSÃO DA ORDEM DE SERVIÇO EM WORD ---
# ==============================================================================
with aba2:
    st.subheader("📄 Emissor de Ordem de Serviço por Flash Point")
    
    if st.session_state.df_filtrado is None or st.session_state.df_filtrado.empty:
        st.info("Aguardando o upload e processamento da planilha na aba anterior para habilitar a geração de OS.")
    else:
        df_base_os = st.session_state.df_filtrado
        
        # Criação dinâmica da lista de seleção com base exclusiva nos Flash Points existentes
        lista_fp_unicos = sorted(list(df_base_os["Flash Point"].unique()))
        
        fp_selecionado = st.selectbox("Selecione o Flash Point para gerar a OS correspondente:", lista_fp_unicos)
        
        # Filtra os dados apenas pertencentes ao Flash Point escolhido
        dados_bloco = df_base_os[df_base_os["Flash Point"] == fp_selecionado]
        
        # Coleta automática do nome do cliente associado àquele bloco
        cliente_sugerido = str(dados_bloco.iloc[0].get("Cliente", "Cliente Não Identificado"))
        
        # Interface de Inputs Cadastrais Organizada em Colunas Duplas
        col1, col2 = st.columns(2)
        with col1:
            nome_cliente_input = st.text_input("Cliente (Preenchido Automaticamente):", value=cliente_sugerido)
            cidade_input = st.text_input("Cidade (Adicionar a critério do usuário):", placeholder="Ex: Cascavel - PR")
        with col2:
            flash_point_confirmacao = st.text_input("Flash Point Relacionado:", value=fp_selecionado, disabled=True)
            contato_input = st.text_input("Contato (Adicionar a critério do usuário):", placeholder="Ex: (45) 99999-9999")
            
        st.write("### Serviços Vinculados que farão parte desta OS:")
        
        # Executa as mesmas regras de remoção de valor de placas duplicadas para a OS individual
        linhas_os_finais = []
        placas_vistas_os = set()
        soma_total_os = 0
        
        for idx, row in dados_bloco.iterrows():
            row_dict = row.to_dict()
            placa = str(row_dict.get("Placa", "")).strip()
            
            if placa in placas_vistas_os and placa != "":
                row_dict["Valor"] = None
            else:
                if placa != "":
                    placas_vistas_os.add(placa)
                if row_dict["Valor"] is not None:
                    soma_total_os += float(row_dict["Valor"])
            
            linhas_os_finais.append(row_dict)
            
        # Preview em tempo real da tabela formatada e tratada antes de exportar
        df_preview_os = pd.DataFrame(linhas_os_finais)
        df_preview_os["Descrição"] = df_preview_os["Descrição"].apply(limpar_descricao_os)
        st.dataframe(df_preview_os[["Nº Mapa", "Data", "Veículo", "Placa", "Descrição", "Valor"]])
        
        st.metric(label="Valor Total Consolidado da OS", value=f"R$ {soma_total_os:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        
        # Botão de Execução e Emissão da Ordem de Serviço
        if st.button("🚀 Gerar Ordem de Serviço"):
            arquivo_word = gerar_docx_os(
                flash_point=fp_selecionado,
                cliente_nome=nome_cliente_input,
                cidade=cidade_input,
                contato=contato_input,
                linhas_tabela=linhas_os_finais,
                total_valor=soma_total_os
            )
            
            st.success(f"Ordem de Serviço para o Flash Point {fp_selecionado} estruturada com sucesso!")
            
            st.download_button(
                label="📥 Baixar Ordem de Serviço Pronta (.docx)",
                data=arquivo_word.getvalue(),
                file_name=f"OS_Hyper_Tork_{fp_selecionado}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )