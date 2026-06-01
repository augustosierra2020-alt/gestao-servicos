import io
import re
import pandas as pd
import streamlit as st
from openpyxl.styles import PatternFill

# Configuração da página do Streamlit
st.set_page_config(
    page_title="Gestão de Serviços - FPF", page_icon="📊", layout="centered"
)

st.title("📊 Gestão de Serviços Realizados")
st.write(
    "Filtragem, cálculo de valores, remoção de duplicadas e formatação avançada por blocos."
)

# 1. Componente para carregar o arquivo
arquivo_carregado = st.file_uploader(
    "Arraste ou selecione a planilha (Padrão: FPF List)",
    type=["xlsx", "xls", "csv"],
)


# Função para definir o valor inicial com base na Descrição (Nome do arquivo) e Veículo (Fabricante)
def calcular_valor_inicial(linha):
    # Usamos .get() com o nome antigo das colunas porque a função roda antes da renomeação final
    descricao = str(linha.get("Nome arquivo", "")).upper().strip()
    veiculo = str(linha.get("Fabricante", "")).upper().strip()

    # Remove múltiplos espaços internos
    descricao = re.sub(r"\s+", " ", descricao)
    veiculo = re.sub(r"\s+", " ", veiculo)

    termos_stg2 = ["STG2", "STG 2", "STAG2", "STAG 2"]
    termos_mod_off = ["MOD", "OFF"]

    fabricantes_especiais = [
        "NEW HOLLAND",
        "VALTRA",
        "CASE IH",
        "CASE",
        "MASSEY FERGUSSON",
        "MASSEY",
        "CLAAS",
        "JHON DEERE",
        "JOHN DEERE",
        "DEERE",
        "FENDT",
        "JACTO",
        "DOPPSTADT",
        "JAN",
        "VOLVO CONSTRUCTION EQUIPMENT",
        "VOLVO CONSTRUCTION",
        "VOLVO CE",
    ]

    eh_especial = any(fab in veiculo for fab in fabricantes_especiais)

    if "VOLVO TRUCK" in veiculo:
        eh_especial = False

    # 1. Regra para STG2
    if any(termo in descricao for termo in termos_stg2):
        return 1400 if eh_especial else 650

    # 2. Regra para MOD / OFF
    elif any(termo in descricao for termo in termos_mod_off):
        return 700 if eh_especial else 350

    return None


if arquivo_carregado is not None:
    df = None

    try:
        # --- LEITURA ROBUSTA ---
        conteudo = arquivo_carregado.read()

        try:
            excel_file = pd.ExcelFile(io.BytesIO(conteudo))
            abas = excel_file.sheet_names

            if len(abas) > 1:
                aba_selecionada = st.selectbox(
                    "Identificamos mais de uma aba no arquivo. Selecione a aba com os dados:",
                    abas,
                )
            else:
                aba_selecionada = abas[0]

            df = pd.read_excel(io.BytesIO(conteudo), sheet_name=aba_selecionada)
        except Exception:
            try:
                df = pd.read_csv(
                    io.BytesIO(conteudo), sep=";", encoding="utf-8"
                )
                if df.shape[1] <= 1:
                    df = pd.read_csv(
                        io.BytesIO(conteudo), sep=",", encoding="utf-8"
                    )
            except Exception:
                df = pd.read_csv(
                    io.BytesIO(conteudo), sep=";", encoding="iso-8859-1"
                )

        # --- PROCESSAMENTO DOS DADOS ---
        if df is None or df.empty or len(df.columns) == 0:
            st.error(
                "Erro: Não foi possível estruturar os dados deste arquivo."
            )
        else:
            df.columns = df.columns.str.strip()

            # 2. Filtrar a coluna "T" (Manter apenas "MOD", excluir "ORI")
            if "T" in df.columns:
                df["T"] = df["T"].astype(str).str.strip()
                df_filtrado = df[df["T"] == "MOD"].copy()
            else:
                st.warning("Aviso: A coluna 'T' não foi encontrada.")
                df_filtrado = df.copy()

            # Mapeamento de colunas originais necessárias
            colunas_originais = [
                "Arquivo ID",
                "Fabricante",
                "Matrícula",
                "FlashPoint",
                "Cliente",
                "Nome arquivo",
                "Dada",  # Lendo como 'Dada' inicialmente
            ]

            colunas_existentes = [
                col for col in colunas_originais if col in df_filtrado.columns
            ]

            df_filtrado = df_filtrado[colunas_existentes].copy()

            # --- GERAÇÃO DA COLUNA VALOR AUTOMÁTICA ---
            df_filtrado["Valor"] = df_filtrado.apply(
                calcular_valor_inicial, axis=1
            )

            # 3. RENOMEAR AS COLUNAS CONFORME SOLICITADO
            # Criamos um dicionário de tradução
            dicionario_renomear = {
                "Arquivo ID": "Nº Mapa",
                "Fabricante": "Veículo",
                "Matrícula": "Placa",
                "Nome arquivo": "Descrição",
                "Dada": "Data",  # Troca 'Dada' por 'Data'
                "FlashPoint": "Flash Point",  # Ajuste de espaço solicitado
            }
            df_filtrado = df_filtrado.rename(columns=dicionario_renomear)

            # 4. DEFINIR A ORDEM EXATA DAS COLUNAS SOLICITADA
            ordem_solicitada = [
                "Nº Mapa",
                "Data",
                "Veículo",
                "Placa",
                "Flash Point",
                "Descrição",
                "Valor",
            ]

            # Mantém apenas as colunas que realmente existem na tabela (evita erros caso falte alguma)
            colunas_finais = [
                col for col in ordem_solicitada if col in df_filtrado.columns
            ]
            df_filtrado = df_filtrado[colunas_finais].copy()

            # 5. Organizar a ordenação interna por: "Flash Point" (Bloco) e depois "Nº Mapa"
            colunas_ordenacao = []
            if "Flash Point" in df_filtrado.columns:
                colunas_ordenacao.append("Flash Point")
            if "Nº Mapa" in df_filtrado.columns:
                colunas_ordenacao.append("Nº Mapa")

            if colunas_ordenacao:
                df_filtrado = df_filtrado.sort_values(
                    by=colunas_ordenacao, ascending=True
                )

            # --- CRIAÇÃO DOS BLOCOS, REMOÇÃO DE REPETIDAS E MAPEAMENTO DE CORES ---
            if not df_filtrado.empty and "Flash Point" in df_filtrado.columns:
                lista_linhas = []
                linhas_amarelas = []  # Para armazenar índices de placas repetidas
                linhas_laranjas = []  # Para armazenar índices das linhas de Total
                contador_linha_excel = 2  # Dados começam na linha 2 do Excel

                # Agrupamos por Flash Point (Bloco do Cliente)
                for fp, bloco in df_filtrado.groupby("Flash Point", sort=False):

                    placas_vistas = set()

                    for idx, linha in bloco.iterrows():
                        linha_dict = linha.to_dict()
                        placa_atual = str(linha.get("Placa", "")).strip()

                        # SE A PLACA JÁ FOI VISTA NESTE CLIENTE:
                        if placa_atual in placas_vistas and placa_atual != "":
                            linha_dict["Valor"] = (
                                None  # Zera o valor para cobrar apenas uma vez
                            )
                            linhas_amarelas.append(contador_linha_excel)
                        else:
                            if placa_atual != "":
                                placas_vistas.add(placa_atual)

                        lista_linhas.append(linha_dict)
                        contador_linha_excel += 1

                    # --- GERAÇÃO DA LINHA DE TOTAL DO BLOCO ---
                    df_bloco_temp = pd.DataFrame(lista_linhas[-len(bloco) :])
                    soma_bloco = pd.to_numeric(
                        df_bloco_temp["Valor"], errors="coerce"
                    ).sum()

                    linha_total = {col: "" for col in colunas_finais}
                    linha_total["Flash Point"] = fp
                    linha_total["Descrição"] = "VALOR TOTAL:"
                    linha_total["Valor"] = (
                        float(soma_bloco) if soma_bloco > 0 else ""
                    )

                    lista_linhas.append(linha_total)
                    linhas_laranjas.append(
                        contador_linha_excel
                    )  # Guarda linha para pintar de Laranja
                    contador_linha_excel += 1

                    # Linha em branco de espaçamento
                    linha_espacamento = {col: "" for col in colunas_finais}
                    lista_linhas.append(linha_espacamento)
                    contador_linha_excel += 1

                if lista_linhas:
                    lista_linhas.pop()  # Remove o último respiro sobressalente

                df_excel_final = pd.DataFrame(
                    lista_linhas, columns=colunas_finais
                )
            else:
                df_excel_final = df_filtrado.copy()
                linhas_amarelas = []
                linhas_laranjas = []

            # --- EXIBIÇÃO E DOWNLOAD ---
            if not df_filtrado.empty:
                st.subheader("📋 Visualização Prévia dos Dados")
                st.dataframe(df_filtrado)

                # --- CONSTRUÇÃO DO EXCEL COM CORES ---
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    df_excel_final.to_excel(
                        writer, index=False, sheet_name="FPF Realizados"
                    )

                    workbook = writer.book
                    worksheet = writer.sheets["FPF Realizados"]

                    # Paletas de cores (Tons claros preservando fonte escura)
                    amarelo_claro = PatternFill(
                        start_color="FFFFCC",
                        end_color="FFFFCC",
                        fill_type="solid",
                    )  # Placas repetidas
                    laranja_claro = PatternFill(
                        start_color="FFE6CC",
                        end_color="FFE6CC",
                        fill_type="solid",
                    )  # Valor Total

                    # 1. Aplica Amarelo nas linhas repetidas
                    for num_linha in linhas_amarelas:
                        for col_idx in range(1, len(colunas_finais) + 1):
                            worksheet.cell(
                                row=num_linha, column=col_idx
                            ).fill = amarelo_claro

                    # 2. Aplica Laranja nas linhas de Valor Total
                    for num_linha in linhas_laranjas:
                        for col_idx in range(1, len(colunas_finais) + 1):
                            worksheet.cell(
                                row=num_linha, column=col_idx
                            ).fill = laranja_claro

                st.success(
                    "Planilha processada! Colunas renomeadas, ordenadas e totais destacados em Laranja Claro."
                )

                st.download_button(
                    label="📥 Baixar Planilha Reestruturada (Excel)",
                    data=buffer.getvalue(),
                    file_name="FPF_Relatorio_Final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.warning(
                    "O resultado do filtro gerou uma tabela vazia (nenhum 'MOD' encontrado)."
                )

    except Exception as e:
        st.error(f"Ocorreu um erro crítico ao processar o arquivo: {e}")
else:
    st.info("Aguardando o upload da planilha para iniciar o processamento...")