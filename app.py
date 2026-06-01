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
    "Filtragem, cálculo inteligente de valores, remoção de cobranças duplicadas por Matrícula e destaque visual."
)

# 1. Componente para carregar o arquivo
arquivo_carregado = st.file_uploader(
    "Arraste ou selecione a planilha (Padrão: FPF List)",
    type=["xlsx", "xls", "csv"],
)


# Função para definir o valor inicial com base no Nome do Arquivo e Fabricante
def calcular_valor_inicial(linha):
    nome_arquivo = str(linha.get("Nome arquivo", "")).upper().strip()
    fabricante = str(linha.get("Fabricante", "")).upper().strip()

    # Remove múltiplos espaços internos
    nome_arquivo = re.sub(r"\s+", " ", nome_arquivo)
    fabricante = re.sub(r"\s+", " ", fabricante)

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

    eh_especial = any(fab in fabricante for fab in fabricantes_especiais)

    if "VOLVO TRUCK" in fabricante:
        eh_especial = False

    # 1. Regra para STG2
    if any(termo in nome_arquivo for termo in termos_stg2):
        return 1400 if eh_especial else 650

    # 2. Regra para MOD / OFF
    elif any(termo in nome_arquivo for termo in termos_mod_off):
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

            # 3. Selecionar colunas desejadas
            colunas_desejadas = [
                "Arquivo ID",
                "Fabricante",
                "Matrícula",
                "FlashPoint",
                "Cliente",
                "Nome arquivo",
                "Dada",
            ]

            colunas_existentes = [
                col for col in colunas_desejadas if col in df_filtrado.columns
            ]
            colunas_faltantes = list(
                set(colunas_desejadas) - set(colunas_existentes)
            )

            if colunas_faltantes:
                st.error(
                    f"Atenção! Colunas obrigatórias ausentes: {colunas_faltantes}"
                )

            df_filtrado = df_filtrado[colunas_existentes].copy()

            # --- GERAÇÃO DA COLUNA VALOR AUTOMÁTICA ---
            df_filtrado["Valor"] = df_filtrado.apply(
                calcular_valor_inicial, axis=1
            )

            todas_colunas_finais = colunas_existentes + ["Valor"]

            # 4. Organizar em ordem: Primeiro por "FlashPoint" (Bloco do Cliente), depois "Arquivo ID"
            colunas_ordenacao = []
            if "FlashPoint" in df_filtrado.columns:
                colunas_ordenacao.append("FlashPoint")
            if "Arquivo ID" in df_filtrado.columns:
                colunas_ordenacao.append("Arquivo ID")

            if colunas_ordenacao:
                df_filtrado = df_filtrado.sort_values(
                    by=colunas_ordenacao, ascending=True
                )

            # --- CRIAÇÃO DOS BLOCOS, REMOÇÃO DE REPETIDAS E FORMATAÇÃO DO EXCEL ---
            if not df_filtrado.empty and "FlashPoint" in df_filtrado.columns:
                lista_linhas = []
                linhas_para_colorir = (
                    []
                )  # Guardará os índices das linhas duplicadas para pintar depois
                contador_linha_excel = (
                    2  # O Excel começa na linha 1 (cabeçalho), dados começam na 2
                )

                # Agrupamos por FlashPoint (Cliente)
                for fp, bloco in df_filtrado.groupby("FlashPoint", sort=False):

                    matriculas_vistas = (
                        set()
                    )  # Controla as matrículas já cobradas DESTE cliente

                    for idx, linha in bloco.iterrows():
                        linha_dict = linha.to_dict()
                        # Limpa o texto da matrícula para evitar erros com espaços
                        matricula_atual = str(linha.get("Matrícula", "")).strip()

                        # SE A MATRÍCULA JÁ FOI VISTA NESTE CLIENTE:
                        if (
                            matricula_atual in matriculas_vistas
                            and matricula_atual != ""
                        ):
                            linha_dict["Valor"] = (
                                None  # Cobra apenas uma vez (zera as repetições)
                            )
                            linhas_para_colorir.append(
                                contador_linha_excel
                            )  # Marca para pintar de amarelo

                        else:
                            # Se for a primeira vez que vê a matrícula, adiciona ao controle de cobrados
                            if matricula_atual != "":
                                matriculas_vistas.add(matricula_atual)

                        lista_linhas.append(linha_dict)
                        contador_linha_excel += 1

                    # --- GERAÇÃO DA LINHA DE TOTAL DO BLOCO ---
                    # Para somar corretamente, transformamos a lista temporária do bloco atual em DataFrame
                    df_bloco_temp = pd.DataFrame(
                        lista_linhas[-len(bloco) :]
                    )  # Pega as últimas N linhas inseridas
                    soma_bloco = pd.to_numeric(
                        df_bloco_temp["Valor"], errors="coerce"
                    ).sum()

                    linha_total = {col: "" for col in todas_colunas_finais}
                    linha_total["FlashPoint"] = fp
                    linha_total["Nome arquivo"] = "VALOR TOTAL:"
                    linha_total["Valor"] = (
                        float(soma_bloco) if soma_bloco > 0 else ""
                    )

                    lista_linhas.append(linha_total)
                    contador_linha_excel += 1  # Conta a linha do Valor Total

                    # Linha em branco de espaçamento
                    linha_espacamento = {
                        col: "" for col in todas_colunas_finais
                    }
                    lista_linhas.append(linha_espacamento)
                    contador_linha_excel += 1  # Conta a linha em branco

                # Remove o último espaçamento sobressalente
                if lista_linhas:
                    lista_linhas.pop()

                df_excel_final = pd.DataFrame(
                    lista_linhas, columns=todas_colunas_finais
                )
            else:
                df_excel_final = df_filtrado.copy()
                linhas_para_colorir = []

            # --- EXIBIÇÃO E DOWNLOAD ---
            if not df_filtrado.empty:
                st.subheader("📋 Visualização Prévia dos Dados")
                st.dataframe(df_filtrado)

                # --- CONSTRUÇÃO DO EXCEL E APLICAÇÃO DA COR ---
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    df_excel_final.to_excel(
                        writer, index=False, sheet_name="FPF Realizados"
                    )

                    # Acessa a estrutura nativa do openpyxl para colorir
                    workbook = writer.book
                    worksheet = writer.sheets["FPF Realizados"]

                    # Define a cor Amarelo Claro (Hex: #FFF2CC ou #FFFFE0)
                    amarelo_claro = PatternFill(
                        start_color="FFFFCC",
                        end_color="FFFFCC",
                        fill_type="solid",
                    )

                    # Colore todas as células das linhas que foram marcadas como repetidas
                    for num_linha in linhas_para_colorir:
                        for col_idx in range(1, len(todas_colunas_finais) + 1):
                            worksheet.cell(
                                row=num_linha, column=col_idx
                            ).fill = amarelo_claro

                st.success(
                    "Valores calculados! Matrículas repetidas foram zeradas e destacadas em amarelo claro."
                )

                st.download_button(
                    label="📥 Baixar Planilha Final com Destaques (Excel)",
                    data=buffer.getvalue(),
                    file_name="FPF_Final_Formatado.xlsx",
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