
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle
)

# ==================================================
# 1. CONFIGURAÇÕES
# ==================================================

arquivo_excel = "PLANILHA DE ATESTADOS DO ANO - SMED 2.xlsx"
planilha = "20252026"

coluna_matricula = "MATRICULA"
coluna_nome = "NOME DO SERVIDOR"
coluna_dias = "QUANT. DIAS DE AFASTAMENTO"

# ==================================================
# 2. LEITURA E TRATAMENTO DOS DADOS
# ==================================================

df = pd.read_excel(
    arquivo_excel,
    sheet_name=planilha,
    header=5
)

# Limpar espaços dos nomes das colunas
df.columns = df.columns.astype(str).str.strip()

# Verificar se as colunas necessárias existem
colunas_necessarias = [
    coluna_matricula,
    coluna_nome,
    coluna_dias
]

faltantes = [
    coluna for coluna in colunas_necessarias
    if coluna not in df.columns
]

if faltantes:
    raise ValueError(
        f"Colunas não encontradas: {faltantes}. "
        f"Colunas disponíveis: {df.columns.tolist()}"
    )

# Converter dias para números
df[coluna_dias] = pd.to_numeric(
    df[coluna_dias],
    errors="coerce"
)

# Remover linhas sem matrícula
df = df.dropna(subset=[coluna_matricula]).copy()

# Padronizar matrícula para evitar diferenças de formato
df[coluna_matricula] = (
    df[coluna_matricula]
    .astype(str)
    .str.strip()
    .str.replace(r"\.0$", "", regex=True)
)

# Excluir matrículas vazias ou inválidas
df = df[
    df[coluna_matricula].ne("")
    & df[coluna_matricula].ne("nan")
].copy()

# ==================================================
# 3. RESUMO DE TODOS OS ATESTADOS POR SERVIDOR
# ==================================================

resumo = (
    df.groupby(coluna_matricula)
    .agg(
        NOME_DO_SERVIDOR=(
            coluna_nome,
            "first"
        ),
        QUANTIDADE_ATESTADOS=(
            coluna_matricula,
            "size"
        ),
        TOTAL_DIAS=(
            coluna_dias,
            "sum"
        ),
        MEDIA_DIAS=(
            coluna_dias,
            "mean"
        )
    )
    .reset_index()
)

resumo = resumo.sort_values(
    by="TOTAL_DIAS",
    ascending=False
)

resumo["MEDIA_DIAS"] = resumo["MEDIA_DIAS"].round(2)

resumo.to_excel(
    "resumo_atestados_por_servidor.xlsx",
    index=False
)

# ==================================================
# 4. RESUMO DOS ATESTADOS COM 11 DIAS OU MAIS
# ==================================================

filtrados = df[
    df[coluna_dias] >= 11
].copy()

resumo_11_dias = (
    filtrados.groupby(coluna_matricula)
    .agg(
        NOME_DO_SERVIDOR=(
            coluna_nome,
            "first"
        ),
        QUANTIDADE_ATESTADOS=(
            coluna_matricula,
            "size"
        ),
        TOTAL_DIAS=(
            coluna_dias,
            "sum"
        ),
        MEDIA_DIAS=(
            coluna_dias,
            "mean"
        )
    )
    .reset_index()
    .sort_values(
        by="TOTAL_DIAS",
        ascending=False
    )
)

resumo_11_dias["MEDIA_DIAS"] = (
    resumo_11_dias["MEDIA_DIAS"].round(2)
)

resumo_11_dias.to_excel(
    "resumo_atestados_11_dias_ou_mais.xlsx",
    index=False
)

# ==================================================
# 5. INDICADORES GERAIS
# ==================================================

total_registros = len(df)
total_servidores = df[coluna_matricula].nunique()
total_dias = df[coluna_dias].sum()

total_atestados_11 = len(filtrados)
servidores_11 = filtrados[coluna_matricula].nunique()
total_dias_11 = filtrados[coluna_dias].sum()

print("\n===== TODOS OS ATESTADOS =====")
print(f"Total de registros: {total_registros}")
print(f"Servidores com registros: {total_servidores}")
print(f"Soma dos dias de afastamento: {total_dias:.0f}")

print("\n===== ATESTADOS COM 11 DIAS OU MAIS =====")
print(f"Total de registros: {total_atestados_11}")
print(f"Servidores com esses registros: {servidores_11}")
print(f"Soma dos dias de afastamento: {total_dias_11:.0f}")

print("\n===== RESUMO POR SERVIDOR =====")
print(resumo.to_string(index=False))

# ==================================================
# 6. GERAR RELATÓRIO PDF
# ==================================================

arquivo_pdf = "relatorio_atestados_por_servidor.pdf"

documento = SimpleDocTemplate(
    arquivo_pdf,
    pagesize=landscape(A4),
    rightMargin=1.2 * cm,
    leftMargin=1.2 * cm,
    topMargin=1.2 * cm,
    bottomMargin=1.2 * cm
)

estilos = getSampleStyleSheet()

titulo = ParagraphStyle(
    "TituloRelatorio",
    parent=estilos["Title"],
    alignment=TA_CENTER,
    fontSize=16,
    spaceAfter=14
)

elementos = []

elementos.append(
    Paragraph(
        "RELATÓRIO DE ATESTADOS POR SERVIDOR",
        titulo
    )
)

elementos.append(
    Paragraph(
        "Resumo geral de afastamentos",
        estilos["Heading2"]
    )
)

elementos.append(
    Paragraph(
        f"Total de registros: {total_registros}",
        estilos["Normal"]
    )
)

elementos.append(
    Paragraph(
        f"Servidores com registros: {total_servidores}",
        estilos["Normal"]
    )
)

elementos.append(
    Paragraph(
        f"Total de dias de afastamento: {total_dias:.0f}",
        estilos["Normal"]
    )
)

elementos.append(Spacer(1, 10))

elementos.append(
    Paragraph(
        "Afastamentos iguais ou superiores a 11 dias",
        estilos["Heading2"]
    )
)

elementos.append(
    Paragraph(
        f"Registros: {total_atestados_11} | "
        f"Servidores: {servidores_11} | "
        f"Total de dias: {total_dias_11:.0f}",
        estilos["Normal"]
    )
)

elementos.append(Spacer(1, 12))

elementos.append(
    Paragraph(
        "Detalhamento por servidor — todos os atestados",
        estilos["Heading2"]
    )
)

# Preparar tabela do PDF
dados = [[
    "Matrícula",
    "Nome do servidor",
    "Qtd. atestados",
    "Total de dias",
    "Média de dias"
]]

for _, linha in resumo.iterrows():
    dados.append([
        str(linha[coluna_matricula]),
        str(linha["NOME_DO_SERVIDOR"]),
        str(int(linha["QUANTIDADE_ATESTADOS"])),
        (
            ""
            if pd.isna(linha["TOTAL_DIAS"])
            else f'{linha["TOTAL_DIAS"]:.0f}'
        ),
        (
            ""
            if pd.isna(linha["MEDIA_DIAS"])
            else f'{linha["MEDIA_DIAS"]:.2f}'
        )
    ])

tabela = Table(
    dados,
    colWidths=[
        3.0 * cm,
        10.0 * cm,
        3.0 * cm,
        3.0 * cm,
        3.0 * cm
    ],
    repeatRows=1
)

tabela.setStyle(TableStyle([
    (
        "BACKGROUND",
        (0, 0),
        (-1, 0),
        colors.HexColor("#24476B")
    ),
    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
    ("FONTSIZE", (0, 0), (-1, -1), 8),
    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ("ALIGN", (1, 1), (1, -1), "LEFT"),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
    (
        "ROWBACKGROUNDS",
        (0, 1),
        (-1, -1),
        [colors.white, colors.HexColor("#EAF0F6")]
    ),
    ("TOPPADDING", (0, 0), (-1, -1), 5),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 5)
]))

elementos.append(tabela)

documento.build(elementos)

print("\nArquivos gerados com sucesso:")
print("- resumo_atestados_por_servidor.xlsx")
print("- resumo_atestados_11_dias_ou_mais.xlsx")
print("- relatorio_atestados_por_servidor.pdf")
