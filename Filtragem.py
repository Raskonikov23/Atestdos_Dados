
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

# 1. Arquivo de entrada e saída
arquivo_excel = "PLANILHA DE ATESTADOS DO ANO - SMED 2.xlsx"
arquivo_pdf = "relatorio_atestados.pdf"

# 2. Leitura da planilha
df = pd.read_excel(
    arquivo_excel,
    sheet_name="20252026",
    header=5
)

# 3. Tratamento da coluna de dias
coluna_dias = "QUANT. DIAS DE AFASTAMENTO"

df[coluna_dias] = pd.to_numeric(
    df[coluna_dias],
    errors="coerce"
)

# 4. Filtrar afastamentos de 11 dias ou mais
filtrados = df[
    df[coluna_dias] >= 11
].copy()

# 5. Calcular indicadores
total_registros = len(filtrados)
total_dias = filtrados[coluna_dias].sum()

if total_registros > 0:
    media_dias = filtrados[coluna_dias].mean()
else:
    media_dias = 0

# 6. Configurar o documento PDF
documento = SimpleDocTemplate(
    arquivo_pdf,
    pagesize=landscape(A4),
    rightMargin=1.5 * cm,
    leftMargin=1.5 * cm,
    topMargin=1.5 * cm,
    bottomMargin=1.5 * cm
)

estilos = getSampleStyleSheet()

titulo = ParagraphStyle(
    "Titulo",
    parent=estilos["Title"],
    alignment=TA_CENTER,
    fontSize=16,
    spaceAfter=15
)

subtitulo = ParagraphStyle(
    "Subtitulo",
    parent=estilos["Heading2"],
    fontSize=12,
    spaceBefore=10,
    spaceAfter=8
)

elementos = []

# 7. Título e resumo
elementos.append(
    Paragraph(
        "RELATÓRIO DE ATESTADOS MÉDICOS",
        titulo
    )
)

elementos.append(
    Paragraph(
        "Afastamentos iguais ou superiores a 11 dias",
        subtitulo
    )
)

elementos.append(Paragraph(
    f"<b>Total de registros:</b> {total_registros}",
    estilos["Normal"]
))

elementos.append(Paragraph(
    f"<b>Total de dias de afastamento:</b> "
    f"{total_dias:.0f}",
    estilos["Normal"]
))

elementos.append(Paragraph(
    f"<b>Média de dias por registro:</b> "
    f"{media_dias:.2f}",
    estilos["Normal"]
))

elementos.append(Spacer(1, 15))

# 8. Preparar a tabela
colunas = [
    "MATRICULA",
    "NOME DO SERVIDOR",
    "DATA AFASTAMENTO",
    "DATA RETORNO",
    coluna_dias
]

relatorio = filtrados[colunas].copy()

# Formatar datas, quando existirem
for coluna in ["DATA AFASTAMENTO", "DATA RETORNO"]:
    relatorio[coluna] = pd.to_datetime(
        relatorio[coluna],
        errors="coerce"
    ).dt.strftime("%d/%m/%Y")

relatorio = relatorio.fillna("")

# Evitar que a matrícula seja apresentada como número decimal
relatorio["MATRICULA"] = (
    relatorio["MATRICULA"].astype(str)
    .str.replace(r"\.0$", "", regex=True)
)

dados = [[
    "Matrícula",
    "Nome do servidor",
    "Data de afastamento",
    "Data de retorno",
    "Dias"
]]

for _, linha in relatorio.iterrows():
    dados.append([
        str(linha["MATRICULA"]),
        str(linha["NOME DO SERVIDOR"]),
        str(linha["DATA AFASTAMENTO"]),
        str(linha["DATA RETORNO"]),
        str(int(linha[coluna_dias]))
    ])

tabela = Table(
    dados,
    colWidths=[2.5*cm, 8*cm, 3.5*cm, 3.5*cm, 1.5*cm],
    repeatRows=1,
    hAlign="LEFT"
)

tabela.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#24476B")),
    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
    ("FONTSIZE", (0, 0), (-1, -1), 8),
    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ("ALIGN", (1, 1), (1, -1), "LEFT"),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
    ("ROWBACKGROUNDS", (0, 1), (-1, -1),
     [colors.white, colors.HexColor("#EAF0F6")]),
    ("TOPPADDING", (0, 0), (-1, -1), 6),
    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
]))

elementos.append(tabela)

# 9. Gerar o PDF
documento.build(elementos)

print("Relatório gerado:", arquivo_pdf)
print("Total de registros:", total_registros)
print("Total de dias de afastamento:", total_dias)