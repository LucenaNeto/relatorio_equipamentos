# -*- coding: utf-8 -*-
"""
Módulo: reports_pdf.py
Gera relatórios PDF por loja a partir do dicionário { loja: DataFrame }.

Requisitos:
    pip install reportlab

Uso (exemplo rápido):
    from src.processing import carregar_e_validar
    from src.pricing import aplicar_preco_sugerido
    from src.reports_pdf import gerar_relatorios_pdf

    df_ok, df_err, grupos = carregar_e_validar("relatorio_equipamentos/input/cadastro_equipamentos.xlsx")
    # preenche preço sugerido onde estiver vazio
    grupos_aj = {lj: aplicar_preco_sugerido(df, base_dir="relatorio_equipamentos")
                 for lj, df in grupos.items()}
    pdfs = gerar_relatorios_pdf(grupos_aj, output_dir="relatorio_equipamentos/output/pdf")
    print(pdfs)
"""

import os
from datetime import datetime
import math
import pandas as pd
import numpy as np

# ReportLab
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
)


# ===== Utilidades =====

def _safe_filename(name: str) -> str:
    """Sanitiza o nome do arquivo para evitar caracteres inválidos."""
    import re
    name = str(name).strip()
    name = re.sub(r"[^\w\-]+", "_", name, flags=re.UNICODE)
    return name[:120] or "sem_nome"


def _brl(v) -> str:
    """Formata número como moeda brasileira (R$ 1.234,56); NaN -> '—'."""
    if v is None or (isinstance(v, float) and (np.isnan(v) or np.isinf(v))):
        return "—"
    try:
        s = f"{float(v):,.2f}"
        # troca separadores para padrão brasileiro
        s = s.replace(",", "X").replace(".", ",").replace("X", ".")
        return f"R$ {s}"
    except Exception:
        return "—"


def _pcent(v) -> str:
    """Formata percentual com 2 casas; NaN -> '—'."""
    if v is None or (isinstance(v, float) and (np.isnan(v) or np.isinf(v))):
        return "—"
    try:
        return f"{float(v)*100:.2f}%"
    except Exception:
        return "—"


# ===== Cálculos por loja =====

def _preparar_itens(df_loja: pd.DataFrame) -> pd.DataFrame:
    """
    Gera colunas calculadas para a tabela Itens:
      - Total sugerido = Qtd * Preço sugerido
      - Total real     = Qtd * (Preço real, ou Preço sugerido quando vazio)
      - Diferença (R$) e Diferença (%)
    """
    df = df_loja.copy()

    # Garantir tipos
    df["Quantidade"] = pd.to_numeric(df["Quantidade"], errors="coerce").fillna(0).astype(int).clip(lower=0)
    df["Preço sugerido"] = pd.to_numeric(df["Preço sugerido"], errors="coerce")
    df["Preço real"]     = pd.to_numeric(df["Preço real"], errors="coerce")

    # Negativos viram NaN (não faz sentido)
    for c in ["Preço sugerido", "Preço real"]:
        df[c] = df[c].where(df[c].isna() | (df[c] >= 0), np.nan)

    ps = df["Preço sugerido"].fillna(0)
    pr_eff = df["Preço real"].where(df["Preço real"].notna(), ps)

    df["Total sugerido"] = df["Quantidade"] * ps
    df["Total real"]     = df["Quantidade"] * pr_eff
    df["Diferença (R$)"] = df["Total real"] - df["Total sugerido"]
    df["Diferença (%)"]  = np.where(df["Total sugerido"] > 0,
                                    (df["Total real"] / df["Total sugerido"]) - 1.0,
                                    np.nan)

    cols = [
        "Equipamento", "Quantidade", "Preço sugerido", "Preço real",
        "Total sugerido", "Total real", "Diferença (R$)", "Diferença (%)"
    ]
    return df[cols].copy()


def _resumo_da_loja(df_itens: pd.DataFrame) -> dict:
    """Calcula agregados para a seção Resumo da loja."""
    total_sugerido = df_itens["Total sugerido"].sum()
    total_real     = df_itens["Total real"].sum()
    return {
        "Itens": int(len(df_itens)),
        "Qtd total": int(df_itens["Quantidade"].sum()),
        "Total sugerido": float(total_sugerido),
        "Total real": float(total_real),
        "Diferença (R$)": float(total_real - total_sugerido),
        "Diferença (%)": (float(total_real / total_sugerido - 1.0) if total_sugerido > 0 else np.nan),
    }


# ===== Renderização do PDF =====

def _tabela_itens(df_itens: pd.DataFrame, styles, largura_util_mm=180, linhas_por_tabela=30):
    """
    Constrói uma lista de elementos (Tables + PageBreak) para os itens,
    quebrando em múltiplas tabelas quando houver muitas linhas.
    """
    elementos = []

    # Cabeçalho da tabela
    header = [
        "Equipamento", "Qtd", "Preço sugerido", "Preço real",
        "Total sugerido", "Total real", "Diferença (R$)", "Diferença (%)"
    ]

    # Converte DataFrame em lista de listas formatada
    dados = []
    for _, r in df_itens.iterrows():
        dados.append([
            str(r["Equipamento"]),
            int(r["Quantidade"]),
            _brl(r["Preço sugerido"]),
            _brl(r["Preço real"]),
            _brl(r["Total sugerido"]),
            _brl(r["Total real"]),
            _brl(r["Diferença (R$)"]),
            _pcent(r["Diferença (%)"]),
        ])

    # Quebra em blocos
    total = len(dados)
    blocos = math.ceil(total / linhas_por_tabela) if total else 1
    largura_util = largura_util_mm * mm

    # Larguras relativas das colunas (somatório = 1.0), depois multiplicamos pela largura útil
    pesos = [0.30, 0.07, 0.12, 0.12, 0.12, 0.12, 0.075, 0.065]
    col_widths = [largura_util * p for p in pesos]

    for i in range(blocos):
        inicio = i * linhas_por_tabela
        fim = inicio + linhas_por_tabela
        bloco_dados = [header] + dados[inicio:fim]

        t = Table(bloco_dados, colWidths=col_widths, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
            ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
            ("ALIGN", (0, 0), (0, -1), "LEFT"),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 9),
            ("FONTSIZE", (0, 1), (-1, -1), 8),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CCCCCC")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#FAFAFA")]),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
        ]))
        elementos.append(t)

        # Quebra de página entre blocos (menos no último)
        if i < blocos - 1:
            elementos.append(Spacer(1, 6 * mm))
            elementos.append(PageBreak())

    return elementos


def _tabela_resumo(resumo: dict, largura_util_mm=100):
    """Cria a pequena tabela de resumo com métricas agregadas."""
    header = ["Métrica", "Valor"]
    dados = [
        ["Itens", resumo["Itens"]],
        ["Qtd total", resumo["Qtd total"]],
        ["Total sugerido", _brl(resumo["Total sugerido"])],
        ["Total real", _brl(resumo["Total real"])],
        ["Diferença (R$)", _brl(resumo["Diferença (R$)"])],
        ["Diferença (%)", _pcent(resumo["Diferença (%)"])],
    ]
    t = Table([header] + dados, colWidths=[(largura_util_mm*0.5)*mm, (largura_util_mm*0.5)*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CCCCCC")),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
    ]))
    return t


def _rodape(canvas, doc):
    """Desenha número de página no rodapé."""
    page_num = canvas.getPageNumber()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.HexColor("#555555"))
    canvas.drawRightString(doc.pagesize[0] - 15*mm, 10*mm, f"Página {page_num}")


def _gerar_pdf_loja(loja: str, df_loja: pd.DataFrame, output_dir: str,
                    titulo: str = "Relatório de Equipamentos por Loja") -> str:
    """Gera um PDF para a loja informada."""
    os.makedirs(output_dir, exist_ok=True)
    fname = f"relatorio_loja_{_safe_filename(loja)}.pdf"
    path = os.path.join(output_dir, fname)

    # Prepara dados
    df_itens = _preparar_itens(df_loja)
    resumo = _resumo_da_loja(df_itens)

    # Documento
    doc = SimpleDocTemplate(
        path,
        pagesize=A4,
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=14*mm, bottomMargin=14*mm
    )
    styles = getSampleStyleSheet()
    estilo_titulo = ParagraphStyle(
        "Titulo",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=14,
        spaceAfter=6,
    )
    estilo_meta = ParagraphStyle(
        "Meta",
        parent=styles["Normal"],
        fontSize=9,
        textColor=colors.HexColor("#555555"),
        spaceAfter=4,
    )

    story = []
    # Cabeçalho
    story.append(Paragraph(titulo, estilo_titulo))
    story.append(Paragraph(f"<b>Loja:</b> {loja}", estilo_meta))
    story.append(Paragraph(f"<b>Gerado em:</b> {datetime.now().strftime('%Y-%m-%d %H:%M')}", estilo_meta))
    story.append(Spacer(1, 4*mm))

    # Tabela de Itens (pode quebrar em múltiplas páginas)
    story.extend(_tabela_itens(df_itens, styles, largura_util_mm=180, linhas_por_tabela=30))
    story.append(Spacer(1, 6*mm))

    # Resumo
    story.append(Paragraph("<b>Resumo</b>", styles["Heading3"]))
    story.append(Spacer(1, 2*mm))
    story.append(_tabela_resumo(resumo, largura_util_mm=120))

    # Build
    doc.build(story, onFirstPage=_rodape, onLaterPages=_rodape)
    return path


# ===== Função pública =====

def gerar_relatorios_pdf(grupos_por_loja: dict[str, pd.DataFrame], output_dir: str,
                         titulo_relatorio: str = "Relatório de Equipamentos por Loja") -> list[str]:
    """
    Gera PDFs para todas as lojas do dicionário {loja: DataFrame}.
    Retorna lista com os caminhos dos arquivos gerados.
    """
    os.makedirs(output_dir, exist_ok=True)
    arquivos = []
    for loja, df_loja in grupos_por_loja.items():
        if df_loja is None or df_loja.empty:
            continue
        path = _gerar_pdf_loja(loja, df_loja, output_dir, titulo=titulo_relatorio)
        arquivos.append(path)
    return arquivos
