# -*- coding: utf-8 -*-
"""
Módulo: reports_pdf.py (versão com layout melhorado + área de logo)
- Colunas com largura balanceada e quebra de linha (Paragraph).
- Cabeçalho com espaço para logo (usa config/logo.png se existir).
- Tabelas longas quebram em múltiplas páginas com cabeçalho repetido.
"""

import os
import math
from datetime import datetime
from pathlib import Path

import pandas as pd
import numpy as np

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_RIGHT, TA_LEFT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
)



# ===== Utilidades =====

def _safe_filename(name: str) -> str:
    import re
    name = str(name).strip()
    name = re.sub(r"[^\w\-]+", "_", name, flags=re.UNICODE)
    return name[:120] or "sem_nome"


def _brl(v) -> str:
    if v is None or (isinstance(v, float) and (np.isnan(v) or np.isinf(v))):
        return "—"
    try:
        s = f"{float(v):,.2f}"
        s = s.replace(",", "X").replace(".", ",").replace("X", ".")
        return f"R$ {s}"
    except Exception:
        return "—"


def _pcent(v) -> str:
    if v is None or (isinstance(v, float) and (np.isnan(v) or np.isinf(v))):
        return "—"
    try:
        return f"{float(v)*100:.2f}%"
    except Exception:
        return "—"


# ===== Estilos =====

def _build_styles():
    styles = getSampleStyleSheet()
    # Título e metadados
    estilo_titulo = ParagraphStyle(
        "Titulo",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=14,
        spaceAfter=4,
    )
    estilo_meta = ParagraphStyle(
        "Meta",
        parent=styles["Normal"],
        fontSize=9,
        textColor=colors.HexColor("#555555"),
        spaceAfter=2,
    )
    # Texto da coluna "Equipamento": quebra palavras longas
    estilo_equip = ParagraphStyle(
        "Equip",
        parent=styles["Normal"],
        fontSize=8,
        leading=10,
        wordWrap="CJK",     # permite quebra mesmo em palavras compridas
        alignment=TA_LEFT,
    )
    # Números alinhados à direita
    estilo_num = ParagraphStyle(
        "Num",
        parent=styles["Normal"],
        fontSize=8,
        leading=10,
        alignment=TA_RIGHT,
    )
    # Cabeçalho de seção
    estilo_h3 = ParagraphStyle(
        "H3",
        parent=styles["Heading3"],
        fontName="Helvetica-Bold",
        fontSize=11,
        spaceBefore=6,
        spaceAfter=2,
    )
    return estilo_titulo, estilo_meta, estilo_equip, estilo_num, estilo_h3


# ===== Cabeçalho com logo =====

def _header_block(loja: str, output_dir: str, titulo: str, estilo_titulo, estilo_meta,
                  logo_path: str | None = None, logo_max_w_mm: float = 35.0, logo_max_h_mm: float = 18.0):
    """
    Retorna um Table com 2 colunas: à esquerda título+metas, à direita logo (se houver).
    Se logo não existir, exibe um placeholder discreto.
    """
    left_cells = [
        Paragraph(titulo, estilo_titulo),
        Paragraph(f"<b>Loja:</b> {loja}", estilo_meta),
        Paragraph(f"<b>Gerado em:</b> {datetime.now().strftime('%Y-%m-%d %H:%M')}", estilo_meta),
    ]

    # Descobrir logo padrão em config/ se não veio por parâmetro
    logo_flowable = None
    if logo_path is None:
        base = Path(output_dir).resolve().parent  # .../relatorio_equipamentos
        cand = base / "config" / "logo.png"
        if cand.exists():
            logo_path = str(cand)

    if logo_path and os.path.exists(logo_path):
        try:
            img = Image(logo_path)
            img._restrictSize(logo_max_w_mm * mm, logo_max_h_mm * mm)
            logo_flowable = img
        except Exception:
            logo_flowable = None

    # Se não tiver logo, cria um placeholder
    if logo_flowable is None:
        ph = Table([["Sua logo aqui"]], colWidths=[logo_max_w_mm * mm], rowHeights=[logo_max_h_mm * mm])
        ph.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#CCCCCC")),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor("#888888")),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
        ]))
        logo_flowable = ph

    # Monta header em 2 colunas
    header_table = Table(
        [[left_cells, logo_flowable]],
        colWidths=[(180 - logo_max_w_mm - 5) * mm, (logo_max_w_mm) * mm],
    )
    header_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    return header_table


# ===== Cálculos por loja =====

def _preparar_itens(df_loja: pd.DataFrame) -> pd.DataFrame:
    df = df_loja.copy()
    df["Quantidade"] = pd.to_numeric(df["Quantidade"], errors="coerce").fillna(0).astype(int).clip(lower=0)
    df["Preço sugerido"] = pd.to_numeric(df["Preço sugerido"], errors="coerce")
    df["Preço real"]     = pd.to_numeric(df["Preço real"], errors="coerce")
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

def _tabela_itens(df_itens: pd.DataFrame, estilos, largura_util_mm=180, linhas_por_tabela=28):
    """
    Constrói Tables para os itens, com quebra em blocos.
    Usa Paragraph para forçar quebra de linha em textos longos e alinhamento correto.
    """
    _, _, estilo_equip, estilo_num, _ = estilos
    elementos = []

    # Cabeçalho
    header = [
        Paragraph("Equipamento", estilo_equip),
        Paragraph("Qtd", estilo_num),
        Paragraph("Preço sugerido", estilo_num),
        Paragraph("Preço real", estilo_num),
        Paragraph("Total sugerido", estilo_num),
        Paragraph("Total real", estilo_num),
        Paragraph("Diferença (R$)", estilo_num),
        Paragraph("Diferença (%)", estilo_num),
    ]

    # Linhas (convertendo para Paragraph quando útil)
    dados = []
    for _, r in df_itens.iterrows():
        dados.append([
            Paragraph(str(r["Equipamento"]), estilo_equip),
            Paragraph(str(int(r["Quantidade"])), estilo_num),
            Paragraph(_brl(r["Preço sugerido"]), estilo_num),
            Paragraph(_brl(r["Preço real"]), estilo_num),
            Paragraph(_brl(r["Total sugerido"]), estilo_num),
            Paragraph(_brl(r["Total real"]), estilo_num),
            Paragraph(_brl(r["Diferença (R$)"]), estilo_num),
            Paragraph(_pcent(r["Diferença (%)"]), estilo_num),
        ])

    total = len(dados)
    blocos = math.ceil(total / linhas_por_tabela) if total else 1
    largura_util = largura_util_mm * mm

    # Distribuição de larguras (somatório = 1.0)
    # Mais espaço para "Equipamento"
    pesos = [0.40, 0.07, 0.10, 0.10, 0.10, 0.10, 0.07, 0.06]
    col_widths = [largura_util * p for p in pesos]

    for i in range(blocos):
        inicio = i * linhas_por_tabela
        fim = inicio + linhas_por_tabela
        bloco_dados = [header] + dados[inicio:fim]

        t = Table(bloco_dados, colWidths=col_widths, repeatRows=1)
        t.setStyle(TableStyle([
            # Cabeçalho
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 9),

            # Corpo
            ("FONTSIZE", (0, 1), (-1, -1), 8),
            ("ALIGN", (0, 1), (0, -1), "LEFT"),
            ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#DDDDDD")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#FAFAFA")]),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ]))
        elementos.append(t)

        if i < blocos - 1:
            elementos.append(Spacer(1, 5 * mm))
            elementos.append(PageBreak())

    return elementos


def _tabela_resumo(resumo: dict, estilos, largura_util_mm=120):
    _, _, _, estilo_num, _ = estilos
    estilo_lbl = ParagraphStyle("Lbl", fontName="Helvetica", fontSize=9, leading=11)
    header = [Paragraph("Métrica", estilo_lbl), Paragraph("Valor", estilo_lbl)]
    dados = [
        [Paragraph("Itens", estilo_lbl), Paragraph(str(resumo["Itens"]), estilo_num)],
        [Paragraph("Qtd total", estilo_lbl), Paragraph(str(resumo["Qtd total"]), estilo_num)],
        [Paragraph("Total sugerido", estilo_lbl), Paragraph(_brl(resumo["Total sugerido"]), estilo_num)],
        [Paragraph("Total real", estilo_lbl), Paragraph(_brl(resumo["Total real"]), estilo_num)],
        [Paragraph("Diferença (R$)", estilo_lbl), Paragraph(_brl(resumo["Diferença (R$)"]), estilo_num)],
        [Paragraph("Diferença (%)", estilo_lbl), Paragraph(_pcent(resumo["Diferença (%)"]), estilo_num)],
    ]
    t = Table([header] + dados, colWidths=[(largura_util_mm*0.5)*mm, (largura_util_mm*0.5)*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#DDDDDD")),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]))
    return t


def _rodape(canvas, doc):
    page_num = canvas.getPageNumber()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.HexColor("#555555"))
    canvas.drawRightString(doc.pagesize[0] - 15*mm, 10*mm, f"Página {page_num}")


def _gerar_pdf_loja(loja: str, df_loja: pd.DataFrame, output_dir: str,
                    titulo: str = "Relatório de Equipamentos por Loja",
                    logo_path: str | None = None) -> str:
    os.makedirs(output_dir, exist_ok=True)
    fname = f"relatorio_loja_{_safe_filename(loja)}.pdf"
    path = os.path.join(output_dir, fname)

    estilos = _build_styles()
    estilo_titulo, estilo_meta, _, _, estilo_h3 = estilos

    # Dados
    df_itens = _preparar_itens(df_loja)
    resumo = _resumo_da_loja(df_itens)

    doc = SimpleDocTemplate(
        path,
        pagesize=A4,
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=14*mm, bottomMargin=14*mm
    )

    story = []
    # Cabeçalho com espaço/uso de logo
    story.append(_header_block(loja, output_dir, titulo, estilo_titulo, estilo_meta, logo_path))
    story.append(Spacer(1, 3*mm))

    # Itens (tabelas com quebras limpas)
    story.extend(_tabela_itens(df_itens, estilos, largura_util_mm=180, linhas_por_tabela=28))
    story.append(Spacer(1, 6*mm))

    # Resumo
    story.append(Paragraph("Resumo", estilo_h3))
    story.append(Spacer(1, 2*mm))
    story.append(_tabela_resumo(resumo, estilos, largura_util_mm=120))

    doc.build(story, onFirstPage=_rodape, onLaterPages=_rodape)
    return path


# ===== Função pública =====

def gerar_relatorios_pdf(grupos_por_loja: dict[str, pd.DataFrame], output_dir: str,
                         titulo_relatorio: str = "Relatório de Equipamentos por Loja",
                         logo_path: str | None = None) -> list[str]:
    """
    Gera PDFs para todas as lojas do dicionário {loja: DataFrame}.
    - Se 'logo_path' não for informado, tenta usar 'config/logo.png' automaticamente (se existir).
    Retorna lista com caminhos dos arquivos gerados.
    """
    os.makedirs(output_dir, exist_ok=True)
    arquivos = []
    for loja, df_loja in grupos_por_loja.items():
        if df_loja is None or df_loja.empty:
            continue
        path = _gerar_pdf_loja(loja, df_loja, output_dir, titulo=titulo_relatorio, logo_path=logo_path)
        arquivos.append(path)
    return arquivos
