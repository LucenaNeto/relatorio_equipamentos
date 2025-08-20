# -*- coding: utf-8 -*-
"""
Módulo: summary_xlsx.py
Gera o arquivo 'resumo_por_loja.xlsx' consolidando os totais por loja.

Uso típico:
    from src.processing import carregar_e_validar
    from src.summary_xlsx import gerar_resumo_consolidado

    df_ok, df_err, grupos = carregar_e_validar("relatorio_equipamentos/input/seu_arquivo.xlsx")
    caminho = gerar_resumo_consolidado(grupos, output_dir="relatorio_equipamentos/output")
    print(caminho)
"""

import os
import pandas as pd
import numpy as np


def _linha_resumo_loja(loja: str, df_loja: pd.DataFrame) -> dict:
    """
    Calcula os agregados de uma loja e devolve um dicionário com as métricas.
    - Quantidade: int >= 0
    - Preço sugerido / real: >= 0 (NaN quando vazio)
    - Total real usa 'Preço sugerido' como fallback quando 'Preço real' estiver NaN
    """
    df = df_loja.copy()

    # Garantir tipos numéricos corretos
    df["Quantidade"] = pd.to_numeric(df["Quantidade"], errors="coerce").fillna(0).astype(int).clip(lower=0)
    df["Preço sugerido"] = pd.to_numeric(df["Preço sugerido"], errors="coerce")
    df["Preço real"]     = pd.to_numeric(df["Preço real"], errors="coerce")

    # Preços negativos não fazem sentido => trata como NaN
    for c in ["Preço sugerido", "Preço real"]:
        df[c] = df[c].where(df[c].isna() | (df[c] >= 0), np.nan)

    # Cálculos por linha
    ps = df["Preço sugerido"].fillna(0)
    pr_eff = df["Preço real"].where(df["Preço real"].notna(), ps)

    total_sugerido = (df["Quantidade"] * ps).sum()
    total_real     = (df["Quantidade"] * pr_eff).sum()

    return {
        "Loja": str(loja),
        "Itens": len(df),
        "Qtd total": int(df["Quantidade"].sum()),
        "Total sugerido": float(total_sugerido),
        "Total real": float(total_real),
        "Diferença (R$)": float(total_real - total_sugerido),
        "Diferença (%)": (float(total_real / total_sugerido - 1.0) if total_sugerido > 0 else np.nan),
    }


def gerar_resumo_consolidado(grupos_por_loja: dict[str, pd.DataFrame], output_dir: str) -> str:
    """
    Gera o arquivo único 'resumo_por_loja.xlsx' no diretório 'output_dir'.
    Retorna o caminho completo do arquivo criado.
    """
    os.makedirs(output_dir, exist_ok=True)
    path = os.path.join(output_dir, "resumo_por_loja.xlsx")

    # Monta tabela de resumo
    linhas = []
    for loja, df_loja in grupos_por_loja.items():
        if df_loja is None or df_loja.empty:
            continue
        linhas.append(_linha_resumo_loja(loja, df_loja))

    df_resumo = pd.DataFrame(linhas, columns=[
        "Loja", "Itens", "Qtd total", "Total sugerido", "Total real", "Diferença (R$)", "Diferença (%)"
    ]).sort_values("Loja")

    # Exporta com formatação adequada
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        df_resumo.to_excel(writer, sheet_name="Resumo", index=False)

        wb = writer.book
        ws = writer.sheets["Resumo"]

        header = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1})
        money  = wb.add_format({"num_format": 'R$ #,##0.00'})
        perc   = wb.add_format({"num_format": '0.00%'})
        iint   = wb.add_format({"num_format": '0'})

        # Reescreve cabeçalho com estilo
        for c, name in enumerate(df_resumo.columns):
            ws.write(0, c, name, header)

        # Larguras e formatos
        ws.set_column("A:A", 12)          # Loja
        ws.set_column("B:B", 10, iint)    # Itens
        ws.set_column("C:C", 12, iint)    # Qtd total
        ws.set_column("D:E", 18, money)   # Totais
        ws.set_column("F:F", 18, money)   # Diferença (R$)
        ws.set_column("G:G", 14, perc)    # Diferença (%)

    return path
