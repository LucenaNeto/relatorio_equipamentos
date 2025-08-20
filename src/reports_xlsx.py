
"""
Módulo: reports_xlsx.py
Gera relatórios XLSX por loja a partir do dicionário { loja: DataFrame }.

Uso esperado (exemplo):
    from src.processing import carregar_e_validar
    from src.reports_xlsx import gerar_relatorios_xlsx

    df_ok, df_err, grupos = carregar_e_validar("relatorio_equipamentos/input/seu_arquivo.xlsx")
    arquivos = gerar_relatorios_xlsx(grupos, output_dir="relatorio_equipamentos/output")
    print(arquivos)
"""

import os
from datetime import datetime
import pandas as pd
import numpy as np


def _safe_filename(name: str) -> str:
    """
    Gera um nome de arquivo seguro (sem caracteres especiais), limitado a 120 chars.
    Ex.: "6402" -> "6402"; "Loja 6402" -> "Loja_6402"
    """
    import re
    name = str(name).strip()
    name = re.sub(r"[^\w\-]+", "_", name, flags=re.UNICODE)
    return name[:120] or "sem_nome"


def _gerar_relatorio_loja(loja: str, df_loja: pd.DataFrame, output_dir: str) -> str:
    """
    Gera 1 arquivo XLSX para a 'loja' recebida com as abas 'Itens' e 'Resumo'.

    Regras de cálculo:
      - Total sugerido = Quantidade * Preço sugerido
      - Total real     = Quantidade * Preço real (quando vazio, usa Preço sugerido como fallback)
      - Diferença (R$) = Total real - Total sugerido
      - Diferença (%)  = (Total real / Total sugerido) - 1, quando Total sugerido > 0; caso contrário, vazio
    """
    os.makedirs(output_dir, exist_ok=True)
    fname = f"relatorio_loja_{_safe_filename(loja)}.xlsx"
    path = os.path.join(output_dir, fname)

    # Cópia para não alterar o DF original
    df = df_loja.copy()

    # Garantir tipos numéricos corretos
    df["Quantidade"] = pd.to_numeric(df["Quantidade"], errors="coerce").fillna(0).astype(int).clip(lower=0)
    df["Preço sugerido"] = pd.to_numeric(df["Preço sugerido"], errors="coerce")
    df["Preço real"]     = pd.to_numeric(df["Preço real"], errors="coerce")

    # Preços negativos viram NaN (não faz sentido no contexto)
    for c in ["Preço sugerido", "Preço real"]:
        df[c] = df[c].where(df[c].isna() | (df[c] >= 0), np.nan)

    # Cálculos de total
    preco_sugerido_eff = df["Preço sugerido"].fillna(0)
    preco_real_eff     = df["Preço real"].where(df["Preço real"].notna(), preco_sugerido_eff)

    df["Total sugerido"] = df["Quantidade"] * preco_sugerido_eff
    df["Total real"]     = df["Quantidade"] * preco_real_eff
    df["Diferença (R$)"] = df["Total real"] - df["Total sugerido"]
    df["Diferença (%)"]  = np.where(df["Total sugerido"] > 0,
                                    (df["Total real"] / df["Total sugerido"]) - 1.0,
                                    np.nan)

    # Ordenar colunas para a aba Itens
    cols_itens = [
        "Equipamento", "Quantidade", "Preço sugerido", "Preço real",
        "Total sugerido", "Total real", "Diferença (R$)", "Diferença (%)"
    ]
    df_itens = df[cols_itens].copy()

    # Resumo agregado
    total_sugerido = df_itens["Total sugerido"].sum()
    total_real     = df_itens["Total real"].sum()
    resumo = pd.DataFrame({
        "Métrica": ["Itens", "Qtd total", "Total sugerido", "Total real", "Diferença (R$)", "Diferença (%)"],
        "Valor": [
            len(df_itens),
            df_itens["Quantidade"].sum(),
            total_sugerido,
            total_real,
            total_real - total_sugerido,
            (total_real / total_sugerido - 1.0) if total_sugerido > 0 else np.nan
        ]
    })

    # Exporta com formatação
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        df_itens.to_excel(writer, sheet_name="Itens", index=False)
        resumo.to_excel(writer, sheet_name="Resumo", index=False)

        wb = writer.book
        ws_itens = writer.sheets["Itens"]
        ws_sum   = writer.sheets["Resumo"]

        # Formatações
        money = wb.add_format({"num_format": 'R$ #,##0.00'})
        perc  = wb.add_format({"num_format": '0.00%'})
        iint  = wb.add_format({"num_format": '0'})
        head  = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1})

        # Cabeçalho estilizado na aba Itens
        for c, name in enumerate(cols_itens):
            ws_itens.write(0, c, name, head)

        # Larguras e formatos
        ws_itens.set_column("A:A", 28)          # Equipamento
        ws_itens.set_column("B:B", 12, iint)    # Quantidade
        ws_itens.set_column("C:D", 16, money)   # Preços
        ws_itens.set_column("E:F", 16, money)   # Totais
        ws_itens.set_column("G:G", 16, money)   # Diferença (R$)
        ws_itens.set_column("H:H", 14, perc)    # Diferença (%)

        # Formatação condicional: diferença negativa (prejuízo) em vermelho
        last_row = len(df_itens)
        ws_itens.conditional_format(1, 6, last_row, 6, {
            "type": "cell", "criteria": "<", "value": 0,
            "format": wb.add_format({"font_color": "#9C0006"})
        })

        # Resumo com formatos adequados por linha
        ws_sum.set_column("A:A", 20)
        ws_sum.set_column("B:B", 20)

        for i in range(1, len(resumo) + 1):
            met = resumo.iloc[i-1, 0]
            val = resumo.iloc[i-1, 1]
            if "Total" in met or "Diferença (R$)" in met:
                ws_sum.write(i, 1, val, money)
            elif "Diferença (%)" in met:
                ws_sum.write(i, 1, val, perc)
            elif "Qtd" in met or "Itens" in met:
                ws_sum.write(i, 1, val, iint)
            else:
                ws_sum.write(i, 1, val)

        # Metadados úteis
        ws_sum.write(0, 3, "Loja:");      ws_sum.write(0, 4, str(loja))
        ws_sum.write(1, 3, "Gerado em:"); ws_sum.write(1, 4, datetime.now().strftime("%Y-%m-%d %H:%M"))

    return path


def gerar_relatorios_xlsx(grupos_por_loja: dict[str, pd.DataFrame], output_dir: str) -> list[str]:
    """
    Gera os relatórios XLSX para todas as lojas do dicionário 'grupos_por_loja'.
    Retorna a lista de caminhos dos arquivos gerados.
    """
    os.makedirs(output_dir, exist_ok=True)
    arquivos = []
    for loja, df_loja in grupos_por_loja.items():
        if df_loja is None or df_loja.empty:
            # pula lojas sem dados
            continue
        caminho = _gerar_relatorio_loja(loja, df_loja, output_dir)
        arquivos.append(caminho)
    return arquivos
