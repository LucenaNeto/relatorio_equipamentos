# -*- coding: utf-8 -*-
"""
Módulo: processing.py (versão inicial)
- Lê o XLSX de entrada (aba 'Cadastro')
- Valida as linhas
- Separa os dados por loja

Obs.: Ainda NÃO gera relatórios. Isso virá no próximo bloco.
"""

import os
import pandas as pd
import numpy as np
from src.settings import LOJAS_VALIDAS  # lista oficial de lojas (Bloco 1)

# Diretórios base (ajuste se necessário)
BASE_DIR  = "relatorio_equipamentos"
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR= os.path.join(BASE_DIR, "output")
LOGS_DIR  = os.path.join(BASE_DIR, "logs")

# Colunas esperadas na aba "Cadastro"
REQUIRED_COLS = ["Loja", "Equipamento", "Quantidade", "Preço sugerido", "Preço real"]


def _read_cadastro_from_xlsx(path: str) -> pd.DataFrame:
    """
    Lê a planilha de entrada.
    - Prioriza a aba 'Cadastro'; se não existir, usa a primeira aba.
    - Normaliza tipos e garante colunas obrigatórias.
    - Remove linhas obviamente vazias (sem Equipamento) e quantidades <= 0.
    """
    if not os.path.exists(path):
        raise FileNotFoundError(f"Arquivo não encontrado: {path}")

    xls = pd.ExcelFile(path)
    sheet = "Cadastro" if "Cadastro" in xls.sheet_names else xls.sheet_names[0]

    # Lê forçando Loja/Equipamento como texto (evita perder zeros ou códigos)
    df = pd.read_excel(path, sheet_name=sheet, dtype={"Loja": str, "Equipamento": str})

    # Normaliza cabeçalhos (remove espaços extras)
    df.rename(columns={c: c.strip() for c in df.columns}, inplace=True)

    # Garante colunas obrigatórias (se faltar alguma, cria vazia)
    for col in REQUIRED_COLS:
        if col not in df.columns:
            df[col] = np.nan

    # Ajusta tipos básicos
    df["Loja"]        = df["Loja"].astype(str).str.strip()
    df["Equipamento"] = df["Equipamento"].astype(str).str.strip()

    # Quantidade: inteiro >= 0 (valores inválidos viram 0)
    df["Quantidade"] = pd.to_numeric(df["Quantidade"], errors="coerce") \
                          .fillna(0).clip(lower=0).astype(int)

    # Preços: float >= 0 (mantém NaN quando vazio)
    for c in ["Preço sugerido", "Preço real"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
        df[c] = df[c].where(df[c].notna(), np.nan)
        df[c] = df[c].clip(lower=0)

    # Remove linhas sem Equipamento
    df = df[df["Equipamento"].str.len() > 0]
    # Remove quantidade 0 (deixa só registros com pelo menos 1 unidade)
    df = df[df["Quantidade"] > 0]

    df.reset_index(drop=True, inplace=True)
    return df


def validar_e_limpar(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Valida as linhas do DataFrame e separa em (válidos, erros).

    Regras:
      - Loja: obrigatória e deve estar em LOJAS_VALIDAS.
      - Equipamento: obrigatório (não vazio).
      - Quantidade: inteiro >= 1.
      - Preço sugerido e Preço real: se informados, devem ser >= 0.

    Retornos:
      - df_validos: somente linhas aprovadas nas regras.
      - df_erros: linhas rejeitadas + coluna 'Erro' explicando o motivo.
    """
    erros, ok = [], []

    for _, row in df.iterrows():
        problemas = []

        loja = str(row.get("Loja", "")).strip()
        equip = str(row.get("Equipamento", "")).strip()
        qtd = row.get("Quantidade", 0)
        ps = row.get("Preço sugerido", np.nan)
        pr = row.get("Preço real", np.nan)

        # Loja
        if not loja:
            problemas.append("Loja vazia")
        elif loja not in LOJAS_VALIDAS:
            problemas.append(f"Loja inválida ({loja})")

        # Equipamento
        if not equip:
            problemas.append("Equipamento vazio")

        # Quantidade
        try:
            qtd_int = int(qtd)
        except Exception:
            problemas.append(f"Quantidade inválida ({qtd})")
            qtd_int = None
        if qtd_int is not None and qtd_int < 1:
            problemas.append(f"Quantidade < 1 ({qtd_int})")

        # Preços (se informados)
        for label, val in [("Preço sugerido", ps), ("Preço real", pr)]:
            if pd.notna(val) and val < 0:
                problemas.append(f"{label} negativo ({val})")

        if problemas:
            r = row.copy()
            r["Erro"] = "; ".join(problemas)
            erros.append(r)
        else:
            ok.append(row)

    df_validos = pd.DataFrame(ok, columns=df.columns) if ok else pd.DataFrame(columns=df.columns)
    df_erros   = pd.DataFrame(erros) if erros else pd.DataFrame(columns=list(df.columns) + ["Erro"])
    return df_validos, df_erros


def separar_por_loja(df_validos: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """
    Separa o DataFrame validado por loja.
    Retorna um dicionário: { '3569': df_loja_3569, '6402': df_loja_6402, ... }
    """
    grupos = {}
    for loja, grupo in df_validos.groupby("Loja", dropna=False):
        grupos[str(loja)] = grupo.reset_index(drop=True).copy()
    return grupos


def carregar_e_validar(input_path: str) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    """
    Função de alto nível para este bloco:
      - Lê o arquivo,
      - Valida as linhas,
      - Salva 'logs/erros_validacao.xlsx' (se houver erros),
      - Retorna (df_validos, df_erros, grupos_por_loja).
    """
    os.makedirs(LOGS_DIR, exist_ok=True)

    df = _read_cadastro_from_xlsx(input_path)

    # Validação
    df_validos, df_erros = validar_e_limpar(df)

    # Exporta erros (se houver)
    if not df_erros.empty:
        caminho_erros = os.path.join(LOGS_DIR, "erros_validacao.xlsx")
        with pd.ExcelWriter(caminho_erros, engine="xlsxwriter") as w:
            df_erros.to_excel(w, sheet_name="Erros", index=False)

    # Separa por loja
    grupos = separar_por_loja(df_validos)

    return df_validos, df_erros, grupos
