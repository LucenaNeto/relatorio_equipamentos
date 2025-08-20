# -*- coding: utf-8 -*-
"""
Módulo: pricing.py
Engine de "Preço Sugerido" com base em histórico local.

Este módulo NÃO altera planilhas nem grava relatórios; ele só:
- carrega/salva o histórico de preços
- calcula sugestões de preços para um DataFrame já validado

Integrações virão no próximo bloco.
"""

import os
from datetime import datetime
import pandas as pd
import numpy as np


# === Configuração de pastas/arquivos ===

def _paths(base_dir: str = "relatorio_equipamentos") -> dict:
    """
    Devolve os caminhos usados pelo módulo, baseados em 'base_dir'.
    """
    config_dir = os.path.join(base_dir, "config")
    os.makedirs(config_dir, exist_ok=True)
    hist_parquet = os.path.join(config_dir, "precos_historico.parquet")
    hist_csv     = os.path.join(config_dir, "precos_historico.csv")
    return {"config_dir": config_dir, "hist_parquet": hist_parquet, "hist_csv": hist_csv}


# === Histórico ===

def carregar_historico(base_dir: str = "relatorio_equipamentos") -> pd.DataFrame:
    """
    Carrega o histórico de preços.
    - Tenta Parquet primeiro; se não existir/der erro, tenta CSV.
    - Se não existir nada, retorna DataFrame vazio com colunas padrão.
    Colunas: Loja | Equipamento | PrecoReal | Fonte | ts
    """
    p = _paths(base_dir)
    if os.path.exists(p["hist_parquet"]):
        try:
            return pd.read_parquet(p["hist_parquet"])
        except Exception:
            pass  # cai para CSV
    if os.path.exists(p["hist_csv"]):
        try:
            return pd.read_csv(p["hist_csv"])
        except Exception:
            pass
    return pd.DataFrame(columns=["Loja", "Equipamento", "PrecoReal", "Fonte", "ts"])


def salvar_historico(df_hist: pd.DataFrame, base_dir: str = "relatorio_equipamentos") -> None:
    """
    Salva o DataFrame de histórico preferindo Parquet.
    Se falhar (ex.: ausência de pyarrow), salva em CSV.
    Mantém no máximo 200k linhas (FIFO) para evitar crescimento infinito.
    """
    p = _paths(base_dir)
    os.makedirs(p["config_dir"], exist_ok=True)

    # Limite de segurança
    if len(df_hist) > 200_000:
        df_hist = df_hist.iloc[-200_000:].copy()

    try:
        df_hist.to_parquet(p["hist_parquet"], index=False)
    except Exception:
        df_hist.to_csv(p["hist_csv"], index=False, encoding="utf-8")


def atualizar_historico(df_validado: pd.DataFrame, base_dir: str = "relatorio_equipamentos") -> int:
    """
    Acrescenta ao histórico as linhas do arquivo atual que têm 'Preço real' > 0.
    Retorna a quantidade de registros inseridos.
    """
    if df_validado is None or df_validado.empty:
        return 0

    # Seleciona apenas as colunas necessárias
    df_add = df_validado.loc[
        df_validado["Preço real"].notna() & (pd.to_numeric(df_validado["Preço real"], errors="coerce") > 0),
        ["Loja", "Equipamento", "Preço real"]
    ].copy()

    if df_add.empty:
        return 0

    df_add.rename(columns={"Preço real": "PrecoReal"}, inplace=True)
    df_add["Fonte"] = "input_atual"
    df_add["ts"] = pd.Timestamp.now().isoformat()

    hist = carregar_historico(base_dir)
    hist = pd.concat([hist, df_add], ignore_index=True)
    salvar_historico(hist, base_dir)
    return len(df_add)


# === Sugestão de Preço ===

def sugerir_precos(df_validado: pd.DataFrame, base_dir: str = "relatorio_equipamentos") -> pd.Series:
    """
    Calcula a coluna 'PrecoSugeridoCalc' para o DataFrame validado.
    Regras em cascata:
      1) Usa 'Preço sugerido' do arquivo se > 0;
      2) Mediana histórica por (Loja, Equipamento) com base em 'PrecoReal';
      3) Mediana do 'Preço real' no arquivo atual por 'Equipamento';
      4) Mediana do 'Preço sugerido' no arquivo atual por 'Equipamento';
      5) Caso tudo falhe, NaN.

    Retorna: pandas.Series alinhada por index com name='PrecoSugeridoCalc'.
    Não altera o DataFrame original (você pode atribuir a uma nova coluna).
    """
    if df_validado is None or df_validado.empty:
        return pd.Series(dtype=float, name="PrecoSugeridoCalc")

    # Garante tipos básicos
    df = df_validado.copy()
    df["Loja"]        = df["Loja"].astype(str).str.strip()
    df["Equipamento"] = df["Equipamento"].astype(str).str.strip()
    df["Preço sugerido"] = pd.to_numeric(df["Preço sugerido"], errors="coerce")
    df["Preço real"]     = pd.to_numeric(df["Preço real"], errors="coerce")

    # 1) mapa de "Preço sugerido" já informado (>0)
    preco_sugerido_informado = df["Preço sugerido"]

    # 2) medianas históricas de PrecoReal por (Loja, Equipamento)
    hist = carregar_historico(base_dir)
    med_hist = {}
    if not hist.empty and {"Loja", "Equipamento", "PrecoReal"}.issubset(hist.columns):
        med_hist = (
            hist.groupby(["Loja", "Equipamento"])["PrecoReal"]
                .median()   # por padrão, ignora NaN
                .to_dict()
        )

    # 3) medianas no arquivo atual por Equipamento (real primeiro, depois sugerido)
    med_atual_real_por_eq = df.groupby("Equipamento")["Preço real"].median().to_dict()
    med_atual_sug_por_eq  = df.groupby("Equipamento")["Preço sugerido"].median().to_dict()

    # Calcula em loop (índice preservado)
    sugeridos = []
    for _, row in df.iterrows():
        # 1) já informado
        if pd.notna(row["Preço sugerido"]) and row["Preço sugerido"] > 0:
            sugeridos.append(row["Preço sugerido"])
            continue

        # 2) histórico (Loja, Equipamento)
        chave = (row["Loja"], row["Equipamento"])
        v = med_hist.get(chave, np.nan)

        # 3) mediana do arquivo atual (Preço real por equipamento)
        if pd.isna(v) or v <= 0:
            v = med_atual_real_por_eq.get(row["Equipamento"], np.nan)

        # 4) mediana do arquivo atual (Preço sugerido por equipamento)
        if pd.isna(v) or v <= 0:
            v = med_atual_sug_por_eq.get(row["Equipamento"], np.nan)

        # 5) mantém NaN se não achou nada válido
        sugeridos.append(v if (pd.notna(v) and v > 0) else np.nan)

    return pd.Series(sugeridos, index=df.index, name="PrecoSugeridoCalc")


def aplicar_preco_sugerido(df_validado: pd.DataFrame, base_dir: str = "relatorio_equipamentos",
                           coluna_destino: str = "Preço sugerido") -> pd.DataFrame:
    """
    (Opcional) Retorna uma cópia do DF onde os 'Preço sugerido' vazios são
    preenchidos com 'PrecoSugeridoCalc'.

    Útil se você quiser persistir/emitir relatórios usando a própria coluna
    'Preço sugerido' como fonte única (mas guardando a origem do cálculo em log).
    """
    df = df_validado.copy()
    sugest = sugerir_precos(df_validado, base_dir=base_dir)
    if coluna_destino not in df.columns:
        df[coluna_destino] = np.nan
    mask = df[coluna_destino].isna() | (pd.to_numeric(df[coluna_destino], errors="coerce") <= 0)
    df.loc[mask, coluna_destino] = sugest[mask]
    return df
