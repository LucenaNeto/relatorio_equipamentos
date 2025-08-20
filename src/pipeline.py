# -*- coding: utf-8 -*-
"""
Módulo: pipeline.py
Integra o fluxo completo:
  1) Carrega e valida o XLSX de entrada;
  2) Calcula e aplica o Preço Sugerido (quando estiver vazio);
  3) Gera relatórios por loja (XLSX);
  4) Gera resumo consolidado por loja (XLSX);
  5) Atualiza o histórico com preços reais informados.

Uso típico:
    from src.pipeline import processar_arquivo
    result = processar_arquivo("relatorio_equipamentos/input/cadastro_equipamentos.xlsx")
    print(result)
"""

from pathlib import Path
import os
import pandas as pd

from src.processing import carregar_e_validar
from src.pricing import sugerir_precos, aplicar_preco_sugerido, atualizar_historico
from src.reports_xlsx import gerar_relatorios_xlsx
from src.summary_xlsx import gerar_resumo_consolidado


def _infer_base_dir(input_path: str) -> Path:
    """
    Infere a pasta base do projeto a partir do caminho do arquivo de entrada.
    Se o arquivo estiver em .../relatorio_equipamentos/input/arquivo.xlsx,
    a base é .../relatorio_equipamentos.
    Caso contrário, usa a pasta 'relatorio_equipamentos' relativa ao cwd.
    """
    p = Path(input_path).resolve()
    if p.parent.name.lower() == "input":
        return p.parent.parent
    # fallback: pasta 'relatorio_equipamentos' no cwd
    return Path("relatorio_equipamentos").resolve()


def _aplicar_precos_sugeridos_por_loja(grupos_por_loja: dict, base_dir: Path) -> dict:
    """
    Para cada loja do dicionário {loja: DataFrame}, calcula os preços sugeridos e
    preenche a coluna 'Preço sugerido' quando estiver vazia ou <= 0.
    Retorna um novo dicionário {loja: DataFrame} já ajustado.
    """
    ajustados = {}
    for loja, df_loja in grupos_por_loja.items():
        if df_loja is None or df_loja.empty:
            continue
        df_aj = aplicar_preco_sugerido(df_loja, base_dir=str(base_dir), coluna_destino="Preço sugerido")
        # (Opcional) também podemos anexar a coluna calculada para transparência:
        df_aj["PrecoSugeridoCalc"] = sugerir_precos(df_loja, base_dir=str(base_dir))
        ajustados[loja] = df_aj
    return ajustados


def processar_arquivo(input_path: str) -> dict:
    """
    Executa o pipeline completo e retorna um dicionário com:
      {
        "status": "ok" | "erro_validacao" | "vazio",
        "arquivos_lojas": [lista de caminhos gerados],
        "resumo_path": "caminho do resumo_por_loja.xlsx",
        "erros_path": "caminho do erros_validacao.xlsx (ou None)"
      }
    """
    base_dir = _infer_base_dir(input_path)
    input_path = str(Path(input_path).resolve())

    # 1) Lê e valida
    df_ok, df_err, grupos = carregar_e_validar(input_path)

    # Se houve erros, aponta o caminho do log (o módulo de validação já escreveu)
    erros_path = None
    if not df_err.empty:
        erros_path = str(base_dir / "logs" / "erros_validacao.xlsx")

    if df_ok is None or df_ok.empty or not grupos:
        return {
            "status": "erro_validacao" if not df_err.empty else "vazio",
            "arquivos_lojas": [],
            "resumo_path": None,
            "erros_path": erros_path,
        }

    # 2) Aplica preço sugerido (preenche 'Preço sugerido' onde estiver vazio/<=0)
    grupos_aj = _aplicar_precos_sugeridos_por_loja(grupos, base_dir)

    # 3) Gera relatórios por loja (XLSX)
    output_dir = base_dir / "output"
    arquivos_lojas = gerar_relatorios_xlsx(grupos_aj, output_dir=str(output_dir))

    # 4) Gera resumo consolidado
    resumo_path = gerar_resumo_consolidado(grupos_aj, output_dir=str(output_dir))

    # 5) Atualiza histórico com preços reais do arquivo atual
    _ = atualizar_historico(df_ok, base_dir=str(base_dir))

    return {
        "status": "ok",
        "arquivos_lojas": arquivos_lojas,
        "resumo_path": resumo_path,
        "erros_path": erros_path,
    }
