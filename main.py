# -*- coding: utf-8 -*-
"""
Arquivo: main.py
CLI para executar as tarefas do projeto sem modificar o pipeline.

Subcomandos:
  template      -> gera estrutura e template XLSX
  validate      -> lê e valida o arquivo (gera logs/erros_validacao.xlsx se houver)
  process       -> executa o pipeline completo (XLSX + resumo + PDF opcional)
  report-xlsx   -> gera somente os relatórios XLSX por loja
  summary       -> gera somente o resumo consolidado por loja
  report-pdf    -> gera somente os PDFs por loja

Exemplos:
  python main.py template
  python main.py validate --input input/cadastro_equipamentos.xlsx
  python main.py process  --input input/cadastro_equipamentos.xlsx --pdf
  python main.py report-xlsx --input input/cadastro_equipamentos.xlsx
  python main.py summary     --input input/cadastro_equipamentos.xlsx
  python main.py report-pdf  --input input/cadastro_equipamentos.xlsx --pdf-dir output/pdf
"""

import argparse
from pathlib import Path
import warnings

# Silencia aviso chato do openpyxl sobre validações do template
warnings.filterwarnings("ignore", message="Data Validation extension is not supported")

# Importações dos módulos já criados nos blocos anteriores
from src.generate_template import criar_template
from src.processing import carregar_e_validar
from src.pricing import aplicar_preco_sugerido, sugerir_precos, atualizar_historico
from src.reports_xlsx import gerar_relatorios_xlsx
from src.summary_xlsx import gerar_resumo_consolidado

# PDF é opcional; se não estiver disponível, tratamos depois
try:
    from src.reports_pdf import gerar_relatorios_pdf  # Bloco 8
    PDF_DISPONIVEL = True
except Exception:
    gerar_relatorios_pdf = None
    PDF_DISPONIVEL = False


# ---------- utilidades ----------

def _infer_base_dir(input_path: Path) -> Path:
    """
    Infere a raiz do projeto:
      .../relatorio_equipamentos/input/arquivo.xlsx -> base = .../relatorio_equipamentos
      caso contrário, usa ./relatorio_equipamentos relativo ao cwd.
    """
    p = input_path.resolve()
    if p.parent.name.lower() == "input":
        return p.parent.parent
    return Path("relatorio_equipamentos").resolve()


def _aplicar_precos_grupos(grupos: dict, base_dir: Path) -> dict:
    """
    Aplica 'Preço sugerido' onde estiver vazio/<=0 e adiciona 'PrecoSugeridoCalc'
    (apenas para transparência). Retorna novo dict {loja: DataFrame}.
    """
    ajustados = {}
    for loja, df_loja in grupos.items():
        if df_loja is None or df_loja.empty:
            continue
        df_aj = aplicar_preco_sugerido(df_loja, base_dir=str(base_dir), coluna_destino="Preço sugerido")
        # Coluna extra só para consulta; os geradores usam 'Preço sugerido' e 'Preço real'
        df_aj["PrecoSugeridoCalc"] = sugerir_precos(df_loja, base_dir=str(base_dir))
        ajustados[loja] = df_aj
    return ajustados


# ---------- comandos ----------

def cmd_template(args):
    base = Path(args.base).resolve()
    path = criar_template(str(base))
    print("Template criado em:", path)


def cmd_validate(args):
    input_path = Path(args.input).resolve()
    base_dir = _infer_base_dir(input_path)
    df_ok, df_err, grupos = carregar_e_validar(str(input_path))

    print(f"Linhas válidas: {len(df_ok)}")
    print(f"Linhas com erro: {len(df_err)}")
    if not df_err.empty:
        print("Arquivo de erros:", base_dir / "logs" / "erros_validacao.xlsx")
    print("Lojas encontradas:", list(grupos.keys()))


def cmd_process(args):
    # Processamento completo, reaproveitando as funções que já temos
    input_path = Path(args.input).resolve()
    base_dir = _infer_base_dir(input_path)
    output_dir = base_dir / "output"

    # 1) Ler + validar
    df_ok, df_err, grupos = carregar_e_validar(str(input_path))
    erros_path = str(base_dir / "logs" / "erros_validacao.xlsx") if not df_err.empty else None

    if df_ok is None or df_ok.empty or not grupos:
        status = "erro_validacao" if not df_err.empty else "vazio"
        print({"status": status, "erros_path": erros_path})
        return

    # 2) Aplicar preços sugeridos
    grupos_aj = _aplicar_precos_grupos(grupos, base_dir)

    # 3) Gerar XLSX por loja
    arquivos_lojas = gerar_relatorios_xlsx(grupos_aj, output_dir=str(output_dir))

    # 4) Resumo consolidado
    resumo_path = gerar_resumo_consolidado(grupos_aj, output_dir=str(output_dir))

    # 5) PDFs (opcional)
    pdf_paths, pdf_error = [], None
    if args.pdf:
        if PDF_DISPONIVEL and callable(gerar_relatorios_pdf):
            pdf_dir = output_dir / (args.pdf_dir or "pdf")
            pdf_paths = gerar_relatorios_pdf(grupos_aj, output_dir=str(pdf_dir))
        else:
            pdf_error = "PDF indisponível (instale 'reportlab' e verifique src/reports_pdf.py)."

    # 6) Atualiza histórico com preços reais do arquivo
    _ = atualizar_historico(df_ok, base_dir=str(base_dir))

    print({
        "status": "ok",
        "arquivos_lojas": arquivos_lojas,
        "resumo_path": resumo_path,
        "pdf_lojas": pdf_paths,
        "pdf_error": pdf_error,
        "erros_path": erros_path,
    })


def cmd_report_xlsx(args):
    input_path = Path(args.input).resolve()
    base_dir = _infer_base_dir(input_path)
    output_dir = base_dir / "output"

    df_ok, df_err, grupos = carregar_e_validar(str(input_path))
    if df_ok.empty or not grupos:
        print("Nada a fazer (sem dados válidos). Verifique logs/erros_validacao.xlsx se houve erros.")
        return

    grupos_aj = _aplicar_precos_grupos(grupos, base_dir)
    arquivos = gerar_relatorios_xlsx(grupos_aj, output_dir=str(output_dir))
    print("Relatórios gerados:", arquivos)


def cmd_summary(args):
    input_path = Path(args.input).resolve()
    base_dir = _infer_base_dir(input_path)
    output_dir = base_dir / "output"

    df_ok, df_err, grupos = carregar_e_validar(str(input_path))
    if df_ok.empty or not grupos:
        print("Nada a resumir (sem dados válidos).")
        return

    grupos_aj = _aplicar_precos_grupos(grupos, base_dir)
    caminho = gerar_resumo_consolidado(grupos_aj, output_dir=str(output_dir))
    print("Resumo gerado em:", caminho)


def cmd_report_pdf(args):
    if not PDF_DISPONIVEL or not callable(gerar_relatorios_pdf):
        print("PDF indisponível: instale 'reportlab' e confira src/reports_pdf.py.")
        return

    input_path = Path(args.input).resolve()
    base_dir = _infer_base_dir(input_path)
    pdf_dir = base_dir / "output" / (args.pdf_dir or "pdf")

    df_ok, df_err, grupos = carregar_e_validar(str(input_path))
    if df_ok.empty or not grupos:
        print("Nada a fazer (sem dados válidos).")
        return

    grupos_aj = _aplicar_precos_grupos(grupos, base_dir)
    pdfs = gerar_relatorios_pdf(grupos_aj, output_dir=str(pdf_dir))
    print("PDFs gerados:", pdfs)


# ---------- entrypoint ----------

def main():
    parser = argparse.ArgumentParser(description="CLI - Relatórios de Equipamentos")
    sub = parser.add_subparsers(dest="command", required=True)

    # template
    p_tpl = sub.add_parser("template", help="Gera estrutura e o template XLSX")
    p_tpl.add_argument("--base", default="relatorio_equipamentos", help="Pasta base do projeto")
    p_tpl.set_defaults(func=cmd_template)

    # validate
    p_val = sub.add_parser("validate", help="Valida o arquivo de entrada e salva erros (se houver)")
    p_val.add_argument("--input", "-i", required=True, help="Caminho do arquivo XLSX (aba 'Cadastro')")
    p_val.set_defaults(func=cmd_validate)

    # process
    p_proc = sub.add_parser("process", help="Executa o pipeline completo (XLSX + resumo + PDF opcional)")
    p_proc.add_argument("--input", "-i", required=True, help="Caminho do arquivo XLSX (aba 'Cadastro')")
    p_proc.add_argument("--pdf", action="store_true", help="Se presente, gera também PDFs por loja")
    p_proc.add_argument("--pdf-dir", default="pdf", help="Subpasta dentro de output/ para salvar os PDFs")
    p_proc.set_defaults(func=cmd_process)

    # report-xlsx
    p_rx = sub.add_parser("report-xlsx", help="Gera apenas os relatórios XLSX por loja")
    p_rx.add_argument("--input", "-i", required=True, help="Caminho do arquivo XLSX (aba 'Cadastro')")
    p_rx.set_defaults(func=cmd_report_xlsx)

    # summary
    p_sum = sub.add_parser("summary", help="Gera apenas o resumo consolidado por loja")
    p_sum.add_argument("--input", "-i", required=True, help="Caminho do arquivo XLSX (aba 'Cadastro')")
    p_sum.set_defaults(func=cmd_summary)

    # report-pdf
    p_pdf = sub.add_parser("report-pdf", help="Gera apenas os PDFs por loja")
    p_pdf.add_argument("--input", "-i", required=True, help="Caminho do arquivo XLSX (aba 'Cadastro')")
    p_pdf.add_argument("--pdf-dir", default="pdf", help="Subpasta dentro de output/ para salvar os PDFs")
    p_pdf.set_defaults(func=cmd_report_pdf)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
