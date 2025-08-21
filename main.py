# -*- coding: utf-8 -*-
"""
Arquivo: main.py (com MENU interativo)
Permite executar as principais funções do projeto escolhendo por números.

Como usar:
    py main.py
ou
    python main.py
"""

import sys
import warnings
from pathlib import Path

# Silencia aviso do openpyxl sobre validações do template (não afeta os dados)
warnings.filterwarnings("ignore", message="Data Validation extension is not supported")

# Importações dos módulos do projeto
from src.generate_template import criar_template
from src.processing import carregar_e_validar
from src.pricing import aplicar_preco_sugerido, sugerir_precos, atualizar_historico
from src.reports_xlsx import gerar_relatorios_xlsx
from src.summary_xlsx import gerar_resumo_consolidado

# PDF é opcional
try:
    from src.reports_pdf import gerar_relatorios_pdf
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
    Preenche 'Preço sugerido' onde estiver vazio/<=0 e adiciona 'PrecoSugeridoCalc'
    (apenas para consulta). Retorna novo dict {loja: DataFrame}.
    """
    ajustados = {}
    for loja, df_loja in grupos.items():
        if df_loja is None or df_loja.empty:
            continue
        df_aj = aplicar_preco_sugerido(df_loja, base_dir=str(base_dir), coluna_destino="Preço sugerido")
        df_aj["PrecoSugeridoCalc"] = sugerir_precos(df_loja, base_dir=str(base_dir))
        ajustados[loja] = df_aj
    return ajustados


def _input_path(prompt: str, default_rel: str = "input/cadastro_equipamentos.xlsx") -> Path:
    """
    Pede um caminho de arquivo ao usuário. Se vazio, usa default_rel relativo à raiz do projeto.
    Repete até existir.
    """
    while True:
        raw = input(f"{prompt} (Enter para '{default_rel}'): ").strip()
        if raw == "":
            p = Path(default_rel)
        else:
            p = Path(raw)
        # Se veio relativo, resolva contra a pasta atual
        p = p if p.is_absolute() else (Path.cwd() / p)
        if p.exists():
            return p
        print(f"[x] Arquivo não encontrado: {p}\n    Tente novamente.")


def _input_yesno(prompt: str, default: bool = True) -> bool:
    """
    Pergunta sim/não. Enter aceita o default.
    """
    suf = "[S/n]" if default else "[s/N]"
    while True:
        resp = input(f"{prompt} {suf}: ").strip().lower()
        if resp == "" and default:
            return True
        if resp == "" and not default:
            return False
        if resp in ("s", "sim", "y", "yes"):
            return True
        if resp in ("n", "nao", "não", "no"):
            return False
        print("Digite 's' para sim ou 'n' para não.")


# ---------- ações do menu ----------

def acao_1_template():
    base = Path.cwd() / "relatorio_equipamentos"
    path = criar_template(str(base))
    print("\n[OK] Template criado em:", path)


def acao_2_validate():
    p = _input_path("Informe o caminho do XLSX (aba 'Cadastro')")
    base_dir = _infer_base_dir(p)
    df_ok, df_err, grupos = carregar_e_validar(str(p))
    print("\n--- RESULTADO ---")
    print(f"Linhas válidas: {len(df_ok)}")
    print(f"Linhas com erro: {len(df_err)}")
    if not df_err.empty:
        print("Arquivo de erros:", base_dir / "logs" / "erros_validacao.xlsx")
    print("Lojas encontradas:", list(grupos.keys()))


def acao_3_processar_completo():
    p = _input_path("Informe o caminho do XLSX (aba 'Cadastro')")
    gerar_pdf = _input_yesno("Gerar PDFs por loja também?", default=True)

    base_dir = _infer_base_dir(p)
    output_dir = base_dir / "output"

    # 1) Ler + validar
    df_ok, df_err, grupos = carregar_e_validar(str(p))
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
    if gerar_pdf:
        if PDF_DISPONIVEL and callable(gerar_relatorios_pdf):
            pdf_dir = output_dir / "pdf"
            pdf_paths = gerar_relatorios_pdf(grupos_aj, output_dir=str(pdf_dir))
        else:
            pdf_error = "PDF indisponível (instale 'reportlab' e verifique src/reports_pdf.py)."

    # 6) Atualiza histórico com preços reais do arquivo
    _ = atualizar_historico(df_ok, base_dir=str(base_dir))

    print("\n--- RESUMO DA EXECUÇÃO ---")
    print("Status: ok")
    print("Relatórios por loja:", *arquivos_lojas, sep="\n - ")
    print("Resumo consolidado:", resumo_path)
    if gerar_pdf:
        if pdf_paths:
            print("PDFs por loja:", *pdf_paths, sep="\n - ")
        if pdf_error:
            print("Aviso PDF:", pdf_error)
    if erros_path:
        print("Erros de validação:", erros_path)


def acao_4_report_xlsx():
    p = _input_path("Informe o caminho do XLSX (aba 'Cadastro')")
    base_dir = _infer_base_dir(p)
    output_dir = base_dir / "output"

    df_ok, df_err, grupos = carregar_e_validar(str(p))
    if df_ok.empty or not grupos:
        print("\n[x] Nada a fazer (sem dados válidos). Verifique logs/erros_validacao.xlsx se houve erros.")
        return

    grupos_aj = _aplicar_precos_grupos(grupos, base_dir)
    arquivos = gerar_relatorios_xlsx(grupos_aj, output_dir=str(output_dir))
    print("\n[OK] Relatórios gerados:", *arquivos, sep="\n - ")


def acao_5_summary():
    p = _input_path("Informe o caminho do XLSX (aba 'Cadastro')")
    base_dir = _infer_base_dir(p)
    output_dir = base_dir / "output"

    df_ok, df_err, grupos = carregar_e_validar(str(p))
    if df_ok.empty or not grupos:
        print("\n[x] Nada a resumir (sem dados válidos).")
        return

    grupos_aj = _aplicar_precos_grupos(grupos, base_dir)
    caminho = gerar_resumo_consolidado(grupos_aj, output_dir=str(output_dir))
    print("\n[OK] Resumo gerado em:", caminho)


def acao_6_report_pdf():
    if not PDF_DISPONIVEL or not callable(gerar_relatorios_pdf):
        print("\n[x] PDF indisponível: instale 'reportlab' e confira src/reports_pdf.py.")
        return

    p = _input_path("Informe o caminho do XLSX (aba 'Cadastro')")
    base_dir = _infer_base_dir(p)
    pdf_dir = base_dir / "output" / "pdf"

    df_ok, df_err, grupos = carregar_e_validar(str(p))
    if df_ok.empty or not grupos:
        print("\n[x] Nada a fazer (sem dados válidos).")
        return

    grupos_aj = _aplicar_precos_grupos(grupos, base_dir)
    pdfs = gerar_relatorios_pdf(grupos_aj, output_dir=str(pdf_dir))
    print("\n[OK] PDFs gerados:", *pdfs, sep="\n - ")


# ---------- menu ----------

def mostrar_menu():
    print("\n===== RELATÓRIO DE EQUIPAMENTOS - MENU =====")
    print("1) Gerar estrutura + template XLSX")
    print("2) Validar arquivo de entrada")
    print("3) Processar tudo (XLSX + resumo + PDF opcional)")
    print("4) Gerar apenas relatórios XLSX por loja")
    print("5) Gerar apenas resumo consolidado")
    print("6) Gerar apenas PDFs por loja")
    print("7) Sair")


def main():
    while True:
        mostrar_menu()
        escolha = input("Escolha uma opção (1-7): ").strip()
        if escolha == "1":
            acao_1_template()
        elif escolha == "2":
            acao_2_validate()
        elif escolha == "3":
            acao_3_processar_completo()
        elif escolha == "4":
            acao_4_report_xlsx()
        elif escolha == "5":
            acao_5_summary()
        elif escolha == "6":
            acao_6_report_pdf()
        elif escolha == "7":
            print("Saindo... até mais!")
            break
        else:
            print("Opção inválida. Tente novamente.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nInterrompido pelo usuário.")
        sys.exit(0)
