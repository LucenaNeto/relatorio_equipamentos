
"""
Módulo: generate_template.py
Cria a estrutura de pastas e um template Excel (com validações) para preenchimento.

Como usar:
    python -m src.generate_template
Isso vai criar (ou atualizar) o arquivo:
    relatorio_equipamentos/input/template_cadastro_equipamentos.xlsx
"""

import os
import pandas as pd
from src.settings import LOJAS_VALIDAS  # usa a lista oficial de lojas (Bloco 1)

def criar_template(base_dir: str = "relatorio_equipamentos") -> str:
    """
    Cria as pastas padrão e grava um template XLSX com 3 abas:
      - Cadastro: onde o usuário preenche os dados.
      - Listas: tabelas auxiliares (Lojas e Equipamentos) para dropdown.
      - Leia-me: instruções de uso.
    Retorna o caminho completo do template gerado.
    """
    # 1) Garante estrutura de pastas
    for d in ("input", "output", "logs", "config", "src"):
        os.makedirs(os.path.join(base_dir, d), exist_ok=True)

    template_path = os.path.join(base_dir, "input", "template_cadastro_equipamentos.xlsx")

    # 2) Define colunas da aba Cadastro
    cadastro_cols = ["Loja", "Equipamento", "Quantidade", "Preço sugerido", "Preço real"]

    # 3) Dados auxiliares para a aba Listas
    df_lojas = pd.DataFrame({"Loja": LOJAS_VALIDAS})
    df_equip = pd.DataFrame({
        "Equipamento": [
            "Notebook", "Impressora", "Monitor", "Roteador", "Nobreak"
        ]
    })

    # 4) Escreve o arquivo com XlsxWriter (permite validações e formatações)
    with pd.ExcelWriter(template_path, engine="xlsxwriter") as writer:
        # --- Aba Listas ---
        # Lojas em A1 (cabeçalho) e A2.. (valores)
        df_lojas.to_excel(writer, sheet_name="Listas", index=False, startrow=0, startcol=0)
        # Equipamentos em C1 (cabeçalho) e C2.. (valores)
        df_equip.to_excel(writer, sheet_name="Listas", index=False, startrow=0, startcol=2)

        # --- Aba Cadastro ---
        pd.DataFrame(columns=cadastro_cols).to_excel(writer, sheet_name="Cadastro", index=False)

        # --- Aba Leia-me ---
        instrucoes = [
            "Como usar este arquivo:",
            "1) Preencha a planilha 'Cadastro' uma linha por equipamento.",
            "2) 'Loja' e 'Equipamento' possuem listas auxiliares em 'Listas'.",
            "3) 'Quantidade' deve ser número inteiro >= 0.",
            "4) 'Preço sugerido' e 'Preço real' devem ser valores >= 0 (em R$).",
            "5) Salve o arquivo preenchido em /input e rode o processamento.",
        ]
        pd.DataFrame({"Instruções": instrucoes}).to_excel(writer, sheet_name="Leia-me", index=False)

        # ===== Formatações e validações =====
        wb        = writer.book
        ws_cad    = writer.sheets["Cadastro"]
        ws_listas = writer.sheets["Listas"]
        ws_help   = writer.sheets["Leia-me"]

        header_fmt = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1})
        money_fmt  = wb.add_format({"num_format": 'R$ #,##0.00'})
        int_fmt    = wb.add_format({"num_format": "0"})
        text_fmt   = wb.add_format({"text_wrap": True})

        # Cabeçalho da aba Cadastro com estilo
        for col, name in enumerate(cadastro_cols):
            ws_cad.write(0, col, name, header_fmt)

        # Largura e formatos de coluna
        ws_cad.set_column("A:A", 16)             # Loja
        ws_cad.set_column("B:B", 26)             # Equipamento
        ws_cad.set_column("C:C", 14, int_fmt)    # Quantidade (inteiro)
        ws_cad.set_column("D:E", 16, money_fmt)  # Preços (R$)

        # Congela o cabeçalho
        ws_cad.freeze_panes(1, 0)

        # Validações de dados (aplicadas até 10.000 linhas)
        max_rows = 10001

        # Intervalo dinâmico de lojas (aba Listas: A2 até A{1+len(lojas)})
        last_loja_row = 1 + len(df_lojas)  # linha 1 é cabeçalho; valores começam na 2
        dv_lojas_range = f"=Listas!$A$2:$A${last_loja_row}"

        # Intervalo de equipamentos (permitindo expansão futura até linha 1000)
        # Se quiser dinâmico exato: use len(df_equip) como fizemos com lojas.
        dv_equip_range = "=Listas!$C$2:$C$1000"

        # Loja: lista suspensa
        ws_cad.data_validation(1, 0, max_rows, 0, {
            "validate": "list",
            "source": dv_lojas_range
        })
        # Equipamento: lista suspensa
        ws_cad.data_validation(1, 1, max_rows, 1, {
            "validate": "list",
            "source": dv_equip_range
        })
        # Quantidade: inteiro >= 0
        ws_cad.data_validation(1, 2, max_rows, 2, {
            "validate": "integer",
            "criteria": ">=",
            "value": 0
        })
        # Preços: decimal >= 0 (aplica a D e E)
        ws_cad.data_validation(1, 3, max_rows, 4, {
            "validate": "decimal",
            "criteria": ">=",
            "value": 0
        })

        # Exemplo de linha preenchida (opcional — ajuda visual)
        exemplo = [["3569", "Notebook", 5, 3500.00, 3299.90]]
        for i, row in enumerate(exemplo, start=1):
            for j, val in enumerate(row):
                ws_cad.write(i, j, val)

        # Títulos da aba Listas e largura
        ws_listas.write(0, 0, "Loja", header_fmt)
        ws_listas.write(0, 2, "Equipamento", header_fmt)
        ws_listas.set_column("A:A", 16)
        ws_listas.set_column("C:C", 26)

        # Leia-me com largura confortável
        ws_help.set_column("A:A", 100, text_fmt)

    return template_path


if __name__ == "__main__":
    path = criar_template()
    print("Template criado em:", path)
