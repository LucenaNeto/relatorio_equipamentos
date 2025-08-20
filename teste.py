"""
from pathlib import Path
from src.processing import carregar_e_validar

# Considera que você está executando a partir da pasta do projeto "relatorio_equipamentos"
BASE = Path(__file__).resolve().parent
input_path = BASE / "input" / "cadastro_equipamentos.xlsx"

df_ok, df_err, grupos = carregar_e_validar(str(input_path))

print("Linhas válidas:", len(df_ok))
print("Erros:", len(df_err))
print("Lojas encontradas:", list(grupos.keys()))
"""
"""
from pathlib import Path
from src.processing import carregar_e_validar
from src.pricing import sugerir_precos, atualizar_historico

BASE = Path(__file__).resolve().parent
input_path = BASE / "input" / "cadastro_equipamentos.xlsx"

# 1) Lê e valida
df_ok, df_err, grupos = carregar_e_validar(str(input_path))

# 2) Calcula o preço sugerido (só para uma loja de exemplo)
loja = "3569"
df_loja = grupos[loja]
sug = sugerir_precos(df_loja, base_dir=str(BASE))
print(sug.head())

# 3) Atualiza histórico com preços reais informados no arquivo atual
qtde = atualizar_historico(df_ok, base_dir=str(BASE))
print("Registros adicionados ao histórico:", qtde)

"""
"""
# teste_pipeline.py
from pathlib import Path
from src.pipeline import processar_arquivo

BASE = Path(__file__).resolve().parent
input_path = BASE / "input" / "cadastro_equipamentos.xlsx"  # coloque o seu arquivo real aqui

resultado = processar_arquivo(str(input_path))
print(resultado)

# Saída esperada (exemplo):
# {
#   'status': 'ok',
#   'arquivos_lojas': ['...\output\relatorio_loja_3569.xlsx', ...],
#   'resumo_path': '...\output\resumo_por_loja.xlsx',
#   'erros_path': '...\logs\erros_validacao.xlsx'  # ou None
# }
"""
from pathlib import Path
from src.processing import carregar_e_validar
from src.pricing import aplicar_preco_sugerido
from src.reports_pdf import gerar_relatorios_pdf

BASE = Path(__file__).resolve().parent
input_path = BASE / "input" / "cadastro_equipamentos.xlsx"

# 1) Lê e valida
df_ok, df_err, grupos = carregar_e_validar(str(input_path))

# 2) Preenche o "Preço sugerido" onde estiver vazio (usa a engine do Bloco 6)
grupos_aj = {lj: aplicar_preco_sugerido(df, base_dir=str(BASE)) for lj, df in grupos.items()}

# 3) Gera PDFs por loja (em output/pdf)
pdfs = gerar_relatorios_pdf(grupos_aj, output_dir=str(BASE / "output" / "pdf"))
print("PDFs gerados:", pdfs)

