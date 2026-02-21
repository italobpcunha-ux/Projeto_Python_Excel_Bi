import os
import pandas as pd

print("Diretório atual:", os.getcwd())
print("Arquivos visíveis:", os.listdir())

# Testa se o arquivo está sendo encontrado
base_path = r"C:\Users\italo\Desktop\Mini_Projeto_python_Excel_BI\databases\Projeto_Python_Excel_Bi"
print(os.listdir(base_path))

# Ler Excel
df = pd.read_excel(
    r"C:\Users\italo\Desktop\Mini_Projeto_python_Excel_BI\databases\Projeto_Python_Excel_Bi\vendas_loja.xlsx",
    sheet_name="Vendas_Brutas"
)

# Criar coluna de valor total
df["Valor_Total"] = df["Quantidade"] * df["Preco_Unitario"]

# Converter data
df["Data"] = pd.to_datetime(df["Data"])

# Criar Ano e Mês
df["Ano"] = df["Data"].dt.year
df["Mes"] = df["Data"].dt.month
df["Nome_Mes"] = df["Data"].dt.month_name()

# Salvar novo arquivo
# Caminho da pasta onde está o script
base_path = os.path.dirname(__file__)

# Caminho completo do arquivo de saída
output_path = os.path.join(base_path, "vendas_tratadas.xlsx")

# Salvar
df.to_excel(output_path, index=False)

print("Arquivo salvo em:", output_path)

print("Arquivo tratado criado com sucesso!")