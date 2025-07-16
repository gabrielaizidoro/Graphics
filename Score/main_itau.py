import pandas as pd
import time
from datetime import datetime
import platform
import getpass

# ============ CONFIGURAÇÃO ============
CAMINHO_INPUT = r"C:\Users\gabrizi\Banco Itaú SA\PARCERIAS_DE_SEGUROS - Documentos\06.ESTUDOS E ANALISES\..."
CAMINHO_OUTPUT = r"C:\Users\gabrizi\Banco Itaú SA\PARCERIAS_DE_SEGUROS - Documentos\06.ESTUDOS E ANALISES\..."

print(f"\n👩‍💻 Usuário: {getpass.getuser()}")
print(f"🖥️ Máquina: {platform.node()}")
print(f"🐍 Python : {platform.python_version()}")
print(f"🌞 Data   : {datetime.today().strftime('%Y-%m-%d %H:%M:%S')}")
print("\n" + "+"*50)

# ============ FUNÇÕES AUXILIARES ============
def nota_score(p):
    """Retorna a nota de 1 a 5 com base no percentual"""
    if p < 0.60: return 1
    elif 0.60 <= p < 0.64: return 2
    elif 0.64 <= p < 0.68: return 2.2
    elif 0.68 <= p < 0.72: return 2.4
    elif 0.72 <= p < 0.76: return 2.6
    elif 0.76 <= p < 0.80: return 2.8
    elif 0.80 <= p < 0.83: return 3
    elif 0.83 <= p < 0.86: return 3.2
    elif 0.86 <= p < 0.89: return 3.4
    elif 0.89 <= p < 0.92: return 3.6
    elif 0.92 <= p < 0.95: return 3.8
    elif 0.95 <= p < 0.958: return 4
    elif 0.958 <= p < 0.966: return 4.2
    elif 0.966 <= p < 0.974: return 4.4
    elif 0.974 <= p < 0.982: return 4.6
    elif 0.982 <= p < 0.99: return 4.8
    elif p >= 0.99: return 5
    else: return 0  # Caso o valor não se enquadre em nenhuma faixa

def formatar_data_completa(data_serie):
    """
    Formata uma série de datas para 'DD/MM/AAAA'.
    Se o dia for inválido, assume '01' como dia.
    """
    datas_dt = pd.to_datetime(data_serie, errors='coerce')
    datas_formatadas = []
    for dt in datas_dt:
        if pd.isna(dt):
            datas_formatadas.append(None)
        else:
            datas_formatadas.append(dt.strftime('%d/%m/%Y'))
    return pd.Series(datas_formatadas)

# ============ EXECUÇÃO ============
inicio = time.time()
print("\n📊 Carregando planilha de Referência...\n")

try:
    df_original = pd.read_excel(CAMINHO_INPUT, sheet_name="BASE GERAL")
    df = df_original.copy()
    print(f"✅ Planilha de Referência '{CAMINHO_INPUT}' carregada com sucesso.")
except FileNotFoundError:
    print(f"❌ ERRO: Arquivo de input '{CAMINHO_INPUT}' não encontrado. Verifique o caminho.")
    exit()
except Exception as e:
    print(f"❌ ERRO ao carregar o arquivo de input: {e}")
    exit()

# ============ AJUSTE DA COLUNA 'Data Referência' ============
df['Data Referencia'] = df['Data Referencia'].apply(
    lambda x: pd.to_datetime(x, errors='coerce')
).apply(
    lambda x: x.strftime('%d/%m/%Y') if pd.notna(x) else None
)
df['Data Referencia Formatada'] = df['Data Referencia']

# ============ CÁLCULOS DAS NOTAS EXISTENTES ============
print("\n✅ Realizando cálculos para as notas de Processamento, Conciliação e Acurácia...")

colunas_necessarias = [
    "Valor de Emissão Processado",
    "Valor de Emissão Não Processado",
    "Quantidade Processado",
    "Quantidade Não Processado",
    "Valor de Comissão (Processado - Líquido)",
    "Valor de Comissão (Pago)",
    "Fora do parametro"
]

for col in colunas_necessarias:
    if col not in df.columns:
        print(f"❌ ERRO: Coluna '{col}' não encontrada na base de dados. Verifique o nome das colunas.")
        exit()

df["%_processado_valor"] = df["Valor de Emissão Processado"].fillna(0) / \
    df["Valor de Emissão Não Processado"].fillna(0).replace(0, 1)

df["%_processado_qtd"] = df["Quantidade Processado"].fillna(0) / \
    df["Quantidade Não Processado"].fillna(0).replace(0, 1)

df["%_processamento_ponderado"] = 0.7 * df["%_processado_valor"].fillna(0) + 0.3 * df["%_processado_qtd"].fillna(0)
df["Nota_processamento"] = df["%_processamento_ponderado"].apply(nota_score)

df["%_conciliação"] = df["Valor de Comissão (Processado - Líquido)"].fillna(0) / \
    df["Valor de Comissão (Pago)"].fillna(0).replace(0, 1)
df["Nota_conciliação"] = df["%_conciliação"].apply(nota_score)

df["%_acuracia"] = 1 - (
    df["Fora do parametro"].fillna(0) /
    (df["Valor de Emissão Processado"].fillna(0) + df["Valor de Emissão Não Processado"].fillna(0)).replace(0, 1)
)
df["Nota_acuracia"] = df["%_acuracia"].apply(nota_score)

print("✅ Cálculos de notas de Processamento, Conciliação e Acurácia finalizados.")

# ============ NOVOS CÁLCULOS: NOTA MÉDIA E NOTA CONSOLIDADA ============
print("\n📊 Calculando Nota Média e Nota Consolidada...")

df["Nota_processamento"] = pd.to_numeric(df["Nota_processamento"], errors='coerce').fillna(0)
df["Nota_conciliação"] = pd.to_numeric(df["Nota_conciliação"], errors='coerce').fillna(0)
df["Nota_acuracia"] = pd.to_numeric(df["Nota_acuracia"], errors='coerce').fillna(0)

df["Nota_média"] = (df["Nota_processamento"] + df["Nota_conciliação"] + df["Nota_acuracia"]) / 3
df["Nota_média"] = df["Nota_média"].round(2)

df["Nota_consolidado"] = (
    df["Nota_processamento"] * 4 +
    df["Nota_conciliação"] * 4 +
    df["Nota_acuracia"] * 2
) / 10
df["Nota_consolidado"] = df["Nota_consolidado"].round(2)

print("✅ Cálculos de Nota Média e Nota Consolidada finalizados.")

# ============ CRIAÇÃO DA NOVA ABA: Média das Notas por Item ============
print("\n📈 Criando a aba 'Média das Notas por Item'...")

df_melted = df.melt(
    id_vars=['Seguradora', 'Frente', 'Data Referencia Formatada'],
    value_vars=['Nota_processamento', 'Nota_conciliação', 'Nota_acuracia'],
    var_name='Item da Nota',
    value_name='Valor da Nota'
)
df_melted['Data Referencia Formatada'] = pd.to_datetime(df_melted['Data Referencia Formatada'], format='%d/%m/%Y', errors='coerce')

df_pivot = pd.pivot_table(
    df_melted,
    index=['Seguradora', 'Frente', 'Item da Nota'],
    columns='Data Referencia Formatada',
    values='Valor da Nota',
    aggfunc='mean'
)

data_cols = [col for col in df_pivot.columns if isinstance(col, pd.Timestamp)]
data_cols_sorted = sorted(data_cols)
df_pivot = df_pivot[data_cols_sorted]
df_pivot.columns = [col.strftime('%d/%m/%Y') for col in df_pivot.columns]
df_pivot['Média'] = df_pivot.mean(axis=1).round(2)
df_pivot = df_pivot.round(2)

# ============ EXPORTAÇÃO ============
try:
    with pd.ExcelWriter(CAMINHO_OUTPUT, engine="openpyxl") as writer:
        df_original.to_excel(writer, sheet_name="Base Original", index=False)

        df_proc = df[["Seguradora", "Frente", "Data Referencia Formatada", "%_processamento_ponderado", "Nota_processamento"]]
        df_proc.to_excel(writer, sheet_name="Nota Processamento", index=False)

        df_conc = df[["Seguradora", "Frente", "Data Referencia Formatada", "%_conciliação", "Nota_conciliação"]]
        df_conc.to_excel(writer, sheet_name="Conciliação financeira", index=False)

        df_acur = df[["Seguradora", "Frente", "Data Referencia Formatada", "%_acuracia", "Nota_acuracia"]]
        df_acur.to_excel(writer, sheet_name="Contrato_Acurácia", index=False)

        df_media = df[["Seguradora", "Frente", "Data Referencia Formatada", "Nota_média"]]
        df_media.to_excel(writer, sheet_name="Nota Média", index=False)

        df_consolidado = df[["Seguradora", "Frente", "Data Referencia Formatada", "Nota_consolidado"]]
        df_consolidado.to_excel(writer, sheet_name="Nota Consolidado", index=False)

        df_pivot.to_excel(writer, sheet_name="Média das Notas por Item")

    print(f"✅ Todas as abas foram exportadas para '{CAMINHO_OUTPUT}'.")

except Exception as e:
    print(f"❌ ERRO ao exportar as abas do Excel: {e}")
    print("Certifique-se de que o arquivo não está aberto e que você tem permissão de escrita.")
    exit()

print("\n📁 Arquivo Excel salvo com sucesso ✅")

# ========== LOG FINAL ==========
print(f"""
------------------------------
⏱️ Tempo : {round(time.time() - inicio, 2)} segundos
📂 Input : {CAMINHO_INPUT}
📁 Output: {CAMINHO_OUTPUT}
------------------------------
""")
