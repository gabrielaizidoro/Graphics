import pandas as pd
import time
from datetime import datetime
import platform
import getpass

# ============ CONFIGURAÇÃO ============
CAMINHO_INPUT = r"C:\Users\gabrizi\Banco Itaú SA\PARCERIAS_DE_SEGUROS - Documentos\06.ESTUDOS E ANALISES\..."
CAMINHO_NOTAS_MANUAIS = r"C:\Users\gabrizi\Banco Itaú SA\PARCERIAS_DE_SEGUROS - Documentos\06.ESTUDOS E ANALISES\Notas_Manuais.xlsx"
CAMINHO_OUTPUT = r"C:\Users\gabrizi\Banco Itaú SA\PARCERIAS_DE_SEGUROS - Documentos\06.ESTUDOS E ANALISES\..."

print(f"\n🧑‍💻 Usuário: {getpass.getuser()}")
print(f"📰 Máquina: {platform.node()}")
print(f"🐇 Python : {platform.python_version()}")
print(f"🌞 Data   : {datetime.today().strftime('%Y-%m-%d %H:%M:%S')}")
print("\n" + "+"*50)

# ============ FUNÇÕES AUXILIARES ============
def nota_score(p):
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
    else: return 0

# ============ EXECUÇÃO ============
inicio = time.time()
print("\n📊 Carregando planilha de Referência...\n")

try:
    df_original = pd.read_excel(CAMINHO_INPUT, sheet_name="BASE GERAL")
    df = df_original.copy()
    print(f"✅ Planilha de Referência '{CAMINHO_INPUT}' carregada com sucesso.")
except Exception as e:
    print(f"❌ ERRO: {e}")
    exit()

df['Data Referencia'] = pd.to_datetime(df['Data Referencia'], errors='coerce')
df['Data Referencia Formatada'] = df['Data Referencia'].dt.strftime('%d/%m/%Y')

# ============ CÁLCULOS EXISTENTES ============
print("\n✅ Realizando cálculos de notas automáticas para mar/25 em diante...")
df['Data Referencia Formatada'] = pd.to_datetime(df['Data Referencia Formatada'], format='%d/%m/%Y', errors='coerce')
df_calculo = df[df['Data Referencia Formatada'] >= pd.Timestamp('2025-03-01')].copy()

# PROCESSAMENTO
df_calculo['%_processado_valor'] = df_calculo['Valor de Emissão Processado'].fillna(0) / df_calculo['Valor de Emissão Não Processado'].fillna(0).replace(0, 1)
df_calculo['%_processado_qtd'] = df_calculo['Quantidade Processado'].fillna(0) / df_calculo['Quantidade Não Processado'].fillna(0).replace(0, 1)
df_calculo['%_processamento_ponderado'] = 0.7 * df_calculo['%_processado_valor'] + 0.3 * df_calculo['%_processado_qtd']
df_calculo['Nota_processamento'] = df_calculo['%_processamento_ponderado'].apply(nota_score)

# CONCILIAÇÃO COL
df_col = df_calculo[df_calculo['Frente'].str.upper() == 'COL'].copy()
df_col['%_conciliação'] = df_col['Valor de Comissão (Processado - Líquido)'].fillna(0) / df_col['Valor de Comissão (Pago)'].fillna(0).replace(0, 1)
df_col['Nota_conciliacao_col'] = df_col['%_conciliação'].apply(nota_score)

# CONCILIAÇÃO NPC
df_npc = df_calculo[df_calculo['Frente'].str.upper() == 'NPC'].copy()
df_npc['%_conciliação'] = df_npc['Valor de Comissão (Pago)'].fillna(0) / df_npc['Valor de Comissão (Processado - Líquido)'].fillna(0).replace(0, 1)
df_npc['Nota_conciliacao_npc'] = df_npc['%_conciliação'].apply(nota_score)

# ACURÁCIA
for df_sub in [df_col, df_npc]:
    df_sub['%_acuracia'] = 1 - (
        df_sub['Fora do parametro'].fillna(0) /
        (df_sub['Valor de Emissão Processado'].fillna(0) + df_sub['Valor de Emissão Não Processado'].fillna(0)).replace(0, 1)
    )
    df_sub['Nota_acuracia'] = df_sub['%_acuracia'].apply(nota_score)

# Unificar os dois frames novamente
df_calculo = pd.concat([df_col, df_npc], ignore_index=True)

# ============ CARREGAR NOTAS MANUAIS ============
print("\n📥 Lendo notas manuais de janeiro e fevereiro de 2025...")
df_manual_raw = pd.read_excel(CAMINHO_NOTAS_MANUAIS, sheet_name="Planilha 1", header=None)
df_manual_raw.columns = df_manual_raw.iloc[0]
df_manual = df_manual_raw[1:]
df_manual = df_manual.melt(
    id_vars=['Seguradora', 'Frente', 'Item da Nota'],
    var_name='Data Referencia Formatada',
    value_name='Nota'
)
df_manual['Data Referencia Formatada'] = pd.to_datetime(df_manual['Data Referencia Formatada'], dayfirst=True, errors='coerce')
df_manual['Nota'] = pd.to_numeric(df_manual['Nota'], errors='coerce')
df_manual['Item da Nota'] = df_manual.apply(
    lambda x: f"Nota_conciliacao_{x['Frente'].lower()}" if x['Item da Nota'] == 'Nota_conciliacao' else x['Item da Nota'],
    axis=1
)

# ============ UNIR NOTAS MANUAIS E AUTOMÁTICAS ============
df_melt_auto = df_calculo.melt(
    id_vars=['Seguradora', 'Frente', 'Data Referencia Formatada'],
    value_vars=['Nota_processamento', 'Nota_acuracia', 'Nota_conciliacao_col', 'Nota_conciliacao_npc'],
    var_name='Item da Nota',
    value_name='Nota'
)
df_notas_final = pd.concat([df_manual, df_melt_auto], ignore_index=True)

# ============ CRIAR PIVOT E EXPORTAR ============
print("\n📈 Criando a aba 'Média das Notas por Item'...")
df_pivot = pd.pivot_table(
    df_notas_final,
    index=['Seguradora', 'Frente', 'Item da Nota'],
    columns='Data Referencia Formatada',
    values='Nota',
    aggfunc='mean'
)

data_cols = sorted([col for col in df_pivot.columns if isinstance(col, pd.Timestamp)])
df_pivot = df_pivot[data_cols]
df_pivot.columns = [col.strftime('%d/%m/%Y') for col in df_pivot.columns]
df_pivot['Média'] = df_pivot.mean(axis=1).round(2)

with pd.ExcelWriter(CAMINHO_OUTPUT, engine="openpyxl") as writer:
    df_original.to_excel(writer, sheet_name="Base Original", index=False)
    df_pivot.to_excel(writer, sheet_name="Média das Notas por Item")

print("\n✅ Exportação concluída com sucesso!")
print(f"\n⏱️ Tempo total: {round(time.time() - inicio, 2)} segundos")
