# -*- coding: utf-8 -*-
"""
📊 Score - Processamento, Conciliação e Acurácia
Autor: Gabriela Izidoro
Versão: 1.0
Descrição: Este script lê um Excel de base de comissões, calcula indicadores e notas conforme critérios definidos,
cria abas com notas de processamento, conciliação e acurácia, além de notas média e consolidada,
e uma nova aba com a média das notas por item em formato de tabela dinâmica,
e exporta o resultado em múltiplas abas separadas.
"""

import pandas as pd
import time
from datetime import datetime
import platform
import getpass

# =================== CONFIGURAÇÃO ===================
CAMINHO_INPUT = r"CAMINHODASUAPASTA\NOMEDOARQUIVO.xlsx"
CAMINHO_OUTPUT = r"CAMINHODASUAPASTA\NOMEDOARQUIVO.xlsx"

# ================ FUNÇÕES AUXILIARES ================
def nota_score(p):
    """Retorna a nota de 1 a 5 com base no percentual"""
    if p < 0.60:
        return 1
    elif p < 0.80:
        return 2
    elif p < 0.95:
        return 3
    elif p < 0.99:
        return 4
    else:
        return 5

# Função para formatar a data como DD/MM/AAAA, tratando datas incompletas/inválidas
def formatar_data_completa(data_serie):
    """
    Formata uma série de datas para 'DD/MM/AAAA'.
    Se o dia for inválido, assume '01' como dia.
    """
    # Converte para datetime, erros viram NaT
    datas_dt = pd.to_datetime(data_serie, errors='coerce')

    # Lista para armazenar as datas formatadas
    datas_formatadas = []

    for dt in datas_dt:
        if pd.isna(dt):
            # Apenas formatamos o que pd.to_datetime conseguiu converter.
            datas_formatadas.append(None)
        else:
            # Formata para DD/MM/AAAA
            datas_formatadas.append(dt.strftime('%d/%m/%Y'))
            
    return pd.Series(datas_formatadas)


# =================== EXECUÇÃO ===================
inicio = time.time()
print("\n🔍 Carregando planilha de referência...\n")

# Carregar base original
try:
    df_original = pd.read_excel(CAMINHO_INPUT)
    df = df_original.copy()
    print(f"✅ Planilha de referência '{CAMINHO_INPUT}' carregada com sucesso.")
except FileNotFoundError:
    print(f"❌ ERRO: Arquivo de input '{CAMINHO_INPUT}' não encontrado. Verifique o caminho.")
    exit()
except Exception as e:
    print(f"❌ ERRO ao carregar o arquivo de input: {e}")
    exit()

# ============ AJUSTE DA COLUNA 'Data Referência' ============
print("\n🔄 Ajustando o formato da coluna 'Data Referência' para DD/MM/AAAA, com dia 01 se ausente...")

# A formatação '%d/%m/%Y' então garante '01/10/2025'.
df['Data Referência'] = df['Data Referência'].apply(
    lambda x: pd.to_datetime(x, errors='coerce')
).apply(
    lambda x: x.strftime('%d/%m/%Y') if pd.notna(x) else None # Usar None ou '' para NaT
)
# Renomeamos para que o nome original da coluna seja usado nas exportações futuras.
df['Data Referência Formatada'] = df['Data Referência']

print("✅ Formato da coluna 'Data Referência' ajustado para DD/MM/AAAA.")


# ============ CÁLCULOS DAS NOTAS EXISTENTES ============
print("\n🔢 Realizando cálculos para as notas de Processamento, Conciliação e Acurácia...")
# Garante que as colunas existam antes de calcular para evitar erros
colunas_necessarias = [
    "Valor Emissão Processado (Correto)",
    "Valor Emissão Não Processado (Erro)",
    "Qtd Processado (Correto)",
    "Qtd Não Processado (Erro)",
    "Valor Comissão processado",
    "Valor Comissão pago",
    "Fora do parametro"
]
for col in colunas_necessarias:
    if col not in df.columns:
        print(f"❌ ERRO: Coluna '{col}' não encontrada na base de dados. Verifique o nome das colunas.")
        exit()

# Calcula nota do Processamento e substitui valores infinitos ou NaN resultantes de divisões por zero ou dados ausentes
# por 0 ou um valor apropriado para não quebrar os cálculos de nota.
df["%_processado_valor"] = df["Valor Emissão Processado (Correto)"].fillna(0) / (
    df["Valor Emissão Processado (Correto)"].fillna(0) + df["Valor Emissão Não Processado (Erro)"].fillna(0)
).replace(0, 1) # Evita divisão por zero, substituindo 0 por 1 no denominador

df["%_processado_qtd"] = df["Qtd Processado (Correto)"].fillna(0) / (
    df["Qtd Processado (Correto)"].fillna(0) + df["Qtd Não Processado (Erro)"].fillna(0)
).replace(0, 1)

df["%_processamento_ponderado"] = 0.7 * df["%_processado_valor"].fillna(0) + 0.3 * df["%_processado_qtd"].fillna(0)
# Nota de Processamento: converte percentual em nota de 1 a 5
df["Nota_processamento"] = df["%_processamento_ponderado"].apply(nota_score)

# Calcula nota da Conciliação 
df["%_conciliação"] = df["Valor Comissão processado"].fillna(0) / df["Valor Comissão pago"].fillna(0).replace(0, 1)
df["Nota_conciliacao"] = df["%_conciliação"].apply(nota_score)

# Calcula nota da Acurácia 
# Acurácia é calculada como 1 - (Fora do parâmetro / (Valor Emissão Processado + Valor Emissão Não Processado))
df["%_acuracia"] = 1 - (
    df["Fora do parametro"].fillna(0) / (
        df["Valor Emissão Processado (Correto)"].fillna(0) + df["Valor Emissão Não Processado (Erro)"].fillna(0)
    ).replace(0, 1)
)
df["Nota_acuracia"] = df["%_acuracia"].apply(nota_score)

print("✅ Cálculos de notas de Processamento, Conciliação e Acurácia finalizados.")

# ============ NOVOS CÁLCULOS: NOTA MÉDIA E NOTA CONSOLIDADA ============
print("\n🔢 Calculando Nota Média e Nota Consolidado...")

# Certifique-se de que as colunas de nota são numéricas para o cálculo
df["Nota_processamento"] = pd.to_numeric(df["Nota_processamento"], errors='coerce').fillna(0)
df["Nota_conciliacao"] = pd.to_numeric(df["Nota_conciliacao"], errors='coerce').fillna(0)
df["Nota_acuracia"] = pd.to_numeric(df["Nota_acuracia"], errors='coerce').fillna(0)

# Nota Média: Média das 3 notas
df["Nota_media"] = (df["Nota_processamento"] + df["Nota_conciliacao"] + df["Nota_acuracia"]) / 3
# Arredondar para duas casas decimais
df["Nota_media"] = df["Nota_media"].round(2) 

# Nota Consolidado com PESO: (Nota_processamento * 4) + (Nota_conciliacao * 4) + (Nota_acuracia * 2) / 10
df["Nota_consolidado"] = (
    (df["Nota_processamento"] * 4) +
    (df["Nota_conciliacao"] * 4) +
    (df["Nota_acuracia"] * 2)
) / 10
df["Nota_consolidado"] = df["Nota_consolidado"].round(2) # Arredondar também para a consolidada

print("✅ Cálculos de Nota Média e Nota Consolidado finalizados.")

# ============ CRIAÇÃO DA NOVA ABA: Média das Notas por Item ============
print("\n📊 Criando a aba 'Média das Notas por Item'...")

# 'Data Referência Formatada' para garantir o formato DD/MM/AAAA
df_melted = df.melt(
    id_vars=['Seguradora', 'Data Referência Formatada'],
    value_vars=['Nota_processamento', 'Nota_conciliacao', 'Nota_acuracia'],
    var_name='Item da Nota',
    value_name='Valor da Nota'
)

# Criar a tabela dinâmica
df_pivot = pd.pivot_table(
    df_melted,
    index=['Seguradora', 'Item da Nota'],
    columns='Data Referência Formatada',
    values='Valor da Nota',
    aggfunc='mean' # Calcula a média dos valores para cada combinação
)

# Adicionar a coluna 'Média' no final
df_pivot['Média'] = df_pivot.mean(axis=1).round(2) # Arredonda a média para 2 casas decimais

# Arredondar os valores (para as colunas de data)
df_pivot = df_pivot.round(2)

print("✅ Aba 'Média das Notas por Item' criada.")

# ========== EXPORTAÇÃO ==========
print("\n📝 Exportando todas as abas para o arquivo Excel de saída...")
try:
    with pd.ExcelWriter(CAMINHO_OUTPUT, engine="openpyxl") as writer:
        df_original.to_excel(writer, sheet_name="Base Original", index=False)
        
        # Usar 'Data Referência Formatada' nas abas para o novo formato de data
        df_proc = df[["Seguradora", "Data Referência Formatada", "%_processamento_ponderado", "Nota_processamento"]]
        df_proc.to_excel(writer, sheet_name="Processamento", index=False)
        
        df_conc = df[["Seguradora", "Data Referência Formatada", "%_conciliação", "Nota_conciliacao"]]
        df_conc.to_excel(writer, sheet_name="Conciliação financeira", index=False)
        
        df_acur = df[["Seguradora", "Data Referência Formatada", "%_acuracia", "Nota_acuracia"]]
        df_acur.to_excel(writer, sheet_name="Contrato_Acurácia", index=False)

        # Novas abas
        df_media = df[["Seguradora", "Data Referência Formatada", "Nota_media"]]
        df_media.to_excel(writer, sheet_name="Nota Media", index=False)

        df_consolidado = df[["Seguradora", "Data Referência Formatada", "Nota_consolidado"]]
        df_consolidado.to_excel(writer, sheet_name="Nota Consolidado", index=False)

        # Exportar a nova aba
        df_pivot.to_excel(writer, sheet_name="Média das Notas por Item")
        
    print(f"✅ Todas as abas foram exportadas para '{CAMINHO_OUTPUT}'.")
except Exception as e:
    print(f"❌ ERRO ao exportar as abas do Excel: {e}")
    print("Certifique-se de que o arquivo não está aberto e que você tem permissão de escrita.")
    exit()

print("\n✅ Arquivo Excel salvo com sucesso.")

# ========== LOG FINAL ==========
print(f"""
---
🧑 Usuário: {getpass.getuser()}
💻 Máquina: {platform.node()}
🐍 Python : {platform.python_version()}
📅 Data   : {datetime.today().strftime('%Y-%m-%d %H:%M:%S')}
⏱️ Tempo   : {round(time.time() - inicio, 2)} segundos

📥 Input  : {CAMINHO_INPUT}
📤 Output : {CAMINHO_OUTPUT}
---
""")