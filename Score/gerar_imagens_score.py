
# ============================================================
# Script: gerar_imagens_score.py
# Descrição: Geração de imagens de ranking com tabela lateral (notas por mês + média)
# Autor: Gabriela Izidoro
# ============================================================

import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from pathlib import Path
import os
import platform
import getpass
from datetime import datetime
import time

# ================================
# INFORMAÇÕES DE EXECUÇÃO
# ================================
inicio = time.time()
print("============================================================")
print(f"🧑 Usuário: {getpass.getuser()}")
print(f"💻 Máquina: {platform.node()}")
print(f"🐍 Python : {platform.python_version()}")
print(f"📅 Data   : {datetime.today().strftime('%Y-%m-%d %H:%M:%S')}")
print("============================================================")

# ================================
# CONFIGURAÇÕES
# ================================
CAMINHO_IMAGEM_BASE = r"CAMINHODAPASTA\imagem_modelo.png"
CAMINHO_PLANILHA = r"CAMINHODAPASTA\arquivo_output.xlsx"
PASTA_SAIDA = r"CAMINHODAPASTA\Output_Score_Imagens"
os.makedirs(PASTA_SAIDA, exist_ok=True)

fonte_padrao = str(Path("C:/Windows/Fonts/arialbd.ttf"))
fonte_tamanho = 26
cor_padrao = (0, 0, 0)
cor_colocacao = (255, 140, 0)

# ================================
# LEITURA DOS DADOS
# ================================
print("📥 Lendo dados do Excel...")
df = pd.read_excel(CAMINHO_PLANILHA, sheet_name="Consolidado")
df = df[["Seguradora", "MEDIA"]].dropna()
df["MEDIA"] = df["MEDIA"].astype(float)
df = df.sort_values(by="MEDIA", ascending=False).reset_index(drop=True)
df["CLASSIFICAÇÃO"] = df.index + 1

coordenadas_ranking = {i+1: (93, 272 + i * 80) for i in range(len(df))}
img_base = Image.open(CAMINHO_IMAGEM_BASE)


# ================================
# CONFIGURAÇÃO DO TÍTULO SUPERIOR
# ================================
MES_ANO_TEXTO = "06/25"
POSICAO_TEXTO = (70, 50)  # (x, y) no topo esquerdo
COR_TEXTO = (255, 255, 255)  # cor laranja, você pode trocar
TAMANHO_TEXTO = 36

# ================================
# FUNÇÃO: desenhar_imagem
# ================================
def desenhar_imagem(df, idx_destacado=None, nome_destacado=None):
    img = img_base.copy()
    draw = ImageDraw.Draw(img)
    font = ImageFont.truetype(fonte_padrao, fonte_tamanho)
    font_tabela = ImageFont.truetype(fonte_padrao, 20)

    # Fonte e texto do título superior (mês/ano)
    font_data = ImageFont.truetype(fonte_padrao, TAMANHO_TEXTO)
    draw.text(POSICAO_TEXTO, MES_ANO_TEXTO, fill=COR_TEXTO, font=font_data)


    for _, row in df.iterrows():
        rank = int(row["CLASSIFICAÇÃO"])
        nome = row["Seguradora"]
        media = row["MEDIA"]
        x, y = coordenadas_ranking[rank]

        draw.text((x + 50, y), f"{rank}", fill=cor_colocacao, font=font)
        espaco = 60
        nome_exibir = nome if (idx_destacado is None or nome == nome_destacado) else "*****"
        draw.text((x + 50 + espaco, y), f"{nome_exibir} ({media:.2f})", fill=cor_padrao, font=font)

    if idx_destacado is not None:
        try:
            df_notas = pd.read_excel(CAMINHO_PLANILHA, sheet_name="Planilha 1")
            df_filtrada = df_notas[df_notas["Seguradora"] == nome_destacado]

            # Considera colunas que representam datas de 2025, mesmo se forem strings
            colunas_dinamicas = []
            for col in df_notas.columns:
                try:
                    data = pd.to_datetime(col, dayfirst=True, errors='raise')
                    if data.year == 2025:
                        colunas_dinamicas.append(col)
                except:
                    continue

            colunas_dinamicas = sorted(colunas_dinamicas, key=lambda x: pd.to_datetime(x, dayfirst=True))


            col_map = {
                "Nota_acuracia": "Acurácia",
                "Nota_conciliacao": "Conciliação",
                "Nota_processamento": "Processamento"
            }

            dados = []
            for cod, nome_exibir in col_map.items():
                linha = df_filtrada[df_filtrada["Item da Nota"] == cod]
                if not linha.empty:
                    notas = [linha[col].values[0] for col in colunas_dinamicas if col in linha.columns]
                    media_val = linha["MEDIA"].values[0] if "MEDIA" in linha.columns else ""
                    dados.append((nome_exibir, notas, media_val))

            
            # Desenhar tabela
            x0, y0 = 1000, 600
            col1_larg = 160
            col_w = 100
            row_h = 50

            # Cabeçalho
            draw.rectangle([x0, y0, x0 + col1_larg, y0 + row_h], outline=cor_padrao)
            draw.text((x0 + 10, y0 + 10), "Item da Nota", fill=cor_padrao, font=font_tabela)

            for j, col in enumerate(colunas_dinamicas):
                draw.rectangle([x0 + col1_larg + j * col_w, y0, x0 + col1_larg + (j + 1) * col_w, y0 + row_h], outline=cor_padrao)
                data_label = pd.to_datetime(str(col), dayfirst=True).strftime("%d/%Y")
                draw.text((x0 + col1_larg + j * col_w + col_w // 2, y0 + row_h // 2), data_label, fill=cor_padrao, font=font_tabela, anchor="mm")


            draw.rectangle([x0 + col1_larg + len(colunas_dinamicas) * col_w, y0, x0 + col1_larg + (len(colunas_dinamicas)+1) * col_w, y0 + row_h], outline=cor_padrao)
            draw.text((x0 + col1_larg + len(colunas_dinamicas) * col_w + 10, y0 + 10), "Média", fill=cor_padrao, font=font)

            # Linhas de dados
            for i, (item_nome, valores, media) in enumerate(dados):
                y = y0 + (i + 1) * row_h
                draw.rectangle([x0, y, x0 + col1_larg, y + row_h], outline=cor_padrao)
                draw.text((x0 + 10, y + 10), item_nome, fill=cor_padrao, font=font_tabela)

                for j, val in enumerate(valores):
                    draw.rectangle([x0 + col1_larg + j * col_w, y, x0 + col1_larg + (j + 1) * col_w, y + row_h], outline=cor_padrao)
                    draw.text((x0 + col1_larg + j * col_w + 20, y + 10), f"{val:.0f}", fill=cor_padrao, font=font)

                draw.rectangle([x0 + col1_larg + len(colunas_dinamicas) * col_w, y, x0 + col1_larg + (len(colunas_dinamicas)+1) * col_w, y + row_h], outline=cor_padrao)
                draw.text((x0 + col1_larg + len(colunas_dinamicas) * col_w + 10, y + 10), f"{media:.2f}", fill=cor_padrao, font=font)

        except Exception as e:
            print(f"⚠️  Erro ao desenhar tabela de '{nome_destacado}': {e}")

    return img

# ================================
# GERAÇÃO DAS IMAGENS
# ================================
print("🖼️  Gerando imagem geral...")
img_total = desenhar_imagem(df)
img_total.save(os.path.join(PASTA_SAIDA, "00_Score_Geral.png"))

print("📊 Gerando imagens por seguradora...")
for i, row in df.iterrows():
    nome = row["Seguradora"]
    print(f"   ▶ {i+1:02d} - {nome}")
    img_individual = desenhar_imagem(df, idx_destacado=i, nome_destacado=nome)
    path_ind = os.path.join(PASTA_SAIDA, f"{i+1:02d}_Score_{nome}.png")
    img_individual.save(path_ind)

# ================================
# FINALIZAÇÃO
# ================================
fim = time.time()
print("✅ Processo finalizado com sucesso!")
print(f"🕒 Tempo de execução: {round(fim - inicio, 2)} segundos")
print(f"📂 Imagens salvas em: {PASTA_SAIDA}")
print("============================================================")
