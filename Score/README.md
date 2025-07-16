# 📊 Análise e Visualização de Score de Desempenho

Este projeto realiza o **cálculo mensal de desempenho**, com base em indicadores operacionais, e gera **visualizações gráficas automáticas** a partir das notas. Ele está dividido em duas etapas principais:

1. **Cálculo das Notas** com base em regras definidas de desempenho
2. **Geração de Imagens** com ranking, notas e quadros ilustrativos

---

## 📈 Indicadores Calculados

As notas de cada entidade são baseadas em três pilares principais:

1. ✅ **Processamento**  
2. 💰 **Conciliação Financeira**  
3. 📄 **Acurácia / Contrato**

Cada métrica recebe uma **nota entre 1 e 5**, conforme a performance relativa.

---

## 🧮 Regras de Cálculo (Resumo)

### 🎯 Processamento (ponderado)
- % processado (valor e quantidade)
- Cálculo ponderado: `0.7 * valor + 0.3 * quantidade`

### 💰 Conciliação
- % conciliação entre valores processados e pagos

### 📄 Acurácia contratual
- Percentual de registros dentro do parâmetro definido

---

## 🏆 Tabela de Notas por Faixa

| Faixa (%)        | Nota | Cor       |
|------------------|------|-----------|
| Abaixo de 60%    | 1    | 🔴 Vermelho claro |
| 60% a 79%        | 2    | 🟠 Laranja claro |
| 80% a 94%        | 3    | 🟡 Amarelo claro |
| 95% a 98%        | 4    | 🟢 Verde claro |
| 99% a 100%       | 5    | ✅ Verde escuro |

---

## 🖼️ Geração de Imagens

Com base nos resultados de notas, o script gera automaticamente imagens contendo:

- 📊 **Ranking Geral** com classificação de todas as entidades
- 🔍 **Ranking Individual** para cada entidade, ocultando as demais (por privacidade)
- 🧾 **Quadro lateral** com as notas mensais e médias

As imagens são personalizadas com:
- Título superior com mês/ano (editável)
- Cores e fontes ajustáveis
- Tabela lateral formatada e centralizada

---

## 📂 Estrutura de Pastas
📦Score
 ┣ 📁 venv/                     # Ambiente virtual Python
 ┣ 📁 Input/                    # Planilhas base e imagem modelo
 ┣ 📁 Output/                   # Arquivo Excel e imagens finais
 ┣ 📄 main.py                  # Cálculo das notas
 ┣ 📄 gerar_imagens_score.py   # Geração automática das imagens
 ┣ 📄 README.md                # Este arquivo
 ┗ 📄 requirements.txt         # Bibliotecas utilizadas

---

## ▶️ Como Executar

### 1. Gerar Notas (Excel)
python main.py

### 2. Gerar Imagens com Ranking e Notas
python gerar_imagens_score.py

As imagens serão salvas na pasta `/Output`, contendo:

- Imagem com o ranking geral
- Imagem individual por entidade com quadro de notas

---

## ✅ Requisitos

- Python 3.9 ou superior
- Bibliotecas:
  - pandas
  - openpyxl
  - Pillow

Instale com:
pip install -r requirements.txt

---

## 👩‍💻 Autora

**Gabriela Izidoro**  
Automação • Dados • Processos Contábeis  
🔗 https://github.com/gabrielaizidoro
