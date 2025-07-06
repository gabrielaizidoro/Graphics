# 📊 Análise de Score de DesempenhoAnálise de Score de Desempenho

Este projeto calcula o **score de desempenho mensal** , com base em três pilares:

1. ✅ **Processamento** (emissão de apólices e parcelas)
2. 💰 **Conciliação financeira** (comissão processada vs. paga)
3. 📄 **Contrato / Acurácia** (emissões fora de parâmetro)

Cada dimensão recebe uma **nota de 1 a 5**, com base em regras pré-estabelecidas de percentual. As notas são salvas em planilhas separadas com **cores visuais** para facilitar a leitura.

---

## 🧮 Lógica de Cálculo

### 🎯 Processamento (ponderado)
- % processado valor = valor emissão processado / total emissão
- % processado quantidade = qtd emissão processada / total qtd
- % ponderado = 0.7 * valor + 0.3 * quantidade

### 💰 Conciliação financeira
- % = comissão processada / comissão paga

### 📄 Acurácia contratual
- % = 1 - (fora do parâmetro / total valor emissão)

---

## 🏆 Tabela de Notas

| Faixa (%)        | Nota | Cor       |
|------------------|------|-----------|
| Abaixo de 60%    | 1    | 🔴 Vermelho claro |
| 60% a 79%        | 2    | 🟠 Laranja claro |
| 80% a 94%        | 3    | 🟡 Amarelo claro |
| 95% a 98%        | 4    | 🟢 Verde claro |
| 99% a 100%       | 5    | ✅ Verde escuro |

---

## 📂 Estrutura de Pastas
```
📦Score
 ┣ 📁 venv/                    # Ambiente virtual Python
 ┣ 📁 Input/                   # Planilhas base com os dados
 ┣ 📁 Output/                  # Arquivo Excel final com scores
 ┣ 📄 main.py                  # Script principal de cálculo
 ┣ 📄 README.md                # Este arquivo
 ┗ 📄 requirements.txt         # Bibliotecas usadas
```
---

## ▶️ Como usar

1. Coloque a planilha base em /Input
2. Execute o script main.py:
   python main.py
3. O resultado será salvo na pasta /Output, com:
   - Aba “Base Original” (sem alterações)
   - Aba “Processamento”
   - Aba “Conciliação financeira”
   - Aba “Contrato_Acurácia”

---

## 📌 Observações
- As cores são aplicadas diretamente nas células de nota.
- A planilha original não é modificada.
- Tipos de dados são mantidos (datas, strings, numéricos).

---

## ✅ Requisitos

- Python 3.9 ou superior
- Bibliotecas:
  - pandas
  - openpyxl

Instale via:
pip install -r requirements.txt

---

## 👩‍💻 Autor


**Gabriela Izidoro**  
Automação • Dados • Processos Contábeis  
[github.com/gabrielaizidoro](https://github.com/gabrielaizidoro)
