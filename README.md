## 🧭 Visão Geral

## 🔗 Cruzamento de Planilhas Excel por MATRICULA (Pandas)

Script de data wrangling para cruzar dois arquivos Excel com base na coluna `MATRICULA`, utilizando `pandas.merge` com `indicator=True`. O resultado é salvo em um novo arquivo Excel com duas abas: `merge` (dados completos) e `summary` (contagem de correspondências).

### ✨ Destaques
- **Chave de junção flexível**: por padrão `MATRICULA`, mas você pode alterar via parâmetro `--on`.
- **Indicator do merge**: identifica se o registro veio do arquivo da esquerda, da direita ou de ambos (`left_only`, `right_only`, `both`).
- **Pós-processamento cuidadoso**: normaliza a coluna-chave como string, remove espaços e aplica `how=outer` para não perder registros.
- **Saída organizada**: Excel final com abas `merge` e `summary`.

### 🚀 Como usar
1) Instale as dependências:

```bash
pip install -r requirements.txt
```

2) Rode o script passando os arquivos Excel de entrada e o caminho de saída:

```bash
python merge_excel.py \
  --left caminho/arquivo_esquerda.xlsx \
  --right caminho/arquivo_direita.xlsx \
  --output resultados/merge_result.xlsx
```

Parâmetros opcionais:
- `--on` (padrão: `MATRICULA`): altera a coluna usada para o cruzamento.
- `--left-sheet` e `--right-sheet`: caso o arquivo tenha múltiplas abas e você precise especificar qual usar.

### 📦 Entrada e Saída
- **Entrada**: dois arquivos Excel (`.xlsx`) com a coluna de chave (por padrão, `MATRICULA`).
- **Saída**: um arquivo Excel contendo:
  - `merge`: resultado completo do `merge` com todas as colunas e a coluna `_merge`.
  - `summary`: contagem de `both`, `left_only` e `right_only` para rápida análise de cobertura.

### 🧠 O que o script faz (em alto nível)
- Lê os dois arquivos Excel (opcionalmente, sheets específicos)
- Padroniza a coluna-chave (string + trim)
- Executa `pd.merge(..., how="outer", indicator=True)`
- Salva o resultado em `merge_result.xlsx` (ou no caminho informado), com abas `merge` e `summary`

### 🗂️ Estrutura do repositório
```text
merge_excel.py      # Script principal (CLI) de cruzamento
README.md           # Este guia
requirements.txt    # Dependências (pandas, openpyxl)
.gitignore          # Itens comuns ignorados no Git
```

### 🛠️ Requisitos
- Python 3.9+
- `pandas` e `openpyxl`

### 🧪 Exemplo rápido
```bash
python merge_excel.py \
  --left dados/funcionarios_a.xlsx \
  --right dados/funcionarios_b.xlsx \
  --on MATRICULA \
  --output resultados/funcionarios_merge.xlsx
```

### 💡 Dicas para portfólio
- Adicione amostras de dados sintéticos na pasta `dados/` e um notebook com análise dos resultados de `summary`.
- Faça screenshots das abas `merge` e `summary` e inclua no README como visual.
- Explique decisões de engenharia (ex.: por que `how=outer`, tratamento da coluna de chave, sufixos de colunas).

---
Made with ❤️ for Data Portfolio

# teste_cursor2.1
