#!/usr/bin/env python3
"""
Merge (cruzamento) de dois arquivos Excel pela coluna 'MATRICULA',
em estilo procedural (sem definição de funções).

Fluxo:
- Leitura dos dois arquivos Excel
- Padronização da coluna-chave (string + trim)
- merge outer com indicator=True
- Geração de arquivo Excel com abas 'merge' e 'summary'

Exemplo:
    python merge_excel.py \
        --left dados/funcionarios_1.xlsx \
        --right dados/funcionarios_2.xlsx \
        --output resultados/merge_result.xlsx
"""

import argparse
from pathlib import Path

import pandas as pd


if __name__ == "__main__":
    # Parser de argumentos (sem funções auxiliares)
    parser = argparse.ArgumentParser(
        description="Cruzamento de dois arquivos Excel pela coluna 'MATRICULA' com indicator=True (estilo procedural)"
    )
    parser.add_argument("--left", required=True, help="Caminho do Excel da esquerda")
    parser.add_argument("--right", required=True, help="Caminho do Excel da direita")
    parser.add_argument(
        "--output",
        required=False,
        default="merge_result.xlsx",
        help="Caminho do arquivo Excel de saída (padrão: merge_result.xlsx)",
    )
    parser.add_argument(
        "--on",
        required=False,
        default="MATRICULA",
        help="Nome da coluna-chave para merge (padrão: MATRICULA)",
    )
    parser.add_argument(
        "--left-sheet",
        required=False,
        default=None,
        help="Nome da planilha (sheet) no arquivo da esquerda (opcional)",
    )
    parser.add_argument(
        "--right-sheet",
        required=False,
        default=None,
        help="Nome da planilha (sheet) no arquivo da direita (opcional)",
    )
    args = parser.parse_args()

    # Caminhos e validações básicas
    left_path = Path(args.left)
    right_path = Path(args.right)
    output_path = Path(args.output)
    key_column = str(args.on)

    if not left_path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {left_path}")
    if not right_path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {right_path}")

    # Leitura dos arquivos Excel
    # (lê sheet específico se informado)
    if args.left_sheet:
        left_df = pd.read_excel(left_path, sheet_name=args.left_sheet)
    else:
        left_df = pd.read_excel(left_path)

    if args.right_sheet:
        right_df = pd.read_excel(right_path, sheet_name=args.right_sheet)
    else:
        right_df = pd.read_excel(right_path)

    # Padroniza cabeçalhos (trim nos nomes das colunas)
    left_df.columns = [str(col).strip() for col in left_df.columns]
    right_df.columns = [str(col).strip() for col in right_df.columns]

    # Garante a existência da coluna-chave
    if key_column not in left_df.columns:
        available = ", ".join(map(str, left_df.columns))
        raise KeyError(
            f"Coluna-chave '{key_column}' não encontrada no arquivo da esquerda. Colunas disponíveis: {available}"
        )
    if key_column not in right_df.columns:
        available = ", ".join(map(str, right_df.columns))
        raise KeyError(
            f"Coluna-chave '{key_column}' não encontrada no arquivo da direita. Colunas disponíveis: {available}"
        )

    # Normaliza coluna-chave como string e remove espaços
    left_df[key_column] = left_df[key_column].astype(str).str.strip()
    right_df[key_column] = right_df[key_column].astype(str).str.strip()

    # Merge (outer) com indicator=True
    merged_df = pd.merge(
        left_df,
        right_df,
        on=key_column,
        how="outer",
        indicator=True,
        suffixes=("_left", "_right"),
    )

    # Summary das categorias do merge
    summary_df = (
        merged_df["_merge"]
        .value_counts(dropna=False)
        .rename_axis("categoria")
        .reset_index(name="quantidade")
    )
    category_order = {"both": 0, "left_only": 1, "right_only": 2}
    summary_df["ord"] = summary_df["categoria"].map(lambda c: category_order.get(c, 99))
    summary_df = summary_df.sort_values(["ord", "categoria"]).drop(columns=["ord"]).reset_index(drop=True)

    # Salva em Excel com duas abas: merge e summary
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        merged_df.to_excel(writer, sheet_name="merge", index=False)
        summary_df.to_excel(writer, sheet_name="summary", index=False)

    print(
        f"Merge concluído com sucesso. Linhas: {len(merged_df)} | Saída: '{output_path.resolve()}'"
    )

