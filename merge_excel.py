#!/usr/bin/env python3
"""
Merge (cruzamento) de dois arquivos Excel pela coluna 'MATRICULA'.

Funcionalidades:
- Lê dois arquivos Excel (planilhas opcionais personalizáveis)
- Faz merge com how=outer e indicator=True, preservando todos os registros
- Padroniza a coluna-chave removendo espaços e convertendo para string
- Gera arquivo Excel de saída com duas abas: 'merge' e 'summary'

Exemplo de uso:
    python merge_excel.py \
        --left dados/funcionarios_1.xlsx \
        --right dados/funcionarios_2.xlsx \
        --output resultados/merge_result.xlsx

Requisitos:
    - pandas
    - openpyxl
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Optional

import pandas as pd


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Cruzamento de dois arquivos Excel pela coluna 'MATRICULA' com indicator=True"
        )
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
    return parser.parse_args()


def read_excel(path: Path, sheet_name: Optional[str] = None) -> pd.DataFrame:
    try:
        if sheet_name:
            df = pd.read_excel(path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(path)
    except Exception as exc:
        raise RuntimeError(f"Falha ao ler '{path}': {exc}") from exc

    # Padroniza nomes de colunas: tira espaços nas pontas
    df.columns = [str(col).strip() for col in df.columns]
    return df


def ensure_key_column(df: pd.DataFrame, key: str) -> None:
    if key not in df.columns:
        available = ", ".join(map(str, df.columns))
        raise KeyError(
            f"Coluna-chave '{key}' não encontrada. Colunas disponíveis: {available}"
        )


def normalize_key_column_as_string(df: pd.DataFrame, key: str) -> None:
    # Converte a coluna-chave para string e remove espaços laterais
    df[key] = df[key].astype(str).str.strip()


def build_summary(df_merged: pd.DataFrame) -> pd.DataFrame:
    counts = df_merged["_merge"].value_counts(dropna=False).rename_axis("categoria").reset_index(name="quantidade")
    # Ordena por categoria para uma visualização consistente
    category_order = {"both": 0, "left_only": 1, "right_only": 2}
    counts["ord"] = counts["categoria"].map(lambda c: category_order.get(c, 99))
    counts = counts.sort_values(["ord", "categoria"]).drop(columns=["ord"]).reset_index(drop=True)
    return counts


def save_to_excel(df_merged: pd.DataFrame, df_summary: pd.DataFrame, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df_merged.to_excel(writer, sheet_name="merge", index=False)
            df_summary.to_excel(writer, sheet_name="summary", index=False)
    except Exception as exc:
        raise RuntimeError(f"Falha ao salvar Excel em '{output_path}': {exc}") from exc


def main() -> None:
    args = parse_args()

    left_path = Path(args.left)
    right_path = Path(args.right)
    output_path = Path(args.output)
    key_column = str(args.on)

    if not left_path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {left_path}")
    if not right_path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {right_path}")

    left_df = read_excel(left_path, args.left_sheet)
    right_df = read_excel(right_path, args.right_sheet)

    ensure_key_column(left_df, key_column)
    ensure_key_column(right_df, key_column)

    normalize_key_column_as_string(left_df, key_column)
    normalize_key_column_as_string(right_df, key_column)

    merged_df = pd.merge(
        left_df,
        right_df,
        on=key_column,
        how="outer",
        indicator=True,
        suffixes=("_left", "_right"),
    )

    summary_df = build_summary(merged_df)
    save_to_excel(merged_df, summary_df, output_path)

    print(
        f"Merge concluído com sucesso. Linhas: {len(merged_df)} | Saída: '{output_path.resolve()}'"
    )


if __name__ == "__main__":
    main()

