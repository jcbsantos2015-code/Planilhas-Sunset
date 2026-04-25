#!/usr/bin/env python3
"""Valida registros pendentes de NFS-e em planilha Excel.

Regras implementadas:
- Processa apenas linhas com STATUS_PROCESSAMENTO == "PENDENTE"
- Valida CPF/CNPJ, Valor Total, Município e DESCRICAO_NF
- Gera arquivos:
  - dados_validos.xlsx
  - erros.xlsx (com coluna MOTIVO_ERRO)
- Atualiza STATUS_PROCESSAMENTO para VALIDADO/ERRO na aba BASE_NF
- Preserva as demais abas inalteradas
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Dict, List

import pandas as pd


def _build_column_lookup(columns: List[str]) -> Dict[str, str]:
    """Mapeia nomes normalizados para nomes reais de colunas."""
    return {str(col).strip().upper(): col for col in columns}


def _require_column(colmap: Dict[str, str], logical_name: str) -> str:
    """Retorna o nome real da coluna ou lança erro amigável."""
    key = logical_name.strip().upper()
    if key not in colmap:
        available = ", ".join(colmap.keys())
        raise ValueError(
            f"Coluna obrigatória não encontrada: '{logical_name}'. "
            f"Colunas disponíveis: {available}"
        )
    return colmap[key]


def _is_blank(value: object) -> bool:
    if pd.isna(value):
        return True
    return str(value).strip() == ""


def _validate_row(
    row: pd.Series,
    cpf_col: str,
    valor_col: str,
    municipio_col: str,
    descricao_col: str,
) -> List[str]:
    erros: List[str] = []

    if _is_blank(row.get(cpf_col)):
        erros.append("CPF/CNPJ vazio")

    valor = row.get(valor_col)
    valor_num = pd.to_numeric(pd.Series([valor]), errors="coerce").iloc[0]
    if pd.isna(valor_num) or float(valor_num) <= 0:
        erros.append("Valor Total deve ser maior que 0")

    if _is_blank(row.get(municipio_col)):
        erros.append("Município não preenchido")

    if _is_blank(row.get(descricao_col)):
        erros.append("DESCRICAO_NF não preenchida")

    return erros


def processar_planilha(
    input_file: Path,
    output_validos: Path,
    output_erros: Path,
    output_atualizado: Path,
) -> None:
    all_sheets = pd.read_excel(input_file, sheet_name=None)

    if "BASE_NF" not in all_sheets:
        raise ValueError("A aba 'BASE_NF' não foi encontrada no arquivo informado.")

    df_base = all_sheets["BASE_NF"].copy()
    colmap = _build_column_lookup(list(df_base.columns))

    status_col = _require_column(colmap, "STATUS_PROCESSAMENTO")
    cpf_col = _require_column(colmap, "CPF/CNPJ")
    valor_col = _require_column(colmap, "Valor Total")
    municipio_col = _require_column(colmap, "Município")
    descricao_col = _require_column(colmap, "DESCRICAO_NF")

    status_normalizado = df_base[status_col].astype(str).str.strip().str.upper()
    mask_pendente = status_normalizado == "PENDENTE"

    pendentes = df_base.loc[mask_pendente].copy()

    motivos: List[str] = []
    status_novo: List[str] = []

    for _, row in pendentes.iterrows():
        erros = _validate_row(row, cpf_col, valor_col, municipio_col, descricao_col)
        if erros:
            motivos.append("; ".join(erros))
            status_novo.append("ERRO")
        else:
            motivos.append("")
            status_novo.append("VALIDADO")

    pendentes["MOTIVO_ERRO"] = motivos
    pendentes[status_col] = status_novo

    validos = pendentes[pendentes[status_col] == "VALIDADO"].drop(columns=["MOTIVO_ERRO"])
    erros_df = pendentes[pendentes[status_col] == "ERRO"].copy()

    # Atualiza o status na BASE_NF original sem alterar as outras abas
    df_base.loc[mask_pendente, status_col] = status_novo
    all_sheets["BASE_NF"] = df_base

    validos.to_excel(output_validos, index=False)
    erros_df.to_excel(output_erros, index=False)

    with pd.ExcelWriter(output_atualizado, engine="openpyxl") as writer:
        for nome_aba, df in all_sheets.items():
            df.to_excel(writer, sheet_name=nome_aba, index=False)

    print(f"Registros pendentes analisados: {len(pendentes)}")
    print(f"Registros válidos: {len(validos)}")
    print(f"Registros com erro: {len(erros_df)}")
    print(f"Arquivo atualizado salvo em: {output_atualizado}")
    print(f"Arquivo de válidos salvo em: {output_validos}")
    print(f"Arquivo de erros salvo em: {output_erros}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Valida BASE_NF e gera arquivos de válidos/erros para NFS-e."
    )
    parser.add_argument(
        "--input",
        default="Tabela residente Atualizado.xlsx",
        help="Arquivo Excel de entrada (padrão: Tabela residente Atualizado.xlsx)",
    )
    parser.add_argument(
        "--validos",
        default="dados_validos.xlsx",
        help="Arquivo de saída com registros válidos",
    )
    parser.add_argument(
        "--erros",
        default="erros.xlsx",
        help="Arquivo de saída com registros inválidos e motivo",
    )
    parser.add_argument(
        "--atualizado",
        default="Tabela residente Atualizado_status.xlsx",
        help="Arquivo de saída com STATUS_PROCESSAMENTO atualizado",
    )

    args = parser.parse_args()

    processar_planilha(
        input_file=Path(args.input),
        output_validos=Path(args.validos),
        output_erros=Path(args.erros),
        output_atualizado=Path(args.atualizado),
    )


if __name__ == "__main__":
    main()
