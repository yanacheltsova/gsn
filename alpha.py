import argparse
import pandas as pd
import re
import sys
import os

from typing import Dict, List, Tuple, Any


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()

    parser.add_argument("--AlphaMissense", required=True)
    parser.add_argument("--mutations", required=True)

    return parser.parse_args()


def read_alphamissense(tsv_path: str) -> pd.DataFrame:
    try:
        return pd.read_csv(tsv_path, sep="\t")
    except Exception as e:
        sys.exit(f"Error reading TSV: {e}")


def get_gene_name(df: pd.DataFrame) -> str:
    if "Gene name" not in df.columns:
        sys.exit("Column 'Gene name' not found")

    gene: str = str(df["Gene name"].iloc[0]).strip()

    if not gene or gene == "nan":
        gene = "UnknownGene"

    gene = re.sub(r'[<>:"/\\|?*]', "_", gene)

    return gene


def generate_output_filename(gene: str) -> str:
    base: str = f"{gene}_result.xlsx"

    if not os.path.exists(base):
        return base

    i: int = 1

    while True:
        candidate: str = f"{gene}_result_{i}.xlsx"

        if not os.path.exists(candidate):
            return candidate

        i += 1


def parse_mutation(mutation: str) -> Tuple[str, int, str]:
    pattern: str = r"^([A-Z])([0-9]+)([A-Z])$"
    match = re.match(pattern, mutation)

    if not match:
        raise ValueError(f"Invalid mutation format: {mutation}")

    ref, pos, alt = match.groups()

    return ref, int(pos), alt


def read_mutations(txt_path: str) -> List[Tuple[str, str, int, str]]:
    mutations: List[Tuple[str, str, int, str]] = []

    try:
        with open(txt_path) as f:
            content: str = f.read()

        raw_mutations: List[str] = content.split(";")

        for mut in raw_mutations:
            mut = mut.strip()

            if not mut:
                continue

            try:
                parsed = parse_mutation(mut)
                mutations.append((mut, *parsed))

            except ValueError as e:
                print(f"Warning: {e}")

        return mutations

    except Exception as e:
        sys.exit(f"Error reading mutations: {e}")


def extract_alt_aas(cell: Any) -> List[str]:
    if pd.isna(cell):
        return []

    cell = str(cell).strip()

    if cell == "None" or cell == "":
        return []

    if ":" in cell:
        cell = cell.split(":", 1)[1]

    return [x.strip() for x in cell.split(",") if x.strip()]


def build_lookup(
    df: pd.DataFrame
) -> Dict[Tuple[str, int], Dict[str, List[str]]]:

    lookup: Dict[Tuple[str, int], Dict[str, List[str]]] = {}

    for _, row in df.iterrows():
        key: Tuple[str, int] = (
            row["a.a."],
            int(row["position"])
        )

        lookup[key] = {
            "benign": extract_alt_aas(row["all benign variants"]),
            "ambiguous": extract_alt_aas(row["all ambiguous variants"]),
            "pathogenic": extract_alt_aas(row["all pathogenic variants"])
        }

    return lookup


def find_category(
    row_data: Dict[str, List[str]],
    alt: str
) -> str:

    if alt in row_data["pathogenic"]:
        return "pathogenic"

    elif alt in row_data["benign"]:
        return "benign"

    elif alt in row_data["ambiguous"]:
        return "ambiguous"

    else:
        return "NOT_FOUND"


def match_mutations(
    lookup: Dict[Tuple[str, int], Dict[str, List[str]]],
    mutations: List[Tuple[str, str, int, str]]
) -> pd.DataFrame:

    results: List[Dict[str, str]] = []

    for mut_str, ref, pos, alt in mutations:
        key: Tuple[str, int] = (ref, pos)

        if key not in lookup:
            category: str = "POSITION_NOT_FOUND"

        else:
            row_data = lookup[key]
            category = find_category(row_data, alt)

        results.append({
            "mutation": mut_str,
            "category": category
        })

    return pd.DataFrame(results)


def main() -> None:
    args: argparse.Namespace = parse_args()

    df: pd.DataFrame = read_alphamissense(args.AlphaMissense)

    mutations: List[Tuple[str, str, int, str]] = read_mutations(
        args.mutations
    )

    gene: str = get_gene_name(df)

    output_file: str = generate_output_filename(gene)

    lookup: Dict[
        Tuple[str, int],
        Dict[str, List[str]]
    ] = build_lookup(df)

    result_df: pd.DataFrame = match_mutations(
        lookup,
        mutations
    )

    try:
        result_df.to_excel(output_file, index=False)
        print(f"Saved to {output_file}")

    except Exception as e:
        sys.exit(f"Error writing Excel: {e}")


if __name__ == "__main__":
    main()