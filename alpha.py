import argparse
import pandas as pd
import re
import sys
import os


def parse_args():
    parser = argparse.ArgumentParser()

    parser.add_argument("--AlphaMissense", required=True)
    parser.add_argument("--mutations", required=True)

    return parser.parse_args()


def read_alphamissense(tsv_path):
    try:
        return pd.read_csv(tsv_path, sep="\t")
    except Exception as e:
        sys.exit(f"Error reading TSV: {e}")


def get_gene_name(df):
    if "Gene name" not in df.columns:
        sys.exit("Column 'Gene name' not found")

    gene = str(df["Gene name"].iloc[0]).strip()

    if not gene or gene == "nan":
        gene = "UnknownGene"

    gene = re.sub(r'[<>:"/\\|?*]', "_", gene)

    return gene


def generate_output_filename(gene):
    base = f"{gene}_result.xlsx"

    if not os.path.exists(base):
        return base

    i = 1
    while True:
        candidate = f"{gene}_result_{i}.xlsx"
        if not os.path.exists(candidate):
            return candidate
        i += 1


def parse_mutation(mutation):
    pattern = r"^([A-Z])([0-9]+)([A-Z])$"
    match = re.match(pattern, mutation)

    if not match:
        raise ValueError(f"Invalid mutation format: {mutation}")

    ref, pos, alt = match.groups()

    return ref, int(pos), alt


def read_mutations(txt_path):
    mutations = []

    try:
        with open(txt_path) as f:
            content = f.read()

        raw_mutations = content.split(";")

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


def extract_alt_aas(cell):
    if pd.isna(cell):
        return []

    cell = str(cell).strip()

    if cell == "None" or cell == "":
        return []

    if ":" in cell:
        cell = cell.split(":", 1)[1]

    return [x.strip() for x in cell.split(",") if x.strip()]


def build_lookup(df):
    lookup = {}

    for _, row in df.iterrows():
        key = (row["a.a."], int(row["position"]))

        lookup[key] = {
            "benign": extract_alt_aas(row["all benign variants"]),
            "ambiguous": extract_alt_aas(row["all ambiguous variants"]),
            "pathogenic": extract_alt_aas(row["all pathogenic variants"])
        }

    return lookup


def find_category(row_data, alt):
    if alt in row_data["pathogenic"]:
        return "pathogenic"
    elif alt in row_data["benign"]:
        return "benign"
    elif alt in row_data["ambiguous"]:
        return "ambiguous"
    else:
        return "NOT_FOUND"


def match_mutations(lookup, mutations):
    results = []

    for mut_str, ref, pos, alt in mutations:
        key = (ref, pos)

        if key not in lookup:
            category = "POSITION_NOT_FOUND"
        else:
            row_data = lookup[key]
            category = find_category(row_data, alt)

        results.append({
            "mutation": mut_str,
            "category": category
        })

    return pd.DataFrame(results)


def main():
    args = parse_args()
    df = read_alphamissense(args.AlphaMissense)
    mutations = read_mutations(args.mutations)

    gene = get_gene_name(df)
    output_file = generate_output_filename(gene)

    lookup = build_lookup(df)

    result_df = match_mutations(lookup, mutations)

    try:
        result_df.to_excel(output_file, index=False)
        print(f"Saved to {output_file}")
    except Exception as e:
        sys.exit(f"Error writing Excel: {e}")


if __name__ == "__main__":
    main()