import argparse
import os
import re
import freesasa
from openpyxl import Workbook


def validate_pdb(path):
    if not os.path.isfile(path):
        raise argparse.ArgumentTypeError(f"File not found: {path}")

    if not path.lower().endswith(".pdb"):
        raise argparse.ArgumentTypeError("File must have a .pdb extension")

    try:
        freesasa.Structure(path)
    except Exception as e:
        raise argparse.ArgumentTypeError(f"Failed to parse PDB file: {e}")

    return path


def parse_chains(value):
    chains = [c.strip() for c in value.split(",") if c.strip()]

    for c in chains:
        if len(c) != 1:
            raise argparse.ArgumentTypeError(f"Invalid chain ID: {c}")

    return chains


res_pattern = re.compile(r"^(\d+)([A-Za-z]?)$")


def parse_residues(value):
    residues = set()
    parts = value.split(",")

    for part in parts:
        part = part.strip()

        if "-" in part:
            start_str, end_str = part.split("-")

            m1 = res_pattern.match(start_str)
            m2 = res_pattern.match(end_str)

            if not m1 or not m2:
                raise argparse.ArgumentTypeError(f"Invalid residue range: {part}")

            start_num = int(m1.group(1))
            end_num = int(m2.group(1))

            if start_num > end_num:
                raise argparse.ArgumentTypeError(f"Invalid residue range: {part}")

            for i in range(start_num, end_num + 1):
                residues.add((i, ""))

        else:
            m = res_pattern.match(part)
            if not m:
                raise argparse.ArgumentTypeError(f"Invalid residue: {part}")

            num = int(m.group(1))
            ic = m.group(2)
            residues.add((num, ic))

    return sorted(residues)


def get_available_residues(pdb_path):
    structure = freesasa.Structure(pdb_path)
    result = freesasa.calc(structure)

    residue_areas = result.residueAreas()

    available = {}

    for chain_id, residues_dict in residue_areas.items():
        available[chain_id] = {}

        for res_id, area in residues_dict.items():
            match = re.match(r"^(\d+)([A-Za-z]?)$", res_id)
            if not match:
                continue

            resnum = int(match.group(1))
            icode = match.group(2)
            resname = area.residueType

            available[chain_id][(resnum, icode)] = resname

    return available


def validate_selection(available, chains, residues):
    available_chains = set(available.keys())

    if chains:
        invalid = [c for c in chains if c not in available_chains]
        if invalid:
            raise ValueError(f"Chains not found: {invalid}")

    if residues:
        selected_chains = chains if chains else available.keys()

        all_residues = set()
        for c in selected_chains:
            all_residues.update(available[c].keys())

        invalid = [r for r in residues if r not in all_residues]
        if invalid:
            formatted = [f"{r[0]}{r[1]}" for r in invalid]
            raise ValueError(f"Residues not found: {formatted}")


def calculate_rsa(pdb_path, chains, residues):
    structure = freesasa.Structure(pdb_path)
    result = freesasa.calc(structure)

    residue_areas = result.residueAreas()

    rsa_data = []
    threshold = 0.25

    for chain_id, residues_dict in residue_areas.items():
        if chains and chain_id not in chains:
            continue

        for res_id, area in residues_dict.items():
            match = re.match(r"^(\d+)([A-Za-z]?)$", res_id)
            if not match:
                continue

            resnum = int(match.group(1))
            icode = match.group(2)

            if residues and (resnum, icode) not in residues:
                continue

            rsa = area.relativeTotal
            resname = area.residueType

            formatted_name = f"{resname.capitalize()}{resnum}{chain_id}"

            category = "На поверхности" if rsa >= threshold else "Внутри"

            rsa_data.append((formatted_name, rsa, category))

    return rsa_data


def generate_output_filename(pdb_path):
    base_name = os.path.splitext(os.path.basename(pdb_path))[0]
    directory = os.path.dirname(pdb_path)

    base_output = os.path.join(directory, f"{base_name}_rsa.xlsx")

    if not os.path.exists(base_output):
        return base_output

    counter = 1
    while True:
        new_name = os.path.join(directory, f"{base_name}_rsa_{counter}.xlsx")
        if not os.path.exists(new_name):
            return new_name
        counter += 1


def save_to_xlsx(data, output_file):
    wb = Workbook()
    ws = wb.active

    ws.append(["residue", "rsa", "category"])

    for residue, rsa, category in data:
        ws.append([residue, rsa, category])

    wb.save(output_file)


parser = argparse.ArgumentParser()

parser.add_argument("--pdb", required=True, type=validate_pdb)
parser.add_argument("--chain", type=parse_chains)
parser.add_argument("--residue", type=parse_residues)

args = parser.parse_args()

available = get_available_residues(args.pdb)

try:
    validate_selection(available, args.chain, args.residue)
except ValueError as e:
    parser.error(str(e))

rsa_data = calculate_rsa(args.pdb, args.chain, args.residue)

output_file = generate_output_filename(args.pdb)

save_to_xlsx(rsa_data, output_file)

print(f"Success! Result saved in {output_file}")
