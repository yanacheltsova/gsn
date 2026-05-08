import argparse
import os
import re
from typing import Dict, List, Optional, Set, Tuple

import freesasa
from openpyxl import Workbook


ResidueID = Tuple[int, str]
RSARecord = Tuple[str, float, str]

res_pattern: re.Pattern[str] = re.compile(r"^(\d+)([A-Za-z]?)$")


def validate_pdb(path: str) -> str:
    if not os.path.isfile(path):
        raise argparse.ArgumentTypeError(f"File not found: {path}")

    if not path.lower().endswith(".pdb"):
        raise argparse.ArgumentTypeError("File must have a .pdb extension")

    try:
        freesasa.Structure(path)
    except Exception as e:
        raise argparse.ArgumentTypeError(f"Failed to parse PDB file: {e}")

    return path


def parse_chains(value: str) -> List[str]:
    chains: List[str] = [c.strip() for c in value.split(",") if c.strip()]

    for c in chains:
        if len(c) != 1:
            raise argparse.ArgumentTypeError(f"Invalid chain ID: {c}")

    return chains


def parse_residues(value: str) -> List[ResidueID]:
    residues: Set[ResidueID] = set()
    parts: List[str] = value.split(",")

    for part in parts:
        part = part.strip()

        if "-" in part:
            start_str, end_str = part.split("-")

            m1 = res_pattern.match(start_str)
            m2 = res_pattern.match(end_str)

            if not m1 or not m2:
                raise argparse.ArgumentTypeError(
                    f"Invalid residue range: {part}"
                )

            start_num: int = int(m1.group(1))
            end_num: int = int(m2.group(1))

            if start_num > end_num:
                raise argparse.ArgumentTypeError(
                    f"Invalid residue range: {part}"
                )

            for i in range(start_num, end_num + 1):
                residues.add((i, ""))

        else:
            m = res_pattern.match(part)

            if not m:
                raise argparse.ArgumentTypeError(
                    f"Invalid residue: {part}"
                )

            num: int = int(m.group(1))
            ic: str = m.group(2)

            residues.add((num, ic))

    return sorted(residues)


def get_available_residues(
    pdb_path: str
) -> Dict[str, Dict[ResidueID, str]]:

    structure = freesasa.Structure(pdb_path)
    result = freesasa.calc(structure)

    residue_areas = result.residueAreas()

    available: Dict[str, Dict[ResidueID, str]] = {}

    for chain_id, residues_dict in residue_areas.items():
        available[chain_id] = {}

        for res_id, area in residues_dict.items():
            match = re.match(r"^(\d+)([A-Za-z]?)$", res_id)

            if not match:
                continue

            resnum: int = int(match.group(1))
            icode: str = match.group(2)
            resname: str = area.residueType

            available[chain_id][(resnum, icode)] = resname

    return available


def validate_selection(
    available: Dict[str, Dict[ResidueID, str]],
    chains: Optional[List[str]],
    residues: Optional[List[ResidueID]]
) -> None:

    available_chains: Set[str] = set(available.keys())

    if chains:
        invalid: List[str] = [
            c for c in chains if c not in available_chains
        ]

        if invalid:
            raise ValueError(f"Chains not found: {invalid}")

    if residues:
        selected_chains = chains if chains else available.keys()

        all_residues: Set[ResidueID] = set()

        for c in selected_chains:
            all_residues.update(available[c].keys())

        invalid: List[ResidueID] = [
            r for r in residues if r not in all_residues
        ]

        if invalid:
            formatted: List[str] = [
                f"{r[0]}{r[1]}" for r in invalid
            ]

            raise ValueError(
                f"Residues not found: {formatted}"
            )


def calculate_rsa(
    pdb_path: str,
    chains: Optional[List[str]],
    residues: Optional[List[ResidueID]]
) -> List[RSARecord]:

    structure = freesasa.Structure(pdb_path)
    result = freesasa.calc(structure)

    residue_areas = result.residueAreas()

    rsa_data: List[RSARecord] = []

    threshold: float = 0.25

    for chain_id, residues_dict in residue_areas.items():

        if chains and chain_id not in chains:
            continue

        for res_id, area in residues_dict.items():

            match = re.match(r"^(\d+)([A-Za-z]?)$", res_id)

            if not match:
                continue

            resnum: int = int(match.group(1))
            icode: str = match.group(2)

            if residues and (resnum, icode) not in residues:
                continue

            rsa: float = area.relativeTotal
            resname: str = area.residueType

            formatted_name: str = (
                f"{resname.capitalize()}{resnum}{chain_id}"
            )

            category: str = (
                "На поверхности"
                if rsa >= threshold
                else "Внутри"
            )

            rsa_data.append(
                (formatted_name, rsa, category)
            )

    return rsa_data


def generate_output_filename(pdb_path: str) -> str:
    base_name: str = os.path.splitext(
        os.path.basename(pdb_path)
    )[0]

    directory: str = os.path.dirname(pdb_path)

    base_output: str = os.path.join(
        directory,
        f"{base_name}_rsa.xlsx"
    )

    if not os.path.exists(base_output):
        return base_output

    counter: int = 1

    while True:
        new_name: str = os.path.join(
            directory,
            f"{base_name}_rsa_{counter}.xlsx"
        )

        if not os.path.exists(new_name):
            return new_name

        counter += 1


def save_to_xlsx(
    data: List[RSARecord],
    output_file: str
) -> None:

    wb = Workbook()
    ws = wb.active

    ws.append(["residue", "rsa", "category"])

    for residue, rsa, category in data:
        ws.append([residue, rsa, category])

    wb.save(output_file)



def main() -> None:

    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--pdb",
        required=True,
        type=validate_pdb
    )

    parser.add_argument(
        "--chain",
        type=parse_chains
    )

    parser.add_argument(
        "--residue",
        type=parse_residues
    )

    args = parser.parse_args()

    available: Dict[str, Dict[ResidueID, str]] = (
        get_available_residues(args.pdb)
    )

    try:
        validate_selection(
            available,
            args.chain,
            args.residue
        )

    except ValueError as e:
        parser.error(str(e))

    rsa_data: List[RSARecord] = calculate_rsa(
        args.pdb,
        args.chain,
        args.residue
    )

    output_file: str = generate_output_filename(
        args.pdb
    )

    save_to_xlsx(rsa_data, output_file)

    print(f"Success! Result saved in {output_file}")


if __name__ == "__main__":
    main()