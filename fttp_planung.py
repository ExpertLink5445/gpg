#!/usr/bin/env python3
"""
FTTP-Planung Prozessierung
- Liest NVT*-Dateien aus input/FTTP-Planung/
- Teilt auf in MFG- und RVT-Dateien
- Splittet Adresse in Straße / Hausnummer / Hausnummernzusatz
- Entfernt '-2024' aus Spalte NAME
- Ergänzt 'zuständig', 'zugehöriges Projekt' und 'Auftraggeber' (interaktive Eingabe)
- Speichert Excel + CSV (UTF-8, Semikolon) in output/FTTP-Planung/
"""

import glob
import os
import sys
import pandas as pd

sys.path.insert(0, os.path.dirname(__file__))
from process_excel import parse_street_address

INPUT_DIR  = "input/FTTP-Planung"
OUTPUT_DIR = "output/FTTP-Planung"


def split_address(addr):
    r = parse_street_address(str(addr))
    return pd.Series({
        "Straße":            r.get("street_name", ""),
        "Hausnummer":        r.get("house_number", ""),
        "Hausnummernzusatz": r.get("house_number_suffix", ""),
    })


def process_file(path, zustaendig, projekt, auftraggeber):
    df_raw = pd.read_excel(path, header=None)

    header = df_raw.iloc[0].tolist()
    data   = df_raw.iloc[1:].copy().reset_index(drop=True)
    data.columns = header

    # Entferne '-2024' aus NAME
    data["NAME"] = data["NAME"].str.replace("-2024", "", regex=False)

    # Adresse aufteilen
    addr_split = data["Adressen"].apply(split_address)
    pos = data.columns.tolist().index("Adressen")
    data = pd.concat([data.iloc[:, :pos], addr_split, data.iloc[:, pos + 1:]], axis=1)

    # Zusatzspalten
    data["Auftraggeber"]        = auftraggeber
    data["zuständig"]           = zustaendig
    data["zugehöriges Projekt"] = projekt

    # Aufteilen nach MFG / RVT
    df_mfg = data[data["NAME"].str.contains("MFG", na=False)].reset_index(drop=True)
    df_rvt = data[data["NAME"].str.contains("RVT", na=False)].reset_index(drop=True)

    stem = os.path.splitext(os.path.basename(path))[0]

    for suffix, df_out in [("MFG", df_mfg), ("RVT", df_rvt)]:
        if df_out.empty:
            print(f"  ⚠️  Keine {suffix}-Einträge gefunden.")
            continue
        xlsx = os.path.join(OUTPUT_DIR, f"{stem}_{suffix}.xlsx")
        csv  = os.path.join(OUTPUT_DIR, f"{stem}_{suffix}.csv")
        df_out.to_excel(xlsx, index=False)
        df_out.to_csv(csv, sep=";", index=False, encoding="utf-8")
        print(f"  ✅ {suffix}: {len(df_out)} Zeilen → {os.path.basename(xlsx)} + .csv")


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    files = glob.glob(os.path.join(INPUT_DIR, "NVT*.xlsx"))
    if not files:
        print(f"Keine NVT*.xlsx-Dateien in {INPUT_DIR} gefunden.")
        sys.exit(1)

    print(f"Gefundene Dateien: {[os.path.basename(f) for f in files]}")
    print()

    auftraggeber = input("Auftraggeber: ").strip()
    zustaendig   = input("zuständig: ").strip()
    projekt      = input("zugehöriges Projekt: ").strip()
    print()

    for path in files:
        print(f"Verarbeite: {os.path.basename(path)}")
        process_file(path, zustaendig, projekt, auftraggeber)

    print("\nFertig.")


if __name__ == "__main__":
    main()
