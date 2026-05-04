# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Excel Data Processing Tool - A Python tool that processes German Excel data with six main functions:
1. **Address Parsing**: Splits combined address data (ZIP/City/Suburb) into separate columns (PLZ/Stadt/Stadtteil)
2. **Name Separation**: Separates combined names into First Name/Last Name (Vorname/Nachname) and identifies organizations
3. **Street Parsing**: Splits street addresses into Street Name/House Number/Suffix (Straßenname/Hausnummer/Hausnummernzusatz)
4. **Street Name Correction**: Fuzzy matches and corrects street names against reference data
5. **Date Formatting**: Auto-detects and formats date columns (DD.MM.YYYY)
6. **Phone Formatting**: Auto-detects and formats phone number columns as text to preserve leading zeros

Uses **simplified German output format** with columns inserted directly next to originals. No metadata columns.

## Common Commands

### Installation
```bash
pip3 install -r requirements.txt
```

### Run the script
```bash
# Mit Projekt (empfohlen):
python3 process_excel.py --project netcom_bw

# Alle verfügbaren Projekte anzeigen:
python3 process_excel.py --list

# Ohne Projekt (automatische Spaltenerkennung, Dateien in input/):
python3 process_excel.py
```

### Neues Projekt anlegen
1. `projects/<projektname>.json` erstellen (Spaltenname, Verzeichnisse, etc.)
2. Excel-Dateien in `input/<Projektname>/` ablegen
3. Script mit `--project <projektname>` aufrufen

## Architecture

### Core Components

- **process_excel.py**: Main script containing all logic
  - `parse_combined_address()`: Multi-pattern parser for German addresses (ZIP/City/Suburb)
  - `find_combined_column()`: Auto-detects address columns
  - `parse_street_address()`: Multi-pattern parser for German street addresses
  - `find_street_column()`: Auto-detects street columns
  - `load_reference_street_names()`: Lazy-loads reference street names from Strassennamen.xlsx
  - `correct_street_name()`: Fuzzy matching to correct street names against reference data
  - `parse_name()`: Multi-pattern parser for German names
  - `find_name_column()`: Auto-detects name columns
  - `is_organization()`: Identifies companies by legal forms and indicators
  - `is_first_name()`: Checks against German first names database
  - `parse_person_name()`: Separates first and last names
  - `process_excel_file()`: Main processing function (handles addresses, names, AND streets)
  - `main()`: Entry point and batch processing orchestration

- **international_first_names.py**: Database module
  - `INTERNATIONAL_FIRST_NAMES`: Set of first names from German and international origins (Turkish, Polish, Czech/Slovak, Serbian/Croatian, Romanian, Bulgarian, Russian/Ukrainian, Greek, Albanian, Arabic) for O(1) lookup

### German Address Parsing Strategy

The parser uses cascading pattern matching with priority order:

1. **Hyphenated format** (highest priority): `80337 München-Schwabing` → Stadt-Stadtteil
2. **Space-separated**: `10115 Berlin Mitte`
3. **Comma-separated**: `60311, Frankfurt, Innenstadt`
4. **Tab/semicolon-separated**: Alternative delimiters
5. **Fallback**: Extracts ZIP and attempts to split remaining text

Handles German characters (ä, ö, ü, ß) and multi-word city/district names.

### German Name Parsing Strategy

The parser uses cascading pattern matching with priority order:

1. **Organization Detection** (highest priority): Checks for legal forms (GmbH, AG, eV), institutional prefixes (Stadt, Gemeinde), business descriptors (Bauunternehmen)
2. **Multi-person format**: `Klein Barbara und Dieter` → Detects "und" pattern
3. **Person with business**: Separates person name from business descriptor
4. **Standard person name**: `Gramer Rainer` → Uses first names database to identify LASTNAME FIRSTNAME format
5. **Reversed format**: `Alexandra Thoma` → Detected via first name lookup
6. **Fallback**: Defaults to German convention (LASTNAME FIRSTNAME) with PARTIAL status

Uses German first names database (2500+ names) and handles name particles (da, de, van, von).

### German Street Parsing Strategy

The parser uses cascading pattern matching with priority order:

1. **Slash suffix with space** (highest priority): `Im Roßhimmel 35 /1` → Street: "Im Roßhimmel", Number: "35", Suffix: "1"
2. **Slash suffix without space**: `Neckartalstraße 39/1` → Street: "Neckartalstraße", Number: "39", Suffix: "1"
3. **Letter suffix with space**: `Hauptstraße 12 a` → Street: "Hauptstraße", Number: "12", Suffix: "a"
4. **Letter suffix without space**: `Hauptstraße 12a` → Street: "Hauptstraße", Number: "12", Suffix: "a"
5. **Standard format** (no suffix): `Starzacher Straße 68` → Street: "Starzacher Straße", Number: "68"
6. **Fallback**: Attempts to extract street name and house number from any remaining patterns

Handles German characters (ä, ö, ü, ß), multi-word street names (Im Roßhimmel, Alte Kelter), and various house number suffix formats.

### Street Name Correction Strategy

The street name correction feature uses fuzzy matching to standardize street names against a reference file (Strassennamen.xlsx). Matching is performed in 4 tiers:

1. **Exact match** (highest priority): Case-sensitive exact match
2. **Case-insensitive match**: Ignores case differences
3. **Normalized fuzzy match**: Converts to lowercase, removes spaces, replaces 'ß' with 'ss'
4. **Partial match**: Matches base street names (e.g., "Starzacher" matches "Starzacher Straße")

Reference data is loaded lazily on first use to avoid unnecessary file I/O.

### Date Formatting Feature

The tool automatically detects and formats date columns in the output Excel file:

- **Auto-detection**: Identifies date columns by keywords ("datum", "termin", "hinzugefügt")
- **Format applied**: DD.MM.YYYY using openpyxl number_format
- **Applies to all rows**: Formats entire column, not just individual cells

### Phone Number Formatting Feature

The tool automatically detects and formats phone number columns as text in the output Excel file:

- **Auto-detection**: Identifies phone columns by keywords ("telefon", "phone", "tel", "mobile", "handy", "fon")
- **Format applied**: Text format (@) using openpyxl number_format
- **Prevents data loss**: Ensures leading zeros in phone numbers are preserved (e.g., "017665817242")
- **Applies to all rows**: Formats entire column, not just individual cells

Both features use pd.ExcelWriter with openpyxl engine for advanced Excel formatting capabilities.

### File Structure

```
process_excel.py         - Generischer Verarbeitungscode (projektunabhängig)
CLAUDE.md                - Tool-Dokumentation (gilt für alle Projekte)
projects/
  netcom_bw.json         - Projektkonfiguration NetCom BW
  generic.json           - Projektkonfiguration Generic (auto-detection)
input/
  NetCom_BW/             - Excel-Dateien für NetCom BW
  Generic/               - Excel-Dateien für generische Verarbeitung
output/
  NetCom_BW/             - Ergebnisse NetCom BW
  Generic/               - Ergebnisse Generic
```

### Project Configuration (JSON Format)

Jede `projects/<name>.json` Datei definiert:
```json
{
  "project_name": "NetCom BW",
  "description": "Beschreibung des Projekts",
  "input_dir": "input/NetCom_BW",
  "output_dir": "output/NetCom_BW",
  "columns": {
    "name": "Kundename",           // null = auto-detect
    "address": null,               // null = auto-detect
    "street": "Anschlussadresse Straße"
  },
  "reference_files": {
    "street_names": "input/NetCom_BW/Strassennamen.xlsx"  // null = disabled
  },
  "date_keywords": ["datum", "termin", "erstellt"],
  "phone_keywords": ["telefon", "tel"]
}
```
- Spalten mit `null` werden automatisch per Keyword erkannt
- Explizite Spaltennamen verhindern Fehlererkennungen

### Output Columns

**Simplified output format** - New columns are inserted directly next to original columns, no metadata columns.

**For Address Data (Ort column):**
- `PLZ`: 5-digit postal code (e.g., "72108")
- `Stadt`: City name (e.g., "Rottenburg am Neckar")
- `Stadtteil`: District name (e.g., "Bieringen")

**For Name Data (Name column):**
- `Vorname`: First name(s) or empty for organizations (e.g., "Monika")
- `Nachname`: Last name for persons, full name for organizations (e.g., "Aha")

**For Street Data (straße column):**
- `Straßenname`: Street name (e.g., "Starzacher Straße", "Im Roßhimmel")
- `Straßenname (korrigiert)`: Corrected street name based on reference file (e.g., "Starzacher Strasse")
- `Hausnummer`: House number (e.g., "68", "35")
- `Hausnummernzusatz`: Optional suffix (e.g., "1" from "/1", "a", "b", or empty)

**Column positioning**: New columns are inserted immediately after the original column using df.insert(). For example, if "Name" is in column C, then "Vorname" will be in column D and "Nachname" in column E.

## Development Notes

- Uses pandas for Excel I/O and openpyxl as the engine
- All regex patterns support German characters (ä, ö, ü, ß)
- Column auto-detection searches for keywords:
  - Address: "address", "adresse", "ort", "location", "plz"
  - Name: "name", "namen", "person", "contact", "kontakt"
  - Street: "street", "strasse", "straße", "addr", "address", "adresse"
- Batch processing: handles all .xlsx files in input/ directory
- Non-destructive: original files are never modified
- Organization detection uses constants: COMPANY_LEGAL_FORMS, INSTITUTIONAL_PREFIXES, BUSINESS_DESCRIPTORS
- First names loaded lazily from international_first_names.py module
- Reference street names loaded lazily from input/Strassennamen.xlsx on first correction attempt
- Handles address-only, name-only, street-only, and mixed files automatically
- Street column detection also uses content analysis: checks for German street suffixes (straße, weg, gasse, platz, allee) and number patterns
- **Simplified output**: No metadata columns (Original_X, Parse_Status, Name_Type, Business_Indicator)
- **German column names**: Uses German terminology (Vorname, Nachname, PLZ, Stadt, Stadtteil, Straßenname, Hausnummer, Hausnummernzusatz)
- **Column positioning**: New columns inserted directly after original column using df.insert()
- **Date formatting**: Auto-detects date columns by keywords and applies DD.MM.YYYY format
- **Phone formatting**: Auto-detects phone columns by keywords and applies text format (@) to preserve leading zeros
- **Street correction**: Fuzzy matches street names against reference file with 4-tier matching algorithm

### Test Results (Verified)

**Individual data type files:**
- **Addresses (Ort.xlsx)**: 107/107 = 100% success rate
- **Names (Namen.xlsx)**: 101/107 success, 6 partial = 94.4% success rate
  - 90 Persons identified
  - 15 Organizations detected
  - 2 Multi-person entries found
  - 0 Failed parsing
- **Streets (Strasse.xlsx)**: 243/243 = 100% success rate
  - 236 Standard format (street name + number)
  - 7 With slash suffixes ("/1")
  - 0 Partial, 0 Failed parsing
  - Correctly handles multi-word street names (Im Roßhimmel, Siegburger Weg)

**Combined file (20251008_Bieringen.xlsx)**: 243 rows with all three data types
- **Addresses**: 243/243 = 100% success rate
  - All in hyphenated format: "72108 Rottenburg am Neckar-Bieringen"
- **Names**: 211 success, 32 partial = 86.8% success rate
  - Correctly parsed common German names (Aha Monika, Ahrens-Diez Bettina, Akyüz Tülin)
- **Streets**: 243/243 = 100% success rate
  - 2 street names corrected using fuzzy matching against reference file
- **Date formatting**: 4 date columns auto-detected and formatted (DD.MM.YYYY)
  - Columns: "Datum der Auftragsstellung", "Plantermin", "Bestätigter Termin", "Hinzugefügt am "
- **Phone formatting**: 1 phone column auto-detected and formatted as text (@)
  - Column: "Telefon" - preserves leading zeros (e.g., "017665817242")
