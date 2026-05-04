# Excel Data Processing - Address and Name Separation

A Python script that automatically processes German Excel data to:
1. **Split combined addresses** (ZIP code, city, suburb) into separate columns
2. **Separate names** (first name, last name) and identify organizations

## Features

### Address Processing
- **Automatic Format Detection**: Handles multiple German address formats
- **German Address Support**: Full support for German characters (ä, ö, ü, ß) and conventions
- **Hyphenated City-District**: Recognizes official German Stadt-Stadtteil format

### Name Processing
- **Smart Name Separation**: Splits German names (typically "LASTNAME FIRSTNAME" format)
- **Organization Detection**: Automatically identifies companies by legal forms (GmbH, AG, eV, etc.)
- **Multi-Person Handling**: Recognizes patterns like "Klein Barbara und Dieter"
- **Business Descriptors**: Identifies technical names like "Bauunternehmen"
- **German First Names Database**: Uses 2500+ German first names for accurate classification

### General
- **Batch Processing**: Process multiple Excel files at once
- **Data Preservation**: Keeps original data for reference
- **Status Tracking**: Shows which rows were successfully parsed
- **Zero Additional Dependencies**: No API calls or external services needed

## Supported Formats

### Address Formats

The script automatically detects and parses these German address formats:

1. **Hyphenated** (Official German standard): `80337 München-Schwabing`
2. **Space-separated**: `10115 Berlin Mitte`
3. **City only**: `20095 Hamburg`
4. **Comma-separated**: `60311, Frankfurt, Innenstadt`
5. **Tab/semicolon-separated**: `12345\tBerlin\tMitte`

### Name Formats

The script recognizes and processes these name patterns:

1. **Standard Person** (German convention): `Gramer Rainer` → Last: Gramer, First: Rainer
2. **Organizations**: `Stadt Rottenburg am Neckar`, `Kath. Kirchengemeinde`, `GmbH`, `AG`, `eV`
3. **Multi-person**: `Klein Barbara und Dieter` → Last: Klein, First: Barbara und Dieter
4. **Person with Business**: `Rist Alexander Bauunternehmen` → Detects business descriptors
5. **Multi-word Last Names**: `Magalhäes da Silva Jose Carlos` → Handles name particles (da, de, van, von)
6. **Reversed Format**: `Alexandra Thoma` → Detected via first name database

### Organization Indicators

The script identifies organizations by detecting:
- **Legal forms**: GmbH, UG, AG, OHG, KG, PartG, GbR, eV, eG, Genossenschaft
- **Institutional prefixes**: Stadt, Gemeinde, Kirchengemeinde, Kath., Evang., Zweckverband
- **Business descriptors**: Bauunternehmen, Firma, Betrieb, Einrichtungen, Kindergarten

## Installation

### Prerequisites

- Python 3.6 or higher
- pip (Python package installer)

### Install Dependencies

```bash
pip3 install -r requirements.txt
```

This will install:
- `pandas` - For Excel file processing
- `openpyxl` - For reading/writing .xlsx files

## Usage

### Quick Start

1. **Place your Excel files** in the `input/` folder
2. **Run the script**:
   ```bash
   python3 process_excel.py
   ```
3. **Find processed files** in the `output/` folder

### Detailed Steps

#### Step 1: Prepare Your Excel Files

Copy your Excel files (.xlsx) containing combined address data into the `input/` directory:

```bash
cp /path/to/your/addresses.xlsx input/
```

#### Step 2: Run the Script

```bash
cd /Users/frank/Library/CloudStorage/Dropbox/Dokumente/Frank/Claude-Test
python3 process_excel.py
```

#### Step 3: Review Results

Processed files will be saved in the `output/` directory with the naming format:
```
{original_filename}_processed_{timestamp}.xlsx
```

Example: `addresses.xlsx` → `addresses_processed_20260214_153045.xlsx`

## Output Format

### Input Example

| ID | Address | Name |
|----|---------|------|
| 1 | 10115 Berlin Mitte | Gramer Rainer |
| 2 | 80337 München-Schwabing | Stadt Rottenburg |
| 3 | 20095 Hamburg | Klein Barbara und Dieter |

### Output Example (with both Address and Name columns)

The script automatically detects and processes both types of columns:

| ID | Address | Name | ZIP | City | Suburb | Original_Address | Address_Parse_Status | First_Name | Last_Name | Name_Type | Business_Indicator | Original_Name | Name_Parse_Status |
|----|---------|------|-----|------|--------|------------------|----------------------|------------|-----------|-----------|-------------------|---------------|-------------------|
| 1 | 10115 Berlin Mitte | Gramer Rainer | 10115 | Berlin | Mitte | 10115 Berlin Mitte | SUCCESS | Rainer | Gramer | PERSON | | Gramer Rainer | SUCCESS |
| 2 | 80337 München-Schwabing | Stadt Rottenburg | 80337 | München | Schwabing | 80337 München-Schwabing | SUCCESS | | Stadt Rottenburg | ORGANIZATION | | Stadt Rottenburg | SUCCESS |
| 3 | 20095 Hamburg | Klein Barbara und Dieter | 20095 | Hamburg | | 20095 Hamburg | SUCCESS | Barbara und Dieter | Klein | MULTI_PERSON | | Klein Barbara und Dieter | SUCCESS |

### New Columns for Addresses

- **ZIP**: The postal code (5 digits)
- **City**: City name (Stadt)
- **Suburb**: Suburb/district name (Stadtteil/Ortsteil) - empty if not present
- **Original_Address**: Preserved original combined value
- **Address_Parse_Status**: Parsing result (SUCCESS/PARTIAL/FAILED/EMPTY)

### New Columns for Names

- **First_Name**: First name(s) - empty for organizations
- **Last_Name**: Last name for persons, full name for organizations
- **Name_Type**: Classification
  - `PERSON`: Individual person
  - `ORGANIZATION`: Company, institution, government entity
  - `MULTI_PERSON`: Multiple persons (e.g., "Barbara und Dieter")
  - `PERSON_WITH_BUSINESS`: Person with business descriptor
- **Business_Indicator**: Business descriptor if present (e.g., "Bauunternehmen", "GmbH")
- **Original_Name**: Preserved original combined value
- **Name_Parse_Status**: Parsing result (SUCCESS/PARTIAL/FAILED/EMPTY)

## How It Works

### Automatic Column Detection

The script automatically identifies which columns to process:

**Address Columns:**
- Searches for keywords: "address", "adresse", "ort", "location", "plz"
- Examines columns for 5-digit postal codes

**Name Columns:**
- Searches for keywords: "name", "namen", "person", "contact", "kontakt"

The script can process files with:
- Only address data
- Only name data
- Both address and name data in the same file

### Address Parsing Strategy (Cascading Pattern Matching)

The script tries multiple patterns in priority order:

1. **Hyphenated City-District** (highest priority): `PLZ Stadt-Stadtteil`
2. **Space-separated**: `PLZ Stadt Stadtteil`
3. **Comma-separated**: `PLZ, Stadt, Stadtteil`
4. **Tab/semicolon-separated**: Alternative delimiters
5. **Fallback**: Extracts ZIP and attempts to split remaining text

### Name Parsing Strategy (Cascading Pattern Matching)

The script uses a sophisticated multi-step approach:

1. **Organization Detection** (highest priority)
   - Checks for legal forms (GmbH, AG, eV, etc.)
   - Checks for institutional prefixes (Stadt, Gemeinde, etc.)
   - Checks for business descriptors (Bauunternehmen, etc.)

2. **Multi-Person Format**
   - Detects "und" (and) pattern
   - Example: "Klein Barbara und Dieter"

3. **Person with Business Descriptor**
   - Identifies business terms in name
   - Separates person name from company descriptor

4. **Standard Person Name** (German convention: LASTNAME FIRSTNAME)
   - Uses German first names database (2500+ names)
   - Identifies which word is the first name
   - Handles name particles: da, de, van, von, der, zu

5. **Reversed Format** (FIRSTNAME LASTNAME)
   - Detected when first word matches first names database

6. **Fallback**
   - Defaults to German convention
   - Marks as PARTIAL status for manual review

### German-Specific Features

**Address Processing:**
- Supports German characters: ä, ö, ü, ß
- Handles multi-word cities: "Bad Homburg", "Frankfurt am Main"
- Handles multi-word districts: "Berlin-Köpenick", "Hamburg-Harburg"
- Preserves capitalization

**Name Processing:**
- German first names database with 2500+ names
- Handles German company legal forms and conventions
- Recognizes German institutional names
- Supports multi-word last names with particles
- Preserves German character encoding

## Troubleshooting

### No Excel files found

**Error**: "No Excel files found in input/ directory"

**Solution**: Make sure your .xlsx files are in the `input/` folder

### Could not identify address column

**Error**: "Warning: Could not identify address column"

**Solution**: The script couldn't find a column with address-like data. Ensure your Excel file contains a column with combined address information that includes 5-digit postal codes.

### Import Error

**Error**: "Required packages not found"

**Solution**: Install dependencies:
```bash
pip3 install pandas openpyxl
```

### Parse Status shows FAILED

If some rows show `FAILED` status:
- Check the `Original_Address` column to see the original value
- Verify the format matches one of the supported formats
- Check for unusual characters or formatting

## Examples

### Example 1: Single File

```bash
# Copy file to input
cp my_addresses.xlsx input/

# Run script
python3 process_excel.py

# Output will be in: output/my_addresses_processed_20260214_153045.xlsx
```

### Example 2: Batch Processing

```bash
# Copy multiple files
cp addresses1.xlsx addresses2.xlsx addresses3.xlsx input/

# Run script once
python3 process_excel.py

# All files will be processed
```

## File Structure

```
Claude-Test/
├── input/              # Place your source Excel files here
├── output/             # Processed files appear here
├── requirements.txt    # Python dependencies
├── process_excel.py    # Main script
├── README.md          # This file
└── CLAUDE.md          # Claude Code guidance
```

## German Address Terminology

- **PLZ (Postleitzahl)**: Postal code (5 digits)
- **Stadt**: City
- **Stadtteil**: District/part of a city
- **Ortsteil**: Locality/quarter within a city
- **Bezirk**: Administrative district (larger subdivision)

## Notes

- All original columns are preserved
- The script is non-destructive (original file is not modified)
- Timestamp in output filename prevents accidental overwrites
- Only .xlsx files are processed (not .xls)
- Files starting with `~` (temporary Excel files) are skipped
- Empty cells are handled gracefully

## Support

If you encounter issues:
1. Check the `Parse_Status` column to identify problematic rows
2. Review the `Original_Address` column to see the original data
3. Verify your Excel file contains address data with 5-digit postal codes
4. Ensure the file is in .xlsx format (not .xls)
