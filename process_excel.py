#!/usr/bin/env python3
"""
Excel Data Processing Tool
Splits combined data into separate columns for German addresses, names, and streets.
- Addresses: ZIP/City/Suburb
- Names: First Name/Last Name (with organization detection)
- Streets: Street Name/House Number/House Number Suffix
"""

import pandas as pd
import re
import os
import sys
import json
import argparse
from datetime import datetime
from pathlib import Path


# ============================================================================
# ORGANIZATION DETECTION CONSTANTS
# ============================================================================

# German company legal forms
COMPANY_LEGAL_FORMS = {
    'gmbh', 'ug', 'ag', 'ohg', 'kg', 'partg', 'gbr',
    'e.v.', 'e.v', 'ev', 'eg', 'genossenschaft', 'weg',
    'mbh',   # standalone "mbH" (typo/variant of GmbH)
}

# Institutional prefixes (government, religious, civic)
INSTITUTIONAL_PREFIXES = {
    'stadt', 'gemeinde', 'kirchengemeinde', 'kath.', 'katholische',
    'evang.', 'evangelische', 'zweckverband', 'verband', 'verein',
    'stiftung', 'bund', 'landesamt', 'bundesamt',
    'wohnungseigentümergemeinschaft',  # full form of WEG
    'immo',  # abbreviated "Immobilien" used as company prefix
}

# Business descriptors and technical names
BUSINESS_DESCRIPTORS = {
    'bauunternehmen', 'unternehmen', 'firma', 'betrieb',
    'einrichtungen', 'einrichtung', 'kindergarten', 'kindergärten',
    'pflege', 'dienst', 'dienstleistungen', 'handel', 'handels',
    'gesellschaft', 'genossenschaft', 'stiftung',
    # Property / real estate
    'immobilien', 'immobilienverwaltung', 'hausverwaltung', 'miethausverwaltung',
    'verwaltung', 'wohnkonzepte', 'wohnkonzept', 'wohnungsbau', 'wohnbau',
    'baugesellschaft', 'bausgesellschaft',
    # Medical / professional practices
    'praxis', 'gemeinschaftspraxis', 'arztpraxis', 'zahnarztpraxis',
    'hno-praxis', 'gemeinützige',
    # Crafts / trades
    'stukkateur', 'stuckateur', 'zimmerer', 'schreiner', 'dachdecker',
    'lackierer', 'bodenleger', 'fliesenleger',
}

# Business suffixes: compound words ending in these are business names
# e.g. "Kosmetikinstitut", "Deckendesign", "Personalservice"
BUSINESS_SUFFIXES = {
    'institut', 'design', 'service', 'studio', 'center', 'gruppe',
    'markt', 'handel', 'technik', 'systems', 'solutions', 'consulting',
    'management', 'marketing', 'logistik',
    'schule',  # e.g. "Heilpraktikerschule", "Musikschule", "Fahrschule"
}

# Titles/honorifics to strip before multi-person name parsing
NAME_TITLES = {'herr', 'frau', 'dr.', 'prof.', 'ing.', 'dipl.', 'mag.', 'hr.', 'fr.'}

# Generational suffixes: stripped from end of name and appended to last name
# e.g. "Schweinbenz Siegfried sr." → Vorname="Siegfried", Nachname="Schweinbenz sr."
GENERATIONAL_SUFFIXES = {'sr.', 'jr.', 'sen.', 'jun.', 'senior', 'junior'}

# Prefixes indicating name particles (multi-word last names)
NAME_PARTICLES = {'da', 'de', 'del', 'della', 'van', 'von', 'der', 'den', 'zu', 'zur'}

# Global variable for lazy loading of first names database
INTERNATIONAL_FIRST_NAMES = None

# Global variable for lazy loading of reference street names
REFERENCE_STREET_NAMES = None


# ============================================================================
# PROJECT CONFIGURATION
# ============================================================================

def load_project_config(project_name):
    """
    Load project configuration from projects/<project_name>.json.

    Args:
        project_name: Name of the project (e.g. 'netcom_bw', 'generic')

    Returns:
        dict: Project configuration, or None if not found
    """
    config_path = os.path.join('projects', f'{project_name}.json')
    if not os.path.exists(config_path):
        print(f"❌ Projektconfig nicht gefunden: {config_path}")
        print(f"\nVerfügbare Projekte:")
        list_projects()
        return None
    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)
    return config


def list_projects():
    """List all available project configurations."""
    projects_dir = 'projects'
    if not os.path.exists(projects_dir):
        print("  (kein 'projects/' Ordner gefunden)")
        return
    configs = [f[:-5] for f in os.listdir(projects_dir) if f.endswith('.json')]
    if not configs:
        print("  (keine Projektkonfigurationen gefunden)")
    for name in sorted(configs):
        config_path = os.path.join(projects_dir, f'{name}.json')
        with open(config_path, 'r', encoding='utf-8') as f:
            cfg = json.load(f)
        print(f"  --project {name:20s}  ({cfg.get('project_name', name)}): {cfg.get('description', '')}")


def parse_combined_address(value):
    """
    Parse combined address string into ZIP, City, and Suburb components.
    Handles multiple German address formats.

    Args:
        value: Combined address string (e.g., "10115 Berlin Mitte")

    Returns:
        dict: {"zip": str, "city": str, "suburb": str, "status": str}
    """
    if pd.isna(value) or value == "":
        return {"zip": "", "city": "", "suburb": "", "status": "EMPTY"}

    # Convert to string and strip whitespace
    value = str(value).strip()

    # Pattern 1: Hyphenated City-District (German Standard)
    # Example: "80337 München-Schwabing"
    pattern1 = r'^(\d{5})\s+([A-Za-zäöüßÄÖÜ\s]+)-([A-Za-zäöüßÄÖÜ\s]+)$'
    match = re.match(pattern1, value)
    if match:
        return {
            "zip": match.group(1).strip(),
            "city": match.group(2).strip(),
            "suburb": match.group(3).strip(),
            "status": "SUCCESS"
        }

    # Pattern 2: Space-separated (PLZ + City + District)
    # Example: "10115 Berlin Mitte" or "20095 Hamburg"
    pattern2 = r'^(\d{5})\s+([A-Za-zäöüßÄÖÜ\s]+?)(?:\s+([A-Za-zäöüßÄÖÜ\s]+))?$'
    match = re.match(pattern2, value)
    if match:
        return {
            "zip": match.group(1).strip(),
            "city": match.group(2).strip(),
            "suburb": match.group(3).strip() if match.group(3) else "",
            "status": "SUCCESS"
        }

    # Pattern 3: Comma-separated
    # Example: "12345, Berlin, Mitte" or "12345, Berlin"
    pattern3 = r'^(\d{5})\s*,\s*([^,]+)(?:\s*,\s*(.+))?$'
    match = re.match(pattern3, value)
    if match:
        return {
            "zip": match.group(1).strip(),
            "city": match.group(2).strip(),
            "suburb": match.group(3).strip() if match.group(3) else "",
            "status": "SUCCESS"
        }

    # Pattern 4: Tab-separated
    # Example: "12345\tBerlin\tMitte"
    if '\t' in value:
        parts = value.split('\t')
        parts = [p.strip() for p in parts if p.strip()]
        if len(parts) >= 2 and re.match(r'^\d{5}$', parts[0]):
            return {
                "zip": parts[0],
                "city": parts[1],
                "suburb": parts[2] if len(parts) > 2 else "",
                "status": "SUCCESS"
            }

    # Pattern 5: Semicolon-separated
    # Example: "12345;Berlin;Mitte"
    if ';' in value:
        parts = value.split(';')
        parts = [p.strip() for p in parts if p.strip()]
        if len(parts) >= 2 and re.match(r'^\d{5}$', parts[0]):
            return {
                "zip": parts[0],
                "city": parts[1],
                "suburb": parts[2] if len(parts) > 2 else "",
                "status": "SUCCESS"
            }

    # Pattern 6: Hyphen with commas (mixed format)
    # Example: "80337, München-Schwabing"
    pattern6 = r'^(\d{5})\s*,\s*([^,]+)-([^,]+)$'
    match = re.match(pattern6, value)
    if match:
        return {
            "zip": match.group(1).strip(),
            "city": match.group(2).strip(),
            "suburb": match.group(3).strip(),
            "status": "SUCCESS"
        }

    # Fallback: Try to extract ZIP and remaining text
    zip_match = re.search(r'\b(\d{5})\b', value)
    if zip_match:
        zip_code = zip_match.group(1)
        remaining = value.replace(zip_code, '').strip()
        # Remove leading punctuation
        remaining = re.sub(r'^[,;\s]+', '', remaining)

        # Try to split remaining into city and suburb
        parts = re.split(r'[,;\s]\s*', remaining, maxsplit=1)
        city = parts[0].strip() if len(parts) > 0 else ""
        suburb = parts[1].strip() if len(parts) > 1 else ""

        return {
            "zip": zip_code,
            "city": city,
            "suburb": suburb,
            "status": "PARTIAL"
        }

    # No pattern matched
    return {
        "zip": "",
        "city": "",
        "suburb": "",
        "status": "FAILED"
    }


def parse_street_address(value):
    """
    Parse street address string into Street Name, House Number, and optional Suffix.
    Handles German street address formats.

    Args:
        value: Street address string (e.g., "Starzacher Straße 68", "Im Roßhimmel 35 /1")

    Returns:
        dict: {
            "street_name": str,
            "house_number": str,
            "house_number_suffix": str,
            "status": str (SUCCESS/PARTIAL/FAILED/EMPTY)
        }
    """
    if pd.isna(value) or value == "":
        return {"street_name": "", "house_number": "", "house_number_suffix": "", "status": "EMPTY"}

    # Convert to string and strip whitespace
    value = str(value).strip()

    # Pattern 1: Slash suffix with space
    # Example: "Im Roßhimmel 35 /1", "Neckartalstraße 22 /1"
    pattern1 = r'^(.+?)\s+(\d+)\s*/\s*(\d+)$'
    match = re.match(pattern1, value)
    if match:
        return {
            "street_name": match.group(1).strip(),
            "house_number": match.group(2).strip(),
            "house_number_suffix": match.group(3).strip(),
            "status": "SUCCESS"
        }

    # Pattern 2: Slash suffix without space
    # Example: "Neckartalstraße 39/1", "Burkhardtstraße 42/1"
    pattern2 = r'^(.+?)\s+(\d+)/(\d+)$'
    match = re.match(pattern2, value)
    if match:
        return {
            "street_name": match.group(1).strip(),
            "house_number": match.group(2).strip(),
            "house_number_suffix": match.group(3).strip(),
            "status": "SUCCESS"
        }

    # Pattern 3: Letter suffix with space
    # Example: "Hauptstraße 12 a"
    pattern3 = r'^(.+?)\s+(\d+)\s+([a-zA-Z])$'
    match = re.match(pattern3, value)
    if match:
        return {
            "street_name": match.group(1).strip(),
            "house_number": match.group(2).strip(),
            "house_number_suffix": match.group(3).strip(),
            "status": "SUCCESS"
        }

    # Pattern 4: Letter suffix without space
    # Example: "Hauptstraße 12a"
    pattern4 = r'^(.+?)\s+(\d+)([a-zA-Z])$'
    match = re.match(pattern4, value)
    if match:
        return {
            "street_name": match.group(1).strip(),
            "house_number": match.group(2).strip(),
            "house_number_suffix": match.group(3).strip(),
            "status": "SUCCESS"
        }

    # Pattern 5a: Numeric suffix with space
    # Example: "Flügelstraße 9 1", "Hauptstraße 47 2"
    pattern5a = r'^(.+?)\s+(\d+)\s+(\d+)$'
    match = re.match(pattern5a, value)
    if match:
        return {
            "street_name": match.group(1).strip(),
            "house_number": match.group(2).strip(),
            "house_number_suffix": match.group(3).strip(),
            "status": "SUCCESS"
        }

    # Pattern 5: Standard format (no suffix)
    # Example: "Starzacher Straße 68", "Taläcker 50"
    pattern5 = r'^(.+?)\s+(\d+)$'
    match = re.match(pattern5, value)
    if match:
        return {
            "street_name": match.group(1).strip(),
            "house_number": match.group(2).strip(),
            "house_number_suffix": "",
            "status": "SUCCESS"
        }

    # Fallback: Try to extract what we can
    # Look for last digits as house number
    fallback_match = re.search(r'(.*?)\s*(\d+)\s*(.*)$', value)
    if fallback_match:
        street = fallback_match.group(1).strip()
        number = fallback_match.group(2).strip()
        suffix = fallback_match.group(3).strip()
        return {
            "street_name": street if street else "",
            "house_number": number,
            "house_number_suffix": suffix if suffix else "",
            "status": "PARTIAL"
        }

    # No pattern matched
    return {
        "street_name": "",
        "house_number": "",
        "house_number_suffix": "",
        "status": "FAILED"
    }


def find_combined_column(df):
    """
    Auto-detect the column containing combined address data.

    Args:
        df: pandas DataFrame

    Returns:
        str: Column name or None if not found
    """
    # Keywords to search for in column names
    keywords = ['address', 'adresse', 'ort', 'location', 'plz', 'addr']

    # Search for keywords in column names (case-insensitive)
    # Validate content: combined addresses start with a 5-digit ZIP code followed by text
    for col in df.columns:
        col_lower = str(col).lower()
        for keyword in keywords:
            if keyword in col_lower:
                if df[col].dtype == 'object':
                    sample = df[col].dropna().head(5)
                    if len(sample) > 0 and any(re.search(r'^\d{5}\s', str(val)) for val in sample):
                        return col
                break  # Keyword matched but column is not a combined address, move to next

    # If no keyword found, examine first non-empty string column
    for col in df.columns:
        if df[col].dtype == 'object':  # String column
            # Check if first few non-null values look like combined addresses
            # (start with a 5-digit ZIP code followed by a space and city name)
            sample = df[col].dropna().head(10)
            if len(sample) > 0:
                has_zip = any(re.search(r'^\d{5}\s', str(val)) for val in sample)
                if has_zip:
                    return col

    return None


def find_street_column(df):
    """
    Auto-detect the column containing street address data.

    Args:
        df: pandas DataFrame

    Returns:
        str: Column name or None if not found
    """
    # Keywords to search for in column names
    keywords = ['street', 'strasse', 'straße', 'addr', 'address', 'adresse']

    # Search for keywords in column names (case-insensitive)
    # Validate content: street columns must contain digits (house numbers), not just city names
    for col in df.columns:
        col_lower = str(col).lower()
        for keyword in keywords:
            if keyword in col_lower:
                if df[col].dtype == 'object':
                    sample = df[col].dropna().head(5)
                    if len(sample) > 0 and any(re.search(r'\d', str(val)) for val in sample):
                        return col
                break  # Keyword matched but column is not a street column, move to next

    # If no keyword found, examine first non-empty string column for street patterns
    for col in df.columns:
        if df[col].dtype == 'object':  # String column
            # Check if first few non-null values look like street addresses
            sample = df[col].dropna().head(10)
            if len(sample) > 0:
                match_count = 0
                for val in sample:
                    val_str = str(val).lower()
                    # Check for German street suffixes
                    has_street_suffix = any(suffix in val_str for suffix in ['straße', 'weg', 'gasse', 'platz', 'allee'])
                    # Check for street pattern: ends with number (possibly with suffix)
                    has_street_pattern = re.search(r'\s+\d+(?:/\d+)?$', str(val))

                    if has_street_suffix or has_street_pattern:
                        match_count += 1

                # If at least 50% of samples match street patterns, consider it a street column
                if match_count >= len(sample) * 0.5:
                    return col

    return None


# ============================================================================
# STREET NAME CORRECTION FUNCTIONS
# ============================================================================

def load_reference_street_names(ref_path=None):
    """Load reference street names from Strassennamen.xlsx (lazy loading).

    Args:
        ref_path: Optional custom path to the reference file (from project config).
                  Falls back to 'input/Strassennamen.xlsx' if not provided.
    """
    global REFERENCE_STREET_NAMES
    if REFERENCE_STREET_NAMES is None:
        try:
            ref_file = ref_path if ref_path else 'input/Strassennamen.xlsx'
            if os.path.exists(ref_file):
                df = pd.read_excel(ref_file, engine='openpyxl')
                if 'Strasse' in df.columns:
                    # Get unique street names and clean them
                    REFERENCE_STREET_NAMES = set(df['Strasse'].dropna().unique())
                else:
                    REFERENCE_STREET_NAMES = set()
                    print("  ⚠️  Warning: 'Strasse' column not found in Strassennamen.xlsx")
            else:
                REFERENCE_STREET_NAMES = set()
                print("  ⚠️  Warning: Strassennamen.xlsx not found. Street name correction disabled.")
        except Exception as e:
            REFERENCE_STREET_NAMES = set()
            print(f"  ⚠️  Warning: Could not load Strassennamen.xlsx: {str(e)}")

    return REFERENCE_STREET_NAMES


def correct_street_name(street_name):
    """
    Correct street name based on reference list from Strassennamen.xlsx.

    Uses fuzzy matching to find the best match:
    1. Exact match (case-sensitive)
    2. Case-insensitive match
    3. Fuzzy match (ignoring spaces and special chars)

    Args:
        street_name: Street name to correct

    Returns:
        str: Corrected street name, or original if no match found
    """
    if pd.isna(street_name) or street_name == "":
        return street_name

    ref_names = load_reference_street_names()

    if not ref_names:
        return street_name

    # 1. Exact match
    if street_name in ref_names:
        return street_name

    # 2. Case-insensitive match
    street_lower = street_name.lower()
    for ref_name in ref_names:
        if ref_name.lower() == street_lower:
            return ref_name

    # 3. Fuzzy match (normalize: lowercase, remove spaces)
    def normalize(s):
        return s.lower().replace(' ', '').replace('ß', 'ss')

    street_normalized = normalize(street_name)

    for ref_name in ref_names:
        if normalize(ref_name) == street_normalized:
            return ref_name

    # 4. Partial match (street name starts with reference name or vice versa)
    # This handles cases like "Starzacher Straße Bushaltestelle" vs "Starzacher Straße"
    for ref_name in ref_names:
        if street_normalized.startswith(normalize(ref_name)) or normalize(ref_name).startswith(street_normalized):
            # Return the shorter one (base street name)
            if len(ref_name) <= len(street_name):
                return ref_name

    # No match found - return original
    return street_name


# ============================================================================
# NAME PARSING FUNCTIONS
# ============================================================================

def load_german_first_names():
    """Load international first names database (lazy loading)."""
    global INTERNATIONAL_FIRST_NAMES
    if INTERNATIONAL_FIRST_NAMES is None:
        try:
            from international_first_names import INTERNATIONAL_FIRST_NAMES as names_set
            INTERNATIONAL_FIRST_NAMES = names_set
        except ImportError:
            # Fallback to empty set if database not available
            INTERNATIONAL_FIRST_NAMES = set()
            print("  ⚠️  Warning: International first names database not found. Name detection may be less accurate.")
    return INTERNATIONAL_FIRST_NAMES


def is_first_name(name):
    """
    Check if a name is in the German first names database.

    Args:
        name: Name to check

    Returns:
        bool: True if name is a German first name
    """
    if not name:
        return False

    names_db = load_german_first_names()
    name_lower = name.lower()
    if name_lower in names_db:
        return True
    # Handle hyphenated first names like "Karl-Heinz" or "Hans-Joachim"
    if '-' in name_lower:
        parts = name_lower.split('-')
        if all(p in names_db for p in parts if len(p) > 1):
            return True
    return False


def is_organization(value):
    """
    Check if a value appears to be an organization/company name.

    Args:
        value: String to check

    Returns:
        bool: True if value appears to be an organization
    """
    value_lower = value.lower()

    # Check for legal forms (whole word match to avoid false positives like "Egon" matching "eg")
    # Use word boundaries to ensure we match complete legal forms
    for legal_form in COMPANY_LEGAL_FORMS:
        # For forms with dots (like e.v.), check as-is
        if '.' in legal_form:
            if legal_form in value_lower:
                return True
        else:
            # For forms without dots, use word boundary regex
            # Check if legal form appears as a separate word (with space, dash, or at end)
            pattern = r'\b' + re.escape(legal_form) + r'\b'
            if re.search(pattern, value_lower):
                return True

    # Check for institutional prefixes (at start of name)
    for prefix in INSTITUTIONAL_PREFIXES:
        if value_lower.startswith(prefix):
            return True
        # Also check with space after prefix
        if value_lower.startswith(prefix + ' '):
            return True

    # Check for business descriptors (whole word or end of word match)
    for descriptor in BUSINESS_DESCRIPTORS:
        # Use word boundary for business descriptors too
        pattern = r'\b' + re.escape(descriptor) + r'\b'
        if re.search(pattern, value_lower):
            return True

    # Check for compound business words ending in known business suffixes
    # e.g. "Kosmetikinstitut" ends with "institut", "Deckendesign" ends with "design"
    for word in value_lower.split():
        for suffix in BUSINESS_SUFFIXES:
            if word.endswith(suffix) and len(word) > len(suffix) + 2:
                return True

    # Check for c/o (care-of) — indicates forwarding to a business
    if ' c/o ' in value_lower or value_lower.startswith('c/o '):
        return True

    return False


def find_business_descriptor(value):
    """
    Find business descriptor/technical name in a string.

    Args:
        value: String to search

    Returns:
        str: Business descriptor if found, empty string otherwise
    """
    value_lower = value.lower()

    for descriptor in BUSINESS_DESCRIPTORS:
        if descriptor in value_lower:
            # Find the actual case-preserving match
            pattern = re.compile(re.escape(descriptor), re.IGNORECASE)
            match = pattern.search(value)
            if match:
                return match.group(0)

    return ""


def parse_person_name(value):
    """
    Parse a person's name (not organization) into first and last name.

    Args:
        value: Name string to parse

    Returns:
        dict: {"first_name": str, "last_name": str, "name_type": str,
               "business_indicator": str, "status": str}
    """
    words = value.split()

    # Detect and strip trailing generational suffix (sr., jr., sen., jun.)
    # Strip it, parse the remaining name normally, then append suffix to last name.
    if len(words) >= 2 and words[-1].lower() in GENERATIONAL_SUFFIXES:
        gen_suffix = words[-1]
        remaining = " ".join(words[:-1])
        result = parse_person_name(remaining)
        if result['last_name']:
            result['last_name'] = result['last_name'] + ' ' + gen_suffix
        return result

    # Single word - treat as last name
    if len(words) == 1:
        return {
            "first_name": "",
            "last_name": words[0],
            "name_type": "PERSON",
            "business_indicator": "",
            "status": "PARTIAL"
        }

    # Handle multi-word last names (da Silva, von Müller, etc.)
    # Find name particles like "da", "de", "van", "von"
    particle_index = -1
    for i, word in enumerate(words):
        if word.lower() in NAME_PARTICLES:
            particle_index = i
            break

    if particle_index >= 0 and len(words) > 2:
        # If the word BEFORE the particle is a first name, it is the first name
        # and the particle + rest form the last name.
        # e.g. "Sabine van Bonn-Schäper" → Sabine (first), van Bonn-Schäper (last)
        # e.g. "Idikó von Ketteler-Boeselager" → Idikó (first), von Ketteler-Boeselager (last)
        if particle_index == 1 and is_first_name(words[0]):
            firstname = words[0]
            lastname = " ".join(words[1:])
            return {
                "first_name": firstname,
                "last_name": lastname,
                "name_type": "PERSON",
                "business_indicator": "",
                "status": "SUCCESS"
            }
        # Otherwise particle is inside the last name
        # e.g. "Magalhäes da Silva Jose Carlos"
        #   particle_index = 1 ("da")
        #   lastname = "Magalhäes da Silva" (indices 0-2)
        #   firstname = "Jose Carlos" (indices 3+)
        lastname_end = particle_index + 2  # Include particle and next word
        lastname = " ".join(words[:lastname_end])
        firstname = " ".join(words[lastname_end:])
        return {
            "first_name": firstname,
            "last_name": lastname,
            "name_type": "PERSON",
            "business_indicator": "",
            "status": "SUCCESS"
        }

    # Check if last word is a first name (German convention: LASTNAME FIRSTNAME...)
    if is_first_name(words[-1]):
        # Scan backwards to collect all consecutive first names
        # e.g. "Wilke Jan Peter Christoph" → lastname="Wilke", firstname="Jan Peter Christoph"
        fn_count = 1
        while fn_count < len(words) - 1 and is_first_name(words[-(fn_count + 1)]):
            fn_count += 1
        firstname = " ".join(words[-fn_count:])
        lastname = " ".join(words[:-fn_count])
        return {
            "first_name": firstname,
            "last_name": lastname,
            "name_type": "PERSON",
            "business_indicator": "",
            "status": "SUCCESS"
        }

    # Check if first word is a first name (reversed format: FIRSTNAME [FIRSTNAME...] LASTNAME)
    if is_first_name(words[0]):
        # Count how many consecutive words from the start are first names
        # e.g. "Hans Günter Diecksmeier" → Hans+Günter are both first names
        # e.g. "Fritz Christian Rolf Bachmann" → Fritz+Christian+Rolf are all first names
        fn_count = 1
        while fn_count < len(words) - 1 and is_first_name(words[fn_count]):
            fn_count += 1
        firstname = " ".join(words[:fn_count])
        lastname = " ".join(words[fn_count:])
        return {
            "first_name": firstname,
            "last_name": lastname,
            "name_type": "PERSON",
            "business_indicator": "",
            "status": "SUCCESS"
        }

    # For multi-word names, check second-to-last word
    if len(words) > 2 and is_first_name(words[-2]):
        # Last two words are first names
        lastname = " ".join(words[:-2])
        firstname = " ".join(words[-2:])
        return {
            "first_name": firstname,
            "last_name": lastname,
            "name_type": "PERSON",
            "business_indicator": "",
            "status": "SUCCESS"
        }

    # Fallback: Unknown name order – take first word as first name (universal FIRSTNAME LASTNAME convention).
    # Applies to foreign names (Turkish, Bosnian, Arabic, etc.) not found in the first names database.
    firstname = words[0]
    lastname = " ".join(words[1:])
    return {
        "first_name": firstname,
        "last_name": lastname,
        "name_type": "PERSON",
        "business_indicator": "",
        "status": "PARTIAL"
    }


def parse_name(value):
    """
    Parse name string into components (first name, last name, type).
    Handles persons, organizations, multi-person entries, and business descriptors.

    Args:
        value: Combined name string (e.g., "Gramer Rainer", "Stadt Rottenburg")

    Returns:
        dict: {
            "first_name": str,
            "last_name": str,
            "name_type": str (PERSON/ORGANIZATION/MULTI_PERSON/PERSON_WITH_BUSINESS),
            "business_indicator": str,
            "status": str (SUCCESS/PARTIAL/FAILED/EMPTY)
        }
    """
    if pd.isna(value) or value == "":
        return {
            "first_name": "",
            "last_name": "",
            "name_type": "",
            "business_indicator": "",
            "status": "EMPTY"
        }

    # Convert to string and strip whitespace
    value = str(value).strip()

    # Normalize legal form typos before any other processing
    # e.g. "GmbHmbH" (doubled) → "GmbH"
    value = re.sub(r'(?i)(gmbh)mbh\b', r'\1', value)
    value = re.sub(r'(?i)\bgmbh\s+gmbh\b', 'GmbH', value)

    # Pattern 1: Organization Detection (Highest Priority)
    if is_organization(value):
        # Extract business descriptor if present
        business_term = find_business_descriptor(value)
        return {
            "first_name": "",
            "last_name": value,
            "name_type": "ORGANIZATION",
            "business_indicator": business_term,
            "status": "SUCCESS"
        }

    # Pattern 1.5: "NACHNAME, VORNAME" format (comma as explicit separator)
    # e.g. "Hirt, Tanja" → Nachname="Hirt", Vorname="Tanja"
    if ',' in value:
        parts = [p.strip() for p in value.split(',', 1)]
        if len(parts) == 2 and parts[0] and parts[1]:
            return {
                "first_name": parts[1],
                "last_name": parts[0],
                "name_type": "PERSON",
                "business_indicator": "",
                "status": "SUCCESS"
            }

    # Pattern 2: Multi-Person Format — separators: "und", "u.", "&", "+"
    # Supports:
    #   FIRSTNAME sep FIRSTNAME LASTNAME  → "Günter und Ursula Pieprzyk"
    #   LASTNAME FIRSTNAME sep FIRSTNAME  → "Klein Barbara und Dieter"
    #   "u." as abbreviation for "und"   → "Johner Sabine u. Franz"
    #   "und" in titles/salutations stripped first → "Gabriele Insenberg und Herr Wilhelm Insenberg"
    multi_sep_found = None
    for sep in [' und ', ' u. ', ' & ', ' + ']:
        if sep in value:
            multi_sep_found = sep
            break

    if multi_sep_found:
        # Strip honorific titles before processing (e.g. "Herr", "Frau", "Dr.")
        clean_words = [w for w in value.split() if w.lower() not in NAME_TITLES]
        clean_value = " ".join(clean_words)

        words = clean_value.split()
        if not words:
            pass  # fall through to other patterns
        else:
            # Detect repeated non-separator words as the shared last name
            # e.g. "Gabrile Insenberg und Wilhelm Insenberg" → "Insenberg" appears twice
            sep_tokens = {'und', 'u.', '&', '+'}
            non_sep = [w for w in words if w.lower() not in sep_tokens]
            from collections import Counter
            counts = Counter(w.lower() for w in non_sep)
            repeated = [w for w, cnt in counts.items() if cnt > 1 and len(w) > 2]

            if repeated:
                # Most frequent repeated word is the shared last name
                shared_lc = repeated[0]
                # Recover original capitalisation from first occurrence
                lastname = next(w for w in words if w.lower() == shared_lc)
                # First names = everything except the repeated last name and separator tokens
                firstnames_words = [w for w in words if w.lower() != shared_lc]
                firstnames = " ".join(firstnames_words)
            elif multi_sep_found in [' & ', ' + ']:
                # Determine orientation: LAST FIRST1 & FIRST2  vs  FIRST1 & FIRST2 LAST
                # If the last word is a known first name but the first word is not → German convention
                if is_first_name(words[-1]) and not is_first_name(words[0]):
                    lastname = words[0]
                    firstnames = " ".join(words[1:])
                else:
                    lastname = words[-1]
                    firstnames = " ".join(words[:-1])
            elif is_first_name(words[0]):
                # "und" with first word being a first name → FIRST und FIRST LAST
                lastname = words[-1]
                firstnames = " ".join(words[:-1])
            else:
                # "und" with first word being a last name → LAST FIRST und FIRST
                lastname = words[0]
                firstnames = " ".join(words[1:])

            # Remove shared last name from firstnames (cleanup pass)
            firstnames = re.sub(r'\b' + re.escape(lastname) + r'\b', '', firstnames)
            firstnames = re.sub(r'\s{2,}', ' ', firstnames).strip()

            return {
                "first_name": firstnames,
                "last_name": lastname,
                "name_type": "MULTI_PERSON",
                "business_indicator": "",
                "status": "SUCCESS"
            }

    # Pattern 3: Person with Business Descriptor
    business_term = find_business_descriptor(value)
    if business_term:
        # Remove business descriptor and parse the person name
        clean_name = value.replace(business_term, "").strip()
        result = parse_person_name(clean_name)
        result["name_type"] = "PERSON_WITH_BUSINESS"
        result["business_indicator"] = business_term
        return result

    # Pattern 4: Standard Person Name
    return parse_person_name(value)


def find_name_column(df):
    """
    Auto-detect the column containing name data.

    Args:
        df: pandas DataFrame

    Returns:
        str: Column name or None if not found
    """
    # Keywords to search for in column names
    keywords = ['name', 'namen', 'person', 'contact', 'kontakt']

    # Search for keywords in column names (case-insensitive)
    for col in df.columns:
        col_lower = str(col).lower()
        for keyword in keywords:
            if keyword in col_lower:
                return col

    return None


def export_import_csv(df, input_path, output_dir, import_cfg):
    """
    Export a GEE-import CSV from the processed DataFrame.
    Format: two header rows (technical field names + German labels), then data rows.
    Encoding: UTF-8 without BOM, semicolon-separated.
    """
    sep       = import_cfg.get('separator', ';')
    encoding  = import_cfg.get('encoding', 'utf-8')
    bom       = import_cfg.get('bom', False)
    suffix    = import_cfg.get('filename_suffix', '_import')
    col_defs  = import_cfg.get('columns', [])

    # Build header rows
    row1 = sep.join(c['field'] for c in col_defs)
    row2 = sep.join(c['label'] for c in col_defs)

    def fmt_cell(value, col_type):
        if value is None or (hasattr(value, '__class__') and value.__class__.__name__ == 'float' and str(value) == 'nan'):
            return ''
        try:
            import math
            if isinstance(value, float) and math.isnan(value):
                return ''
        except Exception:
            pass
        if col_type == 'date':
            if hasattr(value, 'strftime'):
                return value.strftime('%d.%m.%Y')
            import re
            s = str(value)
            if re.match(r'\d{2}\.\d{2}\.\d{4}', s):
                return s
            return s
        if col_type == 'integer':
            s = str(value)
            if s.endswith('.0'):
                s = s[:-2]
            return s
        return str(value)

    data_lines = []
    for _, row in df.iterrows():
        fields = []
        for c in col_defs:
            if 'fixed' in c:
                fields.append(c['fixed'])
            elif not c.get('source'):
                fields.append('')
            else:
                src = c['source']
                val = row.get(src, '')
                fields.append(fmt_cell(val, c.get('type', '')))
        data_lines.append(sep.join(fields))

    content = '\n'.join([row1, row2] + data_lines) + '\n'

    input_stem   = Path(input_path).stem
    csv_filename = f"{input_stem}{suffix}.csv"
    csv_path     = os.path.join(output_dir, csv_filename)

    enc = (encoding + '-sig') if bom else encoding
    with open(csv_path, 'w', encoding=enc) as f:
        f.write(content)

    print(f"  📄 Import-CSV exportiert: {csv_filename}")
    return csv_path


def apply_derived(value, transform):
    """Apply a derived transform to a cell value."""
    if transform == 'split_last:-':
        s = str(value) if value is not None else ''
        parts = [p for p in s.split('-') if p.strip()]
        return parts[-1].strip() if parts else ''
    return str(value) if value is not None else ''


def export_csv(df, input_path, output_dir, export_cfg):
    """
    Export a CSV from the processed DataFrame.
    Supports filename_template ({stem}), single header row, derived fields.
    """
    sep       = export_cfg.get('separator', ';')
    encoding  = export_cfg.get('encoding', 'utf-8-sig')
    col_defs  = export_cfg.get('columns', [])
    template  = export_cfg.get('filename_template', '{stem}_export.csv')

    def fmt_cell(value, col_type):
        if value is None:
            return ''
        try:
            import math
            if isinstance(value, float) and math.isnan(value):
                return ''
        except Exception:
            pass
        if col_type == 'date':
            if hasattr(value, 'strftime'):
                return value.strftime('%d.%m.%Y')
            s = str(value)
            if re.match(r'\d{2}\.\d{2}\.\d{4}', s):
                return s
            return s
        if col_type == 'integer':
            s = str(value)
            if s.endswith('.0'):
                s = s[:-2]
            return s
        return str(value)

    header = sep.join(c['field'] for c in col_defs)
    data_lines = []
    for _, row in df.iterrows():
        fields = []
        for c in col_defs:
            if 'fixed' in c:
                fields.append(c['fixed'])
            elif not c.get('source'):
                fields.append('')
            else:
                val = row.get(c['source'], '')
                if 'derived' in c:
                    val = apply_derived(val, c['derived'])
                else:
                    val = fmt_cell(val, c.get('type', ''))
                fields.append(val)
        data_lines.append(sep.join(fields))

    content = '\n'.join([header] + data_lines) + '\n'

    stem = Path(input_path).stem
    csv_filename = template.replace('{stem}', stem)
    csv_path = os.path.join(output_dir, csv_filename)

    with open(csv_path, 'w', encoding=encoding) as f:
        f.write(content)

    print(f"  📄 CSV exportiert: {csv_filename}")
    return csv_path


def process_excel_file(input_path, output_dir, config=None):
    """
    Process a single Excel file: parse addresses and/or names and create output file.

    Args:
        input_path: Path to input Excel file
        output_dir: Directory for output files
        config: Optional project configuration dict (from projects/<name>.json)

    Returns:
        dict: Statistics about the processing
    """
    print(f"\nProcessing: {os.path.basename(input_path)}")

    try:
        # Read Excel file
        df = pd.read_excel(input_path, engine='openpyxl')

        # Determine column mapping: use explicit config values or fall back to auto-detection
        col_config = config.get('columns', {}) if config else {}

        address_col = col_config.get('address') or find_combined_column(df)
        name_col    = col_config.get('name')    or find_name_column(df)
        street_col  = col_config.get('street')  or find_street_column(df)

        # Validate explicitly configured columns exist in the file
        for col_label, col_name in [('name', name_col), ('street', street_col), ('address', address_col)]:
            if col_name and col_name not in df.columns:
                print(f"  ⚠️  Konfigurierte Spalte '{col_name}' nicht in Datei gefunden. Überspringe.")
                if col_label == 'name':    name_col    = None
                if col_label == 'street':  street_col  = None
                if col_label == 'address': address_col = None

        if address_col is None and name_col is None and street_col is None:
            print(f"  ⚠️  Warning: Could not identify address, name, or street columns. Skipping file.")
            return {"success": False, "error": "No processable columns found"}

        stats = {
            "address_processed": False,
            "name_processed": False,
            "street_processed": False,
            "total_rows": len(df)
        }

        # Process address data if found
        if address_col:
            print(f"  📍 Found address column: '{address_col}'")

            # Parse each address
            address_results = []
            for value in df[address_col]:
                address_results.append(parse_combined_address(value))

            # Insert new address columns directly after the original column
            # Remove existing columns first to avoid duplicates
            for col in ['PLZ', 'Stadt', 'Stadtteil']:
                if col in df.columns:
                    df.drop(columns=[col], inplace=True)
            address_col_idx = df.columns.get_loc(address_col)
            df.insert(address_col_idx + 1, 'PLZ', [r['zip'] for r in address_results])
            df.insert(address_col_idx + 2, 'Stadt', [r['city'] for r in address_results])
            df.insert(address_col_idx + 3, 'Stadtteil', [r['suburb'] for r in address_results])

            # Calculate address statistics
            addr_success = sum(1 for r in address_results if r['status'] == 'SUCCESS')
            addr_partial = sum(1 for r in address_results if r['status'] == 'PARTIAL')
            addr_failed = sum(1 for r in address_results if r['status'] == 'FAILED')
            addr_empty = sum(1 for r in address_results if r['status'] == 'EMPTY')

            print(f"     Address parsing: Success: {addr_success}, Partial: {addr_partial}, Failed: {addr_failed}, Empty: {addr_empty}")
            stats["address_processed"] = True
            stats["address_success"] = addr_success
            stats["address_partial"] = addr_partial
            stats["address_failed"] = addr_failed
            stats["address_empty"] = addr_empty

        # Process name data if found
        if name_col:
            print(f"  👤 Found name column: '{name_col}'")

            # Parse each name
            name_results = []
            for value in df[name_col]:
                name_results.append(parse_name(value))

            # Insert new name columns directly after the original column
            # Remove existing columns first to avoid duplicates
            for col in ['Vorname', 'Nachname']:
                if col in df.columns:
                    df.drop(columns=[col], inplace=True)
            name_col_idx = df.columns.get_loc(name_col)
            df.insert(name_col_idx + 1, 'Vorname', [r['first_name'] for r in name_results])
            df.insert(name_col_idx + 2, 'Nachname', [r['last_name'] for r in name_results])

            # Calculate name statistics
            name_success = sum(1 for r in name_results if r['status'] == 'SUCCESS')
            name_partial = sum(1 for r in name_results if r['status'] == 'PARTIAL')
            name_failed = sum(1 for r in name_results if r['status'] == 'FAILED')
            name_empty = sum(1 for r in name_results if r['status'] == 'EMPTY')

            print(f"     Name parsing: Success: {name_success}, Partial: {name_partial}, Failed: {name_failed}, Empty: {name_empty}")
            stats["name_processed"] = True
            stats["name_success"] = name_success
            stats["name_partial"] = name_partial
            stats["name_failed"] = name_failed
            stats["name_empty"] = name_empty

        # Process street data if found
        if street_col:
            print(f"  🛣️  Found street column: '{street_col}'")

            # Parse each street address
            street_results = []
            for value in df[street_col]:
                street_results.append(parse_street_address(value))

            # Insert new street columns directly after the original column
            # Remove existing columns first to avoid duplicates
            for col in ['Straßenname', 'Straßenname (korrigiert)', 'Hausnummer', 'Hausnummernzusatz']:
                if col in df.columns:
                    df.drop(columns=[col], inplace=True)
            street_col_idx = df.columns.get_loc(street_col)
            df.insert(street_col_idx + 1, 'Straßenname', [r['street_name'] for r in street_results])

            # Correct street names only if a reference file is configured
            ref_path = (config.get('reference_files', {}).get('street_names') if config else None)
            next_col_offset = 2
            if ref_path:
                corrected_names = [correct_street_name(r['street_name']) for r in street_results]
                df.insert(street_col_idx + next_col_offset, 'Straßenname (korrigiert)', corrected_names)
                next_col_offset += 1
            else:
                corrected_names = [r['street_name'] for r in street_results]  # No correction

            df.insert(street_col_idx + next_col_offset,     'Hausnummer',        [r['house_number'] for r in street_results])
            df.insert(street_col_idx + next_col_offset + 1, 'Hausnummernzusatz', [r['house_number_suffix'] for r in street_results])

            # Calculate street statistics
            street_success = sum(1 for r in street_results if r['status'] == 'SUCCESS')
            street_partial = sum(1 for r in street_results if r['status'] == 'PARTIAL')
            street_failed = sum(1 for r in street_results if r['status'] == 'FAILED')
            street_empty = sum(1 for r in street_results if r['status'] == 'EMPTY')

            # Count corrections (only if correction was performed)
            corrections = sum(1 for i, r in enumerate(street_results) if ref_path and r['street_name'] and corrected_names[i] != r['street_name'])

            print(f"     Street parsing: Success: {street_success}, Partial: {street_partial}, Failed: {street_failed}, Empty: {street_empty}")
            if corrections > 0:
                print(f"     Street corrections: {corrections} names corrected")
            stats["street_processed"] = True
            stats["street_success"] = street_success
            stats["street_partial"] = street_partial
            stats["street_failed"] = street_failed
            stats["street_empty"] = street_empty
            stats["street_corrections"] = corrections

        # Generate output filename
        input_filename = Path(input_path).stem
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"{input_filename}_processed_{timestamp}.xlsx"
        output_path = os.path.join(output_dir, output_filename)

        # Write to Excel with date formatting
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')

            # Get the workbook and worksheet
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            # Format date columns with dd.mm.yyyy - use config keywords or defaults
            date_keywords = (config.get('date_keywords') if config else None) or \
                            ['datum', 'termin', 'hinzugefügt', 'erstellt']

            # Format phone columns as text - use config keywords or defaults
            phone_keywords = (config.get('phone_keywords') if config else None) or \
                             ['telefon', 'phone', 'tel', 'mobile', 'handy', 'fon']

            for idx, col_name in enumerate(df.columns):
                col_name_lower = str(col_name).lower()

                # Excel column letter (A, B, C, ...)
                col_letter = chr(65 + idx) if idx < 26 else chr(65 + idx // 26 - 1) + chr(65 + idx % 26)

                # Apply date format
                if any(keyword in col_name_lower for keyword in date_keywords):
                    from openpyxl.styles import numbers
                    for row in range(2, len(df) + 2):  # Start from row 2 (after header)
                        cell = worksheet[f'{col_letter}{row}']
                        cell.number_format = 'DD.MM.YYYY'

                # Apply text format to phone numbers
                elif any(keyword in col_name_lower for keyword in phone_keywords):
                    for row in range(2, len(df) + 2):  # Start from row 2 (after header)
                        cell = worksheet[f'{col_letter}{row}']
                        cell.number_format = '@'  # '@' is Excel's text format code

        print(f"  ✅ Successfully processed: {output_filename}")

        # Export GEE import CSV if configured
        import_cfg = config.get('import_export') if config else None
        if import_cfg and import_cfg.get('enabled'):
            export_import_csv(df, input_path, output_dir, import_cfg)

        # Export multiple CSVs if configured
        for export_cfg in (config.get('exports', []) if config else []):
            export_csv(df, input_path, output_dir, export_cfg)

        stats["success"] = True
        stats["output_file"] = output_path

        return stats

    except Exception as e:
        print(f"  ❌ Error processing file: {str(e)}")
        return {"success": False, "error": str(e)}


def setup_directories():
    """Create input and output directories if they don't exist."""
    os.makedirs("input", exist_ok=True)
    os.makedirs("output", exist_ok=True)


def main():
    """Main entry point for the script."""
    print("=" * 70)
    print("Excel Data Processing - Address, Name, and Street Separation")
    print("Processes German addresses (ZIP/City/Suburb), names (First/Last),")
    print("and street addresses (Street/Number/Suffix)")
    print("=" * 70)

    # Check if pandas and openpyxl are installed
    try:
        import pandas
        import openpyxl
    except ImportError:
        print("\n❌ Error: Required packages not found.")
        print("\nPlease install dependencies:")
        print("  pip3 install -r requirements.txt")
        print("\nOr manually:")
        print("  pip3 install pandas openpyxl")
        sys.exit(1)

    # Parse command line arguments
    parser = argparse.ArgumentParser(
        description='Excel Data Processing Tool - Splits German address, name, and street data',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='Beispiele:\n'
               '  python3 process_excel.py --project netcom_bw\n'
               '  python3 process_excel.py --project generic\n'
               '  python3 process_excel.py --list\n'
    )
    parser.add_argument('--project', '-p', metavar='NAME',
                        help='Projektkonfiguration laden (aus projects/<NAME>.json)')
    parser.add_argument('--list', '-l', action='store_true',
                        help='Verfügbare Projekte anzeigen')
    args = parser.parse_args()

    # Show available projects if requested
    if args.list:
        print("\nVerfügbare Projekte:")
        list_projects()
        print()
        sys.exit(0)

    # Load project configuration
    if args.project:
        config = load_project_config(args.project)
        if config is None:
            sys.exit(1)
        input_dir  = config.get('input_dir', 'input')
        output_dir = config.get('output_dir', 'output')
        print(f"\nProjekt: {config.get('project_name', args.project)}")
        print(f"Input:   {os.path.abspath(input_dir)}")
        print(f"Output:  {os.path.abspath(output_dir)}")

        # Pre-load reference street names from config path (if specified)
        ref_path = config.get('reference_files', {}).get('street_names')
        if ref_path:
            load_reference_street_names(ref_path)
    else:
        config = None
        input_dir  = 'input'
        output_dir = 'output'
        print("\nKein Projekt angegeben – automatische Spaltenerkennung aktiv.")
        print("Tipp: python3 process_excel.py --list  zeigt verfügbare Projekte\n")

    # Setup directories
    os.makedirs(input_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    # Find all Excel files in input directory (skip reference files)
    ref_path = (config.get('reference_files', {}).get('street_names') if config else None) or ''
    ref_basename = os.path.basename(ref_path) if ref_path else ''
    file_filter = set(config.get('files', [])) if config else set()  # optional explicit file list
    excel_files = [
        f for f in os.listdir(input_dir)
        if f.endswith('.xlsx') and not f.startswith('~') and f != ref_basename
        and (not file_filter or f in file_filter)
    ]

    if len(excel_files) == 0:
        print(f"\n⚠️  Keine Excel-Dateien in {input_dir}/ gefunden.")
        print(f"\nBitte Excel-Dateien in '{input_dir}/' ablegen und dann erneut ausführen.")
        print(f"\nInput-Verzeichnis: {os.path.abspath(input_dir)}")
        sys.exit(0)

    print(f"\nFound {len(excel_files)} Excel file(s) to process\n")

    # Process each file
    all_stats = []

    for excel_file in excel_files:
        input_path = os.path.join(input_dir, excel_file)
        stats = process_excel_file(input_path, output_dir, config=config)
        if stats['success']:
            all_stats.append(stats)

    # Print summary
    print("\n" + "=" * 70)
    print("Processing Complete!")
    print("=" * 70)

    if len(all_stats) > 0:
        total_files = len(all_stats)
        total_rows = sum(s['total_rows'] for s in all_stats)

        # Check what was processed
        address_processed = any(s.get('address_processed', False) for s in all_stats)
        name_processed = any(s.get('name_processed', False) for s in all_stats)

        print(f"\n📊 Overall Statistics:")
        print(f"   Files processed: {total_files}")
        print(f"   Total rows: {total_rows}")

        # Address statistics
        if address_processed:
            addr_success = sum(s.get('address_success', 0) for s in all_stats)
            addr_partial = sum(s.get('address_partial', 0) for s in all_stats)
            addr_failed = sum(s.get('address_failed', 0) for s in all_stats)
            addr_empty = sum(s.get('address_empty', 0) for s in all_stats)

            print(f"\n📍 Address Parsing:")
            print(f"   Successfully parsed: {addr_success}")
            print(f"   Partially parsed: {addr_partial}")
            print(f"   Failed to parse: {addr_failed}")
            print(f"   Empty cells: {addr_empty}")

        # Name statistics
        if name_processed:
            name_success = sum(s.get('name_success', 0) for s in all_stats)
            name_partial = sum(s.get('name_partial', 0) for s in all_stats)
            name_failed = sum(s.get('name_failed', 0) for s in all_stats)
            name_empty = sum(s.get('name_empty', 0) for s in all_stats)

            print(f"\n👤 Name Parsing:")
            print(f"   Successfully parsed: {name_success}")
            print(f"   Partially parsed: {name_partial}")
            print(f"   Failed to parse: {name_failed}")
            print(f"   Empty cells: {name_empty}")

        # Street statistics
        street_processed = any(s.get('street_processed', False) for s in all_stats)
        if street_processed:
            street_success = sum(s.get('street_success', 0) for s in all_stats)
            street_partial = sum(s.get('street_partial', 0) for s in all_stats)
            street_failed = sum(s.get('street_failed', 0) for s in all_stats)
            street_empty = sum(s.get('street_empty', 0) for s in all_stats)

            print(f"\n🛣️  Street Parsing:")
            print(f"   Successfully parsed: {street_success}")
            print(f"   Partially parsed: {street_partial}")
            print(f"   Failed to parse: {street_failed}")
            print(f"   Empty cells: {street_empty}")

        print(f"\nOutput files:")
        for stats in all_stats:
            print(f"  - {os.path.basename(stats['output_file'])}")

        print(f"\nOutput directory: {os.path.abspath(output_dir)}")
    else:
        print("\nNo files were successfully processed.")

    print()


if __name__ == "__main__":
    main()
