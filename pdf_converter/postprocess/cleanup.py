"""
Data cleanup and deduplication utilities.
"""

import re
from typing import List, Dict, Any, Optional
import pandas as pd


def clean_instrument_name(name: str) -> str:
    """
    Clean instrument name by removing currency suffixes.
    
    Visual PDFs append currency to instrument names:
        "CEDEAR APPLE INC. - Pesos" -> "CEDEAR APPLE INC."
        "BONO AL30 - Dolar MEP" -> "BONO AL30"
        "CEDEAR NVIDIA - Dolar Cable" -> "CEDEAR NVIDIA"
    
    Args:
        name: Instrument name
    
    Returns:
        Cleaned instrument name
    """
    if not name or not isinstance(name, str):
        return name or ""
    
    # Common suffixes to remove
    suffixes = [
        r'\s*-\s*Pesos\s*$',
        r'\s*-\s*Dolar\s+MEP\s*$',
        r'\s*-\s*Dólar\s+MEP\s*$',
        r'\s*-\s*Dolar\s+Cable\s*$',
        r'\s*-\s*Dólar\s+Cable\s*$',
        r'\s*-\s*USD\s*$',
        r'\s*-\s*ARS\s*$',
    ]
    
    result = name
    for suffix in suffixes:
        result = re.sub(suffix, '', result, flags=re.IGNORECASE)
    
    return result.strip()


def clean_row(row: dict, text_fields: List[str] = None) -> dict:
    """
    Clean a row by stripping whitespace and cleaning instrument names.
    
    Args:
        row: Dictionary representing a row
        text_fields: List of text fields to clean (defaults to common fields)
    
    Returns:
        Cleaned row
    """
    if text_fields is None:
        text_fields = ["instrumento", "especie", "cod_instrumento", "cod_especie"]
    
    result = {}
    
    for key, value in row.items():
        if value is None:
            # Replace None with appropriate default
            if key in text_fields:
                result[key] = ""
            else:
                result[key] = 0
        elif isinstance(value, str):
            cleaned = value.strip()
            # Clean instrument names
            if key in ["instrumento"]:
                cleaned = clean_instrument_name(cleaned)
            result[key] = cleaned
        else:
            result[key] = value
    
    return result


def deduplicate_rows(rows: List[Dict], key_fields: List[str]) -> List[Dict]:
    """
    Remove duplicate rows based on composite key.
    
    Used to handle overlapping chunks where the same row might be
    extracted multiple times.
    
    Args:
        rows: List of row dictionaries
        key_fields: Fields to use for deduplication key
    
    Returns:
        Deduplicated list of rows
    """
    if not rows or not key_fields:
        return rows
    
    # Use pandas for efficient deduplication
    df = pd.DataFrame(rows)
    
    # Check if all key fields exist
    available_keys = [k for k in key_fields if k in df.columns]
    if not available_keys:
        return rows
    
    # Create composite key
    df["_dedup_key"] = df[available_keys].astype(str).agg("|".join, axis=1)
    
    # Keep first occurrence
    df = df.drop_duplicates(subset=["_dedup_key"], keep="first")
    
    # Remove the temporary key column
    df = df.drop(columns=["_dedup_key"])
    
    return df.to_dict("records")


def fill_missing_entity(rows: List[Dict], entity_field: str = "especie", 
                         code_field: str = "cod_especie") -> List[Dict]:
    """
    Fill missing entity values with the previous row's value.
    
    In Gallo PDFs, the especie name only appears on the first row
    of each group. This function propagates it to subsequent rows.
    
    Args:
        rows: List of row dictionaries
        entity_field: Name of the entity field (e.g., "especie")
        code_field: Name of the code field (e.g., "cod_especie")
    
    Returns:
        Rows with filled entity values
    """
    if not rows:
        return rows
    
    result = []
    last_entity = ""
    last_code = ""
    
    for row in rows:
        new_row = row.copy()
        
        # Check if this row has an entity value
        current_entity = str(row.get(entity_field, "")).strip()
        current_code = str(row.get(code_field, "")).strip()
        
        if current_entity:
            # Update last known entity
            last_entity = current_entity
            last_code = current_code
        else:
            # Fill with last known entity
            if entity_field in new_row or entity_field in row:
                new_row[entity_field] = last_entity
            if code_field in new_row or code_field in row:
                new_row[code_field] = last_code
        
        result.append(new_row)
    
    return result


def remove_empty_rows(rows: List[Dict], required_fields: List[str] = None) -> List[Dict]:
    """
    Remove rows where all specified fields are empty or zero.
    
    Args:
        rows: List of row dictionaries
        required_fields: Fields that must have non-empty values
    
    Returns:
        Filtered list of rows
    """
    if not rows:
        return rows
    
    if required_fields is None:
        # Default: require at least one non-empty field
        return [row for row in rows if any(row.values())]
    
    result = []
    for row in rows:
        has_data = False
        for field in required_fields:
            value = row.get(field)
            if value is not None and value != "" and value != 0:
                has_data = True
                break
        if has_data:
            result.append(row)
    
    return result


def normalize_date(date_str: str) -> str:
    """
    Normalize date string to dd/mm/yyyy format.
    
    Args:
        date_str: Date string in various formats
    
    Returns:
        Normalized date string
    """
    if not date_str or not isinstance(date_str, str):
        return date_str or ""
    
    date_str = date_str.strip()
    
    # Already in correct format
    if re.match(r'^\d{2}/\d{2}/\d{4}$', date_str):
        return date_str
    
    # Try to parse common formats
    patterns = [
        (r'^(\d{1,2})/(\d{1,2})/(\d{4})$', r'\1/\2/\3'),  # d/m/yyyy
        (r'^(\d{4})-(\d{2})-(\d{2})$', r'\3/\2/\1'),  # yyyy-mm-dd
        (r'^(\d{2})-(\d{2})-(\d{4})$', r'\1/\2/\3'),  # dd-mm-yyyy
    ]
    
    for pattern, replacement in patterns:
        if re.match(pattern, date_str):
            return re.sub(pattern, replacement, date_str)
    
    return date_str


def cleanup_section_data(
    rows: List[Dict],
    numeric_fields: List[str],
    dedup_keys: List[str],
    report_type: str = "gallo",
    fill_entity: bool = True
) -> List[Dict]:
    """
    Apply all cleanup operations to section data.
    
    Args:
        rows: List of row dictionaries
        numeric_fields: Fields that should be numeric
        dedup_keys: Fields to use for deduplication
        report_type: "gallo" or "visual"
        fill_entity: Whether to fill missing entity values
    
    Returns:
        Cleaned list of rows
    """
    try:
        from .numbers import parse_row_numbers
    except ImportError:
        from postprocess.numbers import parse_row_numbers
    
    if not rows:
        return rows
    
    # Clean each row
    cleaned = [clean_row(row) for row in rows]
    
    # Parse numeric fields
    cleaned = [parse_row_numbers(row, numeric_fields, report_type) for row in cleaned]
    
    # Fill missing entity values (for Gallo transacciones)
    if fill_entity and report_type == "gallo":
        cleaned = fill_missing_entity(cleaned)
    
    # Deduplicate
    if dedup_keys:
        cleaned = deduplicate_rows(cleaned, dedup_keys)
    
    # Remove empty rows
    cleaned = remove_empty_rows(cleaned)
    
    return cleaned
