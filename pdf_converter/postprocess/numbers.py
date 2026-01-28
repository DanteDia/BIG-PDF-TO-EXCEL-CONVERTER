"""
Number parsing and conversion utilities.
Handles European number format (. = thousands, , = decimals) to American format.
"""

import re
from typing import Union, Any


def parse_european_number(value: Any) -> float:
    """
    Parse a number in European format (dot=thousands, comma=decimals).
    
    Examples:
        "1.234,56" -> 1234.56
        "4.000,00" -> 4000.0
        "1,234,567.89" -> 1234567.89 (also handles mixed formats)
        "123" -> 123.0
    
    Args:
        value: The value to parse (string, int, or float)
    
    Returns:
        Float value
    """
    if value is None or value == "":
        return 0.0
    
    if isinstance(value, (int, float)):
        return float(value)
    
    s = str(value).strip()
    
    # Handle empty strings
    if not s:
        return 0.0
    
    # Check if already a valid number (no special formatting)
    try:
        return float(s)
    except ValueError:
        pass
    
    # Detect format based on position of . and ,
    # European: "1.234,56" - comma is decimal separator
    # American: "1,234.56" - period is decimal separator
    
    last_comma = s.rfind(',')
    last_period = s.rfind('.')
    
    if last_comma > last_period:
        # European format: comma is decimal separator
        # Remove all periods (thousands separators), replace comma with period
        s = s.replace('.', '').replace(',', '.')
    else:
        # American format or no decimal: remove commas
        s = s.replace(',', '')
    
    try:
        return float(s)
    except ValueError:
        return 0.0


def convert_trailing_negative(value: Any) -> float:
    """
    Convert Gallo-style trailing negative numbers.
    
    In Gallo PDFs, negative numbers have the minus sign at the END:
        "5,212,573.58-" -> -5212573.58
        "1.234,56-" -> -1234.56
    
    Args:
        value: The value to convert
    
    Returns:
        Float value (negative if trailing minus found)
    """
    if value is None or value == "":
        return 0.0
    
    if isinstance(value, (int, float)):
        return float(value)
    
    s = str(value).strip()
    
    # Check for trailing minus
    is_negative = s.endswith('-')
    if is_negative:
        s = s[:-1].strip()
    
    # Parse the number
    result = parse_european_number(s)
    
    return -result if is_negative else result


def convert_parenthesis_negative(value: Any) -> float:
    """
    Convert Visual-style parenthesis negative numbers.
    
    In Visual PDFs, negative numbers are in parentheses:
        "(42.750,09)" -> -42750.09
        "(1.500,00)" -> -1500.00
    
    Args:
        value: The value to convert
    
    Returns:
        Float value (negative if parentheses found)
    """
    if value is None or value == "":
        return 0.0
    
    if isinstance(value, (int, float)):
        return float(value)
    
    s = str(value).strip()
    
    # Check for parentheses
    match = re.match(r'^\((.+)\)$', s)
    if match:
        inner = match.group(1)
        result = parse_european_number(inner)
        return -result
    
    return parse_european_number(s)


def parse_number_auto(value: Any, report_type: str = "gallo") -> float:
    """
    Automatically parse a number based on report type.
    
    Args:
        value: The value to parse
        report_type: "gallo" (trailing minus) or "visual" (parentheses)
    
    Returns:
        Float value
    """
    if report_type == "gallo":
        return convert_trailing_negative(value)
    else:
        return convert_parenthesis_negative(value)


def parse_row_numbers(row: dict, numeric_fields: list, report_type: str = "gallo") -> dict:
    """
    Parse all numeric fields in a row.
    
    Args:
        row: Dictionary representing a row
        numeric_fields: List of field names that should be numeric
        report_type: "gallo" or "visual"
    
    Returns:
        Row with parsed numeric values
    """
    result = row.copy()
    
    for field in numeric_fields:
        if field in result:
            result[field] = parse_number_auto(result[field], report_type)
    
    return result


def format_number_for_excel(value: float, decimals: int = 2) -> float:
    """
    Format a number for Excel output.
    
    Args:
        value: The value to format
        decimals: Number of decimal places
    
    Returns:
        Rounded float value
    """
    if value is None:
        return 0.0
    
    return round(float(value), decimals)


def is_numeric_string(value: str) -> bool:
    """
    Check if a string represents a numeric value.
    """
    if not value or not isinstance(value, str):
        return False
    
    # Remove common number formatting
    cleaned = value.strip()
    cleaned = cleaned.replace('.', '').replace(',', '').replace('-', '').replace('(', '').replace(')', '')
    
    return cleaned.isdigit() or (cleaned.replace('.', '', 1).isdigit())
