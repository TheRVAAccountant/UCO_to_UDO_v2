"""
Excel utility functions for working with Excel cells and values.

This module provides utility functions for handling Excel cells,
including extracting values from cells with or without formulas.
"""

from typing import Any, Optional, Union
from openpyxl.cell import Cell
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP


def get_cell_value(cell: Cell) -> Any:
    """
    Get the value of a cell, preserving formulas if present.
    
    Args:
        cell: The Excel cell object
        
    Returns:
        The cell value, which can be a formula string or data value
    """
    if cell.data_type == 'f':
        return cell.value
    else:
        return cell.value


def get_calculated_value(cell: Cell) -> Any:
    """
    Get the calculated value of a cell, even if it contains a formula.
    
    Args:
        cell: The Excel cell object
        
    Returns:
        The calculated value of the cell
    """
    if cell.data_type == 'f':
        return cell._value
    else:
        return cell.value


def safe_convert_to_decimal(value: Any, logger: Any) -> Decimal:
    """
    Safely convert a value to Decimal with error handling.
    
    Args:
        value: The value to convert to Decimal
        logger: A logger instance for recording errors
        
    Returns:
        Decimal: The converted value, or Decimal('0') if conversion fails
    """
    try:
        if value is None or value == "":
            return Decimal('0')  # Default to 0 if the value is None or empty

        # Check if the value is a string that starts with '=', indicating a formula
        if isinstance(value, str) and value.startswith('='):
            logger.error(f"Attempted to convert a formula string to Decimal: {value}")
            return Decimal('0')  # Or handle as needed

        # Convert result to Decimal if it's numeric or looks like a number
        return Decimal(str(value)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError) as e:
        logger.error(f"Invalid value for conversion to Decimal: {value} - Error: {e}")
        return Decimal('0')  # Return 0 for any invalid or unconvertible values


def convert_to_number(value: Any) -> Union[Decimal, Any]:
    """
    Convert a value to a number, handling various formats.
    
    Args:
        value: The value to convert to a number
        
    Returns:
        Decimal if conversion is successful, otherwise the original value
    """
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    elif isinstance(value, str):
        try:
            # Remove commas and handle parentheses for negative numbers
            cleaned_value = value.replace(",", "").replace("(", "-").replace(")", "").strip()
            return Decimal(cleaned_value)
        except ValueError:
            return value
    return value


def convert_to_decimal(value: Any, logger: Any) -> Decimal:
    """
    Convert a value to a Decimal, handling various formats, including accounting format.
    
    Args:
        value: The value to convert to Decimal
        logger: A logger instance for recording warnings
        
    Returns:
        Decimal: The converted value, or Decimal('0') if conversion fails
    """
    import math
    
    if value is None:
        logger.warning(f"Encountered None value, converting to Decimal('0')")
        return Decimal('0')
    
    # Handle blank, accounting-format "-", or NaN
    if isinstance(value, str):
        cleaned_value = value.replace(",", "").replace("$", "").strip()
        
        if cleaned_value in ["-", ""]:  # Accounting-style "-" or blank string
            logger.info(f"Detected accounting-format '-' or blank. Treating as Decimal('0').")
            return Decimal('0')
        try:
            if cleaned_value.startswith("(") and cleaned_value.endswith(")"):
                cleaned_value = "-" + cleaned_value[1:-1]
            return Decimal(cleaned_value).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        except InvalidOperation:
            logger.warning(f"Failed to convert string value: {value}")
            return Decimal('0')
    
    # Handle numeric types (int, float)
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            logger.warning(f"Encountered NaN value, converting to Decimal('0')")
            return Decimal('0')
        return Decimal(str(value)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    
    logger.warning(f"Unexpected value type: {type(value)}")
    return Decimal('0')