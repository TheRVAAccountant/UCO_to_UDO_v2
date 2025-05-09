"""
Tests for the Excel utilities module.

This module contains tests for the functions in the excel_utils module.
"""

import pytest
import logging
from decimal import Decimal
from unittest.mock import MagicMock, patch
from src.uco_to_udo_recon.utils.excel_utils import (
    safe_convert_to_decimal,
    convert_to_number,
    convert_to_decimal
)


@pytest.fixture
def mock_logger():
    """Create a mock logger for testing."""
    return MagicMock(spec=logging.Logger)


class TestSafeConvertToDecimal:
    """Tests for the safe_convert_to_decimal function."""
    
    def test_none_value(self, mock_logger):
        """Test converting None value to Decimal."""
        result = safe_convert_to_decimal(None, mock_logger)
        assert result == Decimal('0')
        mock_logger.error.assert_not_called()
    
    def test_empty_string(self, mock_logger):
        """Test converting empty string to Decimal."""
        result = safe_convert_to_decimal("", mock_logger)
        assert result == Decimal('0')
        mock_logger.error.assert_not_called()
    
    def test_formula_string(self, mock_logger):
        """Test converting formula string to Decimal."""
        result = safe_convert_to_decimal("=SUM(A1:A10)", mock_logger)
        assert result == Decimal('0')
        mock_logger.error.assert_called_once()
    
    def test_valid_number(self, mock_logger):
        """Test converting valid number to Decimal."""
        result = safe_convert_to_decimal(123.45, mock_logger)
        assert result == Decimal('123.45')
        mock_logger.error.assert_not_called()
    
    def test_invalid_value(self, mock_logger):
        """Test converting invalid value to Decimal."""
        result = safe_convert_to_decimal("not a number", mock_logger)
        assert result == Decimal('0')
        mock_logger.error.assert_called_once()


class TestConvertToNumber:
    """Tests for the convert_to_number function."""
    
    def test_integer(self):
        """Test converting integer to Decimal."""
        result = convert_to_number(100)
        assert result == Decimal('100')
    
    def test_float(self):
        """Test converting float to Decimal."""
        result = convert_to_number(123.45)
        assert result == Decimal('123.45')
    
    def test_numeric_string(self):
        """Test converting numeric string to Decimal."""
        result = convert_to_number("123.45")
        assert result == Decimal('123.45')
    
    def test_formatted_string(self):
        """Test converting formatted string to Decimal."""
        result = convert_to_number("1,234.56")
        assert result == Decimal('1234.56')
    
    def test_negative_with_parentheses(self):
        """Test converting negative value with parentheses to Decimal."""
        result = convert_to_number("(123.45)")
        assert result == Decimal('-123.45')
    
    def test_non_numeric_string(self):
        """Test converting non-numeric string."""
        result = convert_to_number("not a number")
        assert result == "not a number"


class TestConvertToDecimal:
    """Tests for the convert_to_decimal function."""
    
    def test_none_value(self, mock_logger):
        """Test converting None to Decimal."""
        result = convert_to_decimal(None, mock_logger)
        assert result == Decimal('0')
        mock_logger.warning.assert_called_once()
    
    def test_accounting_format_dash(self, mock_logger):
        """Test converting accounting format dash to Decimal."""
        result = convert_to_decimal("-", mock_logger)
        assert result == Decimal('0')
        mock_logger.info.assert_called_once()
    
    def test_blank_string(self, mock_logger):
        """Test converting blank string to Decimal."""
        result = convert_to_decimal("", mock_logger)
        assert result == Decimal('0')
        mock_logger.info.assert_called_once()
    
    def test_currency_format(self, mock_logger):
        """Test converting currency format to Decimal."""
        result = convert_to_decimal("$1,234.56", mock_logger)
        assert result == Decimal('1234.56')
        mock_logger.info.assert_not_called()
    
    def test_parentheses_negative(self, mock_logger):
        """Test converting parentheses notation for negative values."""
        result = convert_to_decimal("($1,234.56)", mock_logger)
        assert result == Decimal('-1234.56')
        mock_logger.info.assert_not_called()
    
    @patch('math.isnan')
    def test_nan_float(self, mock_isnan, mock_logger):
        """Test converting NaN float to Decimal."""
        mock_isnan.return_value = True
        result = convert_to_decimal(float('nan'), mock_logger)
        assert result == Decimal('0')
        mock_logger.warning.assert_called_once()