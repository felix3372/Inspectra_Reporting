"""
QA Report Helper Package
A modular package for QA report generation functionality.
"""

from .config import Config, ExcelStyling
from .exceptions import ValidationError
from .data_processor import DataProcessor
from .report_generator import ReportGenerator
from .excel_exporter import ExcelExporter
from .email_generator import EmailContentGenerator
from .file_selector import FileSelector
from .data_validator import DataValidator

__version__ = "1.0.2"
__author__ = "Felix Markas Salve"

__all__ = [
    "Config",
    "ExcelStyling", 
    "ValidationError",
    "DataProcessor",
    "ReportGenerator",
    "ExcelExporter",
    "EmailContentGenerator",
    "FileSelector",
    "DataValidator"
]