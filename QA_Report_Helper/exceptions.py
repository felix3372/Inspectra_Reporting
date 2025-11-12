"""
Custom exceptions for QA Report Helper package.
"""


class ValidationError(Exception):
    """Custom exception for validation errors."""
    pass


class FileProcessingError(Exception):
    """Custom exception for file processing errors."""
    pass


class ReportGenerationError(Exception):
    """Custom exception for report generation errors."""
    pass


class ExcelExportError(Exception):
    """Custom exception for Excel export errors."""
    pass