"""
Data processing and validation logic for QA Report Helper package.
Updated to parse data from two sheets: "Qualified" and "Disqualified"
Added support for date-based filtering and reporting.
"""

import logging
from typing import Dict, List, Tuple, Any, Optional
from openpyxl import load_workbook
from datetime import datetime, date
import pandas as pd

from .config import Config
from .exceptions import ValidationError

# Configure logging
logger = logging.getLogger(__name__)


class DataProcessor:
    """Handles data processing and validation logic with date support."""
    
    def __init__(self, config: Config):
        self.config = config
        self.date_column = None
        self.parsed_dates = {}
    
    @staticmethod
    def normalize(value: Any) -> str:
        """Normalize values for comparison."""
        return str(value).strip().lower() if value is not None else ""
    
    def validate_file_size(self, uploaded_file) -> None:
        """Validate file size."""
        if hasattr(uploaded_file, 'size') and uploaded_file.size > self.config.MAX_FILE_SIZE_MB * 1024 * 1024:
            raise ValidationError(f"File size exceeds {self.config.MAX_FILE_SIZE_MB} MB limit")
    
    def load_and_parse_excel(self, uploaded_file) -> Tuple[List[str], List[Dict[str, Any]]]:
        """Load and parse Excel file from two sheets: Qualified and Disqualified."""
        try:
            self.validate_file_size(uploaded_file)
            
            wb = load_workbook(uploaded_file, read_only=True, data_only=True)
            if not wb.worksheets:
                raise ValidationError("No worksheets found in the file")
            
            # Find Qualified and Disqualified sheets (case-insensitive)
            qualified_sheet = None
            disqualified_sheet = None
            
            for sheet in wb.worksheets:
                sheet_name_lower = sheet.title.lower().strip()
                if sheet_name_lower == "qualified":
                    qualified_sheet = sheet
                elif sheet_name_lower == "disqualified":
                    disqualified_sheet = sheet
            
            # Check if at least one sheet exists
            if qualified_sheet is None and disqualified_sheet is None:
                raise ValidationError(
                    'Required sheets not found. File must contain at least one sheet named "Qualified" or "Disqualified" (case-insensitive)'
                )
            
            all_records = []
            all_headers = []
            
            # Process Qualified sheet if it exists
            if qualified_sheet is not None:
                logger.debug(f"Processing 'Qualified' sheet: {qualified_sheet.title}")
                headers, records = self._parse_sheet(qualified_sheet, "Qualified")
                all_headers.extend(headers)
                all_records.extend(records)
                logger.debug(f"Loaded {len(records)} records from Qualified sheet")
            else:
                logger.warning("'Qualified' sheet not found, proceeding without it")
            
            # Process Disqualified sheet if it exists
            if disqualified_sheet is not None:
                logger.debug(f"Processing 'Disqualified' sheet: {disqualified_sheet.title}")
                headers, records = self._parse_sheet(disqualified_sheet, "Disqualified")
                all_headers.extend(headers)
                all_records.extend(records)
                logger.debug(f"Loaded {len(records)} records from Disqualified sheet")
            else:
                logger.warning("'Disqualified' sheet not found, proceeding without it")
            
            if not all_records:
                raise ValidationError("No valid data records found in any sheet")
            
            # Combine all unique headers from both sheets
            unique_headers = []
            seen_headers = set()
            for header in all_headers:
                header_lower = header.lower()
                if header_lower not in seen_headers:
                    unique_headers.append(header)
                    seen_headers.add(header_lower)
            
            logger.debug(f"Successfully loaded {len(all_records)} total records from both sheets")
            return unique_headers, all_records
            
        except ValidationError:
            raise
        except Exception as e:
            logger.error(f"Error loading Excel file: {str(e)}")
            raise ValidationError(f"Error reading file: {str(e)}")
    
    def _parse_sheet(self, worksheet, sheet_type: str) -> Tuple[List[str], List[Dict[str, Any]]]:
        """Parse a single worksheet and return headers and records."""
        if worksheet.max_row < 2:
            logger.warning(f"{sheet_type} sheet contains no data rows")
            return [], []
        
        # Extract data from worksheet
        data = []
        for row in worksheet.iter_rows(values_only=True):
            if any(cell is not None for cell in row):  # Skip completely empty rows
                data.append(row)
        
        if not data:
            logger.warning(f"No valid data found in {sheet_type} sheet")
            return [], []
        
        # Extract headers
        headers = [str(h).strip() if h is not None else f"Column_{i+1}" for i, h in enumerate(data[0])]
        
        # Create records
        records = []
        for row_idx, row in enumerate(data[1:], start=2):
            if not any(cell is not None for cell in row):  # Skip empty rows
                continue
            
            record = {}
            for i, header in enumerate(headers):
                value = row[i] if i < len(row) else None
                record[header] = value
            
            # Add metadata
            record['_row_number'] = row_idx
            record['_sheet_name'] = sheet_type
            
            records.append(record)
        
        return headers, records
    
    def detect_date_column(self, headers: List[str]) -> Optional[str]:
        """Detect if Audit Date column exists, return column name or None."""
        headers_lower = {h.lower(): h for h in headers}
        
        if "audit date" in headers_lower:
            self.date_column = headers_lower["audit date"]
            logger.info(f"Detected date column: {self.date_column}")
            return self.date_column
        
        logger.warning("'Audit Date' column not found")
        return None
    
    def get_date_column_options(self, headers: List[str]) -> List[str]:
        """Get potential date column options from headers."""
        # Filter out metadata columns and system columns
        excluded_lower = ['_row_number', '_sheet_name', 'lead status', 'dq reason', 'agent name']
        
        options = [h for h in headers if h.lower() not in excluded_lower]
        return options
    
    def parse_date(self, date_value: Any) -> Optional[date]:
        """Parse a single date value with multiple format support."""
        if date_value is None or str(date_value).strip() == "" or str(date_value).strip().lower() == 'nan':
            return None
        
        # Handle datetime objects directly (from Excel/pandas)
        from datetime import datetime as dt
        if isinstance(date_value, dt):
            return date_value.date()
        
        # Handle date objects directly
        if isinstance(date_value, date):
            return date_value
        
        date_str = str(date_value).strip()
        
        # Handle datetime strings first (most common from Excel exports)
        if ' ' in date_str and ':' in date_str:
            try:
                date_part = date_str.split(' ')[0]
                return datetime.strptime(date_part, "%Y-%m-%d").date()
            except:
                pass
        
        # Try multiple date formats
        date_formats = [
            "%Y-%m-%d %H:%M:%S",  # 2025-11-06 00:00:00
            "%Y-%m-%d",           # 2025-11-06 (ISO format)
            "%d-%b-%y",           # 06-Nov-25 (primary format)
            "%d-%B-%y",           # 06-November-25 (full month name)
            "%d-%b-%Y",           # 06-Nov-2025 (4-digit year)
            "%d-%B-%Y",           # 06-November-2025
            "%d/%b/%y",           # 06/Nov/25
            "%d/%B/%y",           # 06/November/25
            "%d/%b/%Y",           # 06/Nov/2025
            "%d/%B/%Y",           # 06/November/2025
            "%d.%b.%y",           # 06.Nov.25
            "%d.%B.%y",           # 06.November.25
            "%d-%m-%Y",           # 06-11-2025
            "%d/%m/%Y",           # 06/11/2025
            "%d-%m-%y",           # 06-11-25
            "%d/%m/%y",           # 06/11/25
            "%m/%d/%Y",           # 11/06/2025 (US format)
            "%m-%d-%Y",           # 11-06-2025 (US format)
            "%m/%d/%y",           # 11/06/25 (US format)
            "%m-%d-%y",           # 11-06-25 (US format)
        ]
        
        for fmt in date_formats:
            try:
                parsed_date = datetime.strptime(date_str, fmt).date()
                return parsed_date
            except ValueError:
                continue
        
        # Try Excel serial date numbers
        try:
            if str(date_str).replace('.', '').replace('-', '').isdigit():
                excel_date = float(date_str)
                if excel_date > 59:
                    excel_date -= 1
                base_date = datetime(1900, 1, 1)
                from datetime import timedelta
                parsed_date = base_date + timedelta(days=excel_date-1)
                return parsed_date.date()
        except (ValueError, OverflowError):
            pass
        
        logger.warning(f"Could not parse date: '{date_str}'")
        return None
    
    def parse_dates_from_records(self, records: List[Dict[str, Any]], date_column: str) -> Dict[int, date]:
        """Parse dates from all records and return mapping of record index to parsed date."""
        parsed_dates = {}
        parse_errors = 0
        
        for idx, record in enumerate(records):
            date_value = record.get(date_column)
            parsed_date = self.parse_date(date_value)
            if parsed_date:
                parsed_dates[idx] = parsed_date
            else:
                if date_value is not None and str(date_value).strip() not in ['', 'nan', 'NaN']:
                    parse_errors += 1
        
        self.parsed_dates = parsed_dates
        logger.info(f"Successfully parsed {len(parsed_dates)} dates out of {len(records)} records (Errors: {parse_errors})")
        return parsed_dates
    
    def get_unique_dates(self) -> List[date]:
        """Get sorted list of unique dates from parsed dates."""
        if not self.parsed_dates:
            return []
        
        unique_dates = sorted(set(self.parsed_dates.values()))
        return unique_dates
    
    def filter_records_by_date(self, records: List[Dict[str, Any]], target_date: date) -> List[Dict[str, Any]]:
        """Filter records to only those matching the target date."""
        if not self.date_column:
            logger.error("No date column set for filtering")
            return []
        
        filtered = []
        
        for record in records:
            date_value = record.get(self.date_column)
            parsed_date = self.parse_date(date_value)
            
            if parsed_date and parsed_date == target_date:
                filtered.append(record)
        
        logger.info(f"Filtered {len(filtered)} records for date {target_date}")
        return filtered
    
    def filter_records_mtd(self, records: List[Dict[str, Any]], selected_date: date) -> List[Dict[str, Any]]:
        """
        Filter records for Month-To-Date (from earliest date in dataset to selected date).
        MTD starts from the 1st day of the earliest month with data.
        """
        if not self.date_column:
            logger.error("No date column set for filtering")
            return []
        
        # Find the earliest date in the dataset
        earliest_date = None
        for record in records:
            date_value = record.get(self.date_column)
            parsed_date = self.parse_date(date_value)
            if parsed_date:
                if earliest_date is None or parsed_date < earliest_date:
                    earliest_date = parsed_date
        
        if earliest_date is None:
            logger.warning("No valid dates found in dataset for MTD calculation")
            return []
        
        # Get the 1st day of the earliest month
        mtd_start = date(earliest_date.year, earliest_date.month, 1)
        
        filtered = []
        
        for record in records:
            date_value = record.get(self.date_column)
            parsed_date = self.parse_date(date_value)
            
            if parsed_date and mtd_start <= parsed_date <= selected_date:
                filtered.append(record)
        
        logger.info(f"MTD Filtering: Start={mtd_start} (earliest month), End={selected_date}, Filtered {len(filtered)} records out of {len(records)} total")
        return filtered
    
    def validate_columns(self, headers: List[str]) -> None:
        """Validate required columns are present."""
        headers_lower = [h.lower() for h in headers]
        missing_cols = [
            col for col in self.config.REQUIRED_COLUMNS 
            if col.lower() not in headers_lower
        ]
        
        if missing_cols:
            raise ValidationError(f"Missing required columns: {', '.join(missing_cols)}")
    
    def validate_lead_status(self, records: List[Dict[str, Any]]) -> None:
        """Validate lead status values."""
        unique_statuses = set(
            self.normalize(r.get("Lead Status"))
            for r in records 
            if r.get("Lead Status") is not None
        )
        
        accepted_statuses = [self.normalize(status) for status in self.config.ACCEPTED_LEAD_STATUS]
        unexpected_statuses = [status for status in unique_statuses if status and status not in accepted_statuses]
        
        if unexpected_statuses:
            raise ValidationError(
                f"Invalid Lead Status values: {', '.join(unexpected_statuses)}. "
                f"Allowed values: {', '.join(self.config.ACCEPTED_LEAD_STATUS)}"
            )
    
    def normalize_dq_reason(self, dq_reason: str) -> str:
        """
        Normalize DQ Reason to Proper Case format.
        Accepts all DQ reasons and converts them to title case.
        
        Example: "email bounce back" → "Email Bounce Back"
        Example: "INVALID PHONE NUMBER" → "Invalid Phone Number"
        
        Args:
            dq_reason: Original DQ Reason text
            
        Returns:
            DQ Reason in Proper Case (Title Case)
        """
        if not dq_reason or str(dq_reason).strip() in ["", "-"]:
            return dq_reason
        
        # Clean whitespace and convert to title case
        clean_reason = str(dq_reason).strip()
        return clean_reason.title()
    
    def validate_dq_reasons(self, records: List[Dict[str, Any]]) -> None:
        """
        Normalize DQ reason values for disqualified leads.
        Accepts all DQ reasons but standardizes known variations.
        This is now a normalization step rather than strict validation.
        """
        # No strict validation - just log unique DQ reasons for informational purposes
        unique_reasons = set()
        for record in records:
            lead_status = self.normalize(record.get("Lead Status"))
            
            if lead_status == "disqualified":
                dq_reason = record.get("DQ Reason")
                if dq_reason is not None:
                    clean_reason = str(dq_reason).strip()
                    if clean_reason and clean_reason != "-":
                        unique_reasons.add(clean_reason)
        
        if unique_reasons:
            logger.info(f"Found {len(unique_reasons)} unique DQ Reason values in dataset")
        
        # No ValidationError raised - all DQ reasons are accepted
    
    def check_optional_columns(self, headers: List[str]) -> Dict[str, bool]:
        """Check which optional columns are available in the data."""
        headers_lower = [h.lower() for h in headers]
        optional_availability = {}
        
        for col in self.config.OPTIONAL_COLUMNS:
            optional_availability[col] = col.lower() in headers_lower
        
        return optional_availability
    
    def clean_data(self, records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Clean and standardize data with DQ Reason normalization."""
        cleaned_records = []
        
        for record in records:
            cleaned_record = record.copy()
            
            # Clean key fields
            for field in ["Lead Status", "Agent Name", "DQ Reason"]:
                value = cleaned_record.get(field)
                if value is None or str(value).strip() == "" or str(value).strip() == "-":
                    # For Lead Status, keep empty if missing; for others use (Blank)
                    cleaned_record[field] = "" if field == "Lead Status" else "(Blank)"
                else:
                    cleaned_value = str(value).strip()
                    # Normalize DQ Reason to standardize variations
                    if field == "DQ Reason":
                        cleaned_record[field] = self.normalize_dq_reason(cleaned_value)
                    else:
                        cleaned_record[field] = cleaned_value
            
            cleaned_records.append(cleaned_record)
        
        return cleaned_records
