"""
Excel export functionality for QA Report Helper package.
Updated to include date information in reports.
"""

import io
from datetime import datetime, date
from typing import Dict, List, Any
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.worksheet.worksheet import Worksheet

from .config import ExcelStyling


class ExcelExporter:
    """Handles Excel file generation with professional formatting."""
    
    def __init__(self, styling: ExcelStyling):
        self.styling = styling
    
    def create_excel_report(
        self, 
        reports: Dict[str, List[List[Any]]], 
        campaign_id: str = "",
        selected_date: date = None
    ) -> bytes:
        """Create formatted Excel report with email content and date information."""
        wb = Workbook()
        ws = wb.active
        ws.title = "QA_Report"
        
        current_row = 1
        
        # Add email content at the top if campaign_id is provided
        if campaign_id:
            date_str = selected_date.strftime("%d-%b-%y") if selected_date else self._format_today_date()
            
            # Email greeting and header
            ws.cell(row=current_row, column=1, value="Hi Team,").font = Font(bold=False, size=11)
            current_row += 2
            
            ws.cell(row=current_row, column=1, value=f"PFB QA_Report_{campaign_id}_{date_str}").font = Font(bold=True, size=12)
            current_row += 2
        
        # Add Combined QA Report first
        if "Combined QA Report" in reports:
            current_row = self._add_report_to_worksheet(ws, reports["Combined QA Report"], current_row)
            current_row += 1
        
        # Add optional reports next (Segment and JT Persona)
        optional_reports = ["Segment Wise Qualified Count", "JT Persona Wise Qualified Count"]
        for report_name in optional_reports:
            if report_name in reports:
                current_row = self._add_report_to_worksheet(ws, reports[report_name], current_row)
                current_row += 1  # Add spacing between reports
        
        # Add core reports last (Agent Wise Summary, Primary Reason Disqualified)
        core_reports = ["Agent Wise Summary", "Primary Reason Disqualified"]
        for report_name in core_reports:
            if report_name in reports:
                current_row = self._add_report_to_worksheet(ws, reports[report_name], current_row)
                current_row += 1  # Add spacing between reports
        
        # Add summary at the end if campaign_id is provided
        if campaign_id:
            ws.cell(row=current_row, column=1, value="Best regards,").font = Font(bold=False, size=11)
            current_row += 1
        
        # Auto-fit columns
        self._auto_fit_columns(ws)
        
        # Save to bytes
        output = io.BytesIO()
        wb.save(output)
        return output.getvalue()
    
    def _format_today_date(self) -> str:
        """Format today's date in the required format (e.g., 2nd-Aug-2025)."""
        today = datetime.now()
        
        # Get day with ordinal suffix
        day = today.day
        if 10 <= day % 100 <= 20:
            suffix = 'th'
        else:
            suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
        
        # Format: 2nd-Aug-2025
        formatted_date = f"{day}{suffix}-{today.strftime('%b')}-{today.year}"
        return formatted_date
    
    def _add_report_to_worksheet(self, ws: Worksheet, report_data: List[List[Any]], start_row: int) -> int:
        """Add a single report to the worksheet with formatting."""
        current_row = start_row
        
        for row_idx, row_data in enumerate(report_data):
            # Add data to cells
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=current_row, column=col_idx, value=value)
                self._apply_cell_formatting(cell, row_idx, len(report_data))
            
            current_row += 1
        
        return current_row
    
    def _apply_cell_formatting(self, cell, row_idx: int, total_rows: int) -> None:
        """Apply formatting to a cell based on its position."""
        cell.alignment = self.styling.center_alignment
        cell.border = self.styling.thin_border
        
        # Header row formatting
        if row_idx == 0:
            cell.fill = self.styling.header_fill
            cell.font = self.styling.header_font
        
        # Grand Total row formatting
        elif (isinstance(cell.value, str) and cell.value.lower().strip() == "grand total") or \
             (row_idx == total_rows - 1 and total_rows > 2):
            cell.fill = self.styling.grand_total_fill
            cell.font = self.styling.grand_total_font
    
    def _auto_fit_columns(self, ws: Worksheet) -> None:
        """Auto-fit column widths based on content."""
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            
            for cell in column:
                try:
                    cell_length = len(str(cell.value)) if cell.value is not None else 0
                    max_length = max(max_length, cell_length)
                except:
                    pass
            
            # Set width with padding
            adjusted_width = min(max_length + 3, 50)  # Cap at 50 characters
            ws.column_dimensions[column_letter].width = adjusted_width