"""
Email content generation for QA Report Helper package.
"""

from datetime import datetime
from typing import Dict, List, Any


class EmailContentGenerator:
    """Generates email content for downloading."""
    
    @staticmethod
    def format_today_date() -> str:
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
    
    @staticmethod
    def create_email_content(campaign_id: str, reports: Dict[str, List[List[Any]]]) -> str:
        """Create email content for downloading."""
        today_date = EmailContentGenerator.format_today_date()
        
        email_content = f"""Subject: QA Report - Campaign {campaign_id}

Hi,

PFB QA_Report_{campaign_id}_{today_date}

Please find the detailed QA analysis below:

"""
        
        # Add each report as simple tables
        for report_name, report_data in reports.items():
            if report_data and len(report_data) > 1:
                email_content += f"{report_name}:\n"
                email_content += "=" * 60 + "\n"
                
                # Create simple table format
                for row_idx, row in enumerate(report_data):
                    if row_idx == 0:  # Header
                        row_text = " | ".join(f"{str(cell):^15}" for cell in row)
                        email_content += row_text + "\n"
                        email_content += "-" * len(row_text) + "\n"
                    else:  # Data rows
                        row_text = " | ".join(f"{str(cell):^15}" for cell in row)
                        email_content += row_text + "\n"
                
                email_content += "\n"
        
        email_content += """Summary:
• Total leads processed and qualification rates are shown in the PRE QA / POST QA section
• Agent-wise performance breakdown helps identify training needs
• Primary disqualification reasons highlight common data quality issues

Please review and let me know if you need any clarification.

Best regards,
QA Team"""
        
        return email_content