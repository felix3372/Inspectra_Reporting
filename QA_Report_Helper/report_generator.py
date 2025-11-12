"""
Report generation logic for QA Report Helper package.
Updated to support date-filtered and MTD reporting.
"""

from typing import Dict, List, Any
from collections import defaultdict, Counter

from .data_processor import DataProcessor


class ReportGenerator:
    """Generates various reports from processed data with date filtering support."""
    
    @staticmethod
    def generate_combined_qa_report(
        date_records: List[Dict[str, Any]], 
        mtd_records: List[Dict[str, Any]]
    ) -> List[List[Any]]:
        """
        Generate combined MTD and daily QA report in a single table.
        
        Args:
            date_records: Records filtered to selected date only
            mtd_records: Records filtered from month start to selected date
        """
        # Daily counts (selected date only)
        daily_total = len(date_records)
        daily_qualified = sum(
            1 for r in date_records 
            if DataProcessor.normalize(r.get("Lead Status", "")) == "qualified"
        )
        
        # MTD counts (month-to-date)
        mtd_total = len(mtd_records)
        mtd_qualified = sum(
            1 for r in mtd_records 
            if DataProcessor.normalize(r.get("Lead Status", "")) == "qualified"
        )
        
        return [
            ["MTD PRE QA", "MTD POST QA", "PRE QA", "POST QA"],
            [mtd_total, mtd_qualified, daily_total, daily_qualified]
        ]
    
    @staticmethod
    def generate_agent_breakdown_report(records: List[Dict[str, Any]]) -> List[List[Any]]:
        """Generate agent-wise breakdown report (date-filtered)."""
        agent_data = defaultdict(lambda: {"qualified": 0, "disqualified": 0})
        
        for record in records:
            agent = record.get("Agent Name", "(Blank)")
            status = DataProcessor.normalize(record.get("Lead Status", ""))
            
            if status in ["qualified", "disqualified"]:
                agent_data[agent][status] += 1
        
        # Calculate metrics and sort
        agent_rows = []
        for agent, counts in agent_data.items():
            grand_total = counts["qualified"] + counts["disqualified"]
            error_pct = counts["disqualified"] / grand_total if grand_total > 0 else 0
            
            agent_rows.append([
                agent, counts["disqualified"], counts["qualified"], grand_total, error_pct
            ])
        
        # Sort by error percentage descending
        agent_rows.sort(key=lambda x: x[4], reverse=True)
        
        # Add grand total row
        if agent_rows:
            sum_disqualified = sum(row[1] for row in agent_rows)
            sum_qualified = sum(row[2] for row in agent_rows)
            sum_total = sum(row[3] for row in agent_rows)
            sum_error_pct = sum_disqualified / sum_total if sum_total > 0 else 0
            
            agent_rows.append([
                "Grand Total", sum_disqualified, sum_qualified, sum_total, sum_error_pct
            ])
        
        # Format for display
        report = [["Agent Name", "Disqualified", "Qualified", "Grand Total", "Error%"]]
        for row in agent_rows:
            display_row = row[:4] + [f"{row[4]:.0%}"]  # Round to whole percent
            report.append(display_row)
        
        return report
        
    @staticmethod
    def generate_segment_wise_report(records: List[Dict[str, Any]]) -> List[List[Any]]:
        """Generate segment-wise qualified count report (uses ALL qualified records)."""
        segment_counts = Counter()
        
        for record in records:
            if DataProcessor.normalize(record.get("Lead Status", "")) == "qualified":
                segment = record.get("Segment Tagging", "(Blank)")
                if segment is None or str(segment).strip() == "":
                    segment = "(Blank)"
                else:
                    segment = str(segment).strip()
                segment_counts[segment] += 1
        
        # Sort by count descending
        segment_rows = []
        for segment, count in segment_counts.most_common():
            segment_rows.append([segment, count])
        
        # Add grand total
        if segment_rows:
            total_qualified = sum(row[1] for row in segment_rows)
            segment_rows.append(["Grand Total", total_qualified])
        
        # Create report with headers
        report = [["Segment Wise", "Qualified Count"]]
        report.extend(segment_rows)
        
        return report
    
    @staticmethod
    def generate_jt_persona_wise_report(records: List[Dict[str, Any]]) -> List[List[Any]]:
        """Generate JT persona-wise qualified count report (uses ALL qualified records)."""
        persona_counts = Counter()
        
        for record in records:
            if DataProcessor.normalize(record.get("Lead Status", "")) == "qualified":
                persona = record.get("JT Persona Tagging", "(Blank)")
                if persona is None or str(persona).strip() == "":
                    persona = "(Blank)"
                else:
                    persona = str(persona).strip()
                persona_counts[persona] += 1
        
        # Sort by count descending
        persona_rows = []
        for persona, count in persona_counts.most_common():
            persona_rows.append([persona, count])
        
        # Add grand total
        if persona_rows:
            total_qualified = sum(row[1] for row in persona_rows)
            persona_rows.append(["Grand Total", total_qualified])
        
        # Create report with headers
        report = [["JT Persona Tagging", "Qualified Count"]]
        report.extend(persona_rows)
        
        return report
    
    @staticmethod
    def generate_dq_reason_report(records: List[Dict[str, Any]]) -> List[List[Any]]:
        """Generate primary reason disqualified report (date-filtered)."""
        total_leads = len(records)
        
        # Count DQ reasons
        reason_counts = Counter()
        for record in records:
            if DataProcessor.normalize(record.get("Lead Status", "")) == "disqualified":
                reason = record.get("DQ Reason", "(Blank)")
                reason_counts[reason] += 1
        
        # Calculate metrics and sort
        reason_rows = []
        for reason, count in reason_counts.items():
            error_pct = count / total_leads if total_leads > 0 else 0
            reason_rows.append([reason, count, error_pct])
        
        # Sort by error percentage descending
        reason_rows.sort(key=lambda x: x[2], reverse=True)
        
        # Add grand total
        if reason_rows:
            total_disqualified = sum(row[1] for row in reason_rows)
            grand_error_pct = total_disqualified / total_leads if total_leads > 0 else 0
            reason_rows.append(["Grand Total", total_disqualified, grand_error_pct])
        
        # Format for display
        report = [["DQ Reason", "Disqualified", "Error%"]]
        for row in reason_rows:
            display_row = [row[0], row[1], f"{row[2]:.0%}"]  # Round to whole percent
            report.append(display_row)
        
        return report