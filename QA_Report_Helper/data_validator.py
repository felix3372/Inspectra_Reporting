"""
Data validation and correction module for QA Report Helper package.
Provides intelligent Lead Status validation with fuzzy matching and correction suggestions.
"""

import logging
from typing import Dict, List, Tuple, Any, Optional
from difflib import get_close_matches
from collections import Counter

logger = logging.getLogger(__name__)


class DataValidator:
    """Handles Lead Status validation and intelligent correction suggestions."""
    
    def __init__(self, config):
        self.config = config
        self.corrections_applied = {}
    
    @staticmethod
    def normalize_value(value: Any) -> str:
        """Normalize values for comparison."""
        return str(value).strip().lower() if value is not None else ""
    
    def normalize_lead_status(self, status_text: str) -> str:
        """
        Enhanced lead status normalization to handle variations.
        
        Args:
            status_text: Original status text
            
        Returns:
            Normalized status ('Qualified' or 'Disqualified') if pattern matches
        """
        if not status_text or str(status_text).strip() == "":
            return status_text
        
        status_lower = str(status_text).lower().strip()
        
        # Check for qualified variations
        qualified_patterns = ['qualified', 'qualify', 'qual', 'q']
        if any(pattern == status_lower or status_lower.startswith(pattern) for pattern in qualified_patterns):
            return 'Qualified'
        
        # Check for disqualified variations
        disqualified_patterns = ['disqualified', 'disqualify', 'disqual', 'dq', 'dis']
        if any(pattern == status_lower or status_lower.startswith(pattern) for pattern in disqualified_patterns):
            return 'Disqualified'
        
        return status_text
    
    def find_lead_status_issues(self, records: List[Dict[str, Any]]) -> Tuple[List[Dict], Dict[str, str]]:
        """
        Find invalid Lead Status values and suggest corrections.
        
        Args:
            records: List of data records
            
        Returns:
            Tuple of (issues_list, auto_suggestions_dict)
        """
        valid_statuses = set(self.config.ACCEPTED_LEAD_STATUS)
        status_counts = Counter()
        
        # Count all unique Lead Status values
        for record in records:
            status = record.get('Lead Status')
            if status is not None:
                status_counts[status] += 1
        
        issues = []
        auto_suggestions = {}
        
        # Find invalid statuses
        for status, count in status_counts.items():
            if status not in valid_statuses:
                # Try to normalize it
                normalized = self.normalize_lead_status(status)
                auto_suggestion = normalized if normalized != status and normalized in valid_statuses else None
                
                # Get fuzzy matches
                fuzzy_matches = get_close_matches(str(status), list(valid_statuses), n=2, cutoff=0.6)
                
                issue = {
                    'original': status,
                    'count': count,
                    'auto_suggestion': auto_suggestion,
                    'fuzzy_matches': fuzzy_matches,
                    'valid_options': list(valid_statuses)
                }
                
                issues.append(issue)
                
                # Store auto suggestion if available
                if auto_suggestion:
                    auto_suggestions[status] = auto_suggestion
        
        logger.info(f"Found {len(issues)} Lead Status issues affecting {sum(i['count'] for i in issues)} records")
        return issues, auto_suggestions
    
    def apply_corrections(self, records: List[Dict[str, Any]], corrections: Dict[str, str]) -> Tuple[List[Dict[str, Any]], int]:
        """
        Apply user-selected corrections to the entire dataset.
        
        Args:
            records: List of data records
            corrections: Dictionary mapping original values to corrected values
            
        Returns:
            Tuple of (corrected_records, count_of_corrections)
        """
        if not corrections:
            return records, 0
        
        corrected_records = []
        correction_count = 0
        
        for record in records:
            corrected_record = record.copy()
            
            # Apply Lead Status correction if needed
            original_status = record.get('Lead Status')
            if original_status in corrections:
                corrected_record['Lead Status'] = corrections[original_status]
                correction_count += 1
            
            corrected_records.append(corrected_record)
        
        # Store applied corrections
        self.corrections_applied = corrections.copy()
        
        logger.info(f"Applied {len(corrections)} unique corrections to {correction_count} records")
        return corrected_records, correction_count
    
    def get_correction_summary(self) -> str:
        """
        Get a summary of corrections that were applied.
        
        Returns:
            Formatted string with correction summary
        """
        if not self.corrections_applied:
            return "No corrections were applied."
        
        summary = "Lead Status Corrections Applied:\n"
        for original, corrected in self.corrections_applied.items():
            summary += f"  • '{original}' → '{corrected}'\n"
        
        return summary