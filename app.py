"""
Professional QA Report Generator
A Streamlit application for processing and analyzing lead data with comprehensive validation and reporting.
Production Version - Clean and optimized.
REFACTORED: Clear button removed, functionality integrated into Confirm and Load File button.
"""

import streamlit as st

st.set_page_config(
            page_title="QA Report Generator",
            page_icon="📊",
            layout="wide"
        )
import logging
from typing import Dict, List, Any, Optional
from datetime import date, datetime
import os

# Import helper modules
from QA_Report_Helper import (
    Config, 
    ExcelStyling,
    ValidationError,
    DataProcessor,
    ReportGenerator,
    ExcelExporter,
    EmailContentGenerator,
    DataValidator
)

# Configure logging
logging.basicConfig(level=logging.WARNING)
logger = logging.getLogger(__name__)


class QAReportApp:
    """Main application class."""
    
    def __init__(self):
        self.config = Config()
        self.processor = DataProcessor(self.config)
        self.report_generator = ReportGenerator()
        self.excel_exporter = ExcelExporter(ExcelStyling())
        self.email_generator = EmailContentGenerator()
        self.validator = DataValidator(self.config)
    
    def _clear_all_session_state(self) -> None:
        """Clear all session state related to file processing."""
        keys_to_clear = [
            'processed_data', 'uploaded_file_name', 'date_selected',
            'date_column', 'selected_date', 'reports', 'date_records',
            'mtd_records', 'corrections_reviewed', 'correction_summary',
            'corrected_records', 'force_reprocess',
            'mapping_data', 'column_mapping', 'multi_level_defined_combos'
        ]
        for key in keys_to_clear:
            if key in st.session_state:
                del st.session_state[key]
    
    def run(self) -> None:
        """Run the Streamlit application."""
        
        
        self._add_custom_styling()
        self._render_hero_section()
        
        uploaded_file = self._render_file_upload_section()
        
        if not uploaded_file:
            self._show_instructions()
            return
        
        try:
            file_identifier = uploaded_file.name if uploaded_file else None
            cache_valid = (
                'processed_data' in st.session_state and 
                st.session_state.get('uploaded_file_name') == file_identifier and
                not st.session_state.get('force_reprocess', False)
            )
            
            if not cache_valid:
                if 'force_reprocess' in st.session_state:
                    del st.session_state.force_reprocess
                
                with st.spinner("Processing file..."):
                    headers, records = self.processor.load_and_parse_excel(uploaded_file)
                    
                    # Parse optional Mapping sheet (used by Custom Multi-Level Report)
                    mapping_data = self.processor.parse_mapping_sheet(uploaded_file)
                    st.session_state.mapping_data = mapping_data
                    # Reset any prior column mapping since file changed
                    if 'column_mapping' in st.session_state:
                        del st.session_state['column_mapping']
                    
                    # Lead Status correction interface
                    if not st.session_state.get('corrections_reviewed', False):
                        issues, auto_suggestions = self.validator.find_lead_status_issues(records)
                        
                        if issues:
                            st.warning(f"⚠️ Found {len(issues)} Lead Status values that need correction")
                            user_corrections = self._show_correction_interface(issues)
                            
                            if user_corrections is not None:
                                if st.button("✅ Apply Corrections and Continue", type="primary", key="apply_corrections"):
                                    with st.spinner("Applying corrections..."):
                                        corrected_records, correction_count = self.validator.apply_corrections(records, user_corrections)
                                        records = corrected_records
                                    
                                    st.success(f"✅ Applied {len(user_corrections)} corrections to {correction_count} records!")
                                    st.session_state.corrections_reviewed = True
                                    st.session_state.correction_summary = self.validator.get_correction_summary()
                                    st.session_state.corrected_records = corrected_records
                                    st.rerun()
                                else:
                                    return
                            else:
                                if st.button("⏭️ Skip Corrections and Continue", key="skip_corrections"):
                                    st.session_state.corrections_reviewed = True
                                    st.rerun()
                                else:
                                    return
                        else:
                            st.session_state.corrections_reviewed = True
                    
                    if 'corrected_records' in st.session_state:
                        records = st.session_state.corrected_records
                    
                    self._validate_data(headers, records)
                    cleaned_records = self.processor.clean_data(records)
                    optional_columns = self.processor.check_optional_columns(headers)
                    
                    st.session_state.processed_data = {
                        'headers': headers,
                        'records': records,
                        'cleaned_records': cleaned_records,
                        'optional_columns': optional_columns
                    }
                    st.session_state.uploaded_file_name = uploaded_file.name
                    st.session_state.date_selected = False
            else:
                cached_data = st.session_state.processed_data
                headers = cached_data.get('headers', [])
                records = cached_data.get('records', [])
                cleaned_records = cached_data.get('cleaned_records', [])
                optional_columns = cached_data.get('optional_columns', {})
            
            # Date selection
            if not st.session_state.get('date_selected', False):
                selected_date_column, selected_date = self._handle_date_selection(headers, cleaned_records)
                
                if selected_date_column and selected_date:
                    st.session_state.date_column = selected_date_column
                    st.session_state.selected_date = selected_date
                    st.session_state.date_selected = True
                else:
                    return
            else:
                selected_date_column, selected_date = self._handle_date_selection(headers, cleaned_records)
                
                if selected_date != st.session_state.get('selected_date'):
                    st.session_state.date_column = selected_date_column
                    st.session_state.selected_date = selected_date
                    if 'reports' in st.session_state:
                        del st.session_state['reports']
                    st.rerun()
            
            date_column = st.session_state.get('date_column')
            selected_date = st.session_state.get('selected_date')
            self.processor.date_column = date_column
            
            date_filtered_records = self.processor.filter_records_by_date(cleaned_records, selected_date)
            mtd_filtered_records = self.processor.filter_records_mtd(cleaned_records, selected_date)
            
            self._show_data_summary(date_filtered_records, mtd_filtered_records, selected_date)
            
            if st.session_state.get('correction_summary'):
                with st.expander("✅ Data Corrections Applied"):
                    st.text(st.session_state.correction_summary)
            
            optional_reports = self._show_optional_report_selection(optional_columns, headers=headers)
            
            if st.button("📊 Generate QA Reports", type="primary"):
                self._generate_and_display_reports(
                    date_filtered_records, 
                    mtd_filtered_records, 
                    cleaned_records,
                    optional_reports,
                    selected_date
                )
            elif 'reports' in st.session_state:
                self._show_download_section()
        
        except ValidationError as e:
            # Surface the real reason instead of silently swallowing it.
            # load_and_parse_excel() wraps ALL file-read failures into
            # ValidationError, so a bare `pass` here made any load/validation
            # problem render as a blank page with no feedback anywhere.
            logger.error(f"Validation error: {str(e)}", exc_info=True)
            st.error(f"⚠️ {str(e)}")
        except Exception as e:
            logger.error(f"Unexpected error: {str(e)}", exc_info=True)
            st.error(f"⚠️ An unexpected error occurred: {str(e)}")
    
    def _show_correction_interface(self, issues: List[Dict]) -> Optional[Dict[str, str]]:
        """Show interactive correction interface for Lead Status issues."""
        st.markdown('<div class="section-title">🔧 Lead Status Correction</div>', unsafe_allow_html=True)
        st.info("Please review and correct the invalid Lead Status values below. The system has detected variations that don't match the accepted values: 'Qualified' or 'Disqualified'")
        
        user_corrections = {}
        
        for idx, issue in enumerate(issues):
            with st.expander(f"Issue {idx + 1}: '{issue['original']}' ({issue['count']} records)", expanded=True):
                col1, col2 = st.columns([1, 2])
                
                with col1:
                    st.write(f"**Current Value:** `{issue['original']}`")
                    st.write(f"**Records Affected:** {issue['count']:,}")
                    st.write(f"**Valid Options:** Qualified, Disqualified")
                
                with col2:
                    st.write("**Select Correction:**")
                    options = []
                    
                    if issue['auto_suggestion']:
                        options.append(f"{issue['auto_suggestion']} ⭐ (Auto-suggested)")
                    
                    for valid_option in issue['valid_options']:
                        if valid_option != issue['auto_suggestion']:
                            options.append(valid_option)
                    
                    options.append("Keep as is (no correction)")
                    
                    selected = st.radio(
                        f"Correct '{issue['original']}' to:",
                        options=options,
                        key=f"lead_status_correction_{idx}",
                        help="Select the appropriate Lead Status value",
                        label_visibility="collapsed"
                    )
                    
                    if selected != "Keep as is (no correction)":
                        corrected_value = selected.replace(" ⭐ (Auto-suggested)", "").strip()
                        user_corrections[issue['original']] = corrected_value
        
        return user_corrections if user_corrections else None
    
    def _handle_date_selection(self, headers: List[str], records: List[Dict[str, Any]]) -> tuple:
        """Handle date column detection and date selection."""
        st.markdown('<div class="section-title">📅 Date Selection</div>', unsafe_allow_html=True)
        
        date_column = self.processor.detect_date_column(headers)
        
        if date_column:
            st.success(f"✅ Found date column: '{date_column}'")
        else:
            st.warning("⚠️ 'Audit Date' column not found. Please select which column contains dates:")
            column_options = self.processor.get_date_column_options(headers)
            
            if not column_options:
                st.error("No suitable columns found for date selection.")
                st.stop()
            
            date_column = st.selectbox(
                "Select date column:",
                options=column_options,
                help="Choose the column that contains audit dates"
            )
            
            if not date_column:
                return None, None
        
        with st.spinner("Parsing dates..."):
            parsed_dates = self.processor.parse_dates_from_records(records, date_column)
        
        if not parsed_dates:
            st.error(f"❌ Could not parse any valid dates from column '{date_column}'")
            st.info("Expected formats: 06-Nov-25, 06-11-2025, 2025-11-06, etc.")
            return None, None
        
        unique_dates = self.processor.get_unique_dates()
        
        if not unique_dates:
            st.error("No valid dates found in the selected column.")
            return None, None
        
        st.info(f"📊 Found data for {len(unique_dates)} unique dates: {unique_dates[0].strftime('%d-%b-%Y')} to {unique_dates[-1].strftime('%d-%b-%Y')}")
        
        date_options = [d.strftime('%d-%b-%Y') for d in unique_dates]
        selected_date_str = st.selectbox(
            "Select date for report generation:",
            options=date_options,
            index=len(date_options) - 1,
            help="Reports will be generated for this specific date. MTD reports will include all data from the earliest month to this date."
        )
        
        if not selected_date_str:
            return None, None
        
        selected_date = datetime.strptime(selected_date_str, '%d-%b-%Y').date()
        return date_column, selected_date
    
    def _add_custom_styling(self) -> None:
        """Add custom CSS styling."""
        st.markdown("""
        <style>
            .block-container {
                padding-top: 2rem !important;
                padding-bottom: 1rem !important;
            }
            .inspectra-hero {
                background: linear-gradient(135deg, #00e4d0, #00c3ff);
                padding: 1.2rem 2rem 1rem 2rem;
                border-radius: 20px;
                margin-top: 1rem;
                margin-bottom: 1.3rem;
                box-shadow: 0 8px 22px rgba(0,0,0,0.08);
                display: flex;
                justify-content: center;
                animation: fadeInHero 1.2s;
            }
            @keyframes fadeInHero {
                from { opacity: 0; transform: translateY(-32px);}
                to   { opacity: 1; transform: translateY(0);}
            }
            .inspectra-inline {
                display: inline-flex;
                align-items: center;
                gap: 1.3rem;
                white-space: nowrap;
            }
            .inspectra-title {
                font-size: 2.5rem;
                font-weight: 900;
                margin: 0;
                color: #fff;
                letter-spacing: -1.5px;
                text-shadow: 0 2px 10px rgba(0,0,0,0.08);
            }
            .inspectra-divider {
                font-weight: 400;
                color: #004e66;
                opacity: 0.35;
            }
            .inspectra-tagline {
                font-size: 1.08rem;
                font-weight: 500;
                margin: 0;
                color: #e3feff;
                opacity: 0.94;
                position: relative;
                top: 2px;
                letter-spacing: 0.5px;
            }
            .section {
                background: #f6fafd;
                border-radius: 1.2rem;
                padding: 0.8rem 1.6rem 0.5rem 1.6rem;
                margin-bottom: 1.1rem;
                box-shadow: 0 1px 9px 0 rgba(60,95,246,0.10);
                border-left: 5px solid #00c3ff;
                animation: fadeInSection 0.85s;
            }
            @keyframes fadeInSection {
                from { opacity: 0; transform: translateY(36px);}
                to   { opacity: 1; transform: translateY(0);}
            }
            .section-title {
                font-size: 1.15rem;
                font-weight: 700;
                color: #169bb6;
                margin-bottom: 0rem;
                margin-top: 0;
                letter-spacing: -1px;
                display: flex;
                align-items: center;
                gap: 8px;
            }
            .custom-heading {
                font-size: 1.15rem;
                font-weight: 700;
                color: #169bb6;
                margin-bottom: 1rem;
                margin-top: 0;
                letter-spacing: -1px;
                display: flex;
                align-items: center;
                gap: 8px;
            }
        </style>
        """, unsafe_allow_html=True)
    
    def _render_hero_section(self) -> None:
        """Render the hero section."""
        st.markdown("""
        <div class="inspectra-hero">
            <div class="inspectra-inline">
                <span class="inspectra-title">Inspectra</span>
                <span class="inspectra-divider">|</span>
                <span class="inspectra-tagline">QA Report Generator</span>
            </div>
        </div>
        <div class="section">
            <div class="section-title">📊 What is this?</div>
            <b>QA Report Generator</b> is a powerful tool to generate comprehensive QA reports from lead data.<br>
            Upload your Excel file with audit dates and get detailed date-wise and MTD reports.
        </div>
        """, unsafe_allow_html=True)
    
    def _render_file_upload_section(self):
        """Render the file upload section (upload an Excel file from the computer)."""
        st.markdown('<div class="section-title">📁 Select File</div>', unsafe_allow_html=True)

        uploaded_file = st.file_uploader(
            "Choose an Excel file",
            type=list(self.config.SUPPORTED_EXTENSIONS),
            help=f"Supported formats: {', '.join(self.config.SUPPORTED_EXTENSIONS)}"
        )

        if uploaded_file:
            col_info, col_clear = st.columns([4, 1])
            with col_info:
                st.info(f"📄 Selected: **{uploaded_file.name}** ({uploaded_file.size / (1024*1024):.2f} MB)")
            with col_clear:
                if st.button("🗑️ Clear", key="clear_upload", help="Remove uploaded file"):
                    self._clear_all_session_state()
                    st.rerun()

        st.markdown('</div>', unsafe_allow_html=True)
        return uploaded_file
    
    def _show_instructions(self) -> None:
        """Show usage instructions."""
        st.markdown('<div class="section-title">📋 Instructions</div>', unsafe_allow_html=True)
        st.markdown(f"""
        **Required Columns:** {', '.join(self.config.REQUIRED_COLUMNS)}
        
        **Date Column:** Must have 'Audit Date' column or you'll be prompted to select one
        
        **Valid Lead Status Values:** {', '.join(self.config.ACCEPTED_LEAD_STATUS)}
        
        **File Requirements:**
        - Format: Excel (.xlsx or .xlsm)
        - Size: Maximum {self.config.MAX_FILE_SIZE_MB} MB
        - Must contain header row
        - Must have date column with audit dates
        """)
        st.markdown('</div>', unsafe_allow_html=True)
    
    def _validate_data(self, headers: List[str], records: List[Dict[str, Any]]) -> None:
        """Validate all data requirements."""
        validation_steps = [
            ("Checking required columns", lambda: self.processor.validate_columns(headers)),
            ("Validating lead status values", lambda: self.processor.validate_lead_status(records)),
            ("Validating DQ reasons", lambda: self.processor.validate_dq_reasons(records))
        ]
        
        for step_name, validation_func in validation_steps:
            try:
                validation_func()
                st.success(f"✅ {step_name} - Passed")
            except ValidationError as e:
                st.error(f"⚠️ {step_name} - {str(e)}")
                raise
    
    def _show_data_summary(self, date_records: List[Dict[str, Any]], mtd_records: List[Dict[str, Any]], selected_date: date) -> None:
        """Show data summary statistics."""
        st.markdown('<div class="section-title">📈 Data Summary</div>', unsafe_allow_html=True)
        
        daily_total = len(date_records)
        daily_qualified = sum(1 for r in date_records if DataProcessor.normalize(r.get("Lead Status", "")) == "qualified")
        daily_disqualified = sum(1 for r in date_records if DataProcessor.normalize(r.get("Lead Status", "")) == "disqualified")
        
        mtd_total = len(mtd_records)
        mtd_qualified = sum(1 for r in mtd_records if DataProcessor.normalize(r.get("Lead Status", "")) == "qualified")
        mtd_disqualified = sum(1 for r in mtd_records if DataProcessor.normalize(r.get("Lead Status", "")) == "disqualified")
        
        st.write(f"**Selected Date: {selected_date.strftime('%d-%b-%Y')}**")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.write("**📅 Daily Stats (Selected Date)**")
            subcol1, subcol2, subcol3, subcol4 = st.columns(4)
            with subcol1:
                st.metric("Total", daily_total)
            with subcol2:
                st.metric("Qualified", daily_qualified)
            with subcol3:
                st.metric("Disqualified", daily_disqualified)
            with subcol4:
                qual_rate = (daily_qualified / daily_total * 100) if daily_total > 0 else 0
                st.metric("Qual Rate", f"{qual_rate:.1f}%")
        
        with col2:
            st.write("**📊 MTD Stats (Cumulative)**")
            subcol1, subcol2, subcol3, subcol4 = st.columns(4)
            with subcol1:
                st.metric("Total", mtd_total)
            with subcol2:
                st.metric("Qualified", mtd_qualified)
            with subcol3:
                st.metric("Disqualified", mtd_disqualified)
            with subcol4:
                mtd_qual_rate = (mtd_qualified / mtd_total * 100) if mtd_total > 0 else 0
                st.metric("Qual Rate", f"{mtd_qual_rate:.1f}%")
    
    def _scope_choice(self, key: str) -> str:
        """
        Render a small per-report data-scope selector and return the choice as
        "date" (selected date only) or "full" (whole dataset, all dates).
        """
        choice = st.radio(
            "Data scope",
            options=["Selected date", "Full dataset"],
            horizontal=True,
            key=key,
            label_visibility="collapsed",
            help="'Selected date' uses only the chosen date's records; "
                 "'Full dataset' uses all qualified records across every date."
        )
        return "date" if choice == "Selected date" else "full"

    def _show_optional_report_selection(self, optional_columns: Dict[str, bool], headers: List[str] = None) -> Dict[str, Any]:
        """Show optional report selection UI.

        Returns a dict with keys:
          - segment (bool)
          - jt_persona (bool)
          - custom_multi_level (bool)
          - custom_multi_level_columns (List[str]) — QA columns in Mapping-sheet header order
          - custom_multi_level_expected (Dict[str, List[str]]) — expected values per QA column
          - custom_multi_level_combinations (List[tuple]) — valid Mapping-defined combos
        """
        optional_reports: Dict[str, Any] = {
            "segment": False,
            "segment_scope": "date",
            "jt_persona": False,
            "jt_persona_scope": "date",
            "custom_multi_level": False,
            "custom_multi_level_scope": "date",
            "custom_multi_level_columns": [],
            "custom_multi_level_expected": {},
            "custom_multi_level_combinations": [],
        }

        headers = headers or []
        mapping_data = st.session_state.get('mapping_data') or {}

        st.markdown('<div class="section-title">📋 Optional Reports</div>', unsafe_allow_html=True)

        # Row 1: existing single-column shortcut reports
        col1, col2 = st.columns(2)

        with col1:
            if optional_columns.get("Segment Tagging", False):
                optional_reports["segment"] = st.checkbox(
                    "Include Segment Wise Report",
                    help="Qualified count broken down by Segment Tagging"
                )
                if optional_reports["segment"]:
                    optional_reports["segment_scope"] = self._scope_choice("segment_scope")
            else:
                st.info("📝 Segment Tagging column not found — Segment Wise Report unavailable")

        with col2:
            if optional_columns.get("JT Persona Tagging", False):
                optional_reports["jt_persona"] = st.checkbox(
                    "Include JT Persona Wise Report",
                    help="Qualified count broken down by JT Persona Tagging"
                )
                if optional_reports["jt_persona"]:
                    optional_reports["jt_persona_scope"] = self._scope_choice("jt_persona_scope")
            else:
                st.info("📝 JT Persona Tagging column not found — JT Persona Wise Report unavailable")

        # Row 2: Mapping-sheet-driven multi-level report
        st.markdown("")  # spacer

        if not mapping_data:
            # No Mapping sheet → report unavailable. Disable the checkbox
            # entirely (no manual-entry fallback).
            st.checkbox(
                "Include Custom Multi-Level Report",
                value=False,
                disabled=True,
                help="Requires a 'Mapping' sheet in your uploaded QA file."
            )
            st.info("📝 Add a 'Mapping' sheet to your QA file to enable this report")
            return optional_reports

        optional_reports["custom_multi_level"] = st.checkbox(
            "Include Custom Multi-Level Report — driven by your Mapping sheet",
            help=(
                "Groups qualified counts by the columns of your 'Mapping' sheet, "
                "in left-to-right order (leftmost = outermost group). The report "
                "shows only the parent→child combinations defined in the Mapping "
                "sheet (merged cells = hierarchy), so each parent keeps only its "
                "own children. Missing combinations still appear with a count of 0."
            )
        )

        if optional_reports["custom_multi_level"]:
            optional_reports["custom_multi_level_scope"] = self._scope_choice("custom_multi_level_scope")
            selected_cols, expected_values, combinations = self._render_multi_level_mapping_config(
                headers=headers,
                mapping_data=mapping_data,
            )
            optional_reports["custom_multi_level_columns"] = selected_cols
            optional_reports["custom_multi_level_expected"] = expected_values
            optional_reports["custom_multi_level_combinations"] = combinations

        return optional_reports

    def _render_multi_level_mapping_config(
        self,
        headers: List[str],
        mapping_data: Dict[str, Any],
    ) -> tuple:
        """
        Render the ONLY user input for the Custom Multi-Level Report: map each
        Mapping-sheet column name (business name) to the actual QA data column.

        Everything else is inferred from the Mapping sheet:
          - Headers, left-to-right = grouping levels (outermost first)
          - Each row = one valid parent→child combination (hierarchy)

        Returns (selected_columns_in_order, expected_values_dict,
        expected_combinations). All are returned empty when the mapping is
        incomplete or invalid, which causes the report to be skipped from
        output (with a visible note shown here).
        """
        from difflib import get_close_matches

        mapping_columns: List[str] = mapping_data.get("columns", [])
        mapping_values: Dict[str, List[str]] = mapping_data.get("values", {})
        mapping_combos: List[tuple] = mapping_data.get("combinations", [])

        # Filter out metadata columns from selectable options
        data_columns = [h for h in headers if not h.startswith("_")]

        # ---------- Column mapping: business name -> actual QA column ----------
        column_mapping: Dict[str, str] = st.session_state.get('column_mapping', {})

        with st.expander(
            f"🔗 Column Mapping ({len(mapping_columns)} level(s) from your Mapping sheet)",
            expanded=True
        ):
            st.caption(
                "Map each column from your Mapping sheet to the matching column "
                "in the QA data. Best guesses are pre-selected — adjust if needed. "
                "All columns must be mapped for the report to generate."
            )

            new_mapping: Dict[str, str] = {}
            for logical_name in mapping_columns:
                # Best-guess default via fuzzy match
                matches = get_close_matches(logical_name, data_columns, n=1, cutoff=0.3)
                default_qa_col = matches[0] if matches else None

                # Preserve any previous mapping choice
                prior_choice = column_mapping.get(logical_name)
                if prior_choice and prior_choice in data_columns:
                    default_qa_col = prior_choice

                options = ["— Not mapped —"] + data_columns
                default_idx = (
                    options.index(default_qa_col)
                    if default_qa_col in options else 0
                )

                selected = st.selectbox(
                    f'Map "{logical_name}" to QA column:',
                    options=options,
                    index=default_idx,
                    key=f"colmap_{logical_name}"
                )
                if selected != "— Not mapped —":
                    new_mapping[logical_name] = selected

            st.session_state.column_mapping = new_mapping
            column_mapping = new_mapping

        # ---------- Validate mapping completeness ----------
        unmapped = [name for name in mapping_columns if name not in column_mapping]
        if unmapped:
            st.warning(
                "⚠️ Map every Mapping-sheet column before the Custom Multi-Level "
                f"Report can generate. Still unmapped: {', '.join(unmapped)}"
            )
            return [], {}, []

        # ---------- Validate uniqueness (no QA column used twice) ----------
        mapped_cols = [column_mapping[name] for name in mapping_columns]
        duplicates = sorted({c for c in mapped_cols if mapped_cols.count(c) > 1})
        if duplicates:
            st.error(
                "⚠️ Each QA column can map to only one Mapping column. "
                f"Mapped more than once: {', '.join(duplicates)}"
            )
            return [], {}, []

        # ---------- Derive levels, expected values, and valid combinations ----------
        # Order preserved from Mapping-sheet header order (leftmost = outermost).
        # Only the column IDENTITY is mapped; the level VALUES (and therefore the
        # combinations) are unchanged, so combinations align to selected_cols.
        selected_cols = [column_mapping[name] for name in mapping_columns]
        expected_values = {
            column_mapping[name]: mapping_values.get(name, [])
            for name in mapping_columns
        }
        expected_combinations = list(mapping_combos)

        # Confirmation of what will be generated
        level_preview = " → ".join(
            f'{name} ({column_mapping[name]})' for name in mapping_columns
        )
        st.caption(
            f"🧭 Grouping order: {level_preview} — "
            f"{len(expected_combinations)} defined combination(s)"
        )

        return selected_cols, expected_values, expected_combinations
    
    def _generate_and_display_reports(
        self, 
        date_records: List[Dict[str, Any]], 
        mtd_records: List[Dict[str, Any]],
        all_records: List[Dict[str, Any]],
        optional_reports: Dict[str, Any],
        selected_date: date
    ) -> None:
        """Generate and display all reports."""
        with st.spinner("Generating reports..."):
            reports = {
                "Combined QA Report": self.report_generator.generate_combined_qa_report(date_records, mtd_records),
                "Agent Wise Summary": self.report_generator.generate_agent_breakdown_report(date_records),
                "Primary Reason Disqualified": self.report_generator.generate_dq_reason_report(date_records)
            }
            
            # Optional reports use the scope chosen per report: "date" = the
            # selected date's records only; "full" = the whole dataset (all dates).
            def _scoped(scope):
                return date_records if scope == "date" else all_records

            if optional_reports.get("segment", False):
                seg_records = _scoped(optional_reports.get("segment_scope", "date"))
                reports["Segment Wise Qualified Count"] = self.report_generator.generate_segment_wise_report(seg_records)

            if optional_reports.get("jt_persona", False):
                jt_records = _scoped(optional_reports.get("jt_persona_scope", "date"))
                reports["JT Persona Wise Qualified Count"] = self.report_generator.generate_jt_persona_wise_report(jt_records)
            
            # Reset any highlight combos from a prior generation
            st.session_state.pop('multi_level_defined_combos', None)

            if optional_reports.get("custom_multi_level", False):
                selected_cols = optional_reports.get("custom_multi_level_columns", [])
                expected = optional_reports.get("custom_multi_level_expected", {})
                combinations = optional_reports.get("custom_multi_level_combinations", [])
                if selected_cols:
                    ml_records = _scoped(optional_reports.get("custom_multi_level_scope", "date"))
                    report_data = self.report_generator.generate_multi_level_report(
                        ml_records,
                        level_columns=selected_cols,
                        expected_values=expected,
                        expected_combinations=combinations,
                    )
                    reports["Custom Multi-Level Report"] = report_data

                    # Persist the Mapping-defined combinations so the Excel
                    # export can highlight out-of-snapshot rows on download.
                    st.session_state.multi_level_defined_combos = [
                        tuple(c) for c in combinations
                    ]

                    # Combination-level validation: any qualified pairing not
                    # defined in the Mapping sheet is still included as a row,
                    # but flagged. Extra rows are exactly the table rows whose
                    # combination isn't in the defined set.
                    self._warn_out_of_mapping_combos(report_data, selected_cols, combinations)
            
            st.session_state.reports = reports
            st.session_state.date_records = date_records
            st.session_state.mtd_records = mtd_records
            
            st.success("✅ Reports generated successfully!")
            st.info(f"📅 Reports generated for: **{selected_date.strftime('%d-%b-%Y')}** | MTD: **From earliest data to {selected_date.strftime('%d-%b-%Y')}**")
            
            self._display_combined_qa_report(reports)
            self._display_optional_reports(reports, optional_reports)
            self._display_core_reports(reports)
            self._show_download_section()

    def _warn_out_of_mapping_combos(
        self,
        report_data: List[List[Any]],
        level_columns: List[str],
        expected_combinations: List[tuple],
    ) -> None:
        """
        Surface qualified combinations that aren't defined in the Mapping sheet.

        The report already includes them as rows (Warn + include). Any data row
        whose combination isn't in the defined set is an out-of-Mapping pairing
        with qualified leads — list it so the discrepancy is visible. No-op when
        there are no combinations to validate against.
        """
        if not expected_combinations or not report_data or len(report_data) < 2:
            return

        n = len(level_columns)
        defined_set = {tuple(c) for c in expected_combinations}

        extras: List[str] = []
        for row in report_data[1:]:
            if not row or row[0] == "Grand Total":
                continue
            combo = tuple(row[:n])
            if combo not in defined_set:
                count = row[-1]
                extras.append(f"{' → '.join(str(v) for v in combo)} ({count})")

        if extras:
            st.warning(
                "⚠️ Qualified leads found in combinations not defined in your "
                "Mapping sheet (included as extra rows): " + "; ".join(extras)
            )

    def _display_combined_qa_report(self, reports: Dict[str, List[List[Any]]]) -> None:
        """Display the combined QA report."""
        st.markdown('<div class="custom-heading">📊 QA Summary</div>', unsafe_allow_html=True)
        
        combined_data = reports["Combined QA Report"]
        if combined_data and len(combined_data) > 1:
            table_dict = self._convert_to_table_dict(combined_data)
            st.table(table_dict)
        else:
            st.info("No data available for QA summary.")
    
    def _display_optional_reports(self, reports: Dict[str, List[List[Any]]], optional_reports: Dict[str, Any]) -> None:
        """Display optional reports if they were generated."""
        optional_report_names = [
            "Segment Wise Qualified Count",
            "JT Persona Wise Qualified Count",
            "Custom Multi-Level Report",
        ]
        
        for report_name in optional_report_names:
            if report_name in reports:
                st.markdown(f'<div class="custom-heading">📊 {report_name}</div>', unsafe_allow_html=True)
                report_data = reports[report_name]
                if report_data and len(report_data) > 1:
                    table_dict = self._convert_to_table_dict(report_data)
                    st.table(table_dict)
                else:
                    st.info("No qualified data available for this report.")
    
    def _display_core_reports(self, reports: Dict[str, List[List[Any]]]) -> None:
        """Display core reports (Agent Wise Summary and Primary Reason Disqualified)."""
        core_report_order = ["Agent Wise Summary", "Primary Reason Disqualified"]
        
        for report_name in core_report_order:
            if report_name in reports:
                st.markdown(f'<div class="custom-heading">📊 {report_name}</div>', unsafe_allow_html=True)
                report_data = reports[report_name]
                if report_data and len(report_data) > 1:
                    table_dict = self._convert_to_table_dict(report_data)
                    st.table(table_dict)
                else:
                    st.info("No data available for this report.")
    
    def _convert_to_table_dict(self, report_data: List[List[Any]]) -> Dict[str, List[Any]]:
        """Convert report data to dictionary for st.table() display."""
        table_dict = {}
        headers = report_data[0]
        
        for i, header in enumerate(headers):
            table_dict[header] = [row[i] for row in report_data[1:]]
        
        return table_dict
    
    def _show_download_section(self):
        """Show download section with campaign ID input."""
        st.markdown('<div class="section-title">📥 Download Reports</div>', unsafe_allow_html=True)
        
        campaign_id = st.text_input(
            "Campaign ID",
            placeholder="Enter campaign ID (e.g., CAMP_2024_001 or 6399)",
            help="Enter the campaign ID to include in the Excel report",
            key="campaign_id_input"
        )
        
        valid_campaign_id = self._validate_campaign_id(campaign_id)
        reports = st.session_state.get('reports', {})
        selected_date = st.session_state.get('selected_date')
        
        if reports and valid_campaign_id and selected_date:
            self._render_download_button(reports, valid_campaign_id, selected_date)
        elif reports:
            st.info("💡 Enter a Campaign ID above to enable download")
        else:
            st.info("Generate reports first to enable download")
    
    def _validate_campaign_id(self, campaign_id: str) -> str:
        """Validate and return campaign ID if valid."""
        if campaign_id.strip():
            campaign_id = campaign_id.strip()
            if campaign_id.replace('_', '').replace('-', '').isalnum():
                st.success(f"✅ Campaign ID: {campaign_id}")
                return campaign_id
            else:
                st.warning("⚠️ Campaign ID should contain only letters, numbers, underscores, and hyphens")
        return ""
    
    def _render_download_button(self, reports: Dict[str, List[List[Any]]], campaign_id: str, selected_date: date) -> None:
        """Render the download button and information."""
        multi_level_combos = st.session_state.get('multi_level_defined_combos')
        excel_data = self.excel_exporter.create_excel_report(
            reports, campaign_id, selected_date,
            multi_level_defined_combos=multi_level_combos
        )
        filename = f"QA_Report_{campaign_id}_{selected_date.strftime('%d%b%y')}.xlsx"
        
        st.download_button(
            label="📥 Download Excel Report",
            data=excel_data,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
        st.info(f"""
        **📋 Your Excel report includes:**
        - Email greeting: "Hi Team,"
        - Header: "PFB QA_Report_{campaign_id}_{selected_date.strftime('%d-%b-%y')}"
        - Combined QA summary (MTD PRE QA, MTD POST QA, PRE QA, POST QA)
        - Date-wise Agent Wise Summary (for {selected_date.strftime('%d-%b-%y')})
        - Date-wise Primary Reason Disqualified (for {selected_date.strftime('%d-%b-%y')})
        - Optional Segment, JT Persona, and Custom Multi-Level reports (date-wise or full dataset, per your selection)
        - In the Custom Multi-Level Report, rows **highlighted in red** are combinations not defined in your Mapping snapshot
        - Professional formatting ready to copy and paste!
        """)


def main():
    """Application entry point."""
    app = QAReportApp()
    app.run()
    st.markdown("""
    <hr style="margin-top: 3rem; margin-bottom: 1rem;">
    <div style='text-align: center; font-size: 14px; color: #6c757d;'>
        2025 Interlink. All rights reserved. <br>Built by Felix Markas Salve as an internal innovation project.
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()