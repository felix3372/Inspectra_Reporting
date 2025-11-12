"""
Professional QA Report Generator
A Streamlit application for processing and analyzing lead data with comprehensive validation and reporting.
Production Version - Clean and optimized.
"""

import streamlit as st
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
    
    def run(self) -> None:
        """Run the Streamlit application."""
        st.set_page_config(
            page_title="QA Report Generator",
            page_icon="📊",
            layout="wide"
        )
        
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
                headers = cached_data['headers']
                records = cached_data['records'] 
                cleaned_records = cached_data['cleaned_records']
                optional_columns = cached_data['optional_columns']
            
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
            
            optional_reports = self._show_optional_report_selection(optional_columns)
            
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
            pass
        except Exception as e:
            logger.error(f"Unexpected error: {str(e)}")
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
        """Render the file upload section with hierarchical network path selection."""
        st.markdown('<div class="section-title">📁 Select File</div>', unsafe_allow_html=True)
        
        selection_mode = st.radio(
            "Choose file source:",
            options=["Upload from computer", "Select from network path"],
            horizontal=True,
            help="Upload a file from your computer or select from network shared folder"
        )
        
        uploaded_file = None
        
        if selection_mode == "Upload from computer":
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
                        keys_to_clear = [
                            'processed_data', 'uploaded_file_name', 'date_selected',
                            'date_column', 'selected_date', 'reports', 'date_records',
                            'mtd_records', 'corrections_reviewed', 'correction_summary',
                            'corrected_records'
                        ]
                        for key in keys_to_clear:
                            if key in st.session_state:
                                del st.session_state[key]
                        st.session_state.force_reprocess = True
                        st.rerun()
        else:
            from QA_Report_Helper.file_selector import FileSelector
            
            base_dir = st.text_input(
                "Base Directory:",
                value=r"W:\2025",
                help="Enter the base network path where month folders are stored",
                key="base_dir_input"
            )
            
            if base_dir:
                file_selector = FileSelector(base_dir)
                
                if file_selector.path_exists():
                    st.success(f"✅ Base directory accessible: {base_dir}")
                    
                    months = file_selector.get_month_folders()
                    
                    if months:
                        col1, col2, col3 = st.columns(3)
                        
                        with col1:
                            st.markdown("📅 **Month**")
                            selected_month = st.selectbox(
                                "Select month:",
                                options=months,
                                key="month_selector",
                                label_visibility="collapsed"
                            )
                        
                        if selected_month:
                            campaigns = file_selector.get_campaign_folders(selected_month)
                            
                            if campaigns:
                                with col2:
                                    st.markdown("📂 **Campaign**")
                                    selected_campaign = st.selectbox(
                                        "Select campaign:",
                                        options=campaigns,
                                        key="campaign_selector",
                                        label_visibility="collapsed"
                                    )
                                
                                if selected_campaign:
                                    files = file_selector.get_excel_files(selected_month, selected_campaign)
                                    
                                    if files:
                                        with col3:
                                            st.markdown("📄 **Excel File**")
                                            
                                            file_options = []
                                            for filename, filepath, mod_time, size_mb in files:
                                                display = file_selector.get_file_display_name(filename, mod_time, size_mb)
                                                file_options.append((display, filename, filepath))
                                            
                                            selected_display = st.selectbox(
                                                "Select file:",
                                                options=[opt[0] for opt in file_options],
                                                key="file_selector",
                                                label_visibility="collapsed"
                                            )
                                        
                                        if selected_display:
                                            selected_file_info = next(opt for opt in file_options if opt[0] == selected_display)
                                            _, filename, filepath = selected_file_info
                                            
                                            st.info(f"📂 Path: `{selected_month}` → `{selected_campaign}` → `{filename}`")
                                            
                                            is_valid, message = file_selector.validate_file_access(filepath)
                                            
                                            if is_valid:
                                                file_size_mb = os.path.getsize(filepath) / (1024 * 1024)
                                                file_mod_time = datetime.fromtimestamp(os.path.getmtime(filepath))
                                                
                                                col_info1, col_info2 = st.columns(2)
                                                with col_info1:
                                                    st.write(f"**File Name:** {filename}")
                                                    st.write(f"**Size:** {file_size_mb:.2f} MB")
                                                with col_info2:
                                                    st.write(f"**Modified:** {file_mod_time.strftime('%d-%b-%Y %I:%M %p')}")
                                                    st.write(f"**Location:** `{selected_month}/{selected_campaign}`")
                                                
                                                col_btn1, col_btn2 = st.columns([3, 1])
                                                
                                                with col_btn1:
                                                    if st.button("✅ Confirm and Load File", type="primary", key="confirm_file_load"):
                                                        with st.spinner("Reading file from network..."):
                                                            file_content = file_selector.read_file(filepath)
                                                        
                                                        if file_content:
                                                            import io
                                                            uploaded_file = io.BytesIO(file_content)
                                                            uploaded_file.name = filename
                                                            uploaded_file.size = len(file_content)
                                                            
                                                            st.success(f"✅ File loaded successfully: {filename}")
                                                            st.session_state.network_file = uploaded_file
                                                            st.session_state.network_file_name = filename
                                                            st.session_state.file_loaded = True
                                                        else:
                                                            st.error("❌ Failed to read file from network path")
                                                
                                                with col_btn2:
                                                    if st.button("🗑️ Clear", key="clear_file_selection", help="Clear file selection"):
                                                        keys_to_clear = [
                                                            'network_file', 'network_file_name', 'file_loaded',
                                                            'processed_data', 'uploaded_file_name', 'date_selected',
                                                            'date_column', 'selected_date', 'reports',
                                                            'corrections_reviewed', 'correction_summary', 'corrected_records'
                                                        ]
                                                        for key in keys_to_clear:
                                                            if key in st.session_state:
                                                                del st.session_state[key]
                                                        st.rerun()
                                                
                                                if 'network_file' in st.session_state and st.session_state.get('network_file_name') == filename:
                                                    uploaded_file = st.session_state.network_file
                                                    
                                                    if st.session_state.get('file_loaded'):
                                                        st.success(f"📄 File ready for processing: **{filename}**")
                                                        
                                                        if st.button("❌ Remove File and Start Over", key="remove_loaded_file"):
                                                            keys_to_clear = [
                                                                'network_file', 'network_file_name', 'file_loaded',
                                                                'processed_data', 'uploaded_file_name', 'date_selected',
                                                                'date_column', 'selected_date', 'reports', 'date_records',
                                                                'mtd_records', 'corrections_reviewed', 'correction_summary',
                                                                'corrected_records'
                                                            ]
                                                            for key in keys_to_clear:
                                                                if key in st.session_state:
                                                                    del st.session_state[key]
                                                            st.session_state.force_reprocess = True
                                                            st.success("🔄 File and all cached data removed. Please select a new file.")
                                                            st.rerun()
                                            else:
                                                st.error(f"⚠️ Cannot access file: {message}")
                                    else:
                                        st.warning(f"📂 No Excel files found in: {selected_month}/{selected_campaign}")
                            else:
                                st.warning(f"📂 No campaign folders found in: {selected_month}")
                    else:
                        st.warning(f"📁 No month folders found in: {base_dir}")
                        st.info("Expected structure: `BaseDir/Month/Campaign/ExcelFile.xlsx`")
                else:
                    st.error(f"⚠️ Base directory not accessible: {base_dir}")
                    st.info("Please check:\n- Path exists\n- You have read permissions\n- Network connection is active")
        
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
    
    def _show_optional_report_selection(self, optional_columns: Dict[str, bool]) -> Dict[str, bool]:
        """Show optional report selection checkboxes."""
        optional_reports = {}
        
        if any(optional_columns.values()):
            st.markdown('<div class="section-title">📋 Optional Reports</div>', unsafe_allow_html=True)
            
            col1, col2 = st.columns(2)
            
            with col1:
                if optional_columns.get("Segment Tagging", False):
                    optional_reports["segment"] = st.checkbox(
                        "Include Segment Wise Report (optional)",
                        help="Generate qualified count report by Segment Tagging (uses all qualified records)"
                    )
                else:
                    st.info("📝 Segment Tagging column not found in data")
                    optional_reports["segment"] = False
            
            with col2:
                if optional_columns.get("JT Persona Tagging", False):
                    optional_reports["jt_persona"] = st.checkbox(
                        "Include JT Persona Wise Report (optional)",
                        help="Generate qualified count report by JT Persona Tagging (uses all qualified records)"
                    )
                else:
                    st.info("📝 JT Persona Tagging column not found in data")
                    optional_reports["jt_persona"] = False
        else:
            optional_reports["segment"] = False
            optional_reports["jt_persona"] = False
        
        return optional_reports
    
    def _generate_and_display_reports(
        self, 
        date_records: List[Dict[str, Any]], 
        mtd_records: List[Dict[str, Any]],
        all_records: List[Dict[str, Any]],
        optional_reports: Dict[str, bool],
        selected_date: date
    ) -> None:
        """Generate and display all reports."""
        with st.spinner("Generating reports..."):
            reports = {
                "Combined QA Report": self.report_generator.generate_combined_qa_report(date_records, mtd_records),
                "Agent Wise Summary": self.report_generator.generate_agent_breakdown_report(date_records),
                "Primary Reason Disqualified": self.report_generator.generate_dq_reason_report(date_records)
            }
            
            if optional_reports.get("segment", False):
                reports["Segment Wise Qualified Count"] = self.report_generator.generate_segment_wise_report(all_records)
            
            if optional_reports.get("jt_persona", False):
                reports["JT Persona Wise Qualified Count"] = self.report_generator.generate_jt_persona_wise_report(all_records)
            
            st.session_state.reports = reports
            st.session_state.date_records = date_records
            st.session_state.mtd_records = mtd_records
            
            st.success("✅ Reports generated successfully!")
            st.info(f"📅 Reports generated for: **{selected_date.strftime('%d-%b-%Y')}** | MTD: **From earliest data to {selected_date.strftime('%d-%b-%Y')}**")
            
            self._display_combined_qa_report(reports)
            self._display_optional_reports(reports, optional_reports)
            self._display_core_reports(reports)
            self._show_download_section()
    
    def _display_combined_qa_report(self, reports: Dict[str, List[List[Any]]]) -> None:
        """Display the combined QA report."""
        st.markdown('<div class="custom-heading">📊 QA Summary</div>', unsafe_allow_html=True)
        
        combined_data = reports["Combined QA Report"]
        if combined_data and len(combined_data) > 1:
            table_dict = self._convert_to_table_dict(combined_data)
            st.table(table_dict)
        else:
            st.info("No data available for QA summary.")
    
    def _display_optional_reports(self, reports: Dict[str, List[List[Any]]], optional_reports: Dict[str, bool]) -> None:
        """Display optional reports if they were generated."""
        optional_report_names = ["Segment Wise Qualified Count", "JT Persona Wise Qualified Count"]
        
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
        excel_data = self.excel_exporter.create_excel_report(reports, campaign_id, selected_date)
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
        - Optional segment and JT persona reports (from all qualified records)
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