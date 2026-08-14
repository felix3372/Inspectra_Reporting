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
    def _hierarchical_report(
        records: List[Dict[str, Any]],
        level_columns: List[str],
        expected_combinations: List[tuple],
        count_header: str = "Count",
    ) -> List[List[Any]]:
        """
        Build a report from an explicit set of valid combinations (the tree
        defined by the Mapping sheet). Each defined combination is one row
        (count 0 if no qualified records match). Qualified combinations that
        are NOT defined are appended as extra rows (never hidden), ordered so
        they group under their parent. Values are matched case- and
        whitespace-insensitively and folded to the Mapping spelling.

        `expected_combinations` is an ordered list of value-tuples aligned to
        `level_columns` (tuple[i] is the value for level_columns[i]).
        """
        n = len(level_columns)

        def _val(record, col):
            v = record.get(col)
            if v is None or str(v).strip() == "":
                return "(Blank)"
            return str(v).strip()

        def _norm(s):
            return " ".join(str(s).strip().lower().split())

        # Ordered defined combinations (deduped, Mapping order preserved)
        defined: List[tuple] = list(dict.fromkeys(tuple(c) for c in expected_combinations))
        defined_set = set(defined)

        # Per-level: expected value order (first occurrence) + canonical map
        per_level_order: List[List[str]] = [[] for _ in range(n)]
        per_level_seen: List[set] = [set() for _ in range(n)]
        canonical_maps: List[Dict[str, str]] = [dict() for _ in range(n)]
        for combo in defined:
            for i in range(n):
                v = combo[i]
                nv = _norm(v)
                if nv not in canonical_maps[i]:
                    canonical_maps[i][nv] = v
                if v not in per_level_seen[i]:
                    per_level_seen[i].add(v)
                    per_level_order[i].append(v)

        def _canon(i, record):
            raw = _val(record, level_columns[i])
            return canonical_maps[i].get(_norm(raw), raw)

        # Count qualified records per canonical combination
        combo_counts: Dict[tuple, int] = defaultdict(int)
        for record in records:
            if DataProcessor.normalize(record.get("Lead Status", "")) != "qualified":
                continue
            key = tuple(_canon(i, record) for i in range(n))
            combo_counts[key] += 1

        # All combos to display: defined first, then qualified-present extras
        all_combos: List[tuple] = list(defined)
        for combo in combo_counts:
            if combo not in defined_set:
                all_combos.append(combo)

        # Per-level rank for hierarchical ordering: expected values first (in
        # Mapping order), then any extra values by total count desc, then name.
        per_level_rank: List[Dict[str, int]] = []
        for i in range(n):
            rank = {v: idx for idx, v in enumerate(per_level_order[i])}
            extra_totals: Dict[str, int] = defaultdict(int)
            for combo, c in combo_counts.items():
                v = combo[i]
                if v not in rank:
                    extra_totals[v] += c
            # extras with 0 count (shouldn't happen) still get a stable rank
            for combo in all_combos:
                v = combo[i]
                if v not in rank and v not in extra_totals:
                    extra_totals[v] = 0
            base = len(per_level_order[i])
            for j, v in enumerate(sorted(extra_totals, key=lambda x: (-extra_totals[x], x))):
                rank[v] = base + j
            per_level_rank.append(rank)

        all_combos.sort(key=lambda combo: tuple(per_level_rank[i][combo[i]] for i in range(n)))

        rows: List[List[Any]] = [list(combo) + [combo_counts.get(combo, 0)] for combo in all_combos]

        total = sum(combo_counts.values())
        if rows:
            rows.append(["Grand Total"] + [""] * (n - 1) + [total])

        header = list(level_columns) + [count_header]
        return [header] + rows

    @staticmethod
    def generate_multi_level_report(
        records: List[Dict[str, Any]],
        level_columns: List[str],
        expected_values: Dict[str, List[str]] = None,
        count_header: str = "Count",
        expected_combinations: List[tuple] = None
    ) -> List[List[Any]]:
        """
        Generate a multi-level bifurcation report of qualified counts,
        grouped by 1 to N columns of the caller's choice.

        Two modes:
        - HIERARCHICAL (preferred, when `expected_combinations` is given):
          only the exact combinations defined in the Mapping sheet are shown
          (e.g. PE TAL keeps only its own personas), plus any qualified
          combinations found in the data but not defined (appended, so nothing
          qualified is hidden). Missing defined combinations show 0.
        - CARTESIAN (legacy, when only `expected_values` is given): rows are the
          cross-product of each level's value universe.
        
        Layout for level_columns=["Country", "Segment", "JT Persona"]:
        
            | Country | Segment | JT Persona | Count |
            | US      | Seg A   | Persona 1  | 28    |
            | US      | Seg A   | Persona 2  | 40    |
            | US      | Seg A   | Persona 3  | 0     |
            | Canada  | Seg A   | Persona 1  | 12    |
            ...
            | Grand Total |     |            | 154   |
        
        Each level's "universe" of values comes from:
        - expected_values[column] if provided (ordered as given)
        - Otherwise all values seen in the data for that column, sorted
          by qualified count descending.
        
        When expected_values is provided for a column, values seen in the
        data but NOT in the expected list are still appended (so unexpected
        records are visible for QA). Expected values appear first in the
        mapping-defined order, extras follow sorted by count then name.
        
        Combinations shown are the cartesian product of per-level universes,
        so combinations with zero qualified records still appear as 0 rows.
        
        Only qualified records contribute to the count.
        
        Args:
            records: List of data records
            level_columns: Ordered list of QA data column names to group by
                           (outermost first, innermost last)
            expected_values: Optional {qa_column_name: [ordered expected
                             values]} — usually resolved from a Mapping sheet
                             + user's column mapping
            count_header: Header for the last column (default "Count").
        
        Returns report as list of lists (header row + data rows + grand
        total). Returns just [[count_header]] if level_columns is empty.
        """
        if not level_columns:
            return [[count_header]]

        # Hierarchical mode: only Mapping-defined combinations (+ qualified extras)
        if expected_combinations:
            return ReportGenerator._hierarchical_report(
                records, level_columns, expected_combinations, count_header
            )

        expected_values = expected_values or {}

        def _val(record, col):
            v = record.get(col)
            if v is None or str(v).strip() == "":
                return "(Blank)"
            return str(v).strip()

        def _norm(s):
            # Case- and whitespace-insensitive key for matching data values to
            # the Mapping sheet's expected values, so spelling variants like
            # "Operations & IT / Security" and "Operations & It / Security"
            # collapse into a single Mapping-defined row.
            return " ".join(str(s).strip().lower().split())

        # Per-level canonical map: normalized value -> the exact label from the
        # Mapping sheet (first occurrence wins, preserving Mapping spelling).
        # Levels without expected values are matched as-is (no folding).
        canonical_maps: List[Dict[str, str]] = []
        for col in level_columns:
            cmap: Dict[str, str] = {}
            if col in expected_values:
                for e in expected_values[col]:
                    ne = _norm(e)
                    if ne not in cmap:
                        cmap[ne] = e
            canonical_maps.append(cmap)

        def _canon(i, record, col):
            # Fold a data value to its Mapping label when the normalized forms
            # match; otherwise keep the raw (stripped) value.
            raw = _val(record, col)
            return canonical_maps[i].get(_norm(raw), raw)

        # Walk records once: gather per-level data-derived values and
        # per-level qualified counts (for sorting when no expected values).
        # Only QUALIFIED records contribute — both to the counts and to the
        # value universe — so disqualified-sheet values never appear as rows.
        per_level_data_values: List[set] = [set() for _ in level_columns]
        per_level_counts: List[Dict[str, int]] = [defaultdict(int) for _ in level_columns]

        for record in records:
            is_qualified = (
                DataProcessor.normalize(record.get("Lead Status", "")) == "qualified"
            )
            if not is_qualified:
                continue
            for i, col in enumerate(level_columns):
                cv = _canon(i, record, col)
                per_level_data_values[i].add(cv)
                per_level_counts[i][cv] += 1

        # Build per-level universe (ordered list). Expected first, extras after.
        universes: List[List[str]] = []
        for i, col in enumerate(level_columns):
            if col in expected_values:
                expected = list(expected_values[col])
                seen = set(expected)
                extras = sorted(
                    (v for v in per_level_data_values[i] if v not in seen),
                    key=lambda v: (-per_level_counts[i][v], v)
                )
                universes.append(expected + extras)
            else:
                # Pure data-derived: sort by qualified count desc, then name asc
                universes.append(sorted(
                    per_level_data_values[i],
                    key=lambda v: (-per_level_counts[i][v], v)
                ))

        # If any level has an empty universe (e.g. records is empty),
        # the cartesian product is empty — return just the header
        if any(not u for u in universes):
            header = list(level_columns) + [count_header]
            return [header]

        # Count qualified per combination (canonicalized so case/spacing
        # variants fold into their Mapping-defined row).
        qualified_counts: Dict[tuple, int] = defaultdict(int)
        for record in records:
            if DataProcessor.normalize(record.get("Lead Status", "")) == "qualified":
                key = tuple(
                    _canon(i, record, col) for i, col in enumerate(level_columns)
                )
                qualified_counts[key] += 1
        
        # Cartesian product of universes = full row set (in stable universe order)
        import itertools
        all_combinations = list(itertools.product(*universes))
        
        # Universe index for hierarchical sort — since itertools.product
        # already yields in universe order (leftmost varies slowest), we
        # don't need to re-sort. But we still may want a stable order —
        # itertools.product's order matches what we want:
        # (U0[0], U1[0], U2[0]), (U0[0], U1[0], U2[1]), ...,
        # (U0[0], U1[1], U2[0]), ..., (U0[1], U1[0], U2[0]), ...
        # Exactly outermost-slowest hierarchical order.
        
        rows: List[List[Any]] = []
        for combo in all_combinations:
            rows.append(list(combo) + [qualified_counts[combo]])
        
        # Grand total
        total_qualified = sum(qualified_counts.values())
        if rows:
            grand_total_row = (
                ["Grand Total"]
                + [""] * (len(level_columns) - 1)
                + [total_qualified]
            )
            rows.append(grand_total_row)
        
        header = list(level_columns) + [count_header]
        return [header] + rows
    
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