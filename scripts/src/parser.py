"""Excel parsing module for ISO documents with unified schema format."""

import re
from pathlib import Path
from typing import Optional

import pandas as pd


class ExcelParser:
    """Parse ISO Excel files using unified schema format."""

    # Pattern to extract cross-references like (3.56) from definition text
    CROSS_REF_PATTERN = re.compile(r'\(3\.(\d+)\)')

    def __init__(self, base_path: Path):
        """Initialize parser with base path to schemas directory."""
        self.base_path = Path(base_path)

    def parse_iso27000(self, filename: str, sheet: str = "ISO27000_Vocabulary") -> pd.DataFrame:
        """
        Parse ISO 27000 vocabulary terms from unified schema.

        Returns DataFrame with columns:
        - term_id: Term identifier (e.g., "3.56")
        - term: Term name
        - definition: Term definition
        - notes: Combined notes
        - cross_refs: List of referenced term IDs extracted from definition
        """
        path = self.base_path / filename
        # Read Concept_ID as string to preserve IDs like 3.10 (not 3.1)
        df = pd.read_excel(path, sheet_name=sheet, dtype={'Concept_ID': str, 'Parent_ID': str})

        # Combine note columns
        note_cols = [c for c in df.columns if c.startswith('Note_')]
        df['notes'] = df[note_cols].apply(
            lambda row: ' | '.join([str(v) for v in row if pd.notna(v)]),
            axis=1
        )
        df['notes'] = df['notes'].replace('', None)

        # Extract cross-references from definition
        df['cross_refs'] = df['Definition'].apply(self._extract_cross_refs)

        # Map unified columns to expected output format
        result = pd.DataFrame({
            'term_id': df['Concept_ID'],
            'term': df['Label'],
            'definition': df['Definition'],
            'notes': df['notes'],
            'cross_refs': df['cross_refs']
        })
        result = result.dropna(subset=['term_id', 'term'])

        return result

    def parse_iso27001_clauses(self, filename: str, sheet: str = "ISO27001_Requirements") -> pd.DataFrame:
        """
        Parse ISO 27001 ISMS requirements (clauses) from unified schema.

        Returns DataFrame with columns:
        - clause_id: Hierarchical clause identifier (e.g., "6.1.3")
        - title: Clause title
        - text: Requirement text (definition)
        - sub_items: Sub-items (a, b, c...)
        - notes: Additional notes
        - external_ref: External references
        - requirement_type: "Section" or "Requirement"
        - documented_info: Whether documented info is required
        - broader_id: Parent clause ID
        - level: Hierarchy level (0=top, 1=X.Y, 2=X.Y.Z)
        """
        path = self.base_path / filename
        df = pd.read_excel(path, sheet_name=sheet, dtype={'Concept_ID': str, 'Parent_ID': str})

        # Filter to ISMS Requirements only (exclude Annex A controls)
        df = df[df['Theme'] == 'ISMS Requirements'].copy()

        # Combine note columns
        note_cols = [c for c in df.columns if c.startswith('Note_')]
        df['notes'] = df[note_cols].apply(
            lambda row: ' | '.join([str(v) for v in row if pd.notna(v)]),
            axis=1
        )
        df['notes'] = df['notes'].replace('', None)

        # Derive hierarchy level
        df['level'] = df['Concept_ID'].apply(self._get_hierarchy_level)

        # Map unified columns to expected output format
        result = pd.DataFrame({
            'clause_id': df['Concept_ID'],
            'title': df['Label'],
            'text': df['Definition'],
            'sub_items': df.get('Sub_Items'),
            'notes': df['notes'],
            'external_reference': df.get('External_Ref'),
            'requirement_type': df.get('Requirement_Type'),
            'documented_info': df.get('Doc_Info_Required'),
            'broader_id': df['Parent_ID'].replace('nan', None).replace('', None),
            'level': df['level']
        })
        result = result.dropna(subset=['clause_id'])

        return result

    def parse_iso27001_annexa(self, filename: str, sheet: str = "ISO27001_Requirements") -> pd.DataFrame:
        """
        Parse ISO 27001 Annex A controls from unified schema.

        Returns DataFrame with columns:
        - control_id: Control identifier (e.g., "5.1")
        - name: Control name
        - statement: Control statement
        - theme: Organizational/People/Physical/Technological
        - iso27002_ref: Reference to ISO 27002 clause
        """
        path = self.base_path / filename
        # Read Concept_ID as string to preserve IDs like 5.10 (not 5.1)
        df = pd.read_excel(path, sheet_name=sheet, dtype={'Concept_ID': str, 'Parent_ID': str})

        # Filter to Annex A controls (exclude ISMS Requirements)
        df = df[df['Theme'] != 'ISMS Requirements'].copy()
        df = df[df['Theme'].notna()].copy()

        # Map unified columns to expected output format
        result = pd.DataFrame({
            'control_id': df['Concept_ID'],
            'name': df['Label'],
            'statement': df['Definition'],
            'theme': df['Theme'],
            'iso27002_ref': df['Concept_ID']  # ISO 27002 uses same control IDs
        })
        result = result.dropna(subset=['control_id'])

        return result

    def parse_iso27002(self, filename: str, sheet: str = "ISO27002_Controls") -> pd.DataFrame:
        """
        Parse ISO 27002 controls with attributes from unified schema.

        Returns DataFrame with columns:
        - control_id: Control identifier (e.g., "5.1")
        - name: Control name
        - statement: Control statement
        - purpose: Control purpose
        - guidance: Implementation guidance
        - other_info: Other information
        - theme: Organizational/People/Physical/Technological
        - control_type: Preventive/Detective/Corrective
        - info_security_props: List of CIA properties
        - cybersecurity_concepts: List of NIST functions
        - operational_caps: List of operational capabilities
        - security_domains: List of security domains
        """
        path = self.base_path / filename
        # Read Concept_ID as string to preserve IDs like 5.10 (not 5.1)
        df = pd.read_excel(path, sheet_name=sheet, dtype={'Concept_ID': str, 'Parent_ID': str})

        # Parse attribute columns (format: "#Value1 #Value2")
        attr_mapping = {
            'Attr_CIA': 'info_security_props',
            'Attr_Cyber': 'cybersecurity_concepts',
            'Attr_Ops_Cap': 'operational_caps',
            'Attr_Sec_Domain': 'security_domains'
        }

        for old_col, new_col in attr_mapping.items():
            if old_col in df.columns:
                df[new_col] = df[old_col].apply(self._parse_attributes)

        # Map unified columns to expected output format
        result = pd.DataFrame({
            'control_id': df['Concept_ID'],
            'name': df['Label'],
            'statement': df['Definition'],
            'purpose': df.get('Purpose'),
            'guidance': df.get('Guidance'),
            'other_info': df.get('Other_Info'),
            'theme': df['Theme'],
            'control_type': df.get('Attr_Control_Type'),
            'info_security_props': df.get('info_security_props', pd.Series([[] for _ in range(len(df))])),
            'cybersecurity_concepts': df.get('cybersecurity_concepts', pd.Series([[] for _ in range(len(df))])),
            'operational_caps': df.get('operational_caps', pd.Series([[] for _ in range(len(df))])),
            'security_domains': df.get('security_domains', pd.Series([[] for _ in range(len(df))]))
        })
        result = result.dropna(subset=['control_id'])

        return result

    def _extract_cross_refs(self, text: Optional[str]) -> list[str]:
        """Extract cross-reference IDs from text like (3.56)."""
        if pd.isna(text) or not text:
            return []
        matches = self.CROSS_REF_PATTERN.findall(str(text))
        return [f"3.{m}" for m in matches]

    def _get_broader_id(self, clause_id: str) -> Optional[str]:
        """
        Derive parent clause ID from hierarchical ID.

        Examples:
            "4" -> None (top concept)
            "4.1" -> "4"
            "6.1.3" -> "6.1"
        """
        if pd.isna(clause_id):
            return None
        parts = str(clause_id).split('.')
        if len(parts) <= 1:
            return None
        return '.'.join(parts[:-1])

    def _get_hierarchy_level(self, clause_id: str) -> int:
        """
        Get hierarchy level from clause ID.

        Level 0: "4" (top)
        Level 1: "4.1"
        Level 2: "6.1.3"
        """
        if pd.isna(clause_id):
            return 0
        return str(clause_id).count('.')

    def _parse_attributes(self, cell: Optional[str]) -> list[str]:
        """
        Parse attribute cell with format "#Value1 #Value2".

        Returns list of attribute values without # prefix.
        """
        if pd.isna(cell) or not cell:
            return []
        return [a.strip('#').strip() for a in str(cell).split() if a.startswith('#')]


def load_all_sources(base_path: Path) -> dict[str, pd.DataFrame]:
    """
    Load all ISO source files from unified schema format.

    Returns dict with keys:
    - iso27000: Vocabulary terms
    - iso27001_clauses: ISMS requirements
    - iso27001_annexa: Annex A controls
    - iso27002: Detailed controls with attributes
    """
    parser = ExcelParser(base_path)

    sources = {}

    # ISO 27000
    sources['iso27000'] = parser.parse_iso27000(
        "ISO27000_UNIFIED.xlsx"
    )

    # ISO 27001
    sources['iso27001_clauses'] = parser.parse_iso27001_clauses(
        "ISO27001_UNIFIED.xlsx"
    )
    sources['iso27001_annexa'] = parser.parse_iso27001_annexa(
        "ISO27001_UNIFIED.xlsx"
    )

    # ISO 27002
    sources['iso27002'] = parser.parse_iso27002(
        "ISO27002_UNIFIED.xlsx"
    )

    return sources
