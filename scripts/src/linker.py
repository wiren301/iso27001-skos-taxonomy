"""Cross-reference linking module."""

import re
from typing import Optional

import pandas as pd
from rdflib import Graph, URIRef, Namespace
from rdflib.namespace import SKOS


class CrossRefLinker:
    """Extract and resolve cross-references between SKOS concepts."""

    # Patterns for extracting references
    PATTERNS = {
        'iso27000_internal': re.compile(r'\(3\.(\d+)\)'),
        'iso27001_clause': re.compile(r'[Cc]lause\s+(\d+(?:\.\d+)*)'),
        'iso27002_ref': re.compile(r'ISO[/\s]*(?:IEC\s*)?27002[:\s]+(?:[Cc]lause\s+)?(\d+(?:\.\d+)*)'),
    }

    def __init__(self, graph: Graph, concepts: dict, base_uri: str):
        """
        Initialize linker.

        Args:
            graph: RDF graph to add links to
            concepts: Dict mapping concept keys to URIs
            base_uri: Base URI for namespaces
        """
        self.graph = graph
        self.concepts = concepts
        self.base_uri = base_uri

        self.ISO27000 = Namespace(f"{base_uri}27000/")
        self.ISO27001 = Namespace(f"{base_uri}27001/")
        self.ISO27002 = Namespace(f"{base_uri}27002/")

        self.links_added = 0

    def link_iso27000_cross_refs(self, df: pd.DataFrame) -> int:
        """
        Add skos:related links for cross-references in ISO 27000 definitions.

        Returns count of links added.
        """
        count = 0

        for _, row in df.iterrows():
            term_id = str(row['term_id'])
            source_key = f"iso27000:{term_id}"

            if source_key not in self.concepts:
                continue

            source_uri = self.concepts[source_key]
            cross_refs = row.get('cross_refs', [])

            if isinstance(cross_refs, list):
                for ref_id in cross_refs:
                    target_key = f"iso27000:{ref_id}"
                    if target_key in self.concepts and target_key != source_key:
                        target_uri = self.concepts[target_key]
                        # Add symmetric related links
                        self.graph.add((source_uri, SKOS.related, target_uri))
                        self.graph.add((target_uri, SKOS.related, source_uri))
                        count += 1

        self.links_added += count
        return count

    def link_iso27001_to_iso27002(self, annexa_df: pd.DataFrame) -> int:
        """
        Add skos:exactMatch links between ISO 27001 Annex A and ISO 27002 controls.

        Returns count of links added.
        """
        count = 0

        for _, row in annexa_df.iterrows():
            control_id = str(row['control_id'])
            annexa_key = f"iso27001:control:{control_id}"

            if annexa_key not in self.concepts:
                continue

            annexa_uri = self.concepts[annexa_key]

            # ISO 27002 uses control IDs without A. prefix
            clean_id = control_id.lstrip('A.') if control_id.startswith('A.') else control_id
            iso27002_key = f"iso27002:control:{clean_id}"

            if iso27002_key in self.concepts:
                iso27002_uri = self.concepts[iso27002_key]
                # exactMatch is symmetric
                self.graph.add((annexa_uri, SKOS.exactMatch, iso27002_uri))
                self.graph.add((iso27002_uri, SKOS.exactMatch, annexa_uri))
                count += 1

        self.links_added += count
        return count

    def link_clauses_to_vocabulary(self, clauses_df: pd.DataFrame) -> int:
        """
        Add skos:related links from ISO 27001 clauses to ISO 27000 vocabulary terms
        mentioned in external references.

        Returns count of links added.
        """
        count = 0

        # Key vocabulary terms that clauses reference
        vocab_mappings = {
            'ISO/IEC 27000': ['3.28', '3.41'],  # information security, management system
        }

        for _, row in clauses_df.iterrows():
            clause_id = str(row['clause_id'])
            clause_key = f"iso27001:clause:{clause_id}"

            if clause_key not in self.concepts:
                continue

            clause_uri = self.concepts[clause_key]
            ext_ref = str(row.get('external_reference', '')) if pd.notna(row.get('external_reference')) else ''

            for ref_text, term_ids in vocab_mappings.items():
                if ref_text in ext_ref:
                    for term_id in term_ids:
                        vocab_key = f"iso27000:{term_id}"
                        if vocab_key in self.concepts:
                            vocab_uri = self.concepts[vocab_key]
                            self.graph.add((clause_uri, SKOS.related, vocab_uri))
                            self.graph.add((vocab_uri, SKOS.related, clause_uri))
                            count += 1

        self.links_added += count
        return count

    def add_all_links(self, sources: dict) -> dict:
        """
        Add all cross-reference links.

        Args:
            sources: Dict of DataFrames from parser

        Returns:
            Dict with counts of links added by type
        """
        results = {}

        # ISO 27000 internal cross-references
        if 'iso27000' in sources:
            results['iso27000_internal'] = self.link_iso27000_cross_refs(sources['iso27000'])

        # ISO 27001 Annex A to ISO 27002
        if 'iso27001_annexa' in sources:
            results['annexa_to_27002'] = self.link_iso27001_to_iso27002(sources['iso27001_annexa'])

        # Clauses to vocabulary
        if 'iso27001_clauses' in sources:
            results['clauses_to_vocab'] = self.link_clauses_to_vocabulary(sources['iso27001_clauses'])

        results['total'] = self.links_added
        return results
