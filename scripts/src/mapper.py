"""SKOS concept mapping module."""

from typing import Optional
from datetime import date

import pandas as pd
from rdflib import Graph, Namespace, Literal, URIRef, BNode
from rdflib.namespace import RDF, RDFS, SKOS, DCTERMS, XSD


class SKOSMapper:
    """Map parsed ISO data to SKOS concepts."""

    def __init__(self, base_uri: str = "https://wiren301.github.io/iso27001-skos-taxonomy/"):
        """Initialize mapper with base URI."""
        self.graph = Graph()
        self.base_uri = base_uri

        # Define namespaces
        self.BASE = Namespace(base_uri)
        self.ISO27000 = Namespace(f"{base_uri}27000/")
        self.ISO27001 = Namespace(f"{base_uri}27001/")
        self.ISO27002 = Namespace(f"{base_uri}27002/")

        # Bind prefixes
        self.graph.bind("", self.BASE)
        self.graph.bind("iso27000", self.ISO27000)
        self.graph.bind("iso27001", self.ISO27001)
        self.graph.bind("iso27002", self.ISO27002)
        self.graph.bind("skos", SKOS)
        self.graph.bind("dct", DCTERMS)
        self.graph.bind("rdf", RDF)
        self.graph.bind("rdfs", RDFS)

        # Track created concepts for linking
        self.concepts = {}
        self.collections = {}

    def create_unified_scheme(self) -> URIRef:
        """Create top-level unified concept scheme."""
        scheme = self.BASE["scheme"]
        self.graph.add((scheme, RDF.type, SKOS.ConceptScheme))
        self.graph.add((scheme, SKOS.prefLabel, Literal("ISO Information Security Taxonomy", lang="en")))
        self.graph.add((scheme, DCTERMS.title, Literal("ISO Information Security Taxonomy")))
        self.graph.add((scheme, DCTERMS.description, Literal(
            "Unified SKOS vocabulary for ISO/IEC 27000, 27001, and 27002 standards",
            lang="en"
        )))
        self.graph.add((scheme, DCTERMS.created, Literal(date.today().isoformat(), datatype=XSD.date)))
        self.graph.add((scheme, DCTERMS.creator, Literal("ISO 27001 SKOS Taxonomy Project")))
        return scheme

    def create_scheme(self, ns: Namespace, name: str, label: str, description: str) -> URIRef:
        """Create a concept scheme."""
        scheme = ns[name]
        self.graph.add((scheme, RDF.type, SKOS.ConceptScheme))
        self.graph.add((scheme, SKOS.prefLabel, Literal(label, lang="en")))
        self.graph.add((scheme, DCTERMS.description, Literal(description, lang="en")))
        self.graph.add((scheme, DCTERMS.created, Literal(date.today().isoformat(), datatype=XSD.date)))
        return scheme

    def map_iso27000(self, df: pd.DataFrame) -> URIRef:
        """
        Map ISO 27000 vocabulary terms to SKOS concepts.

        Returns the scheme URI.
        """
        scheme = self.create_scheme(
            self.ISO27000,
            "VocabularyScheme",
            "ISO/IEC 27000:2018 Information Security Vocabulary",
            "77 terms and definitions for information security management systems"
        )

        # Hub concepts to mark as top concepts
        hub_terms = {'3.28', '3.41', '3.64'}  # information security, management system, risk assessment

        for _, row in df.iterrows():
            term_id = str(row['term_id'])
            uri = self.ISO27000[f"term/{term_id}"]

            self.graph.add((uri, RDF.type, SKOS.Concept))
            self.graph.add((uri, SKOS.inScheme, scheme))
            self.graph.add((uri, SKOS.notation, Literal(term_id)))

            if pd.notna(row['term']):
                self.graph.add((uri, SKOS.prefLabel, Literal(str(row['term']).strip(), lang="en")))

            if pd.notna(row['definition']):
                self.graph.add((uri, SKOS.definition, Literal(str(row['definition']).strip(), lang="en")))

            if pd.notna(row.get('notes')) and row['notes']:
                self.graph.add((uri, SKOS.note, Literal(str(row['notes']).strip(), lang="en")))

            # Mark hub concepts as top concepts
            if term_id in hub_terms:
                self.graph.add((scheme, SKOS.hasTopConcept, uri))
                self.graph.add((uri, SKOS.topConceptOf, scheme))

            # Store for cross-reference linking
            self.concepts[f"iso27000:{term_id}"] = uri

        return scheme

    def map_iso27001_clauses(self, df: pd.DataFrame) -> URIRef:
        """
        Map ISO 27001 ISMS clauses to SKOS concepts with hierarchy.

        Returns the scheme URI.
        """
        scheme = self.create_scheme(
            self.ISO27001,
            "ISMSScheme",
            "ISO/IEC 27001:2022 ISMS Requirements",
            "Information Security Management System requirements organized by clause"
        )

        # First pass: create all concepts
        for _, row in df.iterrows():
            clause_id = str(row['clause_id'])
            uri = self.ISO27001[f"clause/{clause_id}"]

            self.graph.add((uri, RDF.type, SKOS.Concept))
            self.graph.add((uri, SKOS.inScheme, scheme))
            self.graph.add((uri, SKOS.notation, Literal(clause_id)))

            if pd.notna(row['title']):
                self.graph.add((uri, SKOS.prefLabel, Literal(str(row['title']).strip(), lang="en")))

            if pd.notna(row.get('text')):
                self.graph.add((uri, SKOS.definition, Literal(str(row['text']).strip(), lang="en")))

            if pd.notna(row.get('sub_items')):
                self.graph.add((uri, SKOS.note, Literal(str(row['sub_items']).strip(), lang="en")))

            if pd.notna(row.get('notes')):
                self.graph.add((uri, SKOS.scopeNote, Literal(str(row['notes']).strip(), lang="en")))

            # Mark top-level clauses (single digit) as top concepts
            if row.get('level', 0) == 0:
                self.graph.add((scheme, SKOS.hasTopConcept, uri))
                self.graph.add((uri, SKOS.topConceptOf, scheme))

            self.concepts[f"iso27001:clause:{clause_id}"] = uri

        # Second pass: add broader/narrower relationships
        for _, row in df.iterrows():
            clause_id = str(row['clause_id'])
            uri = self.concepts[f"iso27001:clause:{clause_id}"]

            if pd.notna(row.get('broader_id')):
                broader_key = f"iso27001:clause:{row['broader_id']}"
                if broader_key in self.concepts:
                    broader_uri = self.concepts[broader_key]
                    self.graph.add((uri, SKOS.broader, broader_uri))
                    self.graph.add((broader_uri, SKOS.narrower, uri))

        return scheme

    def map_iso27001_annexa(self, df: pd.DataFrame) -> URIRef:
        """
        Map ISO 27001 Annex A controls to SKOS concepts with theme collections.

        Returns the scheme URI.
        """
        scheme = self.create_scheme(
            self.ISO27001,
            "AnnexAScheme",
            "ISO/IEC 27001:2022 Annex A Controls",
            "93 information security controls organized by theme"
        )

        # Create theme collections
        themes = df['theme'].dropna().unique()
        theme_collections = {}

        for theme in themes:
            theme_slug = theme.lower().replace(' ', '_')
            coll_uri = self.ISO27001[f"collection/{theme_slug}"]
            self.graph.add((coll_uri, RDF.type, SKOS.Collection))
            self.graph.add((coll_uri, SKOS.prefLabel, Literal(f"{theme} Controls", lang="en")))
            self.graph.add((scheme, SKOS.hasTopConcept, coll_uri))
            theme_collections[theme] = coll_uri
            self.collections[f"iso27001:theme:{theme_slug}"] = coll_uri

        # Create control concepts
        for _, row in df.iterrows():
            control_id = str(row['control_id'])
            uri = self.ISO27001[f"control/{control_id}"]

            self.graph.add((uri, RDF.type, SKOS.Concept))
            self.graph.add((uri, SKOS.inScheme, scheme))
            self.graph.add((uri, SKOS.notation, Literal(control_id)))

            if pd.notna(row['name']):
                self.graph.add((uri, SKOS.prefLabel, Literal(str(row['name']).strip(), lang="en")))

            if pd.notna(row.get('statement')):
                self.graph.add((uri, SKOS.definition, Literal(str(row['statement']).strip(), lang="en")))

            # Add to theme collection
            if pd.notna(row.get('theme')) and row['theme'] in theme_collections:
                self.graph.add((theme_collections[row['theme']], SKOS.member, uri))

            self.concepts[f"iso27001:control:{control_id}"] = uri

        return scheme

    def map_iso27002(self, df: pd.DataFrame) -> URIRef:
        """
        Map ISO 27002 controls with attributes to SKOS concepts and collections.

        Returns the scheme URI.
        """
        scheme = self.create_scheme(
            self.ISO27002,
            "ControlsScheme",
            "ISO/IEC 27002:2022 Information Security Controls",
            "93 controls with detailed guidance and attribute classifications"
        )

        # Create attribute collections
        attribute_configs = {
            'info_security_props': ('cia', 'Information Security Property'),
            'cybersecurity_concepts': ('nist', 'Cybersecurity Function'),
            'operational_caps': ('ops', 'Operational Capability'),
            'security_domains': ('domain', 'Security Domain')
        }

        attr_collections = {}
        for attr_col, (prefix, label_prefix) in attribute_configs.items():
            if attr_col in df.columns:
                all_values = set()
                for vals in df[attr_col].dropna():
                    if isinstance(vals, list):
                        all_values.update(vals)

                for value in all_values:
                    if value:
                        slug = value.lower().replace(' ', '_').replace('/', '_')
                        coll_uri = self.ISO27002[f"attribute/{prefix}/{slug}"]
                        self.graph.add((coll_uri, RDF.type, SKOS.Collection))
                        self.graph.add((coll_uri, SKOS.prefLabel, Literal(f"{label_prefix}: {value}", lang="en")))
                        attr_collections[(attr_col, value)] = coll_uri
                        self.collections[f"iso27002:attr:{prefix}:{slug}"] = coll_uri

        # Create control concepts
        for _, row in df.iterrows():
            control_id = str(row['control_id'])
            uri = self.ISO27002[f"control/{control_id}"]

            self.graph.add((uri, RDF.type, SKOS.Concept))
            self.graph.add((uri, SKOS.inScheme, scheme))
            self.graph.add((uri, SKOS.notation, Literal(control_id)))

            if pd.notna(row['name']):
                self.graph.add((uri, SKOS.prefLabel, Literal(str(row['name']).strip(), lang="en")))

            if pd.notna(row.get('statement')):
                self.graph.add((uri, SKOS.definition, Literal(str(row['statement']).strip(), lang="en")))

            if pd.notna(row.get('purpose')):
                self.graph.add((uri, SKOS.scopeNote, Literal(str(row['purpose']).strip(), lang="en")))

            if pd.notna(row.get('guidance')):
                # Truncate very long guidance to avoid massive literals
                guidance = str(row['guidance']).strip()
                if len(guidance) > 5000:
                    guidance = guidance[:5000] + "..."
                self.graph.add((uri, SKOS.note, Literal(guidance, lang="en")))

            if pd.notna(row.get('other_info')):
                self.graph.add((uri, SKOS.example, Literal(str(row['other_info']).strip(), lang="en")))

            # Add to attribute collections
            for attr_col in attribute_configs.keys():
                if attr_col in df.columns and isinstance(row.get(attr_col), list):
                    for value in row[attr_col]:
                        if (attr_col, value) in attr_collections:
                            self.graph.add((attr_collections[(attr_col, value)], SKOS.member, uri))

            self.concepts[f"iso27002:control:{control_id}"] = uri

        return scheme

    def get_graph(self) -> Graph:
        """Return the RDF graph."""
        return self.graph

    def serialize(self, format: str = "turtle") -> str:
        """Serialize graph to string."""
        return self.graph.serialize(format=format)

    def save(self, path: str, format: str = "turtle"):
        """Save graph to file."""
        self.graph.serialize(destination=path, format=format)
