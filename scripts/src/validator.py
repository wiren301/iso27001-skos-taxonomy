"""SKOS validation module."""

from dataclasses import dataclass, field
from typing import Optional

from rdflib import Graph, URIRef
from rdflib.namespace import RDF, SKOS


@dataclass
class ValidationResult:
    """Result of a validation check."""
    passed: bool
    check_name: str
    message: str
    severity: str = "ERROR"  # ERROR, WARNING, INFO
    details: list = field(default_factory=list)


class SKOSValidator:
    """Validate SKOS graph for consistency and completeness."""

    def __init__(self, graph: Graph):
        """Initialize validator with graph."""
        self.graph = graph
        self.results: list[ValidationResult] = []

    def validate_all(self) -> list[ValidationResult]:
        """Run all validation checks."""
        self.results = []

        # Core SKOS checks
        self.check_concepts_have_preflabel()
        self.check_concepts_in_scheme()
        self.check_hierarchy_consistency()
        self.check_related_symmetry()
        self.check_top_concepts()
        self.check_no_orphan_concepts()

        # Data quality checks
        self.check_no_empty_labels()
        self.check_notation_uniqueness()

        return self.results

    def check_concepts_have_preflabel(self) -> ValidationResult:
        """Check that every skos:Concept has at least one skos:prefLabel."""
        concepts_without_label = []

        for concept in self.graph.subjects(RDF.type, SKOS.Concept):
            labels = list(self.graph.objects(concept, SKOS.prefLabel))
            if not labels:
                concepts_without_label.append(str(concept))

        passed = len(concepts_without_label) == 0
        result = ValidationResult(
            passed=passed,
            check_name="concepts_have_preflabel",
            message=f"All concepts have prefLabel" if passed else f"{len(concepts_without_label)} concepts missing prefLabel",
            severity="ERROR",
            details=concepts_without_label[:10]  # Limit to first 10
        )
        self.results.append(result)
        return result

    def check_concepts_in_scheme(self) -> ValidationResult:
        """Check that every skos:Concept belongs to at least one skos:ConceptScheme."""
        concepts_without_scheme = []

        for concept in self.graph.subjects(RDF.type, SKOS.Concept):
            schemes = list(self.graph.objects(concept, SKOS.inScheme))
            if not schemes:
                concepts_without_scheme.append(str(concept))

        passed = len(concepts_without_scheme) == 0
        result = ValidationResult(
            passed=passed,
            check_name="concepts_in_scheme",
            message=f"All concepts belong to a scheme" if passed else f"{len(concepts_without_scheme)} concepts not in any scheme",
            severity="WARNING",
            details=concepts_without_scheme[:10]
        )
        self.results.append(result)
        return result

    def check_hierarchy_consistency(self) -> ValidationResult:
        """Check that broader/narrower relationships are consistent (no cycles)."""
        # Build hierarchy graph
        hierarchy = {}
        for s, o in self.graph.subject_objects(SKOS.broader):
            if s not in hierarchy:
                hierarchy[s] = set()
            hierarchy[s].add(o)

        # Check for cycles using DFS
        cycles_found = []

        def has_cycle(node, visited, rec_stack):
            visited.add(node)
            rec_stack.add(node)

            for neighbor in hierarchy.get(node, []):
                if neighbor not in visited:
                    if has_cycle(neighbor, visited, rec_stack):
                        return True
                elif neighbor in rec_stack:
                    cycles_found.append((str(node), str(neighbor)))
                    return True

            rec_stack.remove(node)
            return False

        visited = set()
        for node in hierarchy:
            if node not in visited:
                has_cycle(node, visited, set())

        passed = len(cycles_found) == 0
        result = ValidationResult(
            passed=passed,
            check_name="hierarchy_consistency",
            message=f"No cycles in hierarchy" if passed else f"{len(cycles_found)} cycles found in broader/narrower",
            severity="ERROR",
            details=[f"{a} -> {b}" for a, b in cycles_found[:10]]
        )
        self.results.append(result)
        return result

    def check_related_symmetry(self) -> ValidationResult:
        """Check that skos:related is symmetric (if A related B, then B related A)."""
        asymmetric = []

        for s, o in self.graph.subject_objects(SKOS.related):
            # Check reverse exists
            if (o, SKOS.related, s) not in self.graph:
                asymmetric.append((str(s), str(o)))

        passed = len(asymmetric) == 0
        result = ValidationResult(
            passed=passed,
            check_name="related_symmetry",
            message=f"All related links are symmetric" if passed else f"{len(asymmetric)} asymmetric related links",
            severity="WARNING",
            details=[f"{a} related {b} (no reverse)" for a, b in asymmetric[:10]]
        )
        self.results.append(result)
        return result

    def check_top_concepts(self) -> ValidationResult:
        """Check that top concepts have no broader concepts."""
        invalid_top_concepts = []

        for scheme in self.graph.subjects(RDF.type, SKOS.ConceptScheme):
            for top_concept in self.graph.objects(scheme, SKOS.hasTopConcept):
                broader = list(self.graph.objects(top_concept, SKOS.broader))
                if broader:
                    invalid_top_concepts.append(str(top_concept))

        passed = len(invalid_top_concepts) == 0
        result = ValidationResult(
            passed=passed,
            check_name="top_concepts_valid",
            message=f"All top concepts have no broader" if passed else f"{len(invalid_top_concepts)} top concepts have broader concepts",
            severity="ERROR",
            details=invalid_top_concepts[:10]
        )
        self.results.append(result)
        return result

    def check_no_orphan_concepts(self) -> ValidationResult:
        """Check for concepts that are not reachable from any top concept."""
        # Get all concepts
        all_concepts = set(self.graph.subjects(RDF.type, SKOS.Concept))

        # Get all top concepts
        top_concepts = set()
        for scheme in self.graph.subjects(RDF.type, SKOS.ConceptScheme):
            for tc in self.graph.objects(scheme, SKOS.hasTopConcept):
                top_concepts.add(tc)

        # BFS from top concepts
        reachable = set(top_concepts)
        queue = list(top_concepts)

        while queue:
            current = queue.pop(0)
            for narrower in self.graph.objects(current, SKOS.narrower):
                if narrower not in reachable:
                    reachable.add(narrower)
                    queue.append(narrower)

        orphans = all_concepts - reachable
        # Filter out concepts that might be in collections
        collection_members = set()
        for coll in self.graph.subjects(RDF.type, SKOS.Collection):
            for member in self.graph.objects(coll, SKOS.member):
                collection_members.add(member)

        true_orphans = orphans - collection_members - top_concepts

        passed = len(true_orphans) == 0
        result = ValidationResult(
            passed=passed,
            check_name="no_orphan_concepts",
            message=f"No orphan concepts" if passed else f"{len(true_orphans)} orphan concepts found",
            severity="INFO",
            details=[str(o) for o in list(true_orphans)[:10]]
        )
        self.results.append(result)
        return result

    def check_no_empty_labels(self) -> ValidationResult:
        """Check that no prefLabel is empty or whitespace-only."""
        empty_labels = []

        for s, o in self.graph.subject_objects(SKOS.prefLabel):
            label_str = str(o).strip()
            if not label_str:
                empty_labels.append(str(s))

        passed = len(empty_labels) == 0
        result = ValidationResult(
            passed=passed,
            check_name="no_empty_labels",
            message=f"No empty labels" if passed else f"{len(empty_labels)} concepts have empty prefLabel",
            severity="ERROR",
            details=empty_labels[:10]
        )
        self.results.append(result)
        return result

    def check_notation_uniqueness(self) -> ValidationResult:
        """Check that notations are unique within each scheme."""
        duplicates = []

        for scheme in self.graph.subjects(RDF.type, SKOS.ConceptScheme):
            notations = {}
            for concept in self.graph.subjects(SKOS.inScheme, scheme):
                for notation in self.graph.objects(concept, SKOS.notation):
                    notation_str = str(notation)
                    if notation_str in notations:
                        duplicates.append((str(scheme), notation_str, str(concept), str(notations[notation_str])))
                    else:
                        notations[notation_str] = concept

        passed = len(duplicates) == 0
        result = ValidationResult(
            passed=passed,
            check_name="notation_uniqueness",
            message=f"All notations unique within schemes" if passed else f"{len(duplicates)} duplicate notations",
            severity="ERROR",
            details=[f"{scheme}: {notation}" for scheme, notation, _, _ in duplicates[:10]]
        )
        self.results.append(result)
        return result

    def get_statistics(self) -> dict:
        """Get statistics about the graph."""
        stats = {
            'concepts': len(list(self.graph.subjects(RDF.type, SKOS.Concept))),
            'schemes': len(list(self.graph.subjects(RDF.type, SKOS.ConceptScheme))),
            'collections': len(list(self.graph.subjects(RDF.type, SKOS.Collection))),
            'broader_links': len(list(self.graph.subject_objects(SKOS.broader))),
            'narrower_links': len(list(self.graph.subject_objects(SKOS.narrower))),
            'related_links': len(list(self.graph.subject_objects(SKOS.related))) // 2,  # Symmetric
            'exactMatch_links': len(list(self.graph.subject_objects(SKOS.exactMatch))) // 2,
            'closeMatch_links': len(list(self.graph.subject_objects(SKOS.closeMatch))) // 2,
            'total_triples': len(self.graph),
        }
        return stats

    def generate_report(self) -> str:
        """Generate a markdown validation report."""
        lines = [
            "# SKOS Validation Report",
            "",
            "## Statistics",
            ""
        ]

        stats = self.get_statistics()
        for key, value in stats.items():
            lines.append(f"- **{key.replace('_', ' ').title()}**: {value}")

        lines.extend([
            "",
            "## Validation Results",
            ""
        ])

        errors = [r for r in self.results if r.severity == "ERROR" and not r.passed]
        warnings = [r for r in self.results if r.severity == "WARNING" and not r.passed]
        passed = [r for r in self.results if r.passed]

        if errors:
            lines.append("### Errors")
            lines.append("")
            for r in errors:
                lines.append(f"- **{r.check_name}**: {r.message}")
                if r.details:
                    for d in r.details:
                        lines.append(f"  - {d}")
            lines.append("")

        if warnings:
            lines.append("### Warnings")
            lines.append("")
            for r in warnings:
                lines.append(f"- **{r.check_name}**: {r.message}")
            lines.append("")

        lines.append("### Passed Checks")
        lines.append("")
        for r in passed:
            lines.append(f"- {r.check_name}: {r.message}")

        return "\n".join(lines)
