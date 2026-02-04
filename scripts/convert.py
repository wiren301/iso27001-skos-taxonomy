#!/usr/bin/env python3
"""
ISO 27001 SKOS Taxonomy Converter

Converts ISO 27000, 27001, and 27002 Excel documents
into a unified SKOS taxonomy.

Usage:
    python convert.py [--output-dir OUTPUT] [--format FORMAT]
"""

import json
import sys
from pathlib import Path

import click

from src.parser import ExcelParser, load_all_sources
from src.mapper import SKOSMapper
from src.linker import CrossRefLinker
from src.validator import SKOSValidator


BASE_URI = "https://wiren301.github.io/iso27001-skos-taxonomy/"


@click.command()
@click.option('--schemas-dir', default='schemas', help='Directory containing Excel files')
@click.option('--output-dir', default='output', help='Output directory')
@click.option('--format', 'formats', multiple=True, default=['turtle', 'json-ld'],
              help='Output formats (turtle, json-ld)')
@click.option('--validate/--no-validate', default=True, help='Run validation')
@click.option('--verbose', '-v', is_flag=True, help='Verbose output')
def main(schemas_dir: str, output_dir: str, formats: tuple, validate: bool, verbose: bool):
    """Convert ISO Excel documents to SKOS taxonomy."""

    schemas_path = Path(schemas_dir)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    click.echo("=" * 60)
    click.echo("ISO 27001 SKOS Taxonomy Converter")
    click.echo("=" * 60)
    click.echo()

    # Step 1: Parse Excel files
    click.echo("[1/5] Parsing Excel files...")
    try:
        sources = load_all_sources(schemas_path)
        for name, df in sources.items():
            if hasattr(df, '__len__'):
                click.echo(f"  - {name}: {len(df)} rows")
    except Exception as e:
        click.echo(f"Error parsing Excel files: {e}", err=True)
        sys.exit(1)

    # Step 2: Map to SKOS
    click.echo("\n[2/5] Mapping to SKOS concepts...")
    mapper = SKOSMapper(BASE_URI)

    # Create unified scheme
    mapper.create_unified_scheme()

    # Map each source
    if 'iso27000' in sources:
        scheme = mapper.map_iso27000(sources['iso27000'])
        click.echo(f"  - ISO 27000 vocabulary: {len(sources['iso27000'])} terms")

    if 'iso27001_clauses' in sources:
        scheme = mapper.map_iso27001_clauses(sources['iso27001_clauses'])
        click.echo(f"  - ISO 27001 clauses: {len(sources['iso27001_clauses'])} requirements")

    if 'iso27001_annexa' in sources:
        scheme = mapper.map_iso27001_annexa(sources['iso27001_annexa'])
        click.echo(f"  - ISO 27001 Annex A: {len(sources['iso27001_annexa'])} controls")

    if 'iso27002' in sources:
        scheme = mapper.map_iso27002(sources['iso27002'])
        click.echo(f"  - ISO 27002 controls: {len(sources['iso27002'])} controls")

    # Step 3: Add cross-reference links
    click.echo("\n[3/5] Adding cross-reference links...")
    linker = CrossRefLinker(mapper.get_graph(), mapper.concepts, BASE_URI)
    link_results = linker.add_all_links(sources)

    for link_type, count in link_results.items():
        if count > 0:
            click.echo(f"  - {link_type}: {count} links")

    # Step 4: Validate
    if validate:
        click.echo("\n[4/5] Validating SKOS graph...")
        validator = SKOSValidator(mapper.get_graph())
        validation_results = validator.validate_all()

        errors = [r for r in validation_results if r.severity == "ERROR" and not r.passed]
        warnings = [r for r in validation_results if r.severity == "WARNING" and not r.passed]
        passed = [r for r in validation_results if r.passed]

        click.echo(f"  - Passed: {len(passed)}")
        click.echo(f"  - Warnings: {len(warnings)}")
        click.echo(f"  - Errors: {len(errors)}")

        if verbose and (errors or warnings):
            click.echo("\n  Issues found:")
            for r in errors + warnings:
                click.echo(f"    [{r.severity}] {r.check_name}: {r.message}")

        # Save validation report
        report = validator.generate_report()
        report_path = output_path / "validation-report.md"
        report_path.write_text(report)
        click.echo(f"  - Report saved: {report_path}")

        # Save statistics
        stats = validator.get_statistics()
        stats_path = output_path / "statistics.json"
        stats_path.write_text(json.dumps(stats, indent=2))
    else:
        click.echo("\n[4/5] Skipping validation...")

    # Step 5: Serialize and save
    click.echo("\n[5/5] Saving output files...")

    for fmt in formats:
        if fmt in ('turtle', 'ttl'):
            output_file = output_path / "iso-security.ttl"
            mapper.save(str(output_file), format='turtle')
            click.echo(f"  - Turtle: {output_file}")

        elif fmt in ('json-ld', 'jsonld'):
            output_file = output_path / "iso-security.jsonld"
            mapper.save(str(output_file), format='json-ld')
            click.echo(f"  - JSON-LD: {output_file}")

        elif fmt in ('xml', 'rdf-xml', 'rdfxml'):
            output_file = output_path / "iso-security.rdf"
            mapper.save(str(output_file), format='xml')
            click.echo(f"  - RDF/XML: {output_file}")

        elif fmt in ('nt', 'ntriples', 'n-triples'):
            output_file = output_path / "iso-security.nt"
            mapper.save(str(output_file), format='nt')
            click.echo(f"  - N-Triples: {output_file}")

    # Final summary
    click.echo("\n" + "=" * 60)
    click.echo("Conversion complete!")
    click.echo("=" * 60)

    if validate:
        stats = validator.get_statistics()
        click.echo(f"\nSummary:")
        click.echo(f"  - Concepts: {stats['concepts']}")
        click.echo(f"  - Schemes: {stats['schemes']}")
        click.echo(f"  - Collections: {stats['collections']}")
        click.echo(f"  - Total triples: {stats['total_triples']}")


if __name__ == '__main__':
    main()
