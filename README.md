# ISO 27001 SKOS Taxonomy

[![License: CC BY 4.0](https://img.shields.io/badge/License-CC_BY_4.0-lightgrey.svg)](https://creativecommons.org/licenses/by/4.0/)
[![W3C PID](https://img.shields.io/badge/W3ID-wiren301--iso27001--skos-blue.svg)](http://w3id.org/wiren301-iso27001-skos)

Building a Machine-Readable SKOS Taxonomy of ISO/IEC 27001:2022 and related information security standards.

## Overview

This project converts ISO information security standards into a unified SKOS (Simple Knowledge Organization System) taxonomy, enabling semantic web applications and knowledge management systems to work with standardized security concepts.

### Included Standards

- **ISO/IEC 27000:2018** - Information security vocabulary (77 terms)
- **ISO/IEC 27001:2022** - ISMS requirements (44 clauses) and Annex A controls (93 controls)
- **ISO/IEC 27002:2022** - Security controls with guidance (93 controls)

## Output

The taxonomy is available in multiple formats:

| Format | File | Description |
|--------|------|-------------|
| Turtle | `output/iso-security.ttl` | Human-readable RDF format |
| JSON-LD | `output/iso-security.jsonld` | JSON-based linked data |

### Statistics

- **305 Concepts** across 6 ConceptSchemes
- **33 Collections** (themes + attribute classifications)
- **3,139 RDF triples**

## Installation

```bash
# Clone the repository
git clone https://github.com/wiren301/iso27001-skos-taxonomy.git
cd iso27001-skos-taxonomy

# Install dependencies
pip install -r requirements.txt
```

## Usage

### Generate the Taxonomy

```bash
# Run with default settings (Turtle + JSON-LD output)
python convert.py

# Specify output formats
python convert.py --format turtle --format json-ld --format xml

# Skip validation
python convert.py --no-validate

# Verbose output
python convert.py -v
```

### Command Line Options

| Option | Default | Description |
|--------|---------|-------------|
| `--schemas-dir` | `schemas` | Directory containing Excel source files |
| `--output-dir` | `output` | Output directory for generated files |
| `--format` | `turtle`, `json-ld` | Output formats (can specify multiple) |
| `--validate/--no-validate` | `--validate` | Run SKOS validation checks |
| `-v, --verbose` | off | Show detailed output |

## Namespace URIs

```turtle
@prefix : <https://wiren301.github.io/iso27001-skos-taxonomy/> .
@prefix iso27000: <https://wiren301.github.io/iso27001-skos-taxonomy/27000/> .
@prefix iso27001: <https://wiren301.github.io/iso27001-skos-taxonomy/27001/> .
@prefix iso27002: <https://wiren301.github.io/iso27001-skos-taxonomy/27002/> .
```

## Structure

```
iso27001-skos-taxonomy/
├── schemas/                    # Source Excel files
│   ├── ISO27000_UNIFIED.xlsx
│   ├── ISO27001_UNIFIED.xlsx
│   └── ISO27002_UNIFIED.xlsx
├── src/                        # Python modules
│   ├── parser.py              # Excel parsing
│   ├── mapper.py              # SKOS concept creation
│   ├── linker.py              # Cross-reference linking
│   └── validator.py           # SKOS validation
├── config/
│   └── config.yaml            # Configuration
├── output/                     # Generated files
│   ├── iso-security.ttl       # Turtle output
│   ├── iso-security.jsonld    # JSON-LD output
│   ├── context.jsonld         # JSON-LD context
│   ├── validation-report.md   # Validation report
│   └── statistics.json        # Graph statistics
├── convert.py                  # Main CLI script
└── requirements.txt            # Python dependencies
```

## Cross-References

The taxonomy includes semantic links between standards:

- **ISO 27001 Annex A ↔ ISO 27002**: `skos:exactMatch` (same control IDs)
- **ISO 27001/27002 → ISO 27000**: `usesTerm` (vocabulary references)
- **Internal references**: `skos:related` (definition cross-refs like "(3.28)")

## Disclaimer

> **Important Notice**
>
> This taxonomy is an **original academic artifact** developed as part of scholarly research at the University of Twente. It represents the author's interpretation and analysis of information security concepts from the ISO/IEC 27000 series of standards.
>
> **This work:**
> - Is **based on** ISO/IEC 27000:2018, ISO/IEC 27001:2022, and ISO/IEC 27002:2022, but is **not a substitute** for those standards
> - Does **not reproduce** the ISO standards verbatim; all definitions are paraphrased interpretations
> - Should be viewed as a **scholarly analysis** and the author's interpretation of information security concepts
> - Was created in an **academic context** and does not offer guarantees as a reference document
> - Is **not affiliated with, endorsed by, or officially connected to** ISO or IEC
>
> For authoritative definitions and requirements, please consult the official ISO/IEC standards available at [iso.org](https://www.iso.org/).

## License

The SKOS taxonomy structure and tooling are licensed under [CC-BY-4.0](https://creativecommons.org/licenses/by/4.0/).

## Contributing

Contributions welcome. Please ensure any changes pass validation:

```bash
python convert.py --validate -v
```
