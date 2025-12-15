#!/usr/bin/env python3
"""Extract ISO 27000 vocabulary data directly from Excel for D3.js visualization."""

import json
import re
from pathlib import Path

import pandas as pd

# Cross-reference pattern: (3.xx) in definitions
CROSS_REF_PATTERN = re.compile(r'\(3\.(\d+)\)')

BASE_URI = "https://wiren301.github.io/iso27001-skos-taxonomy/"


def extract_cross_refs(text: str) -> list[str]:
    """Extract term IDs referenced in definition text."""
    if not text or pd.isna(text):
        return []
    matches = CROSS_REF_PATTERN.findall(str(text))
    return [f"3.{m}" for m in matches]


def extract_iso27000_from_excel(excel_path: str = "schemas/ISO27000_2018_EXTRACTED.xlsx"):
    """Extract ISO 27000 terms directly from Excel source."""

    df = pd.read_excel(excel_path, dtype={'Term_ID': str})

    # Hub terms marked as top concepts
    hub_terms = {'3.28', '3.41', '3.64'}  # information security, management system, risk assessment

    nodes = []
    node_ids = set()

    for _, row in df.iterrows():
        term_id = str(row['Term_ID']).strip()
        term = str(row['Term']).strip() if pd.notna(row['Term']) else ''
        definition = str(row['Definition']).strip() if pd.notna(row['Definition']) else ''

        if not term_id or not term:
            continue

        # Combine notes
        notes = []
        for i in range(1, 7):
            col = f'Note_{i}'
            if col in df.columns and pd.notna(row.get(col)):
                note_text = str(row[col]).strip()
                if note_text:
                    notes.append(note_text)

        # Extract cross-references from definition
        cross_refs = extract_cross_refs(definition)

        # Build URI
        uri = f"{BASE_URI}27000/term/{term_id}"
        scheme_uri = f"{BASE_URI}27000/VocabularyScheme"

        is_top_concept = term_id in hub_terms

        nodes.append({
            "id": term_id,
            "label": term,
            "definition": definition,
            "notes": notes,
            "uri": uri,
            "inScheme": scheme_uri,
            "isTopConcept": is_top_concept,
            "related": cross_refs
        })
        node_ids.add(term_id)

    # Filter related to only include valid term IDs
    for node in nodes:
        node["related"] = [ref for ref in node["related"] if ref in node_ids and ref != node["id"]]

    # Build links (deduplicated, symmetric)
    links = []
    seen_links = set()

    for node in nodes:
        for related_id in node["related"]:
            link_key = tuple(sorted([node["id"], related_id]))
            if link_key not in seen_links:
                seen_links.add(link_key)
                links.append({
                    "source": link_key[0],
                    "target": link_key[1],
                    "type": "related"
                })

    # Sort nodes by notation number
    def sort_key(x):
        try:
            return float(x["id"].replace("3.", ""))
        except:
            return 0
    nodes.sort(key=sort_key)

    metadata = {
        "source": {
            "file": str(excel_path),
            "type": "Excel"
        },
        "scheme": {
            "uri": f"{BASE_URI}27000/VocabularyScheme",
            "label": "ISO/IEC 27000:2018 Information Security Vocabulary",
            "description": "77 terms and definitions for information security management systems"
        },
        "stats": {
            "totalTerms": len(nodes),
            "totalRelationships": len(links),
            "topConcepts": sum(1 for n in nodes if n["isTopConcept"])
        }
    }

    return {"nodes": nodes, "links": links, "metadata": metadata}


if __name__ == "__main__":
    data = extract_iso27000_from_excel()

    print(f"Extracted {len(data['nodes'])} nodes and {len(data['links'])} links")
    print(f"Top concepts: {data['metadata']['stats']['topConcepts']}")
    print(f"Source: {data['metadata']['source']['file']}")

    # Save to output and docs
    for output_path in ["output/iso27000-graph.json", "docs/iso27000-graph.json"]:
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        with open(output_path, "w") as f:
            json.dump(data, f, indent=2)
        print(f"Saved to {output_path}")
