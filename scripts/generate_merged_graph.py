#!/usr/bin/env python3
"""
Generate merged graph JSON combining ISO 27000, 27001, and 27002.
Uses unified schema format with Concept_ID, Label, Definition, Parent_ID, Theme columns.
"""

import json
import re
import pandas as pd
from pathlib import Path


def parse_iso27000_terms(xlsx_path):
    """Extract ISO 27000 vocabulary terms from unified schema."""
    df = pd.read_excel(xlsx_path, sheet_name='ISO27000_Vocabulary', dtype={'Concept_ID': str, 'Parent_ID': str})
    terms = []

    for _, row in df.iterrows():
        concept_id = str(row['Concept_ID']).strip()
        if not concept_id or concept_id == 'nan':
            continue

        # Collect notes
        notes = []
        for col in ['Note_1', 'Note_2', 'Note_3']:
            if col in row and pd.notna(row[col]) and str(row[col]).strip():
                notes.append(str(row[col]).strip())

        # Extract cross-references from definition (e.g., "(3.28)")
        definition = str(row.get('Definition', '')) if pd.notna(row.get('Definition')) else ''
        related = re.findall(r'\((\d+\.\d+)\)', definition)

        terms.append({
            'id': f'27000-{concept_id}',
            'label': str(row['Label']) if pd.notna(row['Label']) else concept_id,
            'definition': definition,
            'notes': notes,
            'standard': '27000',
            'type': 'term',
            'uri': f'https://wiren301.github.io/iso27001-skos-taxonomy/27000/term/{concept_id}',
            'related': related
        })

    return terms


def parse_iso27001(xlsx_path):
    """Extract ISO 27001 clauses and Annex A controls from unified schema."""
    df = pd.read_excel(xlsx_path, sheet_name='ISO27001_Requirements', dtype={'Concept_ID': str, 'Parent_ID': str})
    clauses = []
    controls = []

    for _, row in df.iterrows():
        concept_id = str(row['Concept_ID']).strip()
        if not concept_id or concept_id == 'nan':
            continue

        theme = str(row.get('Theme', '')) if pd.notna(row.get('Theme')) else ''
        label = str(row['Label']) if pd.notna(row['Label']) else concept_id
        definition = str(row.get('Definition', '')) if pd.notna(row.get('Definition')) else ''
        parent_id = str(row.get('Parent_ID', '')) if pd.notna(row.get('Parent_ID')) else None

        if theme == 'ISMS Requirements':
            # ISMS clause
            clauses.append({
                'id': f'27001-{concept_id}',
                'label': label,
                'definition': definition,
                'standard': '27001',
                'type': 'clause',
                'broader': f'27001-{parent_id}' if parent_id and parent_id != 'nan' else None,
                'uri': f'https://wiren301.github.io/iso27001-skos-taxonomy/27001/clause/{concept_id}'
            })
        else:
            # Annex A control - concept_id already has "A." prefix (e.g., "A.5.1")
            # Strip the "A." to match ISO 27002 control IDs (e.g., "5.1")
            control_num = concept_id[2:] if concept_id.startswith('A.') else concept_id
            controls.append({
                'id': f'27001-{concept_id}',
                'label': label,
                'definition': definition,
                'standard': '27001',
                'type': 'annex-a-control',
                'theme': theme,
                'iso27002Match': f'27002-{control_num}',
                'uri': f'https://wiren301.github.io/iso27001-skos-taxonomy/27001/control/{concept_id}'
            })

    return clauses, controls


def parse_hashtag_list(value):
    """Parse hashtag-separated values like '#Confidentiality #Integrity' into a list."""
    if pd.isna(value) or not str(value).strip():
        return []
    return [v.strip() for v in str(value).replace('#', ' ').split() if v.strip()]


def parse_iso27002_controls(xlsx_path):
    """Extract ISO 27002 controls from unified schema with all attributes."""
    df = pd.read_excel(xlsx_path, sheet_name='ISO27002_Controls', dtype={'Concept_ID': str, 'Parent_ID': str})
    controls = []

    theme_map = {'5': 'Organizational', '6': 'People', '7': 'Physical', '8': 'Technological'}

    for _, row in df.iterrows():
        control_id = str(row['Concept_ID']).strip()
        if not control_id or control_id == 'nan':
            continue

        theme_num = control_id.split('.')[0]
        theme = str(row.get('Theme', '')) if pd.notna(row.get('Theme')) else theme_map.get(theme_num, '')

        # Parse attribute columns
        cia = parse_hashtag_list(row.get('Attr_CIA', ''))
        nist = parse_hashtag_list(row.get('Attr_Cyber', ''))
        ops_cap = parse_hashtag_list(row.get('Attr_Ops_Cap', ''))
        sec_domain = parse_hashtag_list(row.get('Attr_Sec_Domain', ''))

        # Get purpose and guidance
        purpose = str(row.get('Purpose', '')) if pd.notna(row.get('Purpose')) else ''
        guidance = str(row.get('Guidance', '')) if pd.notna(row.get('Guidance')) else ''

        controls.append({
            'id': f'27002-{control_id}',
            'label': str(row['Label']) if pd.notna(row['Label']) else control_id,
            'definition': str(row.get('Definition', '')) if pd.notna(row.get('Definition')) else '',
            'purpose': purpose,
            'guidance': guidance,
            'standard': '27002',
            'type': 'control',
            'theme': theme,
            'uri': f'https://wiren301.github.io/iso27001-skos-taxonomy/27002/control/{control_id}',
            # Practitioner attributes
            'cia': cia,
            'nist': nist,
            'operational_capability': ops_cap,
            'security_domain': sec_domain
        })

    return controls


def build_links(nodes):
    """Build all relationship links."""
    links = []
    node_ids = {n['id'] for n in nodes}

    # 1. ISO 27000 internal related links
    for node in nodes:
        if node['standard'] == '27000' and 'related' in node:
            for rel_id in node['related']:
                target_id = f'27000-{rel_id}'
                if target_id in node_ids and target_id != node['id']:
                    links.append({
                        'source': node['id'],
                        'target': target_id,
                        'type': 'related',
                        'linkType': 'within-standard'
                    })

    # 2. ISO 27001 clause hierarchy (broader/narrower)
    for node in nodes:
        if node['standard'] == '27001' and node['type'] == 'clause' and node.get('broader'):
            if node['broader'] in node_ids:
                links.append({
                    'source': node['broader'],
                    'target': node['id'],
                    'type': 'narrower',
                    'linkType': 'within-standard'
                })

    # 3. ISO 27001 Annex A -> ISO 27002 exactMatch
    for node in nodes:
        if node.get('iso27002Match') and node['iso27002Match'] in node_ids:
            links.append({
                'source': node['id'],
                'target': node['iso27002Match'],
                'type': 'exactMatch',
                'linkType': 'cross-standard'
            })

    # 4. ISO 27001/27002 -> ISO 27000 vocabulary usage (key terms)
    key_terms = {
        'risk': '27000-3.61',
        'control': '27000-3.14',
        'policy': '27000-3.53',
        'process': '27000-3.54',
        'information security': '27000-3.28',
        'organization': '27000-3.50',
        'objective': '27000-3.49',
        'audit': '27000-3.3',
        'management system': '27000-3.41'
    }

    for node in nodes:
        if node['standard'] in ['27001', '27002']:
            label_lower = node['label'].lower()
            definition_lower = node.get('definition', '').lower()

            for term, term_id in key_terms.items():
                if term_id in node_ids:
                    if term in label_lower or term in definition_lower:
                        existing = any(
                            l['source'] == node['id'] and l['target'] == term_id
                            for l in links
                        )
                        if not existing:
                            links.append({
                                'source': node['id'],
                                'target': term_id,
                                'type': 'usesTerm',
                                'linkType': 'vocabulary'
                            })

    return links


def main():
    docs_path = Path('docs')

    print("Parsing standards...")

    # Parse all standards
    iso27000_terms = parse_iso27000_terms('schemas/ISO27000_UNIFIED.xlsx')
    print(f"  ISO 27000: {len(iso27000_terms)} terms")

    iso27001_clauses, iso27001_controls = parse_iso27001('schemas/ISO27001_UNIFIED.xlsx')
    print(f"  ISO 27001 clauses: {len(iso27001_clauses)}")
    print(f"  ISO 27001 Annex A: {len(iso27001_controls)} controls")

    iso27002_controls = parse_iso27002_controls('schemas/ISO27002_UNIFIED.xlsx')
    print(f"  ISO 27002: {len(iso27002_controls)} controls")

    # Combine all nodes
    all_nodes = iso27000_terms + iso27001_clauses + iso27001_controls + iso27002_controls
    print(f"\nTotal nodes: {len(all_nodes)}")

    # Build links
    print("Building relationships...")
    links = build_links(all_nodes)
    print(f"Total links: {len(links)}")

    # Count by type
    link_types = {}
    for link in links:
        lt = link['type']
        link_types[lt] = link_types.get(lt, 0) + 1
    for lt, count in sorted(link_types.items()):
        print(f"  {lt}: {count}")

    # Build output
    graph = {
        'nodes': all_nodes,
        'links': links,
        'metadata': {
            'standards': {
                '27000': {'label': 'ISO/IEC 27000:2018', 'description': 'Vocabulary', 'count': len(iso27000_terms)},
                '27001': {'label': 'ISO/IEC 27001:2022', 'description': 'ISMS Requirements', 'count': len(iso27001_clauses) + len(iso27001_controls)},
                '27002': {'label': 'ISO/IEC 27002:2022', 'description': 'Security Controls', 'count': len(iso27002_controls)}
            },
            'stats': {
                'totalNodes': len(all_nodes),
                'totalLinks': len(links),
                'linkTypes': link_types
            },
            'colors': {
                '27000': '#1f6feb',
                '27001-clause': '#3fb950',
                '27001-control': '#2ea043',
                '27002': '#f0883e'
            }
        }
    }

    # Save
    output_path = docs_path / 'merged-graph.json'
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(graph, f, indent=2, ensure_ascii=False)
    print(f"\nSaved: {output_path}")


if __name__ == '__main__':
    main()
