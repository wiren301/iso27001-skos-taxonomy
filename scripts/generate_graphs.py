#!/usr/bin/env python3
"""
Generate graph JSON files for ISO 27001 and ISO 27002 visualizations.
Extracts data from the main TTL file and creates D3-compatible graph structures.
"""

import re
import json
from pathlib import Path
from collections import defaultdict

def parse_ttl_file(ttl_path):
    """Parse TTL file and extract concepts by namespace."""
    with open(ttl_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # Split into blocks (each concept definition)
    blocks = re.split(r'\n(?=<https://)', content)

    concepts = {
        '27001': [],  # ISO 27001 controls (Annex A)
        '27002': [],  # ISO 27002 controls
    }

    for block in blocks:
        if not block.strip():
            continue

        # Extract URI
        uri_match = re.match(r'<(https://[^>]+)>', block)
        if not uri_match:
            continue
        uri = uri_match.group(1)

        # Determine namespace
        if '/27001/control/' in uri:
            ns = '27001'
        elif '/27002/control/' in uri:
            ns = '27002'
        else:
            continue

        # Skip if it's a Collection (attribute grouping)
        if 'a skos:Collection' in block:
            continue

        # Extract notation
        notation_match = re.search(r'skos:notation "([^"]+)"', block)
        notation = notation_match.group(1) if notation_match else ''

        # Extract prefLabel (take first one if multiple)
        label_match = re.search(r'skos:prefLabel "([^"]+)"@en', block)
        label = label_match.group(1) if label_match else notation

        # Extract definition
        def_match = re.search(r'skos:definition "([^"]+)"@en', block)
        definition = def_match.group(1) if def_match else ''

        # Extract notes
        notes = re.findall(r'skos:note "([^"]+)"@en', block)

        # Extract inScheme
        scheme_match = re.search(r'skos:inScheme ([^;\s]+)', block)
        scheme = scheme_match.group(1) if scheme_match else ''

        # Check if top concept
        is_top = 'skos:topConceptOf' in block

        # Extract broader relationships
        broader = []
        broader_matches = re.findall(r'skos:broader <([^>]+)>', block)
        for b in broader_matches:
            broader.append(b)

        # Extract narrower relationships
        narrower = []
        narrower_matches = re.findall(r'skos:narrower <([^>]+)>', block)
        for n in narrower_matches:
            narrower.append(n)

        # Extract related
        related = []
        related_matches = re.findall(r'skos:related <([^>]+)>', block)
        for r in related_matches:
            related.append(r)

        # Extract exactMatch
        exact_matches = []
        exact_match_matches = re.findall(r'skos:exactMatch <([^>]+)>', block)
        for e in exact_match_matches:
            exact_matches.append(e)

        concept = {
            'uri': uri,
            'notation': notation,
            'label': label,
            'definition': definition,
            'notes': notes,
            'scheme': scheme,
            'isTopConcept': is_top,
            'broader': broader,
            'narrower': narrower,
            'related': related,
            'exactMatch': exact_matches
        }

        concepts[ns].append(concept)

    return concepts


def generate_27002_graph(concepts):
    """Generate ISO 27002 controls graph with theme hierarchy."""

    # Control themes in ISO 27002:2022
    themes = {
        '5': {'label': 'Organizational controls', 'color': '#da3633'},
        '6': {'label': 'People controls', 'color': '#a371f7'},
        '7': {'label': 'Physical controls', 'color': '#3fb950'},
        '8': {'label': 'Technological controls', 'color': '#1f6feb'}
    }

    nodes = []
    links = []
    node_ids = set()

    # Add theme nodes
    for theme_id, theme_info in themes.items():
        nodes.append({
            'id': f'theme-{theme_id}',
            'label': theme_info['label'],
            'notation': f'Clause {theme_id}',
            'definition': f'ISO 27002:2022 {theme_info["label"]}',
            'type': 'theme',
            'isTopConcept': True,
            'uri': f'https://wiren301.github.io/iso27001-skos-taxonomy/27002/theme/{theme_id}',
            'inScheme': 'iso27002:ControlScheme'
        })
        node_ids.add(f'theme-{theme_id}')

    # Add control nodes from 27001 (which maps to 27002)
    for concept in concepts['27001']:
        notation = concept['notation']
        if not notation:
            continue

        # Handle A.x.x format (strip A. prefix) to get theme
        clean_notation = notation.lstrip('A.') if notation.startswith('A.') else notation

        # Determine theme from notation (5.x, 6.x, 7.x, 8.x)
        theme_prefix = clean_notation.split('.')[0] if '.' in clean_notation else ''
        if theme_prefix not in themes:
            continue

        node_id = f'control-{notation}'
        if node_id in node_ids:
            continue

        nodes.append({
            'id': node_id,
            'label': concept['label'],
            'notation': clean_notation,  # Use clean notation without A. prefix
            'definition': concept['definition'],
            'type': 'control',
            'theme': theme_prefix,
            'isTopConcept': False,
            'uri': concept['uri'],
            'inScheme': 'iso27002:ControlScheme'
        })
        node_ids.add(node_id)

        # Link to theme
        links.append({
            'source': f'theme-{theme_prefix}',
            'target': node_id,
            'type': 'narrower'
        })

    # Sort nodes: themes first, then controls by notation
    nodes.sort(key=lambda x: (0 if x['type'] == 'theme' else 1, x.get('notation', '')))

    return {
        'nodes': nodes,
        'links': links,
        'metadata': {
            'source': {'file': 'output/iso-security.ttl', 'type': 'Turtle'},
            'scheme': {
                'uri': 'https://wiren301.github.io/iso27001-skos-taxonomy/27002/ControlScheme',
                'label': 'ISO/IEC 27002:2022 Information Security Controls',
                'description': '93 controls organized in 4 themes'
            },
            'stats': {
                'totalControls': len([n for n in nodes if n['type'] == 'control']),
                'totalThemes': len(themes),
                'totalRelationships': len(links)
            }
        }
    }


def generate_27001_graph(ttl_path):
    """Generate ISO 27001 clauses graph from TTL file."""

    with open(ttl_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # Main clauses in ISO 27001:2022
    main_clauses = {
        '4': {'label': 'Context of the organization', 'color': '#da3633'},
        '5': {'label': 'Leadership', 'color': '#f0883e'},
        '6': {'label': 'Planning', 'color': '#d29922'},
        '7': {'label': 'Support', 'color': '#3fb950'},
        '8': {'label': 'Operation', 'color': '#58a6ff'},
        '9': {'label': 'Performance evaluation', 'color': '#a371f7'},
        '10': {'label': 'Improvement', 'color': '#f778ba'}
    }

    nodes = []
    links = []
    node_ids = set()

    # Parse clauses from TTL
    blocks = re.split(r'\n(?=<https://)', content)
    clauses = []

    for block in blocks:
        if '/27001/clause/' not in block:
            continue
        if 'a skos:Concept' not in block:
            continue

        uri_match = re.match(r'<(https://[^>]+)>', block)
        if not uri_match:
            continue
        uri = uri_match.group(1)

        # Extract clause number from URI
        clause_num = uri.split('/clause/')[-1]

        # Extract prefLabel
        label_match = re.search(r'skos:prefLabel "([^"]+)"@en', block)
        label = label_match.group(1) if label_match else clause_num

        # Extract definition
        def_match = re.search(r'skos:definition "([^"]+)"@en', block)
        definition = def_match.group(1) if def_match else ''

        # Extract broader
        broader_match = re.search(r'skos:broader <([^>]+)>', block)
        broader = broader_match.group(1).split('/clause/')[-1] if broader_match else None

        # Determine main clause (first digit)
        main_clause = clause_num.split('.')[0]

        # Determine depth
        depth = clause_num.count('.') + 1

        clauses.append({
            'id': clause_num,
            'uri': uri,
            'label': label,
            'definition': definition,
            'broader': broader,
            'mainClause': main_clause,
            'depth': depth
        })

    # Add main clause nodes first
    for clause_id, info in main_clauses.items():
        # Check if this main clause exists in parsed data
        existing = next((c for c in clauses if c['id'] == clause_id), None)
        if existing:
            nodes.append({
                'id': f'clause-{clause_id}',
                'label': f'{clause_id}. {info["label"]}',
                'notation': clause_id,
                'definition': existing['definition'],
                'type': 'main-clause',
                'mainClause': clause_id,
                'isTopConcept': True,
                'uri': existing['uri'],
                'inScheme': 'iso27001:ISMSScheme'
            })
            node_ids.add(f'clause-{clause_id}')

    # Add sub-clause nodes
    for clause in clauses:
        if clause['id'] in main_clauses:
            continue  # Skip main clauses (already added)

        node_id = f'clause-{clause["id"]}'
        if node_id in node_ids:
            continue

        nodes.append({
            'id': node_id,
            'label': clause['label'],
            'notation': clause['id'],
            'definition': clause['definition'],
            'type': 'sub-clause',
            'mainClause': clause['mainClause'],
            'depth': clause['depth'],
            'isTopConcept': False,
            'uri': clause['uri'],
            'inScheme': 'iso27001:ISMSScheme'
        })
        node_ids.add(node_id)

        # Create link to broader
        if clause['broader']:
            links.append({
                'source': f'clause-{clause["broader"]}',
                'target': node_id,
                'type': 'narrower'
            })

    # Sort nodes by notation
    def sort_key(n):
        parts = n['notation'].split('.')
        result = []
        for p in parts:
            try:
                result.append((0, int(p)))
            except ValueError:
                result.append((1, p))
        return result

    nodes.sort(key=sort_key)

    return {
        'nodes': nodes,
        'links': links,
        'metadata': {
            'source': {'file': 'output/iso-security.ttl', 'type': 'Turtle'},
            'scheme': {
                'uri': 'https://wiren301.github.io/iso27001-skos-taxonomy/27001/ISMSScheme',
                'label': 'ISO/IEC 27001:2022 ISMS Requirements',
                'description': 'Information security management system requirements organized in clauses 4-10'
            },
            'stats': {
                'totalClauses': len(nodes),
                'mainClauses': len([n for n in nodes if n['type'] == 'main-clause']),
                'totalRelationships': len(links)
            }
        },
        'colors': main_clauses
    }


def main():
    ttl_path = Path('output/iso-security.ttl')
    docs_path = Path('docs')

    print("Parsing TTL file...")
    concepts = parse_ttl_file(ttl_path)

    print(f"Found {len(concepts['27001'])} ISO 27001 controls")
    print(f"Found {len(concepts['27002'])} ISO 27002 controls")

    # Generate ISO 27001 graph (clauses)
    print("\nGenerating ISO 27001 graph...")
    graph_27001 = generate_27001_graph(ttl_path)
    output_27001 = docs_path / 'iso27001-graph.json'
    with open(output_27001, 'w', encoding='utf-8') as f:
        json.dump(graph_27001, f, indent=2, ensure_ascii=False)
    print(f"Saved: {output_27001}")
    print(f"  - {graph_27001['metadata']['stats']['totalClauses']} clauses")
    print(f"  - {graph_27001['metadata']['stats']['mainClauses']} main clauses")

    # Generate ISO 27002 graph
    print("\nGenerating ISO 27002 graph...")
    graph_27002 = generate_27002_graph(concepts)
    output_27002 = docs_path / 'iso27002-graph.json'
    with open(output_27002, 'w', encoding='utf-8') as f:
        json.dump(graph_27002, f, indent=2, ensure_ascii=False)
    print(f"Saved: {output_27002}")
    print(f"  - {graph_27002['metadata']['stats']['totalControls']} controls")
    print(f"  - {graph_27002['metadata']['stats']['totalThemes']} themes")

    print("\nDone!")


if __name__ == '__main__':
    main()
