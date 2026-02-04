// docs-app/src/hooks/useSearch.ts
import { useState, useEffect, useMemo } from 'react';
import type { GraphData, GraphNode } from '../types';

export interface SearchResult {
  id: string;
  label: string;
  notation?: string;
  description: string;
  type: 'vocabulary' | 'clause' | 'control';
  standard: string;
  path: string;
}

interface SearchIndex {
  vocabulary: GraphNode[];
  clauses: GraphNode[];
  controls: GraphNode[];
}

export function useSearch() {
  const [index, setIndex] = useState<SearchIndex | null>(null);
  const [loading, setLoading] = useState(true);
  const [query, setQuery] = useState('');

  // Load all data files on mount
  useEffect(() => {
    const baseUrl = import.meta.env.BASE_URL;

    Promise.all([
      fetch(`${baseUrl}data/iso27000-graph.json`).then(r => r.json()),
      fetch(`${baseUrl}data/iso27001-graph.json`).then(r => r.json()),
      fetch(`${baseUrl}data/iso27002-graph.json`).then(r => r.json()),
    ])
      .then(([vocab, clauses, controls]: GraphData[]) => {
        setIndex({
          vocabulary: vocab.nodes,
          clauses: clauses.nodes,
          controls: controls.nodes,
        });
        setLoading(false);
      })
      .catch(() => setLoading(false));
  }, []);

  const results = useMemo((): SearchResult[] => {
    if (!index || !query.trim()) return [];

    const q = query.toLowerCase().trim();
    const matches: SearchResult[] = [];

    // Search vocabulary - goes to vocabulary list view
    index.vocabulary.forEach(node => {
      if (
        node.label.toLowerCase().includes(q) ||
        node.definition?.toLowerCase().includes(q)
      ) {
        matches.push({
          id: node.id,
          label: node.label,
          notation: node.id,
          description: node.definition?.slice(0, 100) || '',
          type: 'vocabulary',
          standard: 'ISO 27000',
          path: '/vocabulary',
        });
      }
    });

    // Search clauses - goes to clauses view
    index.clauses.forEach(node => {
      if (
        node.label.toLowerCase().includes(q) ||
        node.notation?.toLowerCase().includes(q) ||
        node.definition?.toLowerCase().includes(q)
      ) {
        matches.push({
          id: node.id,
          label: node.label,
          notation: node.notation,
          description: node.definition?.slice(0, 100) || '',
          type: 'clause',
          standard: 'ISO 27001',
          path: '/clauses',
        });
      }
    });

    // Search controls - goes to controls list view
    index.controls.forEach(node => {
      if (
        node.label.toLowerCase().includes(q) ||
        node.notation?.toLowerCase().includes(q) ||
        node.definition?.toLowerCase().includes(q) ||
        node.purpose?.toLowerCase().includes(q)
      ) {
        matches.push({
          id: node.id,
          label: node.label,
          notation: node.notation,
          description: node.purpose || node.definition?.slice(0, 100) || '',
          type: 'control',
          standard: 'ISO 27002',
          path: '/controls',
        });
      }
    });

    return matches.slice(0, 20); // Limit results
  }, [index, query]);

  return { query, setQuery, results, loading };
}
