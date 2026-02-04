import React, { useState, useMemo, useCallback } from 'react';
import { useGraphData } from '../hooks/useGraphData';
import type { GraphNode } from '../types';
import './Practitioner.css';

const CLAUSE_INFO: Record<string, { label: string; color: string }> = {
  '4': { label: 'Context of the organization', color: '#2dd4bf' },
  '5': { label: 'Leadership', color: '#f0883e' },
  '6': { label: 'Planning', color: '#d29922' },
  '7': { label: 'Support', color: '#3fb950' },
  '8': { label: 'Operation', color: '#58a6ff' },
  '9': { label: 'Performance evaluation', color: '#a371f7' },
  '10': { label: 'Improvement', color: '#f778ba' },
  'A': { label: 'Annex A Controls', color: '#f0883e' },
};

interface VocabTerm {
  id: string;
  label: string;
  definition?: string;
}

export function Practitioner() {
  const { data: mergedData, loading, error } = useGraphData('merged-graph.json');
  const [selectedClause, setSelectedClause] = useState<string>('4');
  const [expandedItem, setExpandedItem] = useState<string | null>(null);
  const [hoveredTerm, setHoveredTerm] = useState<VocabTerm | null>(null);

  // Extract vocabulary terms (ISO 27000)
  const vocabTerms = useMemo(() => {
    if (!mergedData) return new Map<string, VocabTerm>();
    const terms = new Map<string, VocabTerm>();
    for (const node of mergedData.nodes) {
      if (node.standard === '27000') {
        terms.set(node.id, {
          id: node.id,
          label: node.label.toLowerCase(),
          definition: node.definition,
        });
      }
    }
    return terms;
  }, [mergedData]);

  // Extract main clause from ID
  const getMainClause = (node: GraphNode): string => {
    const id = node.id;
    const match = id.match(/^27001-(\d+)/);
    if (match) return match[1];
    if (node.type === 'annex-a-control') return 'A';
    return '';
  };

  // Extract ISO 27001 items grouped by clause
  const clauses = useMemo(() => {
    if (!mergedData) return {};
    const grouped: Record<string, GraphNode[]> = {};

    for (const node of mergedData.nodes) {
      if (node.standard !== '27001') continue;
      const mainClause = getMainClause(node);
      if (!mainClause || !CLAUSE_INFO[mainClause]) continue;
      if (!grouped[mainClause]) grouped[mainClause] = [];
      grouped[mainClause].push(node);
    }

    for (const key of Object.keys(grouped)) {
      grouped[key].sort((a, b) => {
        const aMatch = a.id.match(/27001-(.+)/);
        const bMatch = b.id.match(/27001-(.+)/);
        const aId = aMatch ? aMatch[1] : a.id;
        const bId = bMatch ? bMatch[1] : b.id;
        const aParts = aId.split('.').map(p => p.replace(/\D/g, '')).map(Number);
        const bParts = bId.split('.').map(p => p.replace(/\D/g, '')).map(Number);
        for (let i = 0; i < Math.max(aParts.length, bParts.length); i++) {
          const diff = (aParts[i] ?? 0) - (bParts[i] ?? 0);
          if (diff !== 0) return diff;
        }
        return 0;
      });
    }
    return grouped;
  }, [mergedData]);

  // Build usesTerm map
  const usesTermMap = useMemo(() => {
    if (!mergedData) return new Map<string, string[]>();
    const map = new Map<string, string[]>();
    for (const link of mergedData.links) {
      if (link.type === 'usesTerm') {
        const source = typeof link.source === 'string' ? link.source : link.source.id;
        const target = typeof link.target === 'string' ? link.target : link.target.id;
        if (!map.has(source)) map.set(source, []);
        map.get(source)!.push(target);
      }
    }
    return map;
  }, [mergedData]);

  // Build exactMatch map
  const exactMatchMap = useMemo(() => {
    if (!mergedData) return new Map<string, GraphNode>();
    const map = new Map<string, GraphNode>();
    const nodeById = new Map(mergedData.nodes.map(n => [n.id, n]));

    for (const link of mergedData.links) {
      if (link.type === 'exactMatch') {
        const source = typeof link.source === 'string' ? link.source : link.source.id;
        const target = typeof link.target === 'string' ? link.target : link.target.id;
        const sourceNode = nodeById.get(source);
        const targetNode = nodeById.get(target);
        if (sourceNode?.standard === '27001' && targetNode?.standard === '27002') {
          map.set(source, targetNode);
        } else if (targetNode?.standard === '27001' && sourceNode?.standard === '27002') {
          map.set(target, sourceNode);
        }
      }
    }
    return map;
  }, [mergedData]);

  // Highlight vocabulary terms in text
  const highlightTerms = useCallback((text: string, itemId: string) => {
    const usedTermIds = usesTermMap.get(itemId) || [];
    if (usedTermIds.length === 0) return <span>{text}</span>;

    const termsToHighlight: VocabTerm[] = [];
    for (const termId of usedTermIds) {
      const term = vocabTerms.get(termId);
      if (term) termsToHighlight.push(term);
    }
    if (termsToHighlight.length === 0) return <span>{text}</span>;

    termsToHighlight.sort((a, b) => b.label.length - a.label.length);
    const pattern = new RegExp(
      `\\b(${termsToHighlight.map(t => t.label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|')})\\b`,
      'gi'
    );

    const parts: (string | React.ReactElement)[] = [];
    let lastIndex = 0;
    let match;

    while ((match = pattern.exec(text)) !== null) {
      if (match.index > lastIndex) {
        parts.push(text.slice(lastIndex, match.index));
      }
      const matchedText = match[1];
      const term = termsToHighlight.find(t => t.label.toLowerCase() === matchedText.toLowerCase());
      parts.push(
        <span
          key={match.index}
          className="vocab-term"
          onMouseEnter={() => term && setHoveredTerm(term)}
          onMouseLeave={() => setHoveredTerm(null)}
        >
          {matchedText}
        </span>
      );
      lastIndex = pattern.lastIndex;
    }
    if (lastIndex < text.length) {
      parts.push(text.slice(lastIndex));
    }
    return <span>{parts}</span>;
  }, [usesTermMap, vocabTerms]);

  if (loading) {
    return <div className="loading"><p>Loading guide...</p></div>;
  }

  if (error || !mergedData) {
    return <div className="error"><p>Error: {error ?? 'Failed to load data'}</p></div>;
  }

  const selectedClauses = clauses[selectedClause] ?? [];

  return (
    <div className="practitioner">
      <aside className="practitioner-sidebar">
        <h3>ISMS Requirements</h3>
        <ul className="clause-nav">
          {Object.entries(CLAUSE_INFO).map(([id, { label, color }]) => (
            <li key={id}>
              <button
                className={selectedClause === id ? 'active' : ''}
                onClick={() => {
                  setSelectedClause(id);
                  setExpandedItem(null);
                }}
                style={{ borderLeftColor: color }}
              >
                <span className="clause-num">{id}</span>
                <span className="clause-label">{label}</span>
              </button>
            </li>
          ))}
        </ul>
      </aside>

      <div className="practitioner-content">
        <header className="content-header">
          <h2>
            Clause {selectedClause}: {CLAUSE_INFO[selectedClause]?.label}
          </h2>
          <span className="clause-count">
            {selectedClauses.length} requirements
          </span>
        </header>

        <div className="clauses-list">
          {selectedClauses.map((item) => {
            const isExpanded = expandedItem === item.id;
            const linkedControl = exactMatchMap.get(item.id);
            const usedTerms = usesTermMap.get(item.id) || [];

            return (
              <article
                key={item.id}
                className={`clause-card ${isExpanded ? 'expanded' : ''}`}
              >
                <div
                  className="clause-header"
                  onClick={() => setExpandedItem(isExpanded ? null : item.id)}
                >
                  <span className="toggle-icon">{isExpanded ? '▼' : '▶'}</span>
                  <span className="notation">{item.notation || item.id.replace('27001-', '')}</span>
                  <h3>{item.label}</h3>
                  {linkedControl && (
                    <span className="has-control" title="Has linked ISO 27002 control">⟷</span>
                  )}
                </div>

                {isExpanded && (
                  <div className="clause-details">
                    {item.definition && (
                      <div className="detail-section">
                        <h4>Definition</h4>
                        <p>{highlightTerms(item.definition, item.id)}</p>
                      </div>
                    )}

                    {usedTerms.length > 0 && (
                      <div className="detail-section">
                        <h4>Vocabulary Terms Used</h4>
                        <div className="vocab-list">
                          {usedTerms.map(termId => {
                            const term = vocabTerms.get(termId);
                            return term ? (
                              <div key={termId} className="vocab-item">
                                <span className="vocab-label">{term.label}</span>
                                {term.definition && <span className="vocab-def">{term.definition}</span>}
                              </div>
                            ) : null;
                          })}
                        </div>
                      </div>
                    )}

                    {linkedControl && (
                      <div className="detail-section linked-control">
                        <h4>ISO 27002 Control (exactMatch)</h4>
                        <div className="control-card">
                          <div className="control-header">
                            <span className="control-notation">{linkedControl.notation || linkedControl.id.replace('27002-', '')}</span>
                            {linkedControl.theme && <span className="control-theme">{linkedControl.theme}</span>}
                          </div>
                          <h5>{linkedControl.label}</h5>

                          {linkedControl.purpose && (
                            <div className="control-section">
                              <h6>Purpose</h6>
                              <p>{linkedControl.purpose}</p>
                            </div>
                          )}

                          {linkedControl.definition && (
                            <div className="control-section">
                              <h6>Control</h6>
                              <p>{linkedControl.definition}</p>
                            </div>
                          )}

                          {/* Attributes */}
                          <div className="control-attributes">
                            {linkedControl.cia && linkedControl.cia.length > 0 && (
                              <div className="attribute-group">
                                <span className="attribute-label">CIA:</span>
                                {linkedControl.cia.map(attr => (
                                  <span key={attr} className="attribute-tag cia">{attr[0]}</span>
                                ))}
                              </div>
                            )}
                            {linkedControl.nist && linkedControl.nist.length > 0 && (
                              <div className="attribute-group">
                                <span className="attribute-label">NIST:</span>
                                {linkedControl.nist.map(attr => (
                                  <span key={attr} className="attribute-tag nist">{attr}</span>
                                ))}
                              </div>
                            )}
                            {linkedControl.operational_capability && linkedControl.operational_capability.length > 0 && (
                              <div className="attribute-group">
                                <span className="attribute-label">Capability:</span>
                                {linkedControl.operational_capability.map(attr => (
                                  <span key={attr} className="attribute-tag capability">{attr.replace(/_/g, ' ')}</span>
                                ))}
                              </div>
                            )}
                            {linkedControl.security_domain && linkedControl.security_domain.length > 0 && (
                              <div className="attribute-group">
                                <span className="attribute-label">Domain:</span>
                                {linkedControl.security_domain.map(attr => (
                                  <span key={attr} className="attribute-tag domain">{attr.replace(/_/g, ' ')}</span>
                                ))}
                              </div>
                            )}
                          </div>

                          {linkedControl.guidance && (
                            <details className="control-guidance">
                              <summary>Implementation Guidance</summary>
                              <p>{linkedControl.guidance}</p>
                            </details>
                          )}
                        </div>
                      </div>
                    )}
                  </div>
                )}
              </article>
            );
          })}
        </div>
      </div>

      {/* Vocabulary Tooltip */}
      {hoveredTerm && (
        <div className="vocab-tooltip">
          <strong>{hoveredTerm.label}</strong>
          {hoveredTerm.definition && <p>{hoveredTerm.definition}</p>}
        </div>
      )}
    </div>
  );
}
