import { useState, useCallback, useMemo, useEffect } from 'react';
import { useLocation } from 'react-router-dom';
import { Graph } from '../components/Graph';
import { NodeDetails } from '../components/NodeDetails';
import { useGraphData } from '../hooks/useGraphData';
import type { GraphNode } from '../types';
import './Clauses.css';

type ViewMode = 'list' | 'graph';

const CLAUSE_INFO: Record<string, { label: string; color: string }> = {
  '4': { label: 'Context of the organization', color: '#2dd4bf' },
  '5': { label: 'Leadership', color: '#f0883e' },
  '6': { label: 'Planning', color: '#d29922' },
  '7': { label: 'Support', color: '#3fb950' },
  '8': { label: 'Operation', color: '#58a6ff' },
  '9': { label: 'Performance evaluation', color: '#a371f7' },
  '10': { label: 'Improvement', color: '#f778ba' },
};

export function Clauses() {
  const location = useLocation();
  const { data, loading, error } = useGraphData('iso27001-graph.json');
  const [selectedNode, setSelectedNode] = useState<GraphNode | null>(null);
  const [viewMode, setViewMode] = useState<ViewMode>('list');
  const [selectedClause, setSelectedClause] = useState<string>('4');
  const [expandedItem, setExpandedItem] = useState<string | null>(null);

  // Handle navigation with selectedId
  useEffect(() => {
    const state = location.state as { selectedId?: string } | null;
    if (state?.selectedId && data) {
      const node = data.nodes.find(n => n.id === state.selectedId);
      if (node) {
        const mainClause = node.mainClause ?? node.notation?.split('.')[0] ?? '4';
        setSelectedClause(mainClause);
        setExpandedItem(node.id);
        setSelectedNode(node);
      }
    }
  }, [location.state, data]);

  // Group clauses by main clause
  const clausesByMain = useMemo(() => {
    if (!data) return {};
    const groups: Record<string, GraphNode[]> = {};

    data.nodes.forEach(node => {
      if (node.type === 'main-clause') return;
      const mainClause = node.mainClause ?? node.notation?.split('.')[0];
      if (!mainClause || !CLAUSE_INFO[mainClause]) return;
      if (!groups[mainClause]) groups[mainClause] = [];
      groups[mainClause].push(node);
    });

    // Sort by notation
    Object.keys(groups).forEach(clause => {
      groups[clause].sort((a, b) => {
        const aParts = (a.notation ?? '').split('.').map(p => parseInt(p) || 0);
        const bParts = (b.notation ?? '').split('.').map(p => parseInt(p) || 0);
        for (let i = 0; i < Math.max(aParts.length, bParts.length); i++) {
          const diff = (aParts[i] ?? 0) - (bParts[i] ?? 0);
          if (diff !== 0) return diff;
        }
        return 0;
      });
    });

    return groups;
  }, [data]);

  const getNodeColor = useCallback((node: GraphNode) => {
    const mainClause = node.mainClause ?? node.notation?.split('.')[0];
    return CLAUSE_INFO[mainClause ?? '']?.color ?? '#6e7681';
  }, []);

  const getNodeRadius = useCallback((node: GraphNode) => {
    if (node.type === 'main-clause') return 18;
    const depth = (node.notation?.match(/\./g) || []).length;
    return depth >= 3 ? 6 : 10;
  }, []);

  const chargeStrength = useCallback((node: GraphNode) => {
    return node.type === 'main-clause' ? -400 : -150;
  }, []);

  const collisionRadius = useCallback((node: GraphNode) => {
    return node.type === 'main-clause' ? 40 : 25;
  }, []);

  const radialConfig = useMemo(() => ({
    radius: (node: GraphNode) => node.type === 'main-clause' ? 0 : 150,
    strength: 0.3,
  }), []);

  if (loading) {
    return <div className="loading"><p>Loading ISO 27001 clauses...</p></div>;
  }

  if (error || !data) {
    return <div className="error"><p>Error: {error ?? 'Failed to load data'}</p></div>;
  }

  const selectedClauses = clausesByMain[selectedClause] ?? [];

  return (
    <div className="clauses-page">
      <div className="clauses-header">
        <div className="header-left">
          <h1>ISO 27001 Clauses</h1>
          <span className="clause-count">{data.nodes.filter(n => n.type !== 'main-clause').length} requirements</span>
        </div>
        <div className="header-right">
          <div className="view-toggle">
            <button
              className={viewMode === 'list' ? 'active' : ''}
              onClick={() => setViewMode('list')}
            >
              List
            </button>
            <span>/</span>
            <button
              className={viewMode === 'graph' ? 'active' : ''}
              onClick={() => setViewMode('graph')}
            >
              Graph
            </button>
          </div>
        </div>
      </div>

      {viewMode === 'list' ? (
        <div className="clauses-list-view">
          <aside className="clause-sidebar">
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
                    <span className="clause-item-count">{clausesByMain[id]?.length ?? 0}</span>
                  </button>
                </li>
              ))}
            </ul>
          </aside>

          <div className="clauses-content">
            <header className="content-header">
              <h2 style={{ color: CLAUSE_INFO[selectedClause]?.color }}>
                Clause {selectedClause}: {CLAUSE_INFO[selectedClause]?.label}
              </h2>
              <span className="item-count">{selectedClauses.length} requirements</span>
            </header>

            <div className="clauses-list">
              {selectedClauses.map(clause => (
                <article
                  key={clause.id}
                  className={`clause-card ${expandedItem === clause.id ? 'expanded' : ''}`}
                >
                  <div
                    className="clause-item-header"
                    onClick={() => setExpandedItem(expandedItem === clause.id ? null : clause.id)}
                  >
                    <span className="toggle-icon">{expandedItem === clause.id ? '▼' : '▶'}</span>
                    <span className="notation">{clause.notation}</span>
                    <h3>{clause.label}</h3>
                  </div>

                  {expandedItem === clause.id && (
                    <div className="clause-details">
                      {clause.definition && (
                        <div className="detail-section">
                          <h4>Requirement</h4>
                          <p>{clause.definition}</p>
                        </div>
                      )}
                    </div>
                  )}
                </article>
              ))}
            </div>
          </div>
        </div>
      ) : (
        <div className="clauses-graph-view">
          <Graph
            initialNodes={data.nodes}
            initialLinks={data.links}
            width={window.innerWidth}
            height={window.innerHeight - 120}
            getNodeColor={getNodeColor}
            getNodeRadius={getNodeRadius}
            selectedNode={selectedNode}
            onNodeSelect={setSelectedNode}
            chargeStrength={chargeStrength}
            linkDistance={60}
            collisionRadius={collisionRadius}
            radial={radialConfig}
          />
          <NodeDetails node={selectedNode} onClose={() => setSelectedNode(null)} />
          <div className="legend">
            {Object.entries(CLAUSE_INFO).map(([id, { color }]) => (
              <div key={id} className="legend-item">
                <span className="legend-color" style={{ background: color }} />
                <span>Clause {id}</span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}
