import { useState, useCallback, useMemo, useEffect } from 'react';
import { useLocation } from 'react-router-dom';
import { Graph } from '../components/Graph';
import { NodeDetails } from '../components/NodeDetails';
import { useGraphData } from '../hooks/useGraphData';
import type { GraphNode } from '../types';
import './Controls.css';

type ViewMode = 'list' | 'graph';

const THEME_INFO: Record<string, { label: string; color: string }> = {
  '5': { label: 'Organizational', color: '#f0883e' },
  '6': { label: 'People', color: '#a371f7' },
  '7': { label: 'Physical', color: '#3fb950' },
  '8': { label: 'Technological', color: '#1f6feb' },
};

const CLUSTERS = {
  '5': { x: 0.35, y: 0.35 },
  '6': { x: 0.65, y: 0.35 },
  '7': { x: 0.35, y: 0.65 },
  '8': { x: 0.65, y: 0.65 },
};

export function Controls() {
  const location = useLocation();
  const { data, loading, error } = useGraphData('iso27002-graph.json');
  const [selectedNode, setSelectedNode] = useState<GraphNode | null>(null);
  const [viewMode, setViewMode] = useState<ViewMode>('list');
  const [selectedTheme, setSelectedTheme] = useState<string>('5');
  const [expandedItem, setExpandedItem] = useState<string | null>(null);

  // Handle navigation with selectedId
  useEffect(() => {
    const state = location.state as { selectedId?: string } | null;
    if (state?.selectedId && data) {
      const node = data.nodes.find(n => n.id === state.selectedId);
      if (node) {
        const theme = node.theme ?? node.notation?.split('.')[0] ?? '5';
        setSelectedTheme(theme);
        setExpandedItem(node.id);
        setSelectedNode(node);
      }
    }
  }, [location.state, data]);

  // Group controls by theme
  const controlsByTheme = useMemo(() => {
    if (!data) return {};
    const groups: Record<string, GraphNode[]> = {};

    data.nodes.forEach(node => {
      if (node.type === 'theme') return;
      const theme = node.theme ?? node.notation?.split('.')[0];
      if (!theme || !THEME_INFO[theme]) return;
      if (!groups[theme]) groups[theme] = [];
      groups[theme].push(node);
    });

    // Sort by notation
    Object.keys(groups).forEach(theme => {
      groups[theme].sort((a, b) => {
        const aNum = parseFloat(a.notation?.replace(/[^\d.]/g, '') ?? '0');
        const bNum = parseFloat(b.notation?.replace(/[^\d.]/g, '') ?? '0');
        return aNum - bNum;
      });
    });

    return groups;
  }, [data]);

  const getClusterKey = useCallback((node: GraphNode): string | undefined => {
    if (node.type === 'theme') return node.id.replace('theme-', '');
    return node.theme ?? node.notation?.split('.')[0];
  }, []);

  const getNodeColor = useCallback((node: GraphNode) => {
    if (node.type === 'theme') {
      const themeId = node.id.replace('theme-', '');
      return THEME_INFO[themeId]?.color ?? '#6e7681';
    }
    const theme = node.theme ?? node.notation?.split('.')[0];
    return THEME_INFO[theme ?? '']?.color ?? '#6e7681';
  }, []);

  const getNodeRadius = useCallback((node: GraphNode) => {
    return node.type === 'theme' ? 16 : 8;
  }, []);

  if (loading) {
    return <div className="loading"><p>Loading ISO 27002 controls...</p></div>;
  }

  if (error || !data) {
    return <div className="error"><p>Error: {error ?? 'Failed to load data'}</p></div>;
  }

  const selectedControls = controlsByTheme[selectedTheme] ?? [];

  return (
    <div className="controls-page">
      <div className="controls-header">
        <div className="header-left">
          <h1>ISO 27002 Controls</h1>
          <span className="control-count">{data.nodes.filter(n => n.type !== 'theme').length} controls</span>
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
        <div className="controls-list-view">
          <aside className="theme-sidebar">
            <h3>Themes</h3>
            <ul className="theme-nav">
              {Object.entries(THEME_INFO).map(([id, { label, color }]) => (
                <li key={id}>
                  <button
                    className={selectedTheme === id ? 'active' : ''}
                    onClick={() => {
                      setSelectedTheme(id);
                      setExpandedItem(null);
                    }}
                    style={{ borderLeftColor: color }}
                  >
                    <span className="theme-num">{id}</span>
                    <span className="theme-label">{label}</span>
                    <span className="theme-count">{controlsByTheme[id]?.length ?? 0}</span>
                  </button>
                </li>
              ))}
            </ul>
          </aside>

          <div className="controls-content">
            <header className="content-header">
              <h2 style={{ color: THEME_INFO[selectedTheme]?.color }}>
                {THEME_INFO[selectedTheme]?.label} Controls
              </h2>
              <span className="item-count">{selectedControls.length} controls</span>
            </header>

            <div className="controls-list">
              {selectedControls.map(control => (
                <article
                  key={control.id}
                  className={`control-card ${expandedItem === control.id ? 'expanded' : ''}`}
                >
                  <div
                    className="control-header"
                    onClick={() => setExpandedItem(expandedItem === control.id ? null : control.id)}
                  >
                    <span className="toggle-icon">{expandedItem === control.id ? '▼' : '▶'}</span>
                    <span className="notation">{control.notation}</span>
                    <h3>{control.label}</h3>
                  </div>

                  {expandedItem === control.id && (
                    <div className="control-details">
                      {control.purpose && (
                        <div className="detail-section">
                          <h4>Purpose</h4>
                          <p>{control.purpose}</p>
                        </div>
                      )}

                      {control.definition && (
                        <div className="detail-section">
                          <h4>Control</h4>
                          <p>{control.definition}</p>
                        </div>
                      )}

                      <div className="control-attributes">
                        {control.cia && control.cia.length > 0 && (
                          <div className="attribute-group">
                            <span className="attribute-label">CIA:</span>
                            {control.cia.map(attr => (
                              <span key={attr} className="attribute-tag cia">{attr[0]}</span>
                            ))}
                          </div>
                        )}
                        {control.nist && control.nist.length > 0 && (
                          <div className="attribute-group">
                            <span className="attribute-label">NIST:</span>
                            {control.nist.map(attr => (
                              <span key={attr} className="attribute-tag nist">{attr}</span>
                            ))}
                          </div>
                        )}
                        {control.operational_capability && control.operational_capability.length > 0 && (
                          <div className="attribute-group">
                            <span className="attribute-label">Capability:</span>
                            {control.operational_capability.map(attr => (
                              <span key={attr} className="attribute-tag capability">{attr.replace(/_/g, ' ')}</span>
                            ))}
                          </div>
                        )}
                      </div>

                      {control.guidance && (
                        <details className="guidance-section">
                          <summary>Implementation Guidance</summary>
                          <p>{control.guidance}</p>
                        </details>
                      )}
                    </div>
                  )}
                </article>
              ))}
            </div>
          </div>
        </div>
      ) : (
        <div className="controls-graph-view">
          <Graph
            initialNodes={data.nodes}
            initialLinks={data.links}
            width={window.innerWidth}
            height={window.innerHeight - 120}
            getNodeColor={getNodeColor}
            getNodeRadius={getNodeRadius}
            selectedNode={selectedNode}
            onNodeSelect={setSelectedNode}
            chargeStrength={-30}
            linkDistance={50}
            linkStrength={0.3}
            collisionRadius={12}
            clusters={CLUSTERS}
            getClusterKey={getClusterKey}
            clusterStrength={0.08}
          />
          <NodeDetails node={selectedNode} onClose={() => setSelectedNode(null)} />
          <div className="legend">
            {Object.entries(THEME_INFO).map(([id, { label, color }]) => (
              <div key={id} className="legend-item">
                <span className="legend-color" style={{ background: color }} />
                <span>{label}</span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}
