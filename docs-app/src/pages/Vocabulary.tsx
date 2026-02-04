import { useState, useCallback, useMemo, useEffect } from 'react';
import { useLocation } from 'react-router-dom';
import { Graph } from '../components/Graph';
import { NodeDetails } from '../components/NodeDetails';
import { useGraphData } from '../hooks/useGraphData';
import type { GraphNode } from '../types';
import './Vocabulary.css';

type ViewMode = 'list' | 'graph';

export function Vocabulary() {
  const location = useLocation();
  const { data, loading, error } = useGraphData('iso27000-graph.json');
  const [selectedNode, setSelectedNode] = useState<GraphNode | null>(null);
  const [viewMode, setViewMode] = useState<ViewMode>('list');
  const [expandedItem, setExpandedItem] = useState<string | null>(null);
  const [searchFilter, setSearchFilter] = useState('');

  // Handle navigation with selectedId
  useEffect(() => {
    const state = location.state as { selectedId?: string } | null;
    if (state?.selectedId && data) {
      const node = data.nodes.find(n => n.id === state.selectedId);
      if (node) {
        setExpandedItem(node.id);
        setSelectedNode(node);
      }
    }
  }, [location.state, data]);

  // Group terms by first letter (filtered for list view)
  const groupedTerms = useMemo(() => {
    if (!data) return {};
    const filtered = searchFilter
      ? data.nodes.filter(n =>
          n.label.toLowerCase().includes(searchFilter.toLowerCase()) ||
          n.definition?.toLowerCase().includes(searchFilter.toLowerCase())
        )
      : data.nodes;

    const groups: Record<string, GraphNode[]> = {};
    filtered.forEach(node => {
      const letter = node.label[0].toUpperCase();
      if (!groups[letter]) groups[letter] = [];
      groups[letter].push(node);
    });

    // Sort within each group
    Object.keys(groups).forEach(letter => {
      groups[letter].sort((a, b) => a.label.localeCompare(b.label));
    });

    return groups;
  }, [data, searchFilter]);

  const filteredCount = useMemo(() => {
    if (!data || !searchFilter) return data?.nodes.length ?? 0;
    return data.nodes.filter(n =>
      n.label.toLowerCase().includes(searchFilter.toLowerCase()) ||
      n.definition?.toLowerCase().includes(searchFilter.toLowerCase())
    ).length;
  }, [data, searchFilter]);

  const sortedLetters = useMemo(() =>
    Object.keys(groupedTerms).sort(),
    [groupedTerms]
  );

  const getNodeColor = useCallback((node: GraphNode) => {
    return node.isTopConcept ? '#238636' : '#58a6ff';
  }, []);

  const getNodeRadius = useCallback((node: GraphNode) => {
    return node.isTopConcept ? 14 : 8;
  }, []);

  const radialConfig = useMemo(() => ({
    radius: Math.min(window.innerWidth, window.innerHeight - 60) * 0.35,
    strength: 0.15,
  }), []);

  if (loading) {
    return <div className="loading"><p>Loading ISO 27000 vocabulary...</p></div>;
  }

  if (error || !data) {
    return <div className="error"><p>Error: {error ?? 'Failed to load data'}</p></div>;
  }

  return (
    <div className="vocabulary-page">
      <div className="vocabulary-header">
        <div className="header-left">
          <h1>ISO 27000 Vocabulary</h1>
          <span className="term-count">
            {searchFilter ? `${filteredCount} of ${data.nodes.length}` : data.nodes.length} terms
          </span>
        </div>
        <div className="header-right">
          <input
            type="text"
            className="filter-input"
            placeholder="Filter terms..."
            value={searchFilter}
            onChange={e => setSearchFilter(e.target.value)}
          />
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
        <div className="vocabulary-list">
          <nav className="letter-nav">
            {sortedLetters.map(letter => (
              <a key={letter} href={`#letter-${letter}`}>{letter}</a>
            ))}
          </nav>
          <div className="terms-container">
            {sortedLetters.map(letter => (
              <section key={letter} id={`letter-${letter}`} className="letter-section">
                <h2 className="letter-heading">{letter}</h2>
                <div className="terms-list">
                  {groupedTerms[letter].map(term => (
                    <article
                      key={term.id}
                      className={`term-card ${expandedItem === term.id ? 'expanded' : ''} ${term.isTopConcept ? 'top-concept' : ''}`}
                    >
                      <div
                        className="term-header"
                        onClick={() => setExpandedItem(expandedItem === term.id ? null : term.id)}
                      >
                        <span className="toggle-icon">{expandedItem === term.id ? '▼' : '▶'}</span>
                        <h3>{term.label}</h3>
                        {term.isTopConcept && <span className="top-badge">Top Concept</span>}
                      </div>
                      {expandedItem === term.id && (
                        <div className="term-details">
                          {term.definition && (
                            <div className="detail-section">
                              <h4>Definition</h4>
                              <p>{term.definition}</p>
                            </div>
                          )}
                        </div>
                      )}
                    </article>
                  ))}
                </div>
              </section>
            ))}
          </div>
        </div>
      ) : (
        <div className="vocabulary-graph">
          <Graph
            initialNodes={data.nodes}
            initialLinks={data.links}
            width={window.innerWidth}
            height={window.innerHeight - 120}
            getNodeColor={getNodeColor}
            getNodeRadius={getNodeRadius}
            selectedNode={selectedNode}
            onNodeSelect={setSelectedNode}
            chargeStrength={-80}
            linkDistance={60}
            linkStrength={0.4}
            collisionRadius={20}
            radial={radialConfig}
            highlightFilter={searchFilter}
          />
          <NodeDetails node={selectedNode} onClose={() => setSelectedNode(null)} />
        </div>
      )}
    </div>
  );
}
