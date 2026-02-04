import type { GraphNode } from '../../types';
import './NodeDetails.css';

interface NodeDetailsProps {
  node: GraphNode | null;
  onClose: () => void;
}

export function NodeDetails({ node, onClose }: NodeDetailsProps) {
  if (!node) return null;

  return (
    <div className="node-details">
      <div className="node-details-header">
        <h3>{node.label}</h3>
        <button className="close-btn" onClick={onClose}>×</button>
      </div>

      <div className="node-details-content">
        {node.notation && (
          <div className="detail-row">
            <span className="detail-label">Notation</span>
            <span className="detail-value code">{node.notation}</span>
          </div>
        )}

        {node.type && (
          <div className="detail-row">
            <span className="detail-label">Type</span>
            <span className="detail-value">{node.type}</span>
          </div>
        )}

        {node.theme && (
          <div className="detail-row">
            <span className="detail-label">Theme</span>
            <span className="detail-value">{node.theme}</span>
          </div>
        )}

        {node.definition && (
          <div className="detail-row definition">
            <span className="detail-label">Definition</span>
            <p className="detail-value">{node.definition}</p>
          </div>
        )}

        {node.uri && (
          <div className="detail-row">
            <span className="detail-label">URI</span>
            <a href={node.uri} target="_blank" rel="noopener noreferrer" className="detail-value uri">
              {node.uri}
            </a>
          </div>
        )}

        {node.isTopConcept && (
          <div className="badge top-concept">Top Concept</div>
        )}
      </div>
    </div>
  );
}
