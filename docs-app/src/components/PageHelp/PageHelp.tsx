// docs-app/src/components/PageHelp/PageHelp.tsx
import { useState, useEffect } from 'react';
import './PageHelp.css';

interface PageHelpProps {
  title: string;
  description: string;
  hints: string;
  storageKey: string;
}

export function PageHelp({ title, description, hints, storageKey }: PageHelpProps) {
  const [isExpanded, setIsExpanded] = useState(() => {
    const saved = localStorage.getItem(`pagehelp-${storageKey}`);
    return saved === null ? true : saved === 'expanded';
  });

  useEffect(() => {
    localStorage.setItem(`pagehelp-${storageKey}`, isExpanded ? 'expanded' : 'collapsed');
  }, [isExpanded, storageKey]);

  return (
    <div className={`page-help ${isExpanded ? 'expanded' : 'collapsed'}`}>
      <div className="page-help-header">
        <h2 className="page-help-title">{title}</h2>
        <div className="page-help-actions">
          <button
            className="page-help-toggle"
            onClick={() => setIsExpanded(!isExpanded)}
            title={isExpanded ? 'Collapse' : 'Expand'}
          >
            {isExpanded ? '−' : '?'}
          </button>
        </div>
      </div>
      {isExpanded && (
        <div className="page-help-content">
          <p className="page-help-description">{description}</p>
          <p className="page-help-hints">{hints}</p>
        </div>
      )}
    </div>
  );
}
