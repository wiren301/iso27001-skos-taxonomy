import { useState, useRef, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { useSearch } from '../../hooks/useSearch';
import type { SearchResult } from '../../hooks/useSearch';
import './Search.css';

export function Search() {
  const { query, setQuery, results, loading } = useSearch();
  const [isOpen, setIsOpen] = useState(false);
  const [selectedIndex, setSelectedIndex] = useState(0);
  const inputRef = useRef<HTMLInputElement>(null);
  const navigate = useNavigate();

  // Close dropdown when clicking outside
  useEffect(() => {
    const handleClick = (e: MouseEvent) => {
      if (!(e.target as Element).closest('.search-container')) {
        setIsOpen(false);
      }
    };
    document.addEventListener('click', handleClick);
    return () => document.removeEventListener('click', handleClick);
  }, []);

  // Reset selection when results change
  useEffect(() => {
    setSelectedIndex(0);
  }, [results]);

  const handleSelect = (result: SearchResult) => {
    navigate(result.path, { state: { selectedId: result.id } });
    setQuery('');
    setIsOpen(false);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (!isOpen || results.length === 0) return;

    switch (e.key) {
      case 'ArrowDown':
        e.preventDefault();
        setSelectedIndex(i => Math.min(i + 1, results.length - 1));
        break;
      case 'ArrowUp':
        e.preventDefault();
        setSelectedIndex(i => Math.max(i - 1, 0));
        break;
      case 'Enter':
        e.preventDefault();
        handleSelect(results[selectedIndex]);
        break;
      case 'Escape':
        setIsOpen(false);
        inputRef.current?.blur();
        break;
    }
  };

  const groupedResults = results.reduce((acc, result) => {
    if (!acc[result.type]) acc[result.type] = [];
    acc[result.type].push(result);
    return acc;
  }, {} as Record<string, SearchResult[]>);

  const typeLabels = {
    control: 'Controls (ISO 27002)',
    vocabulary: 'Vocabulary (ISO 27000)',
    clause: 'Clauses (ISO 27001)',
  };

  return (
    <div className="search-container">
      <input
        ref={inputRef}
        type="text"
        className="search-input"
        placeholder="Search..."
        value={query}
        onChange={e => {
          setQuery(e.target.value);
          setIsOpen(true);
        }}
        onFocus={() => setIsOpen(true)}
        onKeyDown={handleKeyDown}
      />
      {isOpen && query.trim() && (
        <div className="search-dropdown">
          {loading ? (
            <div className="search-loading">Loading...</div>
          ) : results.length === 0 ? (
            <div className="search-empty">No results found</div>
          ) : (
            Object.entries(groupedResults).map(([type, items]) => (
              <div key={type} className="search-group">
                <div className="search-group-label">
                  {typeLabels[type as keyof typeof typeLabels]}
                </div>
                {items.map((result) => {
                  const globalIndex = results.indexOf(result);
                  return (
                    <button
                      key={result.id}
                      className={`search-result ${globalIndex === selectedIndex ? 'selected' : ''}`}
                      onClick={() => handleSelect(result)}
                      onMouseEnter={() => setSelectedIndex(globalIndex)}
                    >
                      <span className="search-result-label">
                        {result.notation && (
                          <span className="search-result-notation">{result.notation}</span>
                        )}
                        {result.label}
                      </span>
                      <span className="search-result-desc">{result.description}</span>
                    </button>
                  );
                })}
              </div>
            ))
          )}
        </div>
      )}
    </div>
  );
}
