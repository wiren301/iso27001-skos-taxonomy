// docs-app/src/components/Layout/Layout.tsx
import { NavLink, Outlet, useLocation } from 'react-router-dom';
import { useTheme } from '../../hooks/useTheme';
import { Search } from '../Search';
import './Layout.css';

export function Layout() {
  const { theme, toggleTheme } = useTheme();
  const location = useLocation();

  return (
    <div className="layout">
      <header className="header">
        <div className="header-brand">
          <NavLink to="/" className="brand-link">
            <h1>ISO 27K Viewer</h1>
          </NavLink>
        </div>
        <Search />
        <nav className="header-nav">
          <NavLink to="/" end>Home</NavLink>
          <NavLink to="/vocabulary">Vocabulary (27000)</NavLink>
          <NavLink to="/clauses">Clauses (27001)</NavLink>
          <NavLink to="/controls">Controls (27002)</NavLink>
          <NavLink to="/relationships">Relationships</NavLink>
          <NavLink to="/about">About</NavLink>
        </nav>
        <div className="header-links">
          <button className="theme-btn" onClick={toggleTheme} title={`Switch to ${theme === 'dark' ? 'light' : 'dark'} mode`}>
            {theme === 'dark' ? 'Light' : 'Dark'}
          </button>
          <a
            href="https://github.com/wiren301/iso27001-skos-taxonomy"
            target="_blank"
            rel="noopener noreferrer"
          >
            GitHub
          </a>
        </div>
      </header>
      <main className="main">
        <div className="page-content" key={location.pathname}>
          <Outlet />
        </div>
      </main>
    </div>
  );
}
