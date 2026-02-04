import { Link } from 'react-router-dom';
import { BackgroundGraph } from '../components/BackgroundGraph';
import './Home.css';

const navLinks = [
  { to: '/vocabulary', title: 'Vocabulary', description: '77 terms from ISO 27000' },
  { to: '/clauses', title: 'Clauses', description: 'Requirements from ISO 27001' },
  { to: '/controls', title: 'Controls', description: '93 security controls from ISO 27002' },
  { to: '/relationships', title: 'Relationships', description: 'All standards connected' },
  { to: '/about', title: 'About', description: 'Learn about the standards' },
];

export function Home() {
  return (
    <div className="home">
      <div className="home-background">
        <BackgroundGraph />
      </div>
      <div className="home-content">
        <div className="home-hero">
          <h1 className="home-title">ISO 27000<br />Taxonomy Viewer</h1>
          <p className="home-description">
            This tool presents the ISO 27000 family of information security standards
            as a linked, searchable knowledge base. Browse vocabulary terms, requirements,
            and security controls with visual relationship mapping powered by SKOS.
          </p>
        </div>
        <nav className="home-nav">
          <h2 className="home-nav-title">Explore</h2>
          <ul className="home-nav-list">
            {navLinks.map(link => (
              <li key={link.to}>
                <Link to={link.to} className="home-nav-link">
                  <span className="home-nav-link-title">{link.title}</span>
                  <span className="home-nav-link-desc">{link.description}</span>
                </Link>
              </li>
            ))}
          </ul>
        </nav>
      </div>
      <footer className="home-footer">
        <p>
          Want to see other ISO documents added?{' '}
          <a
            href="https://github.com/wiren301/iso27001-skos-taxonomy"
            target="_blank"
            rel="noopener noreferrer"
          >
            Open a pull request or contact me on GitHub
          </a>.
        </p>
      </footer>
    </div>
  );
}
