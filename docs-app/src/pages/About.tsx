import { BackgroundGraph } from '../components/BackgroundGraph';
import './About.css';

export function About() {
  return (
    <div className="about">
      <div className="about-background">
        <BackgroundGraph />
      </div>
      <div className="about-content">
        <div className="about-left">
          <div className="about-hero">
            <h1 className="about-title">Understanding the<br />ISO 27000 Series</h1>
            <p className="about-intro">
              The ISO 27000 series is a family of international standards for information
              security management, published by ISO and IEC.
            </p>
          </div>

          <section className="about-section">
            <h2>How They Connect</h2>
            <ul className="connection-list">
              <li>
                <strong>ISO 27001</strong> requirements reference <strong>ISO 27000</strong> vocabulary
                terms for precise language.
              </li>
              <li>
                <strong>ISO 27001 Annex A</strong> maps directly to <strong>ISO 27002</strong> controls.
              </li>
              <li>
                <strong>ISO 27002</strong> provides detailed implementation guidance for each control.
              </li>
            </ul>
          </section>

          <section className="about-section">
            <h2>What Is This Tool?</h2>
            <p>
              This viewer presents these standards as a linked, searchable knowledge base
              using SKOS — a W3C standard for representing taxonomies and thesauri.
            </p>
          </section>

        </div>

        <div className="about-right">
          <div className="about-section disclaimer">
            <h2>Disclaimer</h2>
            <p>
              This taxonomy is an <strong>original academic artifact</strong> developed as scholarly
              research at the University of Twente. It represents the author's interpretation and
              analysis of information security concepts.
            </p>
            <p>
              This work is <strong>based on</strong> the ISO/IEC 27000 series but is <strong>not a
              substitute</strong> for those standards. All definitions are paraphrased interpretations,
              not verbatim reproductions. This tool is <strong>not affiliated with or endorsed by</strong> ISO or IEC.
            </p>
            <p>
              For authoritative definitions, consult the official standards at{' '}
              <a href="https://www.iso.org/" target="_blank" rel="noopener noreferrer">iso.org</a>.
            </p>
          </div>
        </div>
      </div>
    </div>
  );
}
