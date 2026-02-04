export interface GraphNode {
  id: string;
  label: string;
  notation?: string;
  definition?: string;
  type?: string;
  theme?: string;
  mainClause?: string;
  isTopConcept?: boolean;
  uri?: string;
  standard?: string;
  // ISO 27002 specific fields
  purpose?: string;
  guidance?: string;
  cia?: string[];
  nist?: string[];
  operational_capability?: string[];
  security_domain?: string[];
  // D3 simulation properties
  x?: number;
  y?: number;
  vx?: number;
  vy?: number;
  fx?: number | null;
  fy?: number | null;
}

export interface GraphLink {
  source: string | GraphNode;
  target: string | GraphNode;
  type?: string;
}

export interface GraphData {
  nodes: GraphNode[];
  links: GraphLink[];
  metadata?: {
    source?: { file: string; type: string };
    scheme?: { uri: string; label: string; description: string };
    stats?: Record<string, number>;
  };
  colors?: Record<string, { label: string; color: string }>;
}
