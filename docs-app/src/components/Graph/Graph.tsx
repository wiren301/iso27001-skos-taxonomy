import { useRef, useEffect, useCallback } from 'react';
import * as d3 from 'd3';
import type { GraphNode, GraphLink } from '../../types';
import { useForceSimulation } from '../../hooks/useForceSimulation';
import './Graph.css';

interface ClusterConfig {
  [key: string]: { x: number; y: number };
}

interface RadialConfig {
  radius: number | ((node: GraphNode) => number);
  strength?: number;
  x?: number;
  y?: number;
}

interface GraphProps {
  initialNodes: GraphNode[];
  initialLinks: GraphLink[];
  width: number;
  height: number;
  getNodeColor: (node: GraphNode) => string;
  getNodeRadius: (node: GraphNode) => number;
  selectedNode: GraphNode | null;
  onNodeSelect: (node: GraphNode | null) => void;
  // Force simulation config
  chargeStrength?: number | ((node: GraphNode) => number);
  linkDistance?: number;
  linkStrength?: number;
  collisionRadius?: number | ((node: GraphNode) => number);
  clusters?: ClusterConfig;
  getClusterKey?: (node: GraphNode) => string | undefined;
  clusterStrength?: number;
  // Filter which links affect physics (all links still rendered)
  physicsLinkFilter?: (link: GraphLink) => boolean;
  // Radial force
  radial?: RadialConfig;
  // Highlight filter - nodes matching this string are highlighted
  highlightFilter?: string;
}

export function Graph({
  initialNodes,
  initialLinks,
  width,
  height,
  getNodeColor,
  getNodeRadius,
  selectedNode,
  onNodeSelect,
  chargeStrength = -40,
  linkDistance = 30,
  linkStrength = 0.6,
  collisionRadius = 8,
  clusters,
  getClusterKey,
  clusterStrength,
  physicsLinkFilter,
  radial,
  highlightFilter,
}: GraphProps) {
  const svgRef = useRef<SVGSVGElement>(null);
  const gRef = useRef<SVGGElement>(null);

  const { nodes, links, dragStart, drag, dragEnd } = useForceSimulation(
    initialNodes,
    initialLinks,
    {
      width,
      height,
      chargeStrength,
      linkDistance,
      linkStrength,
      collisionRadius,
      clusters,
      getClusterKey,
      clusterStrength,
      physicsLinkFilter,
      radial,
    }
  );

  const zoomRef = useRef<d3.ZoomBehavior<SVGSVGElement, unknown> | null>(null);

  // Setup zoom
  useEffect(() => {
    if (!svgRef.current || !gRef.current) return;

    const svg = d3.select(svgRef.current);
    const g = d3.select(gRef.current);

    const zoom = d3.zoom<SVGSVGElement, unknown>()
      .scaleExtent([0.1, 4])
      .on('zoom', (event) => {
        g.attr('transform', event.transform);
      });

    zoomRef.current = zoom;
    svg.call(zoom);

    // Reset zoom on double click
    svg.on('dblclick.zoom', () => {
      svg.transition().duration(300).call(zoom.transform, d3.zoomIdentity);
    });
  }, []);

  // Zoom to selected node when it changes
  useEffect(() => {
    if (!selectedNode || !svgRef.current || !zoomRef.current) return;

    const zoomToNode = () => {
      const node = nodes.find(n => n.id === selectedNode.id);
      if (!node || node.x === undefined || node.y === undefined) return false;

      const svg = d3.select(svgRef.current!);
      const zoom = zoomRef.current!;

      // Calculate transform to center on the node
      const scale = 1.5;
      const x = width / 2 - node.x * scale;
      const y = height / 2 - node.y * scale;

      svg.transition()
        .duration(500)
        .call(zoom.transform, d3.zoomIdentity.translate(x, y).scale(scale));
      return true;
    };

    // Try immediately, then retry after simulation settles
    if (!zoomToNode()) {
      const timer = setTimeout(zoomToNode, 500);
      return () => clearTimeout(timer);
    }
  }, [selectedNode?.id, nodes, width, height]);

  const handleMouseDown = useCallback((e: React.MouseEvent, node: GraphNode) => {
    e.stopPropagation();
    dragStart(node);

    const handleMouseMove = (moveEvent: MouseEvent) => {
      const svg = svgRef.current;
      if (!svg) return;
      const rect = svg.getBoundingClientRect();
      const transform = d3.zoomTransform(svg);
      const x = (moveEvent.clientX - rect.left - transform.x) / transform.k;
      const y = (moveEvent.clientY - rect.top - transform.y) / transform.k;
      drag(node, x, y);
    };

    const handleMouseUp = () => {
      dragEnd(node);
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };

    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
  }, [dragStart, drag, dragEnd]);

  const getSourceCoords = (link: GraphLink) => {
    const source = link.source;
    if (typeof source === 'string') {
      const node = nodes.find(n => n.id === source);
      return { x: node?.x ?? 0, y: node?.y ?? 0 };
    }
    return { x: (source as GraphNode).x ?? 0, y: (source as GraphNode).y ?? 0 };
  };

  const getTargetCoords = (link: GraphLink) => {
    const target = link.target;
    if (typeof target === 'string') {
      const node = nodes.find(n => n.id === target);
      return { x: node?.x ?? 0, y: node?.y ?? 0 };
    }
    return { x: (target as GraphNode).x ?? 0, y: (target as GraphNode).y ?? 0 };
  };

  const matchesFilter = useCallback((node: GraphNode) => {
    if (!highlightFilter?.trim()) return true;
    const q = highlightFilter.toLowerCase();
    return (
      node.label.toLowerCase().includes(q) ||
      node.definition?.toLowerCase().includes(q) ||
      node.notation?.toLowerCase().includes(q)
    );
  }, [highlightFilter]);

  const hasFilter = Boolean(highlightFilter?.trim());

  return (
    <svg ref={svgRef} width={width} height={height} className="graph-svg">
      <g ref={gRef}>
        {/* Links */}
        <g className="links">
          {links.map((link, i) => {
            const source = getSourceCoords(link);
            const target = getTargetCoords(link);
            return (
              <line
                key={i}
                x1={source.x}
                y1={source.y}
                x2={target.x}
                y2={target.y}
                className={`link ${link.type ?? ''}`}
              />
            );
          })}
        </g>

        {/* Nodes */}
        <g className="nodes">
          {nodes.map((node) => {
            const isMatch = matchesFilter(node);
            const nodeClass = [
              'node',
              selectedNode?.id === node.id ? 'selected' : '',
              hasFilter ? (isMatch ? 'highlighted' : 'dimmed') : '',
            ].filter(Boolean).join(' ');

            return (
            <g
              key={node.id}
              className={nodeClass}
              transform={`translate(${node.x ?? 0}, ${node.y ?? 0})`}
              onMouseDown={(e) => handleMouseDown(e, node)}
              onClick={(e) => {
                e.stopPropagation();
                onNodeSelect(selectedNode?.id === node.id ? null : node);
              }}
            >
              <circle
                r={getNodeRadius(node)}
                fill={getNodeColor(node)}
                className={node.isTopConcept ? 'top-concept' : ''}
              />
              <text
                dy={getNodeRadius(node) + 12}
                textAnchor="middle"
                className="node-label"
              >
                {node.notation && !node.notation.startsWith('http')
                  ? node.notation
                  : node.label.slice(0, 20)}
              </text>
            </g>
            );
          })}
        </g>
      </g>
    </svg>
  );
}
