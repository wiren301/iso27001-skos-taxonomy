import { useEffect, useRef, useState } from 'react';
import * as d3 from 'd3';
import type { GraphNode, GraphLink } from '../types';

interface ClusterConfig {
  [key: string]: { x: number; y: number };
}

interface RadialConfig {
  radius: number | ((node: GraphNode) => number);
  strength?: number;
  x?: number;
  y?: number;
}

interface SimulationConfig {
  width: number;
  height: number;
  chargeStrength?: number | ((node: GraphNode) => number);
  linkDistance?: number;
  linkStrength?: number;
  collisionRadius?: number | ((node: GraphNode) => number);
  clusters?: ClusterConfig;
  getClusterKey?: (node: GraphNode) => string | undefined;
  clusterStrength?: number; // Default 0.15
  // Filter which links are used for physics (all links still rendered)
  physicsLinkFilter?: (link: GraphLink) => boolean;
  // Radial force configuration
  radial?: RadialConfig;
}

export function useForceSimulation(
  initialNodes: GraphNode[],
  initialLinks: GraphLink[],
  config: SimulationConfig
) {
  const [nodes, setNodes] = useState<GraphNode[]>([]);
  const [links, setLinks] = useState<GraphLink[]>([]);
  const simulationRef = useRef<d3.Simulation<GraphNode, GraphLink> | null>(null);

  useEffect(() => {
    if (initialNodes.length === 0) return;

    // Deep copy to avoid mutating original data
    const nodesCopy = initialNodes.map(n => ({ ...n }));
    const linksCopy = initialLinks.map(l => ({ ...l }));

    // Filter links for physics if filter provided
    const physicsLinks = config.physicsLinkFilter
      ? linksCopy.filter(config.physicsLinkFilter)
      : linksCopy;

    const { width, height, clusters, getClusterKey } = config;

    const linkForce = d3.forceLink<GraphNode, GraphLink>(physicsLinks)
      .id(d => d.id)
      .distance(config.linkDistance ?? 30)
      .strength(config.linkStrength ?? 0.6);

    // Handle charge strength - can be number or function
    const chargeForce = d3.forceManyBody<GraphNode>();
    if (typeof config.chargeStrength === 'function') {
      chargeForce.strength(config.chargeStrength);
    } else {
      chargeForce.strength(config.chargeStrength ?? -40);
    }

    // Handle collision radius - can be number or function
    const collisionForce = d3.forceCollide<GraphNode>();
    if (typeof config.collisionRadius === 'function') {
      collisionForce.radius(config.collisionRadius);
    } else {
      collisionForce.radius(config.collisionRadius ?? 8);
    }

    const simulation = d3.forceSimulation<GraphNode>(nodesCopy)
      .force('link', linkForce)
      .force('charge', chargeForce)
      .force('center', d3.forceCenter(width / 2, height / 2))
      .force('collision', collisionForce);

    // Add radial force if configured
    if (config.radial) {
      const { radius, strength = 0.3, x = width / 2, y = height / 2 } = config.radial;
      const radialForce = d3.forceRadial<GraphNode>(
        typeof radius === 'function' ? radius : () => radius,
        x,
        y
      ).strength(strength);
      simulation.force('radial', radialForce);
    }

    // Add cluster force if configured - applied on each tick
    if (clusters && getClusterKey) {
      const strength = config.clusterStrength ?? 0.15;
      simulation.force('cluster', () => {
        const alpha = simulation.alpha();
        nodesCopy.forEach(d => {
          const key = getClusterKey(d);
          const c = key ? clusters[key] : undefined;
          if (c) {
            d.vx = (d.vx ?? 0) + (c.x * width - (d.x ?? 0)) * alpha * strength;
            d.vy = (d.vy ?? 0) + (c.y * height - (d.y ?? 0)) * alpha * strength;
          }
        });
      });
    }

    simulation.on('tick', () => {
      setNodes([...nodesCopy]);
      setLinks([...linksCopy]);
    });

    simulationRef.current = simulation;

    return () => {
      simulation.stop();
    };
  }, [initialNodes, initialLinks, config.width, config.height, config.chargeStrength, config.linkDistance, config.linkStrength, config.collisionRadius, config.clusters, config.getClusterKey, config.clusterStrength, config.physicsLinkFilter, config.radial]);

  const reheat = () => {
    simulationRef.current?.alpha(0.3).restart();
  };

  const dragStart = (node: GraphNode) => {
    if (!simulationRef.current) return;
    simulationRef.current.alphaTarget(0.3).restart();
    node.fx = node.x;
    node.fy = node.y;
  };

  const drag = (node: GraphNode, x: number, y: number) => {
    node.fx = x;
    node.fy = y;
  };

  const dragEnd = (node: GraphNode) => {
    if (!simulationRef.current) return;
    simulationRef.current.alphaTarget(0);
    node.fx = null;
    node.fy = null;
  };

  return { nodes, links, reheat, dragStart, drag, dragEnd };
}
