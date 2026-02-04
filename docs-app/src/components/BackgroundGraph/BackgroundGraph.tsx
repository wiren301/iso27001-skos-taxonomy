import { useRef, useEffect, useCallback } from 'react';

interface Node {
  x: number;
  y: number;
  vx: number;
  vy: number;
}

interface BackgroundGraphProps {
  nodeCount?: number;
  maxDistance?: number;
  mouseRepelRadius?: number;
  mouseRepelStrength?: number;
}

export function BackgroundGraph({
  nodeCount = 45,
  maxDistance = 160,
  mouseRepelRadius = 150,
  mouseRepelStrength = 0.3,
}: BackgroundGraphProps) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const nodesRef = useRef<Node[]>([]);
  const mouseRef = useRef<{ x: number; y: number } | null>(null);
  const animationRef = useRef<number>(0);

  const initNodes = useCallback((width: number, height: number) => {
    const nodes: Node[] = [];
    for (let i = 0; i < nodeCount; i++) {
      nodes.push({
        x: Math.random() * width,
        y: Math.random() * height,
        vx: (Math.random() - 0.5) * 0.3,
        vy: (Math.random() - 0.5) * 0.3,
      });
    }
    return nodes;
  }, [nodeCount]);

  useEffect(() => {
    const canvas = canvasRef.current;
    const container = containerRef.current;
    if (!canvas || !container) return;

    const ctx = canvas.getContext('2d');
    if (!ctx) return;

    const resize = () => {
      const rect = container.getBoundingClientRect();
      if (rect.width === 0 || rect.height === 0) return;

      canvas.width = rect.width;
      canvas.height = rect.height;

      if (nodesRef.current.length === 0) {
        nodesRef.current = initNodes(rect.width, rect.height);
      }
    };

    resize();

    // Use ResizeObserver for more reliable size detection
    const resizeObserver = new ResizeObserver(resize);
    resizeObserver.observe(container);

    const colors = [
      '#58a6ff',  // blue
      '#8b5cf6',  // purple
      '#2dd4bf',  // teal
      '#3b82f6',  // darker blue
    ];

    // Assign colors to nodes once
    const nodeColors = nodesRef.current.map(() =>
      colors[Math.floor(Math.random() * colors.length)]
    );

    const draw = () => {
      if (!ctx || !canvas || canvas.width === 0) {
        animationRef.current = requestAnimationFrame(draw);
        return;
      }

      ctx.clearRect(0, 0, canvas.width, canvas.height);

      // Update node positions
      nodesRef.current.forEach(node => {
        // Apply mouse repulsion
        if (mouseRef.current) {
          const dx = node.x - mouseRef.current.x;
          const dy = node.y - mouseRef.current.y;
          const dist = Math.sqrt(dx * dx + dy * dy);
          if (dist < mouseRepelRadius && dist > 0) {
            const force = (1 - dist / mouseRepelRadius) * mouseRepelStrength;
            node.vx += (dx / dist) * force;
            node.vy += (dy / dist) * force;
          }
        }

        // Add gentle random drift
        node.vx += (Math.random() - 0.5) * 0.02;
        node.vy += (Math.random() - 0.5) * 0.02;

        // Apply velocity with damping
        node.vx *= 0.98;
        node.vy *= 0.98;
        node.x += node.vx;
        node.y += node.vy;

        // Keep nodes in bounds
        if (node.x < 0) { node.x = 0; node.vx *= -0.5; }
        if (node.x > canvas.width) { node.x = canvas.width; node.vx *= -0.5; }
        if (node.y < 0) { node.y = 0; node.vy *= -0.5; }
        if (node.y > canvas.height) { node.y = canvas.height; node.vy *= -0.5; }
      });

      // Calculate center fade - less opacity near content area
      const centerX = canvas.width / 2;
      const centerY = canvas.height / 2;
      const fadeRadius = Math.min(canvas.width, canvas.height) * 0.5;

      const getPositionFade = (x: number, y: number) => {
        const dx = x - centerX;
        const dy = y - centerY;
        const dist = Math.sqrt(dx * dx + dy * dy);
        // Fade from 0.2 at center to 1.0 at edges
        return Math.min(1, 0.2 + (dist / fadeRadius) * 0.8);
      };

      // Draw edges
      ctx.lineWidth = 1;
      for (let i = 0; i < nodesRef.current.length; i++) {
        for (let j = i + 1; j < nodesRef.current.length; j++) {
          const a = nodesRef.current[i];
          const b = nodesRef.current[j];
          const dx = a.x - b.x;
          const dy = a.y - b.y;
          const dist = Math.sqrt(dx * dx + dy * dy);
          if (dist < maxDistance) {
            const midX = (a.x + b.x) / 2;
            const midY = (a.y + b.y) / 2;
            const fade = getPositionFade(midX, midY);
            const opacity = (1 - dist / maxDistance) * 0.5 * fade;
            ctx.globalAlpha = opacity;
            ctx.strokeStyle = nodeColors[i];
            ctx.beginPath();
            ctx.moveTo(a.x, a.y);
            ctx.lineTo(b.x, b.y);
            ctx.stroke();
          }
        }
      }

      // Draw nodes
      nodesRef.current.forEach((node, i) => {
        const fade = getPositionFade(node.x, node.y);
        ctx.globalAlpha = 0.7 * fade;
        ctx.fillStyle = nodeColors[i];
        ctx.beginPath();
        ctx.arc(node.x, node.y, 4, 0, Math.PI * 2);
        ctx.fill();
      });

      ctx.globalAlpha = 1;
      animationRef.current = requestAnimationFrame(draw);
    };

    draw();

    return () => {
      resizeObserver.disconnect();
      cancelAnimationFrame(animationRef.current);
    };
  }, [initNodes, maxDistance, mouseRepelRadius, mouseRepelStrength]);

  const handleMouseMove = useCallback((e: React.MouseEvent) => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const rect = canvas.getBoundingClientRect();
    mouseRef.current = {
      x: e.clientX - rect.left,
      y: e.clientY - rect.top,
    };
  }, []);

  const handleMouseLeave = useCallback(() => {
    mouseRef.current = null;
  }, []);

  return (
    <div
      ref={containerRef}
      onMouseMove={handleMouseMove}
      onMouseLeave={handleMouseLeave}
      style={{
        position: 'absolute',
        top: 0,
        left: 0,
        width: '100%',
        height: '100%',
      }}
    >
      <canvas
        ref={canvasRef}
        style={{
          display: 'block',
          width: '100%',
          height: '100%',
        }}
      />
    </div>
  );
}
