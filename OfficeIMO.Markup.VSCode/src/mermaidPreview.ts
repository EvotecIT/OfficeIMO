import mermaid from 'mermaid';

type SimpleEdge = {
  from: string;
  fromLabel: string;
  to: string;
  toLabel: string;
};

const blocks = Array.from(document.querySelectorAll<HTMLElement>('.mermaid-diagram'));
if (blocks.length > 0) {
  const background = getComputedStyle(document.body).backgroundColor || '';
  const rgb = /rgb\((\d+),\s*(\d+),\s*(\d+)\)/i.exec(background);
  const brightness = rgb
    ? (Number(rgb[1]) * 0.299 + Number(rgb[2]) * 0.587 + Number(rgb[3]) * 0.114)
    : 255;

  mermaid.initialize({
    startOnLoad: false,
    securityLevel: 'loose',
    theme: brightness < 140 ? 'dark' : 'default',
    flowchart: {
      htmlLabels: false,
      useMaxWidth: true
    }
  });

  const nodes = blocks
    .map((block) => block.querySelector<HTMLElement>('.mermaid'))
    .filter((node): node is HTMLElement => Boolean(node));

  mermaid.run({ nodes }).then(() => {
    blocks.forEach((block) => completeBlock(block));
  }).catch(() => {
    blocks.forEach((block) => {
      if (renderSimpleFlowchartFallback(block)) {
        completeBlock(block, 'simple-fallback');
      } else {
        block.classList.remove('pending');
        block.classList.add('render-failed');
      }
    });
  });
}

function completeBlock(block: HTMLElement, extraClass?: string): void {
  const mermaidElement = block.querySelector<HTMLElement>('.mermaid');
  if (!mermaidElement?.querySelector('svg') && !renderSimpleFlowchartFallback(block)) {
    block.classList.remove('pending');
    block.classList.add('render-failed');
    return;
  }

  block.classList.remove('pending');
  block.classList.remove('render-failed');
  block.classList.add('rendered');
  if (extraClass) {
    block.classList.add(extraClass);
  }
}

function renderSimpleFlowchartFallback(block: HTMLElement): boolean {
  const mermaidElement = block.querySelector<HTMLElement>('.mermaid');
  const source = mermaidElement?.textContent ?? '';
  const edges = parseSimpleEdges(source);
  if (!mermaidElement || edges.length === 0) {
    return false;
  }

  const orderedIds: string[] = [];
  const labels = new Map<string, string>();
  for (const edge of edges) {
    addNode(orderedIds, labels, edge.from, edge.fromLabel);
    addNode(orderedIds, labels, edge.to, edge.toLabel);
  }

  const nodeWidth = 150;
  const nodeHeight = 46;
  const gap = 54;
  const padding = 28;
  const width = Math.max(360, padding * 2 + orderedIds.length * nodeWidth + Math.max(0, orderedIds.length - 1) * gap);
  const height = 150;
  const y = 54;
  const nodeCenters = new Map<string, number>();
  const nodes = orderedIds.map((id, index) => {
    const x = padding + index * (nodeWidth + gap);
    nodeCenters.set(id, x + nodeWidth / 2);
    return `<g>
      <rect x="${x}" y="${y}" width="${nodeWidth}" height="${nodeHeight}" rx="8" fill="#eff6ff" stroke="#2563eb" stroke-width="1.5" />
      <text x="${x + nodeWidth / 2}" y="${y + 29}" text-anchor="middle" font-family="Aptos, Arial, sans-serif" font-size="14" font-weight="700" fill="#172033">${escapeSvg(labels.get(id) ?? id)}</text>
    </g>`;
  }).join('');
  const arrows = edges.map((edge) => {
    const from = nodeCenters.get(edge.from);
    const to = nodeCenters.get(edge.to);
    if (from === undefined || to === undefined) {
      return '';
    }

    const x1 = from + nodeWidth / 2;
    const x2 = to - nodeWidth / 2;
    const lineY = y + nodeHeight / 2;
    return `<line x1="${x1}" y1="${lineY}" x2="${x2}" y2="${lineY}" stroke="#64748b" stroke-width="2" marker-end="url(#arrow)" />`;
  }).join('');

  mermaidElement.innerHTML = `<svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Mermaid flowchart preview">
    <defs>
      <marker id="arrow" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto" markerUnits="strokeWidth">
        <path d="M0,0 L0,6 L9,3 z" fill="#64748b" />
      </marker>
    </defs>
    <rect x="0" y="0" width="${width}" height="${height}" fill="#ffffff" />
    ${arrows}
    ${nodes}
  </svg>`;
  return true;
}

function parseSimpleEdges(source: string): SimpleEdge[] {
  return source
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0 && !/^(flowchart|graph)\s+/i.test(line))
    .map((line) => {
      const match = /^([A-Za-z0-9_]+)(?:\[([^\]]+)\])?\s*-->\s*([A-Za-z0-9_]+)(?:\[([^\]]+)\])?/.exec(line);
      if (!match) {
        return undefined;
      }

      return {
        from: match[1],
        fromLabel: match[2] ?? match[1],
        to: match[3],
        toLabel: match[4] ?? match[3]
      };
    })
    .filter((edge): edge is SimpleEdge => Boolean(edge));
}

function addNode(ids: string[], labels: Map<string, string>, id: string, label: string): void {
  if (!labels.has(id)) {
    ids.push(id);
    labels.set(id, label);
  }
}

function escapeSvg(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
