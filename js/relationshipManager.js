const RELATIONSHIP_COLORS = {
  holeToHole: "#0f766e",
  rowToRow: "#1d4ed8",
  offset: "#b45309",
};

const RELATIONSHIP_LABELS = {
  holeToHole: "Hole-to-hole",
  rowToRow: "Row-to-row",
  offset: "Offset",
};

export function relationshipColor(type) {
  return RELATIONSHIP_COLORS[type] || "#475569";
}

export function relationshipLabel(type) {
  return RELATIONSHIP_LABELS[type] || "Relationship";
}

export function ensureRelationshipState(state) {
  if (!state.relationships || typeof state.relationships !== "object") {
    state.relationships = { originHoleId: null, edges: [], nextId: 1 };
  }
  if (!Array.isArray(state.relationships.edges)) state.relationships.edges = [];
  state.relationships.edges = state.relationships.edges.map((edge) => {
    if (edge?.type !== "offset") return edge;
    const fallback = Number.isFinite(Number(edge.offsetMs)) ? Number(edge.offsetMs) : null;
    const minOffsetMs = Number.isFinite(Number(edge.minOffsetMs)) ? Number(edge.minOffsetMs) : fallback ?? 17;
    const maxOffsetMs = Number.isFinite(Number(edge.maxOffsetMs)) ? Number(edge.maxOffsetMs) : fallback ?? 42;
    return {
      ...edge,
      minOffsetMs: Math.min(minOffsetMs, maxOffsetMs),
      maxOffsetMs: Math.max(minOffsetMs, maxOffsetMs),
    };
  });
  if (!Number.isFinite(Number(state.relationships.nextId)) || Number(state.relationships.nextId) < 1) {
    state.relationships.nextId = 1;
  }
  return state.relationships;
}

export function clearRelationships(state) {
  const relationships = ensureRelationshipState(state);
  relationships.edges = [];
  relationships.originHoleId = null;
  relationships.nextId = 1;
}

export function setOriginHole(state, holeId) {
  const relationships = ensureRelationshipState(state);
  relationships.originHoleId = holeId || null;
}

export function addRelationship(state, input) {
  const relationships = ensureRelationshipState(state);
  const minOffsetMs = Number.isFinite(Number(input.minOffsetMs))
    ? Number(input.minOffsetMs)
    : Number.isFinite(Number(input.offsetMs))
      ? Number(input.offsetMs)
      : 17;
  const maxOffsetMs = Number.isFinite(Number(input.maxOffsetMs))
    ? Number(input.maxOffsetMs)
    : Number.isFinite(Number(input.offsetMs))
      ? Number(input.offsetMs)
      : 42;
  const edge = {
    id: `rel-${relationships.nextId++}`,
    type: input.type,
    fromHoleId: input.fromHoleId,
    toHoleId: input.toHoleId,
    sign: input.sign ?? 1,
    minOffsetMs: Math.min(minOffsetMs, maxOffsetMs),
    maxOffsetMs: Math.max(minOffsetMs, maxOffsetMs),
  };
  relationships.edges.push(edge);
  return edge;
}

export function updateRelationship(state, relationshipId, patch) {
  const relationships = ensureRelationshipState(state);
  const edge = relationships.edges.find((item) => item.id === relationshipId);
  if (!edge) return null;
  Object.assign(edge, patch);
  edge.sign = edge.sign === -1 ? -1 : 1;
  const minOffsetMs = Number.isFinite(Number(edge.minOffsetMs))
    ? Number(edge.minOffsetMs)
    : Number.isFinite(Number(edge.offsetMs))
      ? Number(edge.offsetMs)
      : 17;
  const maxOffsetMs = Number.isFinite(Number(edge.maxOffsetMs))
    ? Number(edge.maxOffsetMs)
    : Number.isFinite(Number(edge.offsetMs))
      ? Number(edge.offsetMs)
      : 42;
  edge.minOffsetMs = Math.min(minOffsetMs, maxOffsetMs);
  edge.maxOffsetMs = Math.max(minOffsetMs, maxOffsetMs);
  delete edge.offsetMs;
  return edge;
}

export function deleteRelationship(state, relationshipId) {
  const relationships = ensureRelationshipState(state);
  relationships.edges = relationships.edges.filter((item) => item.id !== relationshipId);
}

export function describeRelationship(edge, holesById) {
  const fromLabel = holesById.get(edge.fromHoleId)?.holeNumber || holesById.get(edge.fromHoleId)?.id || edge.fromHoleId;
  const toLabel = holesById.get(edge.toHoleId)?.holeNumber || holesById.get(edge.toHoleId)?.id || edge.toHoleId;
  if (edge.type === "offset") {
    const minOffsetMs = Number.isFinite(Number(edge.minOffsetMs))
      ? Number(edge.minOffsetMs)
      : Number.isFinite(Number(edge.offsetMs))
        ? Number(edge.offsetMs)
        : 17;
    const maxOffsetMs = Number.isFinite(Number(edge.maxOffsetMs))
      ? Number(edge.maxOffsetMs)
      : Number.isFinite(Number(edge.offsetMs))
        ? Number(edge.offsetMs)
        : 42;
    return `${relationshipLabel(edge.type)}: ${fromLabel} -> ${toLabel} (${Math.min(minOffsetMs, maxOffsetMs)}-${Math.max(minOffsetMs, maxOffsetMs)} ms)`;
  }
  const signText = edge.sign === -1 ? "-" : "+";
  return `${relationshipLabel(edge.type)}: ${fromLabel} -> ${toLabel} (${signText})`;
}

export function relationToolLabel(mode) {
  switch (mode) {
    case "origin":
      return "Select Origin";
    case "holeRelationshipPositive":
      return "Hole-to-Hole +";
    case "holeRelationshipNegative":
      return "Hole-to-Hole -";
    case "rowRelationshipPositive":
      return "Row-to-Row +";
    case "rowRelationshipNegative":
      return "Row-to-Row -";
    case "offsetRelationship":
      return "Offset";
    default:
      return "Select";
  }
}
