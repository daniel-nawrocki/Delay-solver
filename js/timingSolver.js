function generateValues(min, max, maxSamples = 15) {
  const absMin = Math.abs(Math.floor(Number(min) || 0));
  const absMax = Math.abs(Math.floor(Number(max) || 0));
  const a = Math.max(0, Math.min(absMin, absMax));
  const b = Math.max(a, Math.max(absMin, absMax));
  if (a === b) return [a];
  const span = b - a;
  const step = Math.max(1, Math.ceil(span / (maxSamples - 1)));
  const values = [];
  for (let value = a; value <= b; value += step) values.push(value);
  if (values[values.length - 1] !== b) values.push(b);
  return values;
}

function maxHolesInWindow(times, windowMs = 8) {
  if (!times.length) return 0;
  const sorted = [...times].sort((a, b) => a - b);
  let start = 0;
  let maxCount = 1;
  for (let end = 0; end < sorted.length; end += 1) {
    while (sorted[end] - sorted[start] > windowMs) start += 1;
    maxCount = Math.max(maxCount, end - start + 1);
  }
  return maxCount;
}

function edgeDelay(edge, holeDelay, rowDelay, offsetAssignments) {
  if (edge.type === "offset") return offsetAssignments.get(edge.id) ?? 17;
  if (edge.type === "rowToRow") return (edge.sign === -1 ? -1 : 1) * rowDelay;
  return (edge.sign === -1 ? -1 : 1) * holeDelay;
}

function offsetRangeValues(edge) {
  const fallback = Number.isFinite(Number(edge.offsetMs)) ? Number(edge.offsetMs) : null;
  const min = Number.isFinite(Number(edge.minOffsetMs)) ? Number(edge.minOffsetMs) : fallback ?? 17;
  const max = Number.isFinite(Number(edge.maxOffsetMs)) ? Number(edge.maxOffsetMs) : fallback ?? 42;
  return generateValues(Math.min(min, max), Math.max(min, max), 26);
}

function buildGraphState(state) {
  const edges = state.relationships?.edges || [];
  const originHoleId = state.relationships?.originHoleId || null;
  if (!originHoleId) return { valid: false, reason: "Select an origin hole before solving." };
  if (!state.holesById.has(originHoleId)) return { valid: false, reason: "The selected origin hole no longer exists." };
  if (!edges.length) return { valid: false, reason: "Create at least one timing relationship before solving." };

  const adjacency = new Map();
  edges.forEach((edge) => {
    if (!state.holesById.has(edge.fromHoleId) || !state.holesById.has(edge.toHoleId)) return;
    if (!adjacency.has(edge.fromHoleId)) adjacency.set(edge.fromHoleId, []);
    adjacency.get(edge.fromHoleId).push(edge);
  });

  return { valid: true, originHoleId, edges, adjacency };
}

export function validateTimingGraph(state) {
  const graph = buildGraphState(state);
  if (!graph.valid) return graph;

  const visited = new Set();
  const queue = [graph.originHoleId];
  while (queue.length) {
    const holeId = queue.shift();
    if (visited.has(holeId)) continue;
    visited.add(holeId);
    for (const edge of graph.adjacency.get(holeId) || []) {
      queue.push(edge.toHoleId);
    }
  }

  const unreachable = state.holes.filter((hole) => !visited.has(hole.id));
  if (unreachable.length) {
    const sample = unreachable.slice(0, 3).map((hole) => hole.holeNumber || hole.id).join(", ");
    return {
      valid: false,
      reason: `Every hole must be reachable from the origin through directed relationships. Unreachable: ${sample}${unreachable.length > 3 ? "..." : ""}`,
    };
  }

  return graph;
}

function buildSchedule(state, graph, holeDelay, rowDelay, offsetAssignments = new Map()) {
  const holeTimes = new Map([[graph.originHoleId, 0]]);
  const queue = [graph.originHoleId];
  const epsilon = 0.0001;

  while (queue.length) {
    const holeId = queue.shift();
    const baseTime = holeTimes.get(holeId);
    for (const edge of graph.adjacency.get(holeId) || []) {
      const nextTime = baseTime + edgeDelay(edge, holeDelay, rowDelay, offsetAssignments);
      const existing = holeTimes.get(edge.toHoleId);
      if (existing === undefined) {
        holeTimes.set(edge.toHoleId, nextTime);
        queue.push(edge.toHoleId);
        continue;
      }
      if (Math.abs(existing - nextTime) > epsilon) {
        return {
          valid: false,
          reason: `Conflicting timing paths detected at hole ${edge.toHoleId}.`,
        };
      }
    }
  }

  if (holeTimes.size !== state.holes.length) {
    return { valid: false, reason: "Some holes could not be assigned a derived firing time." };
  }

  const times = [...holeTimes.values()];
  const minTime = times.length ? Math.min(...times) : 0;
  const maxTime = times.length ? Math.max(...times) : 0;
  const endTime = maxTime - minTime;
  const density8ms = maxHolesInWindow(times, 8);

  return {
    valid: true,
    holeDelay,
    rowDelay,
    offsetAssignments: new Map(offsetAssignments),
    holeTimes,
    times,
    endTime,
    density8ms,
  };
}

export function solveTimingCombinations(state) {
  const graph = validateTimingGraph(state);
  if (!graph.valid) return [];

  const holeValues = generateValues(state.timing.holeToHole.min, state.timing.holeToHole.max);
  const rowValues = generateValues(state.timing.rowToRow.min, state.timing.rowToRow.max);
  const offsetEdges = graph.edges.filter((edge) => edge.type === "offset");
  const candidates = [];

  function exploreOffsets(index, offsetAssignments, holeDelay, rowDelay) {
    if (index >= offsetEdges.length) {
      const schedule = buildSchedule(state, graph, holeDelay, rowDelay, offsetAssignments);
      if (schedule.valid) candidates.push(schedule);
      return;
    }

    const edge = offsetEdges[index];
    const values = offsetRangeValues(edge);
    values.forEach((value) => {
      const nextAssignments = new Map(offsetAssignments);
      nextAssignments.set(edge.id, value);
      exploreOffsets(index + 1, nextAssignments, holeDelay, rowDelay);
    });
  }

  holeValues.forEach((holeDelay) => {
    rowValues.forEach((rowDelay) => {
      exploreOffsets(0, new Map(), holeDelay, rowDelay);
    });
  });

  candidates.sort((a, b) => {
    if (a.density8ms !== b.density8ms) return a.density8ms - b.density8ms;
    if (a.endTime !== b.endTime) return a.endTime - b.endTime;
    return (a.holeDelay + a.rowDelay) - (b.holeDelay + b.rowDelay);
  });

  return candidates.slice(0, 12);
}

export function formatTimingResult(result, index) {
  const offsetSummary = result.offsetAssignments?.size
    ? ` | offsets: ${[...result.offsetAssignments.values()].join(",")}ms`
    : "";
  return `${index + 1}. H2H ${result.holeDelay}ms | R2R ${result.rowDelay}ms${offsetSummary} | peak in 8ms: ${result.density8ms} holes | total duration: ${result.endTime.toFixed(1)}ms`;
}
