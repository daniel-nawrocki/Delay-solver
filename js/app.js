import { parseCsvText, buildHolesFromMapping } from "./csvParser.js";
import { DiagramRenderer } from "./diagramRenderer.js";
import { initTimingControls } from "./timingControls.js";
import { solveTimingCombinations, formatTimingResult, validateTimingGraph } from "./timingSolver.js";
import {
  addRelationship,
  clearRelationships,
  deleteRelationship,
  describeRelationship,
  ensureRelationshipState,
  relationToolLabel,
  setOriginHole,
  updateRelationship,
} from "./relationshipManager.js";

const TOOL_TO_RELATIONSHIP_TYPE = {
  holeRelationshipPositive: "holeToHole",
  holeRelationshipNegative: "holeToHole",
  rowRelationshipPositive: "rowToRow",
  rowRelationshipNegative: "rowToRow",
  offsetRelationship: "offset",
};

const TOOL_TO_SIGN = {
  holeRelationshipPositive: 1,
  holeRelationshipNegative: -1,
  rowRelationshipPositive: 1,
  rowRelationshipNegative: -1,
};

const state = {
  holes: [],
  holesById: new Map(),
  selection: new Set(),
  ui: {
    showGrid: true,
    showRelationships: true,
    showOverlayText: true,
    toolMode: "origin",
    coordView: "collar",
    activeTimingPreviewIndex: -1,
    relationshipDraft: null,
  },
  timing: {
    holeToHole: { min: 16, max: 34 },
    rowToRow: { min: 84, max: 142 },
    offset: { min: 17, max: 42 },
  },
  relationships: { originHoleId: null, edges: [], nextId: 1 },
  csvCache: null,
  timingResults: [],
  solverMessage: "",
};

const printState = {
  holes: [],
  holesById: new Map(),
  selection: new Set(),
  ui: {
    showGrid: false,
    showRelationships: true,
    showOverlayText: true,
    activeTimingPreviewIndex: 0,
    relationshipDraft: null,
    textScale: 1,
    orientation: "landscape",
  },
  relationships: { originHoleId: null, edges: [], nextId: 1 },
  timingResults: [],
};

const els = {
  csvInput: document.getElementById("csvInput"),
  mappingPanel: document.getElementById("mappingPanel"),
  coordTypeSelect: document.getElementById("coordTypeSelect"),
  xColumnSelect: document.getElementById("xColumnSelect"),
  yColumnSelect: document.getElementById("yColumnSelect"),
  toeXColumnSelect: document.getElementById("toeXColumnSelect"),
  toeYColumnSelect: document.getElementById("toeYColumnSelect"),
  idColumnSelect: document.getElementById("idColumnSelect"),
  importMappedBtn: document.getElementById("importMappedBtn"),
  gridToggle: document.getElementById("gridToggle"),
  relationshipVisibilityToggle: document.getElementById("relationshipVisibilityToggle"),
  relationshipVisibilityToggleSecondary: document.getElementById("relationshipVisibilityToggleSecondary"),
  fitViewBtn: document.getElementById("fitViewBtn"),
  coordViewSelect: document.getElementById("coordViewSelect"),
  rotateLeftBtn: document.getElementById("rotateLeftBtn"),
  rotateRightBtn: document.getElementById("rotateRightBtn"),
  rotateFineLeftBtn: document.getElementById("rotateFineLeftBtn"),
  rotateFineRightBtn: document.getElementById("rotateFineRightBtn"),
  rotateResetBtn: document.getElementById("rotateResetBtn"),
  originStatus: document.getElementById("originStatus"),
  toolModeStatus: document.getElementById("toolModeStatus"),
  clearRelationshipsBtn: document.getElementById("clearRelationshipsBtn"),
  clearOriginBtn: document.getElementById("clearOriginBtn"),
  relationshipList: document.getElementById("relationshipList"),
  holeDelayMin: document.getElementById("holeDelayMinInput"),
  holeDelayMax: document.getElementById("holeDelayMaxInput"),
  rowDelayMin: document.getElementById("rowDelayMinInput"),
  rowDelayMax: document.getElementById("rowDelayMaxInput"),
  offsetDelayMin: document.getElementById("offsetDelayMinInput"),
  offsetDelayMax: document.getElementById("offsetDelayMaxInput"),
  solveTimingBtn: document.getElementById("solveTimingBtn"),
  timingResults: document.getElementById("timingResults"),
  timingResultsMenuWrap: document.getElementById("timingResultsMenuWrap"),
  helpBtn: document.getElementById("helpBtn"),
  csvExportBtn: document.getElementById("csvExportBtn"),
  exportPdfBtn: document.getElementById("exportPdfBtn"),
  originToolBtn: document.getElementById("originToolBtn"),
  holeRelationPositiveToolBtn: document.getElementById("holeRelationPositiveToolBtn"),
  holeRelationNegativeToolBtn: document.getElementById("holeRelationNegativeToolBtn"),
  rowRelationPositiveToolBtn: document.getElementById("rowRelationPositiveToolBtn"),
  rowRelationNegativeToolBtn: document.getElementById("rowRelationNegativeToolBtn"),
  offsetRelationToolBtn: document.getElementById("offsetRelationToolBtn"),
  menuToggles: [...document.querySelectorAll("[data-menu-toggle]")],
  menuPanels: [...document.querySelectorAll(".menu-panel")],
  printWorkspace: document.getElementById("printWorkspace"),
  printCanvas: document.getElementById("printCanvas"),
  printPaperFrame: document.getElementById("printPaperFrame"),
  printBackBtn: document.getElementById("printBackBtn"),
  printActionBtn: document.getElementById("printActionBtn"),
  printFitBtn: document.getElementById("printFitBtn"),
  printTextScaleInput: document.getElementById("printTextScaleInput"),
  printOrientationSelect: document.getElementById("printOrientationSelect"),
  printColorModeToggle: document.getElementById("printColorModeToggle"),
  printRelationshipToggle: document.getElementById("printRelationshipToggle"),
  helpWorkspace: document.getElementById("helpWorkspace"),
  helpBackBtn: document.getElementById("helpBackBtn"),
};

const renderer = new DiagramRenderer(document.getElementById("diagramCanvas"), {
  stateRef: state,
  onHoleClick: handleHoleClick,
  onHoleHover: handleHoleHover,
  onPointerUp: handlePointerUp,
  onHoleContextMenu: () => {},
});

const printRenderer = new DiagramRenderer(document.getElementById("printCanvas"), {
  stateRef: printState,
  onHoleClick: () => {},
  onHoleHover: () => {},
  onPointerUp: () => {},
  onHoleContextMenu: () => {},
});

initTimingControls(state, els, () => {
  resetTimingResults();
  renderer.render();
});

function closeAllMenus() {
  els.menuPanels.forEach((panel) => panel.classList.add("hidden"));
  els.menuToggles.forEach((button) => button.classList.remove("active"));
}

function toggleMenu(menuId) {
  const panel = document.getElementById(menuId);
  const button = els.menuToggles.find((item) => item.dataset.menuToggle === menuId);
  if (!panel || !button) return;
  const opening = panel.classList.contains("hidden");
  closeAllMenus();
  if (!opening) return;
  panel.classList.remove("hidden");
  button.classList.add("active");
}

function initMenuToggles() {
  els.menuToggles.forEach((button) => {
    button.addEventListener("click", (event) => {
      event.stopPropagation();
      toggleMenu(button.dataset.menuToggle);
    });
  });

  els.menuPanels.forEach((panel) => {
    panel.addEventListener("click", (event) => event.stopPropagation());
  });

  document.addEventListener("click", () => closeAllMenus());
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") closeAllMenus();
  });
}

function openMenu(menuId) {
  if (!menuId) return;
  const panel = document.getElementById(menuId);
  const button = els.menuToggles.find((item) => item.dataset.menuToggle === menuId);
  if (!panel || !button) return;
  closeAllMenus();
  panel.classList.remove("hidden");
  button.classList.add("active");
}

function cloneSelectedTiming(selectedTiming) {
  if (!selectedTiming) return [];
  return [{
    ...selectedTiming,
    holeTimes: new Map(selectedTiming.holeTimes),
    offsetAssignments: selectedTiming.offsetAssignments ? new Map(selectedTiming.offsetAssignments) : new Map(),
  }];
}

function loadPrintState(selectedTiming) {
  printState.holes = state.holes.map((hole) => ({ ...hole }));
  printState.holesById = new Map(printState.holes.map((hole) => [hole.id, hole]));
  printState.selection = new Set();
  printState.relationships = {
    originHoleId: state.relationships.originHoleId,
    edges: state.relationships.edges.map((edge) => ({ ...edge })),
    nextId: state.relationships.nextId,
  };
  printState.timingResults = cloneSelectedTiming(selectedTiming);
  printState.ui.activeTimingPreviewIndex = printState.timingResults.length ? 0 : -1;
  printState.ui.showGrid = false;
  printState.ui.showRelationships = state.ui.showRelationships;
  printState.ui.showOverlayText = true;
  printState.ui.textScale = Number(els.printTextScaleInput.value) || 1;
  printState.ui.orientation = "landscape";
}

function syncPrintControls() {
  els.printTextScaleInput.value = String(printState.ui.textScale || 1);
  els.printOrientationSelect.value = printState.ui.orientation || "landscape";
  els.printRelationshipToggle.checked = printState.ui.showRelationships !== false;
  els.printColorModeToggle.checked = !els.printPaperFrame.classList.contains("greyscale");
}

function applyPrintOrientation() {
  const orientation = els.printOrientationSelect.value === "portrait" ? "portrait" : "landscape";
  printState.ui.orientation = orientation;
  els.printWorkspace.classList.toggle("portrait", orientation === "portrait");
  els.printWorkspace.classList.toggle("landscape", orientation === "landscape");
  els.printPaperFrame.classList.toggle("portrait", orientation === "portrait");
  els.printPaperFrame.classList.toggle("landscape", orientation === "landscape");
}

function openPrintWorkspace() {
  const selectedTiming = state.timingResults[state.ui.activeTimingPreviewIndex] || null;
  if (!selectedTiming) {
    window.alert("Select a timing result first, then open print preview.");
    return;
  }
  loadPrintState(selectedTiming);
  els.printPaperFrame.classList.remove("greyscale");
  els.printColorModeToggle.checked = true;
  syncPrintControls();
  applyPrintOrientation();
  closeHelpWorkspace();
  document.body.classList.add("print-preview-active");
  els.printWorkspace.classList.remove("hidden");
  closeAllMenus();
  requestAnimationFrame(() => {
    printRenderer.resize();
    printRenderer.rotationDeg = renderer.rotationDeg;
    printRenderer.fitToData();
  });
}

function closePrintWorkspace() {
  document.body.classList.remove("print-preview-active");
  els.printWorkspace.classList.add("hidden");
}

function openHelpWorkspace() {
  closeAllMenus();
  closePrintWorkspace();
  document.body.classList.add("help-active");
  els.helpWorkspace.classList.remove("hidden");
}

function closeHelpWorkspace() {
  document.body.classList.remove("help-active");
  els.helpWorkspace.classList.add("hidden");
}

function applyPrintSettings() {
  printState.ui.textScale = Number(els.printTextScaleInput.value) || 1;
  printState.ui.showRelationships = els.printRelationshipToggle.checked;
  els.printPaperFrame.classList.toggle("greyscale", !els.printColorModeToggle.checked);
  applyPrintOrientation();
  printRenderer.render();
}

function csvEscape(value) {
  const text = value === null || value === undefined ? "" : String(value);
  if (!/[",\n\r]/.test(text)) return text;
  return `"${text.replace(/"/g, "\"\"")}"`;
}

function exportSelectedTimingCsv() {
  const selectedTiming = state.timingResults[state.ui.activeTimingPreviewIndex] || null;
  if (!selectedTiming) {
    window.alert("Select a timing result first, then export CSV.");
    return;
  }

  const rows = state.holes.map((hole) => {
    const originalX = hole.collar?.original?.x ?? hole.original?.x ?? "";
    const originalY = hole.collar?.original?.y ?? hole.original?.y ?? "";
    const delayTime = selectedTiming.holeTimes instanceof Map ? selectedTiming.holeTimes.get(hole.id) : undefined;
    return [
      hole.holeNumber || hole.id,
      originalX,
      originalY,
      Number.isFinite(delayTime) ? delayTime : "",
    ];
  });

  const csvText = [
    ["hole_number", "x", "y", "delay_time_ms"],
    ...rows,
  ].map((row) => row.map(csvEscape).join(",")).join("\r\n");

  const blob = new Blob([csvText], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "timing-delays.csv";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function resetTimingResults(message = "") {
  state.timingResults = [];
  state.ui.activeTimingPreviewIndex = -1;
  state.solverMessage = message;
  renderTimingResults();
}

function uniqueHoleIds(holes, records, idColumn) {
  const seen = new Set();
  holes.forEach((hole) => {
    let id = String(hole.id);
    if (idColumn && records[hole.sourceIndex]?.[idColumn]) id = String(records[hole.sourceIndex][idColumn]);
    while (seen.has(id)) id = `${id}_dup`;
    hole.id = id;
    seen.add(id);
  });
}

function inferHeaderByPriority(headers, priorityGroups) {
  const lower = headers.map((header) => ({ raw: header, low: header.toLowerCase() }));
  for (const group of priorityGroups) {
    const match = lower.find((entry) => group.every((needle) => entry.low.includes(needle)));
    if (match) return match.raw;
  }
  return "";
}

function setColumnOptions(headers) {
  [els.xColumnSelect, els.yColumnSelect, els.toeXColumnSelect, els.toeYColumnSelect, els.idColumnSelect].forEach((select) => {
    select.innerHTML = "";
    if (select === els.idColumnSelect || select === els.toeXColumnSelect || select === els.toeYColumnSelect) {
      const none = document.createElement("option");
      none.value = "";
      none.textContent = select === els.idColumnSelect ? "(Auto)" : "(None)";
      select.appendChild(none);
    }
    headers.forEach((header) => {
      const option = document.createElement("option");
      option.value = header;
      option.textContent = header;
      select.appendChild(option);
    });
  });

  const xGuess = inferHeaderByPriority(headers, [["start", "point", "easting"], ["start", "easting"], ["easting"], ["longitude"], ["x"]]);
  const yGuess = inferHeaderByPriority(headers, [["start", "point", "northing"], ["start", "northing"], ["northing"], ["latitude"], ["y"]]);
  const toeXGuess = inferHeaderByPriority(headers, [["toe", "easting"], ["end", "point", "easting"], ["toe", "longitude"], ["end", "point", "longitude"], ["toe", "x"]]);
  const toeYGuess = inferHeaderByPriority(headers, [["toe", "northing"], ["end", "point", "northing"], ["toe", "latitude"], ["end", "point", "latitude"], ["toe", "y"]]);
  const idGuess = inferHeaderByPriority(headers, [["hole"], ["id"]]);

  if (xGuess) els.xColumnSelect.value = xGuess;
  if (yGuess) els.yColumnSelect.value = yGuess;
  if (toeXGuess) els.toeXColumnSelect.value = toeXGuess;
  if (toeYGuess) els.toeYColumnSelect.value = toeYGuess;
  if (idGuess) els.idColumnSelect.value = idGuess;

  const lowerHeaders = headers.map((header) => header.toLowerCase());
  if (lowerHeaders.some((header) => header.includes("lat")) && lowerHeaders.some((header) => header.includes("lon"))) {
    els.coordTypeSelect.value = "latlon";
  }
}

function rebuildHolesById() {
  state.holesById = new Map(state.holes.map((hole) => [hole.id, hole]));
}

function normalizeHoleCoordinateSets(hole) {
  if (!hole.collar || !Number.isFinite(hole.collar.x) || !Number.isFinite(hole.collar.y)) {
    hole.collar = { x: Number.isFinite(hole.x) ? hole.x : 0, y: Number.isFinite(hole.y) ? hole.y : 0, original: hole.original || null };
  }
  if (hole.toe && (!Number.isFinite(hole.toe.x) || !Number.isFinite(hole.toe.y))) hole.toe = null;
}

function hasAnyToeCoordinates() {
  return state.holes.some((hole) => hole.toe && Number.isFinite(hole.toe.x) && Number.isFinite(hole.toe.y));
}

function applyCoordinateView(view, { fit = false } = {}) {
  const hasToe = hasAnyToeCoordinates();
  const targetView = view === "toe" && hasToe ? "toe" : "collar";
  state.ui.coordView = targetView;

  state.holes.forEach((hole) => {
    normalizeHoleCoordinateSets(hole);
    const target = targetView === "toe" && hole.toe ? hole.toe : hole.collar;
    hole.x = target.x;
    hole.y = target.y;
  });

  els.coordViewSelect.disabled = !hasToe;
  els.coordViewSelect.value = targetView;
  renderer.render();
  if (fit) renderer.fitToData();
}

function renderOriginStatus() {
  const hole = state.holesById.get(state.relationships.originHoleId || "");
  els.originStatus.textContent = hole ? `Origin: ${hole.holeNumber || hole.id}` : "Origin: not set";
}

function renderRelationshipList() {
  if (!state.relationships.edges.length) {
    els.relationshipList.innerHTML = "<div>No relationships defined</div>";
    return;
  }
  els.relationshipList.innerHTML = state.relationships.edges.map((edge) => {
    const description = describeRelationship(edge, state.holesById);
    const actions = edge.type === "offset"
      ? `<button data-rel-action="delete" data-rel-id="${edge.id}">Delete</button>`
      : `<button data-rel-action="edit" data-rel-id="${edge.id}">Edit</button>
        <button data-rel-action="delete" data-rel-id="${edge.id}">Delete</button>`;
    return `<div class="relationship-row">
      <div>${description}</div>
      <div class="row-actions">
        ${actions}
      </div>
    </div>`;
  }).join("");
}

function syncRelationshipVisibilityUi() {
  els.relationshipVisibilityToggle.checked = state.ui.showRelationships;
  if (els.relationshipVisibilityToggleSecondary) {
    els.relationshipVisibilityToggleSecondary.checked = state.ui.showRelationships;
  }
}

function renderTimingResults() {
  const hasResults = state.timingResults.length > 0;
  els.timingResultsMenuWrap.classList.toggle("hidden", !hasResults);
  if (!hasResults) {
    const resultsButton = els.menuToggles.find((button) => button.dataset.menuToggle === "timingResultsMenu");
    document.getElementById("timingResultsMenu")?.classList.add("hidden");
    resultsButton?.classList.remove("active");
  }
  if (!state.timingResults.length) {
    els.timingResults.innerHTML = `<div>${state.solverMessage || "Run solver to see best delay combinations."}</div>`;
    return;
  }
  els.timingResults.innerHTML = state.timingResults.map((result, index) => {
    const active = index === state.ui.activeTimingPreviewIndex ? "active" : "";
    return `<button class="timing-item ${active}" data-timing-index="${index}">${formatTimingResult(result, index)}</button>`;
  }).join("");
}

function fullRefresh({ fit = false } = {}) {
  renderOriginStatus();
  renderRelationshipList();
  renderTimingResults();
  renderer.render();
  if (fit) renderer.fitToData();
}

function resetGraphState() {
  ensureRelationshipState(state);
  clearRelationships(state);
  state.ui.relationshipDraft = null;
}

function applyImportedHoles(holes) {
  holes.forEach((hole) => normalizeHoleCoordinateSets(hole));
  state.holes = holes;
  state.selection = new Set();
  state.ui.coordView = "collar";
  state.ui.relationshipDraft = null;
  rebuildHolesById();
  resetGraphState();
  resetTimingResults();
  applyCoordinateView("collar");
}

function setToolMode(mode) {
  state.ui.toolMode = mode;
  state.ui.relationshipDraft = null;
  els.toolModeStatus.textContent = `Tool: ${relationToolLabel(mode)}`;
  els.originToolBtn.classList.toggle("active", mode === "origin");
  els.holeRelationPositiveToolBtn.classList.toggle("active", mode === "holeRelationshipPositive");
  els.holeRelationNegativeToolBtn.classList.toggle("active", mode === "holeRelationshipNegative");
  els.rowRelationPositiveToolBtn.classList.toggle("active", mode === "rowRelationshipPositive");
  els.rowRelationNegativeToolBtn.classList.toggle("active", mode === "rowRelationshipNegative");
  els.offsetRelationToolBtn.classList.toggle("active", mode === "offsetRelationship");
  renderer.render();
}

function promptRelationshipConfig(type, existing = null) {
  const input = window.prompt("Enter relationship sign: + or -", existing?.sign === -1 ? "-" : "+");
  if (input === null) return null;
  const normalized = input.trim().toLowerCase();
  if (!["+", "-", "positive", "negative"].includes(normalized)) {
    window.alert("Enter + or -.");
    return null;
  }
  return { sign: normalized === "-" || normalized === "negative" ? -1 : 1 };
}

function finalizeRelationshipPath(holeIds, relationshipType, sign) {
  const uniquePath = [];
  holeIds.forEach((holeId) => {
    if (!holeId) return;
    if (uniquePath[uniquePath.length - 1] === holeId) return;
    uniquePath.push(holeId);
  });
  if (uniquePath.length < 2) return false;

  for (let index = 0; index < uniquePath.length - 1; index += 1) {
    const fromHoleId = uniquePath[index];
    const toHoleId = uniquePath[index + 1];
    if (fromHoleId === toHoleId) continue;
    addRelationship(state, { type: relationshipType, fromHoleId, toHoleId, sign });
  }
  return true;
}

function finalizeOffsetRelationship(toHoleId) {
  const draft = state.ui.relationshipDraft;
  const fromHoleId = draft?.holeIds?.[0] || null;
  if (!fromHoleId || !toHoleId || fromHoleId === toHoleId) {
    state.ui.relationshipDraft = null;
    renderer.render();
    return;
  }
  state.ui.relationshipDraft = null;
  addRelationship(state, { type: draft.type, fromHoleId, toHoleId });
  resetTimingResults();
  fullRefresh();
}

function handleHoleClick(hole, event) {
  if (state.ui.toolMode === "origin") {
    setOriginHole(state, hole.id);
    resetTimingResults();
    fullRefresh();
    return;
  }

  const relationshipType = TOOL_TO_RELATIONSHIP_TYPE[state.ui.toolMode];
  if (relationshipType) {
    if (!state.ui.relationshipDraft?.holeIds?.length) {
      state.ui.relationshipDraft = { type: relationshipType, sign: TOOL_TO_SIGN[state.ui.toolMode] ?? 1, holeIds: [hole.id] };
    } else if (state.ui.relationshipDraft.type === relationshipType) {
      const holeIds = state.ui.relationshipDraft.holeIds;
      if (holeIds[holeIds.length - 1] !== hole.id) holeIds.push(hole.id);
    } else {
      state.ui.relationshipDraft = { type: relationshipType, sign: TOOL_TO_SIGN[state.ui.toolMode] ?? 1, holeIds: [hole.id] };
    }
    renderer.render();
    return;
  }

  if (!event.shiftKey) state.selection = new Set([hole.id]);
  else if (state.selection.has(hole.id)) state.selection.delete(hole.id);
  else state.selection.add(hole.id);
  renderer.render();
}

function handleHoleHover(hole) {
  if (!state.ui.relationshipDraft?.holeIds?.length) return;
  if (state.ui.relationshipDraft.type === "offset") return;
  const holeIds = state.ui.relationshipDraft.holeIds;
  if (holeIds[holeIds.length - 1] === hole.id) return;
  holeIds.push(hole.id);
  renderer.render();
}

function handlePointerUp(payload) {
  const draft = state.ui.relationshipDraft;
  if (!draft?.holeIds?.length) return;

  if (draft.type === "offset") {
    finalizeOffsetRelationship(payload?.hole?.id || null);
    return;
  }

  const created = finalizeRelationshipPath(draft.holeIds, draft.type, draft.sign);
  state.ui.relationshipDraft = null;
  if (created) {
    resetTimingResults();
    fullRefresh();
    return;
  }
  renderer.render();
}

function editRelationship(edge) {
  if (edge.type === "offset") return;
  const config = promptRelationshipConfig(edge.type, edge);
  if (!config) return;
  updateRelationship(state, edge.id, config);
  resetTimingResults();
  fullRefresh();
}

els.csvInput.addEventListener("change", async () => {
  const file = els.csvInput.files[0];
  if (!file) return;
  const text = await file.text();
  const parsed = parseCsvText(text);
  state.csvCache = parsed;
  setColumnOptions(parsed.headers);
  els.mappingPanel.classList.remove("hidden");
});

els.importMappedBtn.addEventListener("click", () => {
  if (!state.csvCache) return;
  const { headers, records } = state.csvCache;
  if (!headers.length || !records.length) return;

  const idColumn = els.idColumnSelect.value || null;
  const toeXColumn = els.toeXColumnSelect.value || null;
  const toeYColumn = els.toeYColumnSelect.value || null;
  if ((toeXColumn && !toeYColumn) || (!toeXColumn && toeYColumn)) {
    window.alert("Select both Toe X and Toe Y columns, or leave both empty.");
    return;
  }

  const holes = buildHolesFromMapping({
    records,
    coordType: els.coordTypeSelect.value,
    xColumn: els.xColumnSelect.value,
    yColumn: els.yColumnSelect.value,
    idColumn,
  });
  if (!holes.length) {
    window.alert("No valid collar coordinates found for selected columns.");
    return;
  }

  let toeBySource = new Map();
  if (toeXColumn && toeYColumn) {
    const toeHoles = buildHolesFromMapping({
      records,
      coordType: els.coordTypeSelect.value,
      xColumn: toeXColumn,
      yColumn: toeYColumn,
      idColumn,
    });
    toeBySource = new Map(toeHoles.map((hole) => [hole.sourceIndex, { x: hole.x, y: hole.y, original: hole.original }]));
  }

  holes.forEach((hole) => {
    hole.collar = { x: hole.x, y: hole.y, original: hole.original };
    hole.toe = toeBySource.get(hole.sourceIndex) || null;
  });

  uniqueHoleIds(holes, records, idColumn);
  applyImportedHoles(holes);
  fullRefresh({ fit: true });
});

els.gridToggle.addEventListener("change", () => {
  state.ui.showGrid = els.gridToggle.checked;
  renderer.render();
});

els.relationshipVisibilityToggle.addEventListener("change", () => {
  state.ui.showRelationships = els.relationshipVisibilityToggle.checked;
  syncRelationshipVisibilityUi();
  renderer.render();
});

if (els.relationshipVisibilityToggleSecondary) {
  els.relationshipVisibilityToggleSecondary.addEventListener("change", () => {
    state.ui.showRelationships = els.relationshipVisibilityToggleSecondary.checked;
    syncRelationshipVisibilityUi();
    renderer.render();
  });
}

els.fitViewBtn.addEventListener("click", () => renderer.fitToData());
els.coordViewSelect.addEventListener("change", () => applyCoordinateView(els.coordViewSelect.value, { fit: true }));
els.rotateLeftBtn.addEventListener("click", () => renderer.rotateBy(-15));
els.rotateRightBtn.addEventListener("click", () => renderer.rotateBy(15));
els.rotateFineLeftBtn.addEventListener("click", () => renderer.rotateBy(-1));
els.rotateFineRightBtn.addEventListener("click", () => renderer.rotateBy(1));
els.rotateResetBtn.addEventListener("click", () => renderer.resetRotation());

els.originToolBtn.addEventListener("click", () => setToolMode("origin"));
els.holeRelationPositiveToolBtn.addEventListener("click", () => setToolMode("holeRelationshipPositive"));
els.holeRelationNegativeToolBtn.addEventListener("click", () => setToolMode("holeRelationshipNegative"));
els.rowRelationPositiveToolBtn.addEventListener("click", () => setToolMode("rowRelationshipPositive"));
els.rowRelationNegativeToolBtn.addEventListener("click", () => setToolMode("rowRelationshipNegative"));
els.offsetRelationToolBtn.addEventListener("click", () => setToolMode("offsetRelationship"));
els.clearRelationshipsBtn.addEventListener("click", () => {
  state.relationships.edges = [];
  state.ui.relationshipDraft = null;
  resetTimingResults();
  fullRefresh();
});

els.clearOriginBtn.addEventListener("click", () => {
  setOriginHole(state, null);
  resetTimingResults();
  fullRefresh();
});

els.relationshipList.addEventListener("click", (event) => {
  const button = event.target.closest("[data-rel-action]");
  if (!button) return;
  const edge = state.relationships.edges.find((item) => item.id === button.getAttribute("data-rel-id"));
  if (!edge) return;
  const action = button.getAttribute("data-rel-action");
  if (action === "edit") editRelationship(edge);
  if (action === "delete") {
    deleteRelationship(state, edge.id);
    resetTimingResults();
    fullRefresh();
  }
});

els.solveTimingBtn.addEventListener("click", () => {
  const validation = validateTimingGraph(state);
  if (!validation.valid) {
    resetTimingResults(validation.reason);
    renderer.render();
    return;
  }
  state.timingResults = solveTimingCombinations(state);
  state.ui.activeTimingPreviewIndex = state.timingResults.length ? 0 : -1;
  state.solverMessage = state.timingResults.length ? "" : "No valid timing combinations were produced for the current graph.";
  renderTimingResults();
  if (state.timingResults.length) openMenu("timingResultsMenu");
  renderer.render();
});

els.timingResults.addEventListener("click", (event) => {
  const target = event.target.closest("[data-timing-index]");
  if (!target) return;
  const index = Number(target.getAttribute("data-timing-index"));
  if (!Number.isFinite(index)) return;
  state.ui.activeTimingPreviewIndex = index;
  renderTimingResults();
  renderer.render();
});

els.exportPdfBtn.addEventListener("click", () => {
  openPrintWorkspace();
});

els.helpBtn.addEventListener("click", () => openHelpWorkspace());
els.csvExportBtn.addEventListener("click", () => exportSelectedTimingCsv());
els.helpBackBtn.addEventListener("click", () => closeHelpWorkspace());
els.printBackBtn.addEventListener("click", () => closePrintWorkspace());
els.printFitBtn.addEventListener("click", () => printRenderer.fitToData());
els.printTextScaleInput.addEventListener("input", () => applyPrintSettings());
els.printOrientationSelect.addEventListener("change", () => {
  applyPrintSettings();
  printRenderer.fitToData();
});
els.printColorModeToggle.addEventListener("change", () => applyPrintSettings());
els.printRelationshipToggle.addEventListener("change", () => applyPrintSettings());
els.printActionBtn.addEventListener("click", () => {
  window.print();
  closePrintWorkspace();
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && !els.helpWorkspace.classList.contains("hidden")) closeHelpWorkspace();
});

ensureRelationshipState(state);
setToolMode(state.ui.toolMode);
els.coordViewSelect.value = state.ui.coordView;
els.coordViewSelect.disabled = true;
syncRelationshipVisibilityUi();
renderOriginStatus();
renderRelationshipList();
renderTimingResults();
initMenuToggles();
renderer.render();
