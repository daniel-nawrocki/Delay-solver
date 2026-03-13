export function saveProject(state) {
  const snapshot = {
    holes: state.holes,
    rows: state.rows,
    timing: state.timing,
    centerPull: state.centerPull,
  };

  const blob = new Blob([JSON.stringify(snapshot, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "blast-project.json";
  a.click();
  URL.revokeObjectURL(url);
}

export async function loadProjectFile(file) {
  const text = await file.text();
  return JSON.parse(text);
}

export function hydrateStateFromProject(state, project) {
  state.holes = Array.isArray(project.holes) ? project.holes : [];
  state.holesById = new Map(state.holes.map((h) => [h.id, h]));
  state.rows = project.rows || {};
  state.timing = project.timing || state.timing;
  state.centerPull = project.centerPull || state.centerPull;
  state.selection = new Set();
}
