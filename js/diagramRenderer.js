import { relationshipColor } from "./relationshipManager.js";

function timingColor(value, min, max) {
  if (!Number.isFinite(value)) return "#64748b";
  if (max <= min) return "#0ea5e9";
  const t = (value - min) / (max - min);
  const hue = 210 - t * 210;
  return `hsl(${hue} 78% 46%)`;
}

function pointToSegmentDistance(point, start, end) {
  const dx = end.x - start.x;
  const dy = end.y - start.y;
  if (dx === 0 && dy === 0) return Math.hypot(point.x - start.x, point.y - start.y);
  const t = Math.max(0, Math.min(1, ((point.x - start.x) * dx + (point.y - start.y) * dy) / (dx * dx + dy * dy)));
  const px = start.x + t * dx;
  const py = start.y + t * dy;
  return Math.hypot(point.x - px, point.y - py);
}

export class DiagramRenderer {
  constructor(canvas, options = {}) {
    this.canvas = canvas;
    this.ctx = canvas.getContext("2d");
    this.onHoleClick = options.onHoleClick || (() => {});
    this.onHoleHover = options.onHoleHover || (() => {});
    this.onPointerUp = options.onPointerUp || (() => {});
    this.onHoleContextMenu = options.onHoleContextMenu || (() => {});
    this.stateRef = options.stateRef;
    this.zoom = 1;
    this.panX = 0;
    this.panY = 0;
    this.dragging = false;
    this.lastMouse = null;
    this.pointerScreen = null;
    this.holeRadius = 5;
    this.rotationDeg = 0;

    this.resize();
    this.attachEvents();
    window.addEventListener("resize", () => this.resize());
  }

  resize() {
    const rect = this.canvas.getBoundingClientRect();
    this.canvas.width = Math.max(400, Math.floor(rect.width));
    this.canvas.height = Math.max(300, Math.floor(rect.height));
    this.render();
  }

  textScale() {
    return Number(this.stateRef?.ui?.textScale) || 1;
  }

  rotatePoint(x, y) {
    const theta = (this.rotationDeg * Math.PI) / 180;
    const c = Math.cos(theta);
    const s = Math.sin(theta);
    return { x: x * c - y * s, y: x * s + y * c };
  }

  inverseRotatePoint(x, y) {
    const theta = (-this.rotationDeg * Math.PI) / 180;
    const c = Math.cos(theta);
    const s = Math.sin(theta);
    return { x: x * c - y * s, y: x * s + y * c };
  }

  worldToScreen(x, y) {
    const rotated = this.rotatePoint(x, y);
    return {
      x: rotated.x * this.zoom + this.panX,
      y: this.canvas.height - (rotated.y * this.zoom + this.panY),
    };
  }

  screenToWorld(x, y) {
    const xr = (x - this.panX) / this.zoom;
    const yr = (this.canvas.height - y - this.panY) / this.zoom;
    return this.inverseRotatePoint(xr, yr);
  }

  fitToData() {
    const holes = this.stateRef.holes;
    if (!holes.length) return;
    const rotated = holes.map((hole) => this.rotatePoint(hole.x, hole.y));
    const xs = rotated.map((hole) => hole.x);
    const ys = rotated.map((hole) => hole.y);
    const minX = Math.min(...xs);
    const maxX = Math.max(...xs);
    const minY = Math.min(...ys);
    const maxY = Math.max(...ys);
    const width = Math.max(1, maxX - minX);
    const height = Math.max(1, maxY - minY);
    const margin = 80;
    this.zoom = Math.max(0.02, Math.min((this.canvas.width - margin) / width, (this.canvas.height - margin) / height));
    this.panX = -minX * this.zoom + margin / 2;
    this.panY = -minY * this.zoom + margin / 2;
    this.render();
  }

  drawGrid() {
    if (!this.stateRef.ui.showGrid) return;
    const stepPx = 50;
    this.ctx.save();
    this.ctx.strokeStyle = "#edf2f7";
    this.ctx.lineWidth = 1;
    for (let x = this.panX % stepPx; x < this.canvas.width; x += stepPx) {
      this.ctx.beginPath();
      this.ctx.moveTo(x, 0);
      this.ctx.lineTo(x, this.canvas.height);
      this.ctx.stroke();
    }
    for (let y = this.panY % stepPx; y < this.canvas.height; y += stepPx) {
      this.ctx.beginPath();
      this.ctx.moveTo(0, y);
      this.ctx.lineTo(this.canvas.width, y);
      this.ctx.stroke();
    }
    this.ctx.restore();
  }

  drawNorthArrow() {
    const x = this.canvas.width - 50;
    const y = 65;
    const theta = (this.rotationDeg * Math.PI) / 180;
    const ux = Math.sin(theta);
    const uy = -Math.cos(theta);
    const tx = x + ux * 20;
    const ty = y + uy * 20;
    const bx = x - ux * 20;
    const by = y - uy * 20;
    const nx = x + ux * 28;
    const ny = y + uy * 28;
    this.ctx.save();
    this.ctx.fillStyle = "#0f172a";
    this.ctx.strokeStyle = "#0f172a";
    this.ctx.lineWidth = 2;
    this.ctx.beginPath();
    this.ctx.moveTo(nx, ny);
    this.ctx.lineTo(tx - uy * 6, ty + ux * 6);
    this.ctx.lineTo(tx + uy * 6, ty - ux * 6);
    this.ctx.closePath();
    this.ctx.fill();
    this.ctx.beginPath();
    this.ctx.moveTo(tx, ty);
    this.ctx.lineTo(bx, by);
    this.ctx.stroke();
    this.ctx.font = `bold ${Math.max(10, Math.round(13 * this.textScale()))}px Segoe UI`;
    this.ctx.fillText("N", nx - 5, ny - 8);
    this.ctx.restore();
  }

  drawArrow(start, end, color, dashed = false) {
    const angle = Math.atan2(end.y - start.y, end.x - start.x);
    const head = 10;
    this.ctx.save();
    this.ctx.strokeStyle = color;
    this.ctx.fillStyle = color;
    this.ctx.lineWidth = 2;
    if (dashed) this.ctx.setLineDash([6, 4]);
    this.ctx.beginPath();
    this.ctx.moveTo(start.x, start.y);
    this.ctx.lineTo(end.x, end.y);
    this.ctx.stroke();
    this.ctx.setLineDash([]);
    this.ctx.beginPath();
    this.ctx.moveTo(end.x, end.y);
    this.ctx.lineTo(end.x - head * Math.cos(angle - Math.PI / 6), end.y - head * Math.sin(angle - Math.PI / 6));
    this.ctx.lineTo(end.x - head * Math.cos(angle + Math.PI / 6), end.y - head * Math.sin(angle + Math.PI / 6));
    this.ctx.closePath();
    this.ctx.fill();
    this.ctx.restore();
  }

  drawRelationships() {
    if (this.stateRef.ui.showRelationships === false) return;
    const edges = this.stateRef.relationships?.edges || [];
    edges.forEach((edge) => {
      const fromHole = this.stateRef.holesById.get(edge.fromHoleId);
      const toHole = this.stateRef.holesById.get(edge.toHoleId);
      if (!fromHole || !toHole) return;
      const start = this.worldToScreen(fromHole.x, fromHole.y);
      const end = this.worldToScreen(toHole.x, toHole.y);
      const color = relationshipColor(edge.type);
      this.drawArrow(start, end, color);
    });
  }

  drawRelationshipDraft() {
    if (this.stateRef.ui.showRelationships === false) return;
    const draft = this.stateRef.ui.relationshipDraft;
    if (!draft?.holeIds?.length) return;
    const points = draft.holeIds
      .map((holeId) => this.stateRef.holesById.get(holeId))
      .filter(Boolean)
      .map((hole) => this.worldToScreen(hole.x, hole.y));
    if (!points.length) return;
    if (this.pointerScreen && draft.type !== "offset") points.push(this.pointerScreen);
    if (draft.type === "offset" && draft.holeIds.length === 1 && this.pointerScreen) points.push(this.pointerScreen);
    if (points.length < 2) return;

    this.ctx.save();
    this.ctx.strokeStyle = relationshipColor(draft.type || "holeToHole");
    this.ctx.lineWidth = 2;
    this.ctx.setLineDash([6, 4]);
    this.ctx.beginPath();
    points.forEach((point, index) => {
      if (index === 0) this.ctx.moveTo(point.x, point.y);
      else this.ctx.lineTo(point.x, point.y);
    });
    this.ctx.stroke();
    this.ctx.restore();
  }

  drawHoles(showLabels = true, showTiming = true) {
    const preview = this.stateRef.timingResults?.[this.stateRef.ui.activeTimingPreviewIndex] || null;
    const times = preview ? this.stateRef.holes.map((hole) => preview.holeTimes.get(hole.id)).filter((v) => Number.isFinite(v)) : [];
    const minT = times.length ? Math.min(...times) : 0;
    const maxT = times.length ? Math.max(...times) : 0;
    const originHoleId = this.stateRef.relationships?.originHoleId || null;

    for (const hole of this.stateRef.holes) {
      const point = this.worldToScreen(hole.x, hole.y);
      const selected = this.stateRef.selection.has(hole.id);
      const isOrigin = hole.id === originHoleId;
      const time = preview ? preview.holeTimes.get(hole.id) : null;
      this.ctx.beginPath();
      this.ctx.arc(point.x, point.y, this.holeRadius, 0, Math.PI * 2);
      this.ctx.fillStyle = preview ? timingColor(time, minT, maxT) : "#475569";
      this.ctx.fill();
      this.ctx.lineWidth = isOrigin ? 4 : selected ? 3 : 1;
      this.ctx.strokeStyle = isOrigin ? "#f59e0b" : selected ? "#0f172a" : "#dbe4ee";
      this.ctx.stroke();

      if (isOrigin) {
        this.ctx.beginPath();
        this.ctx.arc(point.x, point.y, this.holeRadius + 6, 0, Math.PI * 2);
        this.ctx.strokeStyle = "rgba(245, 158, 11, 0.55)";
        this.ctx.lineWidth = 2;
        this.ctx.stroke();
      }

      if (showLabels) {
        const label = hole.holeNumber || hole.id;
        this.ctx.fillStyle = "#111827";
        const labelSize = Math.max(9, Math.round(11 * this.textScale()));
        this.ctx.font = selected || isOrigin ? `bold ${labelSize}px Segoe UI` : `${labelSize}px Segoe UI`;
        this.ctx.fillText(label, point.x + 8, point.y - 6);
      }

      if (showTiming && preview && Number.isFinite(time)) {
        this.ctx.fillStyle = "#334155";
        this.ctx.font = `${Math.max(8, Math.round(10 * this.textScale()))}px Segoe UI`;
        this.ctx.fillText(`${time.toFixed(0)}ms`, point.x + 8, point.y + 8);
      }
    }
  }

  clear() {
    this.ctx.clearRect(0, 0, this.canvas.width, this.canvas.height);
    this.ctx.fillStyle = "#ffffff";
    this.ctx.fillRect(0, 0, this.canvas.width, this.canvas.height);
  }

  drawTimingPreviewInfo() {
    const preview = this.stateRef.timingResults?.[this.stateRef.ui.activeTimingPreviewIndex] || null;
    if (!preview) return;
    const topOverlayOffset = 92;
    this.ctx.save();
    this.ctx.fillStyle = "#0f172a";
    this.ctx.font = `${Math.max(10, Math.round(12 * this.textScale()))}px Segoe UI`;
    this.ctx.fillText(
      `Timing Preview: H2H ${preview.holeDelay}ms | R2R ${preview.rowDelay}ms | Peak(8ms): ${preview.density8ms}`,
      14,
      topOverlayOffset
    );
    this.ctx.restore();
  }

  render() {
    this.clear();
    this.drawGrid();
    this.drawRelationships();
    this.drawRelationshipDraft();
    this.drawHoles();
    this.drawNorthArrow();
    if (this.stateRef.ui.showOverlayText !== false) this.drawTimingPreviewInfo();
  }

  rotateBy(deltaDeg) {
    this.setRotation(this.rotationDeg + deltaDeg);
  }

  setRotation(deg) {
    if (!Number.isFinite(deg)) return;
    this.rotationDeg = ((deg % 360) + 360) % 360;
    if (this.rotationDeg > 180) this.rotationDeg -= 360;
    this.render();
  }

  resetRotation() {
    this.setRotation(0);
  }

  findHoleAtScreen(x, y) {
    for (const hole of this.stateRef.holes) {
      const point = this.worldToScreen(hole.x, hole.y);
      if (Math.hypot(x - point.x, y - point.y) <= this.holeRadius + 4) return hole;
    }
    return null;
  }

  findRelationshipAtScreen(x, y) {
    const point = { x, y };
    return (this.stateRef.relationships?.edges || []).find((edge) => {
      const fromHole = this.stateRef.holesById.get(edge.fromHoleId);
      const toHole = this.stateRef.holesById.get(edge.toHoleId);
      if (!fromHole || !toHole) return false;
      const start = this.worldToScreen(fromHole.x, fromHole.y);
      const end = this.worldToScreen(toHole.x, toHole.y);
      return pointToSegmentDistance(point, start, end) <= 8;
    }) || null;
  }

  attachEvents() {
    this.canvas.addEventListener("mousedown", (event) => {
      const rect = this.canvas.getBoundingClientRect();
      const x = event.clientX - rect.left;
      const y = event.clientY - rect.top;
      this.pointerScreen = { x, y };
      const hole = this.findHoleAtScreen(x, y);
      if (hole) {
        this.onHoleClick(hole, event);
        return;
      }
      this.dragging = true;
      this.lastMouse = { x: event.clientX, y: event.clientY };
    });

    window.addEventListener("mousemove", (event) => {
      const rect = this.canvas.getBoundingClientRect();
      this.pointerScreen = { x: event.clientX - rect.left, y: event.clientY - rect.top };
      const hoverHole = this.findHoleAtScreen(this.pointerScreen.x, this.pointerScreen.y);
      if ((event.buttons & 1) === 1 && hoverHole) this.onHoleHover(hoverHole, event);
      if (!this.dragging || !this.lastMouse) {
        if (this.stateRef.ui.relationshipDraft?.holeIds?.length) this.render();
        return;
      }
      const dx = event.clientX - this.lastMouse.x;
      const dy = event.clientY - this.lastMouse.y;
      this.panX += dx;
      this.panY -= dy;
      this.lastMouse = { x: event.clientX, y: event.clientY };
      this.render();
    });

    window.addEventListener("mouseup", (event) => {
      const rect = this.canvas.getBoundingClientRect();
      const x = event.clientX - rect.left;
      const y = event.clientY - rect.top;
      this.pointerScreen = { x, y };
      this.dragging = false;
      this.lastMouse = null;
      this.onPointerUp({ hole: this.findHoleAtScreen(x, y), event });
    });

    this.canvas.addEventListener("contextmenu", (event) => {
      const rect = this.canvas.getBoundingClientRect();
      const x = event.clientX - rect.left;
      const y = event.clientY - rect.top;
      const hole = this.findHoleAtScreen(x, y);
      if (hole) {
        event.preventDefault();
        this.onHoleContextMenu(hole, event);
        return;
      }
      const relationship = this.findRelationshipAtScreen(x, y);
      if (relationship) event.preventDefault();
    });

    this.canvas.addEventListener("wheel", (event) => {
      event.preventDefault();
      if (event.ctrlKey) {
        this.rotateBy(event.deltaY < 0 ? 1 : -1);
        return;
      }
      const rect = this.canvas.getBoundingClientRect();
      const mouseX = event.clientX - rect.left;
      const mouseY = event.clientY - rect.top;
      const before = this.screenToWorld(mouseX, mouseY);
      this.zoom = Math.max(0.01, Math.min(300, this.zoom * (event.deltaY < 0 ? 1.1 : 0.9)));
      const after = this.screenToWorld(mouseX, mouseY);
      this.panX += (after.x - before.x) * this.zoom;
      this.panY += (after.y - before.y) * this.zoom;
      this.render();
    }, { passive: false });
  }
}
