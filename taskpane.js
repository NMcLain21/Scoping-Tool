// ─────────────────────────────────────────────────────────────────
// Default palettes — each palette is independent
// ─────────────────────────────────────────────────────────────────
const DEFAULTS = {
  fill: [
    { hex: "#1F3864", name: "Dark Navy" },
    { hex: "#2E75B6", name: "Corporate Blue" },
    { hex: "#70AD47", name: "Green" },
    { hex: "#FFC000", name: "Amber" },
    { hex: "#C00000", name: "Red" },
    { hex: "#FFFFFF", name: "White" },
    { hex: "#F2F2F2", name: "Light Gray" },
  ],
  border: [
    { hex: "#1F3864", name: "Dark Navy" },
    { hex: "#2E75B6", name: "Corporate Blue" },
    { hex: "#000000", name: "Black" },
    { hex: "#595959", name: "Dark Gray" },
    { hex: "#C00000", name: "Red" },
    { hex: "#FFFFFF", name: "White" },
  ],
  text: [
    { hex: "#000000", name: "Black" },
    { hex: "#FFFFFF", name: "White" },
    { hex: "#1F3864", name: "Dark Navy" },
    { hex: "#2E75B6", name: "Corporate Blue" },
    { hex: "#70AD47", name: "Green" },
    { hex: "#FFC000", name: "Amber" },
    { hex: "#C00000", name: "Red" },
  ],
};

const STORAGE_KEYS = {
  fill:   "scopingToolPalette_fill",
  border: "scopingToolPalette_border",
  text:   "scopingToolPalette_text",
};

// ─────────────────────────────────────────────────────────────────
// Palette persistence
// ─────────────────────────────────────────────────────────────────
function loadPalette(type) {
  try {
    const saved = localStorage.getItem(STORAGE_KEYS[type]);
    if (saved) return JSON.parse(saved);
  } catch (_) {}
  return [...DEFAULTS[type]];
}

function savePalette(type, palette) {
  try {
    localStorage.setItem(STORAGE_KEYS[type], JSON.stringify(palette));
  } catch (_) {}
}

// ─────────────────────────────────────────────────────────────────
// Shape capability matrix
// ─────────────────────────────────────────────────────────────────
const CAPABILITIES = {
  geometricShape: { fill: true,  border: true,  text: true,  lineEnds: false },
  textBox:        { fill: true,  border: true,  text: true,  lineEnds: false },
  placeholder:    { fill: true,  border: true,  text: true,  lineEnds: false },
  callout:        { fill: true,  border: true,  text: true,  lineEnds: false },
  freeform:       { fill: true,  border: true,  text: true,  lineEnds: false },
  group:          { fill: true,  border: true,  text: true,  lineEnds: false },
  line:           { fill: false, border: true,  text: false, lineEnds: true  },
  image:          { fill: false, border: true,  text: false, lineEnds: false },
  table:          { fill: false, border: true,  text: true,  lineEnds: false },
};

const TYPE_META = {
  geometricShape: { icon: "◻", label: "Geometric Shape" },
  textBox:        { icon: "T",  label: "Text Box" },
  placeholder:    { icon: "⊞", label: "Placeholder" },
  callout:        { icon: "💬", label: "Callout" },
  freeform:       { icon: "✏", label: "Freeform" },
  group:          { icon: "⊡", label: "Group" },
  line:           { icon: "╱", label: "Line" },
  image:          { icon: "⬜", label: "Image" },
  table:          { icon: "⊟", label: "Table" },
};

// Track which inline editors are open
const editorOpen = { fill: false, border: false, text: false };

// ─────────────────────────────────────────────────────────────────
// Office ready
// ─────────────────────────────────────────────────────────────────
Office.onReady(() => {

  // Render all three swatch rows from saved palettes
  renderSwatches("fill");
  renderSwatches("border");
  renderSwatches("text");

  // Border weight buttons
  document.querySelectorAll("[data-weight]").forEach(btn =>
    btn.addEventListener("click", () => applyBorderWeight(parseFloat(btn.dataset.weight))));

  // Border dash style buttons
  document.querySelectorAll("[data-dash]").forEach(btn =>
    btn.addEventListener("click", () => applyBorderDash(btn.dataset.dash)));

  // Line ends apply button
  document.getElementById("btn-apply-ends")
    .addEventListener("click", applyLineEnds);

  // Edit palette buttons (✏) — one per palette type
  document.querySelectorAll(".edit-palette-btn").forEach(btn =>
    btn.addEventListener("click", () => toggleEditor(btn.dataset.palette)));

  // Add color buttons — one per palette type
  document.querySelectorAll(".add-color-btn").forEach(btn =>
    btn.addEventListener("click", () => addColor(btn.dataset.palette)));

  // Allow Enter in name fields to submit
  ["fill", "border", "text"].forEach(type => {
    const nameField = document.getElementById(`name-${type}`);
    if (nameField) {
      nameField.addEventListener("keydown", e => { if (e.key === "Enter") addColor(type); });
    }
  });

  // Reset buttons — one per palette type
  document.querySelectorAll(".reset-palette-btn").forEach(btn =>
    btn.addEventListener("click", () => resetPalette(btn.dataset.palette)));

  // Selection change listener
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    () => inspectSelection()
  );

  // Initial inspection on open
  inspectSelection();
});

// ─────────────────────────────────────────────────────────────────
// Swatch rendering
// ─────────────────────────────────────────────────────────────────
function renderSwatches(type) {
  const palette   = loadPalette(type);
  const container = document.getElementById(`swatches-${type}`);
  if (!container) return;

  // Build color swatches
  const colorBtns = palette.map(({ hex, name }) => {
    const light = isLightColor(hex);
    let attr = "";
    if (type === "fill")   attr = `data-fill="${hex}"`;
    if (type === "border") attr = `data-border="${hex}"`;
    if (type === "text")   attr = `data-textcolor="${hex}"`;
    return `<button class="swatch${light ? " swatch-light" : ""}" style="background:${hex};" ${attr} title="${escapeHtml(name)}"></button>`;
  }).join("");

  // "None" button for fill and border only
  let noneBtn = "";
  if (type === "fill")   noneBtn = `<button class="swatch no-fill" data-fill="none"   title="No Fill">∅</button>`;
  if (type === "border") noneBtn = `<button class="swatch no-fill" data-border="none" title="No Border">∅</button>`;

  container.innerHTML = colorBtns + noneBtn;

  // Re-attach click handlers after innerHTML rebuild
  container.querySelectorAll("[data-fill]").forEach(btn =>
    btn.addEventListener("click", () => applyFill(btn.dataset.fill)));
  container.querySelectorAll("[data-border]").forEach(btn =>
    btn.addEventListener("click", () => applyBorderColor(btn.dataset.border)));
  container.querySelectorAll("[data-textcolor]").forEach(btn =>
    btn.addEventListener("click", () => applyTextColor(btn.dataset.textcolor)));
}

// ─────────────────────────────────────────────────────────────────
// Inline palette editor — toggle open/close
// ─────────────────────────────────────────────────────────────────
function toggleEditor(type) {
  editorOpen[type] = !editorOpen[type];
  const editorEl = document.getElementById(`editor-${type}`);
  const editBtn  = document.querySelector(`.edit-palette-btn[data-palette="${type}"]`);

  if (editorOpen[type]) {
    editorEl.classList.remove("hidden");
    editBtn.classList.add("active");
    editBtn.title = "Close editor";
    renderPaletteList(type);
  } else {
    editorEl.classList.add("hidden");
    editBtn.classList.remove("active");
    editBtn.title = `Edit ${type} palette`;
  }
}

// ─────────────────────────────────────────────────────────────────
// Inline palette list rendering
// ─────────────────────────────────────────────────────────────────
function renderPaletteList(type) {
  const palette   = loadPalette(type);
  const container = document.getElementById(`palette-list-${type}`);
  if (!container) return;

  if (palette.length === 0) {
    container.innerHTML = `<p class="palette-empty">No colors yet. Add one below.</p>`;
    return;
  }

  container.innerHTML = palette.map(({ hex, name }, index) => `
    <div class="palette-item">
      <div class="palette-preview" style="background:${hex};${isLightColor(hex) ? "border:1px solid #ddd;" : ""}"></div>
      <div class="palette-info">
        <span class="palette-name">${escapeHtml(name)}</span>
        <span class="palette-hex">${hex.toUpperCase()}</span>
      </div>
      <button class="palette-delete" data-type="${type}" data-index="${index}" title="Remove">✕</button>
    </div>
  `).join("");

  // Wire up delete buttons
  container.querySelectorAll(".palette-delete").forEach(btn =>
    btn.addEventListener("click", () => deleteColor(btn.dataset.type, parseInt(btn.dataset.index))));
}

// ─────────────────────────────────────────────────────────────────
// Add color to a palette
// ─────────────────────────────────────────────────────────────────
function addColor(type) {
  const hex     = document.getElementById(`picker-${type}`).value;
  const rawName = document.getElementById(`name-${type}`).value.trim();
  const name    = rawName || hex.toUpperCase();

  const palette = loadPalette(type);

  if (palette.some(c => c.hex.toLowerCase() === hex.toLowerCase())) {
    setStatus(`That color is already in the ${type} palette.`);
    return;
  }

  palette.push({ hex, name });
  savePalette(type, palette);

  // Refresh both the list and the swatch row
  renderPaletteList(type);
  renderSwatches(type);

  document.getElementById(`name-${type}`).value = "";
  setStatus(`Added to ${type}: ${name} (${hex.toUpperCase()})`);
}

// ─────────────────────────────────────────────────────────────────
// Delete color from a palette
// ─────────────────────────────────────────────────────────────────
function deleteColor(type, index) {
  const palette = loadPalette(type);
  const removed = palette.splice(index, 1)[0];
  savePalette(type, palette);
  renderPaletteList(type);
  renderSwatches(type);
  setStatus(`Removed from ${type}: ${removed.name}`);
}

// ─────────────────────────────────────────────────────────────────
// Reset a palette to defaults
// ─────────────────────────────────────────────────────────────────
function resetPalette(type) {
  const label = type.charAt(0).toUpperCase() + type.slice(1);
  if (!confirm(`Reset the ${label} palette to its original defaults? This cannot be undone.`)) return;
  savePalette(type, [...DEFAULTS[type]]);
  renderPaletteList(type);
  renderSwatches(type);
  setStatus(`${label} palette reset to defaults.`);
}

// ─────────────────────────────────────────────────────────────────
// Shape inspection
// ─────────────────────────────────────────────────────────────────
async function inspectSelection() {
  try {
    await PowerPoint.run(async (context) => {
      const selected = context.presentation.getSelectedShapes();
      selected.load("items/type,items/name");
      await context.sync();

      const items = selected.items;
      if (items.length === 0) { renderEmptyState(); return; }

      const merged = { fill: false, border: false, text: false, lineEnds: false };
      let allSameType = true;
      const firstType = toKey(items[0].type);
      const firstName = items[0].name || "";

      items.forEach(shape => {
        const key = toKey(shape.type);
        if (key !== firstType) allSameType = false;
        const caps = CAPABILITIES[key] || {};
        if (caps.fill)     merged.fill     = true;
        if (caps.border)   merged.border   = true;
        if (caps.text)     merged.text     = true;
        if (caps.lineEnds) merged.lineEnds = true;
      });

      const anySupported = Object.values(merged).some(Boolean);
      const meta = allSameType
        ? (TYPE_META[firstType] || { icon: "?", label: capitalize(firstType || "Shape") })
        : { icon: "⊕", label: "Mixed Selection" };

      renderUI({
        capabilities: merged,
        icon:         meta.icon,
        typeLabel:    meta.label,
        shapeName:    items.length === 1 ? firstName : `${items.length} shapes`,
        count:        items.length,
        anySupported,
        isLineOnly:   allSameType && firstType === "line",
      });
    });
  } catch (_) {
    renderEmptyState();
  }
}

// ─────────────────────────────────────────────────────────────────
// Render: main UI sections based on capabilities
// ─────────────────────────────────────────────────────────────────
function renderUI({ capabilities, icon, typeLabel, shapeName, count, anySupported, isLineOnly }) {
  setVisible("empty-state",       false);
  setVisible("unsupported-state", false);
  setVisible("shape-banner",      true);
  setVisible("main-content",      true);

  el("shape-icon").textContent       = icon;
  el("shape-type-label").textContent = typeLabel;
  el("shape-name").textContent       = shapeName ? `"${shapeName}"` : "";

  const badge = el("shape-count-badge");
  badge.textContent = `×${count}`;
  setVisible("shape-count-badge", count > 1);

  setVisible("section-fill",      capabilities.fill);
  setVisible("section-border",    capabilities.border);
  setVisible("section-line-ends", capabilities.lineEnds);
  setVisible("section-text",      capabilities.text);

  el("border-heading").textContent = isLineOnly ? "Line Style" : "Border";

  if (!anySupported) setVisible("unsupported-state", true);
}

function renderEmptyState() {
  setVisible("empty-state",        true);
  setVisible("shape-banner",       false);
  setVisible("unsupported-state",  false);
  setVisible("main-content",       true);
  ["section-fill","section-border","section-line-ends","section-text"]
    .forEach(id => setVisible(id, false));
  setStatus("");
}

// ─────────────────────────────────────────────────────────────────
// DOM helpers
// ─────────────────────────────────────────────────────────────────
function el(id) { return document.getElementById(id); }

function setVisible(id, visible) {
  const node = el(id);
  if (!node) return;
  node.classList.toggle("hidden", !visible);
}

function setStatus(msg) { el("status").textContent = msg; }

function toKey(raw) {
  if (!raw) return "unsupported";
  const s = typeof raw === "string" ? raw : String(raw);
  return s.charAt(0).toLowerCase() + s.slice(1);
}

function capitalize(s) { return s.charAt(0).toUpperCase() + s.slice(1); }

function isLightColor(hex) {
  try {
    const r = parseInt(hex.slice(1, 3), 16);
    const g = parseInt(hex.slice(3, 5), 16);
    const b = parseInt(hex.slice(5, 7), 16);
    return (r * 299 + g * 587 + b * 114) / 1000 > 180;
  } catch (_) { return false; }
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g,"&amp;").replace(/</g,"&lt;")
    .replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}

// ─────────────────────────────────────────────────────────────────
// Shared: get selected shapes inside a PowerPoint.run context
// ─────────────────────────────────────────────────────────────────
async function getSelected(context) {
  const col = context.presentation.getSelectedShapes();
  col.load("items");
  await context.sync();
  return col.items;
}

// ─────────────────────────────────────────────────────────────────
// Apply: Fill color
// ─────────────────────────────────────────────────────────────────
async function applyFill(color) {
  await PowerPoint.run(async (context) => {
    const shapes = await getSelected(context);
    if (!shapes.length) return setStatus("No shape selected.");
    shapes.forEach(s => {
      if (color === "none") s.fill.clear();
      else s.fill.setSolidColor(color);
    });
    await context.sync();
    setStatus(`Fill → ${color === "none" ? "cleared" : color}`);
  });
}

// ─────────────────────────────────────────────────────────────────
// Apply: Border color
// ─────────────────────────────────────────────────────────────────
async function applyBorderColor(color) {
  await PowerPoint.run(async (context) => {
    const shapes = await getSelected(context);
    if (!shapes.length) return setStatus("No shape selected.");
    shapes.forEach(s => {
      if (color === "none") {
        s.lineFormat.visible = false;
      } else {
        s.lineFormat.visible = true;
        s.lineFormat.color   = color;
      }
    });
    await context.sync();
    setStatus(`Border color → ${color === "none" ? "removed" : color}`);
  });
}

// ─────────────────────────────────────────────────────────────────
// Apply: Border weight
// ─────────────────────────────────────────────────────────────────
async function applyBorderWeight(pts) {
  await PowerPoint.run(async (context) => {
    const shapes = await getSelected(context);
    if (!shapes.length) return setStatus("No shape selected.");
    shapes.forEach(s => {
      s.lineFormat.weight  = pts;
      s.lineFormat.visible = true;
    });
    await context.sync();
    setStatus(`Border weight → ${pts}pt`);
  });
}

// ─────────────────────────────────────────────────────────────────
// Apply: Border dash style
// ─────────────────────────────────────────────────────────────────
async function applyBorderDash(style) {
  const map = {
    solid: PowerPoint.LineDashStyle.solid,
    dash:  PowerPoint.LineDashStyle.dash,
    dot:   PowerPoint.LineDashStyle.dot,
  };
  await PowerPoint.run(async (context) => {
    const shapes = await getSelected(context);
    if (!shapes.length) return setStatus("No shape selected.");
    shapes.forEach(s => {
      s.lineFormat.dashStyle = map[style] ?? PowerPoint.LineDashStyle.solid;
      s.lineFormat.visible   = true;
    });
    await context.sync();
    setStatus(`Border style → ${style}`);
  });
}

// ─────────────────────────────────────────────────────────────────
// Apply: Text color
// ─────────────────────────────────────────────────────────────────
async function applyTextColor(color) {
  await PowerPoint.run(async (context) => {
    const shapes = await getSelected(context);
    if (!shapes.length) return setStatus("No shape selected.");
    shapes.forEach(s => {
      try { s.textFrame.textRange.font.color = color; } catch (_) {}
    });
    await context.sync();
    setStatus(`Text color → ${color}`);
  });
}

// ─────────────────────────────────────────────────────────────────
// Apply: Line ends (Line shapes only)
// ─────────────────────────────────────────────────────────────────
async function applyLineEnds() {
  const startVal = el("line-start").value;
  const endVal   = el("line-end").value;

  const arrowMap = {
    none:      PowerPoint.ArrowheadStyle.none,
    arrow:     PowerPoint.ArrowheadStyle.arrow,
    openArrow: PowerPoint.ArrowheadStyle.openArrow,
    diamond:   PowerPoint.ArrowheadStyle.diamond,
    oval:      PowerPoint.ArrowheadStyle.oval,
  };

  await PowerPoint.run(async (context) => {
    const shapes = await getSelected(context);
    if (!shapes.length) return setStatus("No shape selected.");
    shapes.forEach(s => {
      if (toKey(s.type) !== "line") return;
      try {
        s.lineFormat.beginArrowheadStyle = arrowMap[startVal] ?? PowerPoint.ArrowheadStyle.none;
        s.lineFormat.endArrowheadStyle   = arrowMap[endVal]   ?? PowerPoint.ArrowheadStyle.none;
      } catch (_) {}
    });
    await context.sync();
    setStatus(`Line ends → ${startVal} / ${endVal}`);
  });
}
