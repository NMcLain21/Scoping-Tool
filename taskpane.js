// ─────────────────────────────────────────────────────────────────
// Default palette & localStorage persistence
// ─────────────────────────────────────────────────────────────────
const DEFAULT_PALETTE = [
  { hex: "#1F3864", name: "Dark Navy" },
  { hex: "#2E75B6", name: "Corporate Blue" },
  { hex: "#70AD47", name: "Green" },
  { hex: "#FFC000", name: "Amber" },
  { hex: "#C00000", name: "Red" },
  { hex: "#FFFFFF", name: "White" },
  { hex: "#000000", name: "Black" },
];

const STORAGE_KEY = "scopingToolPalette";

function loadPalette() {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) return JSON.parse(saved);
  } catch (_) {}
  return [...DEFAULT_PALETTE];
}

function savePalette(palette) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(palette));
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

// ─────────────────────────────────────────────────────────────────
// App state
// ─────────────────────────────────────────────────────────────────
let paletteEditorOpen = false;

// ─────────────────────────────────────────────────────────────────
// Office ready — wire up static controls, render swatches, start
// ─────────────────────────────────────────────────────────────────
Office.onReady(() => {

  // Border weight
  document.querySelectorAll("[data-weight]").forEach(btn =>
    btn.addEventListener("click", () => applyBorderWeight(parseFloat(btn.dataset.weight))));

  // Border dash style
  document.querySelectorAll("[data-dash]").forEach(btn =>
    btn.addEventListener("click", () => applyBorderDash(btn.dataset.dash)));

  // Line ends
  document.getElementById("btn-apply-ends")
    .addEventListener("click", applyLineEnds);

  // Palette toggle
  document.getElementById("btn-toggle-palette")
    .addEventListener("click", togglePaletteEditor);

  // Add color button
  document.getElementById("btn-add-color")
    .addEventListener("click", addColor);

  // Reset defaults button
  document.getElementById("btn-reset-palette")
    .addEventListener("click", resetToDefaults);

  // Allow Enter key in name field to submit
  document.getElementById("new-color-name")
    .addEventListener("keydown", e => { if (e.key === "Enter") addColor(); });

  // Render dynamic swatches from saved palette
  renderAllSwatches();

  // Selection change listener — skip if palette editor is open
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    () => { if (!paletteEditorOpen) inspectSelection(); }
  );

  // Initial inspection
  inspectSelection();
});

// ─────────────────────────────────────────────────────────────────
// Swatch rendering — builds swatch buttons dynamically from palette
// ─────────────────────────────────────────────────────────────────
function renderAllSwatches() {
  renderSectionSwatches("fill");
  renderSectionSwatches("border");
  renderSectionSwatches("text");
}

function renderSectionSwatches(type) {
  const palette  = loadPalette();
  const container = document.getElementById(`swatches-${type}`);
  if (!container) return;

  // Build swatch buttons from palette
  const buttons = palette.map(({ hex, name }) => {
    const light = isLightColor(hex);
    let attr = "";
    if (type === "fill")   attr = `data-fill="${hex}"`;
    if (type === "border") attr = `data-border="${hex}"`;
    if (type === "text")   attr = `data-textcolor="${hex}"`;
    return `<button class="swatch${light ? " swatch-light" : ""}" style="background:${hex};" ${attr} title="${name}"></button>`;
  }).join("");

  // "No fill" / "No border" always appended
  let noneBtn = "";
  if (type === "fill")   noneBtn = `<button class="swatch no-fill" data-fill="none"   title="No Fill">∅</button>`;
  if (type === "border") noneBtn = `<button class="swatch no-fill" data-border="none" title="No Border">∅</button>`;

  container.innerHTML = buttons + noneBtn;

  // Re-attach click handlers after innerHTML replacement
  container.querySelectorAll("[data-fill]").forEach(btn =>
    btn.addEventListener("click", () => applyFill(btn.dataset.fill)));
  container.querySelectorAll("[data-border]").forEach(btn =>
    btn.addEventListener("click", () => applyBorderColor(btn.dataset.border)));
  container.querySelectorAll("[data-textcolor]").forEach(btn =>
    btn.addEventListener("click", () => applyTextColor(btn.dataset.textcolor)));
}

// Returns true if the hex color is light (needs a border to be visible on white)
function isLightColor(hex) {
  try {
    const r = parseInt(hex.slice(1, 3), 16);
    const g = parseInt(hex.slice(3, 5), 16);
    const b = parseInt(hex.slice(5, 7), 16);
    return (r * 299 + g * 587 + b * 114) / 1000 > 180;
  } catch (_) { return false; }
}

// ─────────────────────────────────────────────────────────────────
// Palette editor — toggle, render, add, delete, reset
// ─────────────────────────────────────────────────────────────────
function togglePaletteEditor() {
  paletteEditorOpen = !paletteEditorOpen;
  const btn = document.getElementById("btn-toggle-palette");

  if (paletteEditorOpen) {
    // Hide main UI, show palette panel
    setVisible("empty-state",        false);
    setVisible("shape-banner",       false);
    setVisible("main-content",       false);
    setVisible("unsupported-state",  false);
    setVisible("palette-panel",      true);
    btn.textContent = "← Back to Styling";
    renderPaletteManager();
  } else {
    // Hide palette panel, restore main UI
    setVisible("palette-panel", false);
    btn.textContent = "🎨 Edit Palette";
    inspectSelection();
  }
}

function renderPaletteManager() {
  const palette   = loadPalette();
  const container = document.getElementById("palette-list");

  if (palette.length === 0) {
    container.innerHTML = `<p class="palette-empty">No colors yet. Add one below.</p>`;
    return;
  }

  container.innerHTML = palette.map((color, index) => `
    <div class="palette-item">
      <div class="palette-preview"
           style="background:${color.hex};${isLightColor(color.hex) ? "border:1px solid #ddd;" : ""}">
      </div>
      <div class="palette-info">
        <span class="palette-name">${escapeHtml(color.name)}</span>
        <span class="palette-hex">${color.hex.toUpperCase()}</span>
      </div>
      <button class="palette-delete" data-index="${index}" title="Remove color">✕</button>
    </div>
  `).join("");

  container.querySelectorAll(".palette-delete").forEach(btn =>
    btn.addEventListener("click", () => deleteColor(parseInt(btn.dataset.index))));
}

function addColor() {
  const hex     = document.getElementById("new-color-picker").value;
  const rawName = document.getElementById("new-color-name").value.trim();
  const name    = rawName || hex.toUpperCase();

  const palette = loadPalette();

  // Prevent exact duplicate hex values
  if (palette.some(c => c.hex.toLowerCase() === hex.toLowerCase())) {
    setStatus("That color is already in the palette.");
    return;
  }

  palette.push({ hex, name });
  savePalette(palette);
  renderPaletteManager();
  renderAllSwatches();

  // Reset the name field but keep the color picker on the last chosen color
  document.getElementById("new-color-name").value = "";
  setStatus(`Added: ${name} (${hex.toUpperCase()})`);
}

function deleteColor(index) {
  const palette = loadPalette();
  const removed = palette.splice(index, 1)[0];
  savePalette(palette);
  renderPaletteManager();
  renderAllSwatches();
  setStatus(`Removed: ${removed.name}`);
}

function resetToDefaults() {
  if (!confirm("Reset palette to the original default colors? This cannot be undone.")) return;
  savePalette([...DEFAULT_PALETTE]);
  renderPaletteManager();
  renderAllSwatches();
  setStatus("Palette reset to defaults.");
}

// Prevent XSS from user-supplied color names
function escapeHtml(str) {
  return str.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}

// ─────────────────────────────────────────────────────────────────
// Shape inspection — reads type of selected shapes, updates UI
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

function toKey(raw) {
  if (!raw) return "unsupported";
  const s = typeof raw === "string" ? raw : String(raw);
  return s.charAt(0).toLowerCase() + s.slice(1);
}

function capitalize(s) {
  return s.charAt(0).toUpperCase() + s.slice(1);
}

// ─────────────────────────────────────────────────────────────────
// Render: main styling UI based on capability flags
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
