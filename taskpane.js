// =============================================================
// Scoping Tool — Office.js Task Pane Add-in
// PowerPoint shape styling with automatic type detection
// =============================================================

// ── Shape capability matrix ──────────────────────────────────
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

// ── Display metadata per type ─────────────────────────────────
const TYPE_META = {
  geometricShape: { icon: "◻", label: "Geometric Shape" },
  textBox:        { icon: "T",  label: "Text Box"        },
  placeholder:    { icon: "⊞", label: "Placeholder"      },
  callout:        { icon: "💬", label: "Callout"          },
  freeform:       { icon: "✏", label: "Freeform"         },
  group:          { icon: "⊡", label: "Group"            },
  line:           { icon: "╱", label: "Line"             },
  image:          { icon: "⬜", label: "Image"            },
  table:          { icon: "⊟", label: "Table"            },
};

// =============================================================
// Office ready — wire up all controls + selection listener
// =============================================================
Office.onReady(() => {
  // Fill swatches
  document.querySelectorAll("[data-fill]").forEach(btn =>
    btn.addEventListener("click", () => applyFill(btn.dataset.fill)));

  // Border color swatches
  document.querySelectorAll("[data-border]").forEach(btn =>
    btn.addEventListener("click", () => applyBorderColor(btn.dataset.border)));

  // Border weight buttons
  document.querySelectorAll("[data-weight]").forEach(btn =>
    btn.addEventListener("click", () => applyBorderWeight(parseFloat(btn.dataset.weight))));

  // Border dash-style buttons
  document.querySelectorAll("[data-dash]").forEach(btn =>
    btn.addEventListener("click", () => applyBorderDash(btn.dataset.dash)));

  // Text color swatches
  document.querySelectorAll("[data-textcolor]").forEach(btn =>
    btn.addEventListener("click", () => applyTextColor(btn.dataset.textcolor)));

  // Line ends apply button
  document.getElementById("btn-apply-ends")
    .addEventListener("click", applyLineEnds);

  // Listen for selection changes
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    () => inspectSelection()
  );

  // Initial inspection on pane open
  inspectSelection();
});

// =============================================================
// Inspect the current selection and update the UI
// =============================================================
async function inspectSelection() {
  try {
    await PowerPoint.run(async (context) => {
      const selected = context.presentation.getSelectedShapes();
      selected.load("items/type,items/name");
      await context.sync();

      const items = selected.items;
      if (items.length === 0) { renderEmptyState(); return; }

      // Accumulate merged capabilities across all selected shapes
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

      // Banner metadata
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
    // getSelectedShapes() throws when the selection is not on a shape
    renderEmptyState();
  }
}

// ── Normalize type string: "GeometricShape" → "geometricShape"
function toKey(raw) {
  if (!raw) return "unsupported";
  const s = typeof raw === "string" ? raw : String(raw);
  return s.charAt(0).toLowerCase() + s.slice(1);
}

function capitalize(s) {
  return s.charAt(0).toUpperCase() + s.slice(1);
}

// =============================================================
// Render UI — show/hide sections based on capabilities
// =============================================================
function renderUI({ capabilities, icon, typeLabel, shapeName, count, anySupported, isLineOnly }) {
  setVisible("empty-state",        false);
  setVisible("unsupported-state",  false);
  setVisible("shape-banner",       true);

  // Banner
  el("shape-icon").textContent       = icon;
  el("shape-type-label").textContent = typeLabel;
  el("shape-name").textContent       = shapeName ? `"${shapeName}"` : "";

  const badge = el("shape-count-badge");
  badge.textContent = `×${count}`;
  setVisible("shape-count-badge", count > 1);

  // Sections
  setVisible("section-fill",      capabilities.fill);
  setVisible("section-border",    capabilities.border);
  setVisible("section-line-ends", capabilities.lineEnds);
  setVisible("section-text",      capabilities.text);

  // Rename border heading for pure line selections
  el("border-heading").textContent = isLineOnly ? "Line Style" : "Border";

  if (!anySupported) setVisible("unsupported-state", true);
}

function renderEmptyState() {
  setVisible("empty-state",        true);
  setVisible("shape-banner",       false);
  setVisible("unsupported-state",  false);
  ["section-fill","section-border","section-line-ends","section-text"]
    .forEach(id => setVisible(id, false));
  setStatus("");
}

// =============================================================
// DOM helpers
// =============================================================
function el(id)                   { return document.getElementById(id); }
function setVisible(id, visible)  { const n = el(id); if (n) n.classList.toggle("hidden", !visible); }
function setStatus(msg)           { el("status").textContent = msg; }

// =============================================================
// Shared: get selected shapes inside a PowerPoint.run context
// =============================================================
async function getSelected(context) {
  const col = context.presentation.getSelectedShapes();
  col.load("items");
  await context.sync();
  return col.items;
}

// =============================================================
// Apply: Fill color
// =============================================================
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

// =============================================================
// Apply: Border color
// =============================================================
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

// =============================================================
// Apply: Border weight
// =============================================================
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

// =============================================================
// Apply: Border dash style
// =============================================================
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

// =============================================================
// Apply: Text color
// =============================================================
async function applyTextColor(color) {
  await PowerPoint.run(async (context) => {
    const shapes = await getSelected(context);
    if (!shapes.length) return setStatus("No shape selected.");
    shapes.forEach(s => {
      try { s.textFrame.textRange.font.color = color; } catch (_) { /* no text frame */ }
    });
    await context.sync();
    setStatus(`Text color → ${color}`);
  });
}

// =============================================================
// Apply: Line ends (Line shapes only)
// Requires API requirement set 1.5+
// =============================================================
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
      } catch (_) { /* ArrowheadStyle unavailable on older requirement sets */ }
    });
    await context.sync();
    setStatus(`Line ends → ${startVal} / ${endVal}`);
  });
}
