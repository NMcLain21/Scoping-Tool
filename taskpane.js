'use strict';

// ─────────────────────────────────────────────────────────────────
// Constants
// ─────────────────────────────────────────────────────────────────

const BRAND_COLORS = [
  { hex: '#151F37', name: 'Navy'  },
  { hex: '#2A3E6D', name: 'Blue'  },
  { hex: '#D50032', name: 'Red'   },
  { hex: '#008579', name: 'Teal'  },
  { hex: '#E2E3E2', name: 'Gray'  },
  { hex: '#FFFFFF', name: 'White' },
  { hex: '#000000', name: 'Black' },
];

const STANDARD_COLORS = [
  { hex: '#C00000', name: 'Dark Red'    },
  { hex: '#FF0000', name: 'Red'         },
  { hex: '#FFC000', name: 'Orange'      },
  { hex: '#FFFF00', name: 'Yellow'      },
  { hex: '#92D050', name: 'Light Green' },
  { hex: '#00B050', name: 'Green'       },
  { hex: '#00B0F0', name: 'Light Blue'  },
  { hex: '#0070C0', name: 'Blue'        },
  { hex: '#002060', name: 'Dark Blue'   },
  { hex: '#7030A0', name: 'Purple'      },
];

const DEFAULT_PALETTES = {
  fill: [
    { hex: '#151F37', name: 'Navy'       },
    { hex: '#2A3E6D', name: 'Blue'       },
    { hex: '#D50032', name: 'Red'        },
    { hex: '#008579', name: 'Teal'       },
    { hex: '#FFFFFF', name: 'White'      },
    { hex: '#E2E3E2', name: 'Light Gray' },
  ],
  border: [
    { hex: '#151F37', name: 'Navy'  },
    { hex: '#2A3E6D', name: 'Blue'  },
    { hex: '#D50032', name: 'Red'   },
    { hex: '#000000', name: 'Black' },
    { hex: '#FFFFFF', name: 'White' },
  ],
  text: [
    { hex: '#151F37', name: 'Navy'  },
    { hex: '#2A3E6D', name: 'Blue'  },
    { hex: '#FFFFFF', name: 'White' },
    { hex: '#D50032', name: 'Red'   },
    { hex: '#000000', name: 'Black' },
  ],
};

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
  geometricShape: { icon: '◻', label: 'Geometric Shape' },
  textBox:        { icon: 'T',  label: 'Text Box'        },
  placeholder:    { icon: '⊞', label: 'Placeholder'     },
  callout:        { icon: '💬', label: 'Callout'         },
  freeform:       { icon: '✏', label: 'Freeform'        },
  group:          { icon: '⊡', label: 'Group'            },
  line:           { icon: '╱', label: 'Line'             },
  image:          { icon: '🖼', label: 'Image'            },
  table:          { icon: '⊟', label: 'Table'            },
};

const NO_FILL_LABELS = { fill: 'No Fill', border: 'No Border', text: 'No Color' };

const SHADE_NAMES = ['Light 1', 'Light 2', 'Light 3', 'Dark 1', 'Dark 2'];

// ─────────────────────────────────────────────────────────────────
// Color utilities
// ─────────────────────────────────────────────────────────────────

function hexToRgb(hex) {
  const h = hex.replace('#', '');
  return {
    r: parseInt(h.slice(0, 2), 16),
    g: parseInt(h.slice(2, 4), 16),
    b: parseInt(h.slice(4, 6), 16),
  };
}

function rgbToHex(r, g, b) {
  return '#' + [r, g, b]
    .map(x => Math.round(Math.min(255, Math.max(0, x))).toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
}

function isLight(hex) {
  const { r, g, b } = hexToRgb(hex);
  return (r * 299 + g * 587 + b * 114) / 1000 > 155;
}

function tintColor(hex, pct) {
  const { r, g, b } = hexToRgb(hex);
  return rgbToHex(r + (255 - r) * pct, g + (255 - g) * pct, b + (255 - b) * pct);
}

function shadeColor(hex, pct) {
  const { r, g, b } = hexToRgb(hex);
  return rgbToHex(r * (1 - pct), g * (1 - pct), b * (1 - pct));
}

// Generates 5 shade variants: 3 tints (light→base) + 2 shades (base→dark)
const SHADE_FNS = [
  h => tintColor(h, 0.80),
  h => tintColor(h, 0.55),
  h => tintColor(h, 0.30),
  h => shadeColor(h, 0.25),
  h => shadeColor(h, 0.50),
];

// ─────────────────────────────────────────────────────────────────
// Storage helpers
// ─────────────────────────────────────────────────────────────────

function getCustomColors(key) {
  try {
    const s = localStorage.getItem(`sct_custom_${key}`);
    return s ? JSON.parse(s) : [...DEFAULT_PALETTES[key]];
  } catch { return [...DEFAULT_PALETTES[key]]; }
}

function saveCustomColors(key, colors) {
  localStorage.setItem(`sct_custom_${key}`, JSON.stringify(colors));
}

function addCustomColor(key, hex, name) {
  const colors = getCustomColors(key);
  const norm = hex.toUpperCase();
  if (!colors.find(c => c.hex.toUpperCase() === norm)) {
    colors.push({ hex: norm, name: name || norm });
    saveCustomColors(key, colors);
  }
}

function deleteCustomColor(key, index) {
  const colors = getCustomColors(key);
  colors.splice(index, 1);
  saveCustomColors(key, colors);
}

function getRecentColors(key) {
  try {
    return JSON.parse(localStorage.getItem(`sct_recent_${key}`)) || [];
  } catch { return []; }
}

function trackRecent(key, hex) {
  if (!hex || hex === 'none') return;
  const norm = hex.toUpperCase();
  let arr = getRecentColors(key).filter(h => h !== norm);
  arr.unshift(norm);
  localStorage.setItem(`sct_recent_${key}`, JSON.stringify(arr.slice(0, 10)));
}

// ─────────────────────────────────────────────────────────────────
// Color picker renderer
// ─────────────────────────────────────────────────────────────────

function makeSwatch(hex, name, cls) {
  const light = isLight(hex) ? ' cp-light' : '';
  const safeName = (name || hex).replace(/"/g, '&quot;');
  return `<div class="cp-swatch ${cls}${light}" style="background:${hex}"
               data-color="${hex}" title="${safeName}&#10;${hex}" tabindex="0" role="button"></div>`;
}

function renderPicker(containerId, key, applyFn) {
  const container = document.getElementById(containerId);
  if (!container) return;

  const custom = getCustomColors(key);
  const recent = getRecentColors(key);

  // ── APi Group Colors: base row + 5 shade rows ──
  const baseRow = `<div class="cp-row">${BRAND_COLORS.map(c => makeSwatch(c.hex, c.name, 'cp-sm')).join('')}</div>`;
  const shadeRows = SHADE_FNS.map((fn, i) =>
    `<div class="cp-row">${BRAND_COLORS.map(c => makeSwatch(fn(c.hex), `${c.name} — ${SHADE_NAMES[i]}`, 'cp-sm')).join('')}</div>`
  ).join('');

  // ── Custom colors ──
  const customSwatches = custom.map((c, i) => `
    <div class="cp-custom-wrap" data-idx="${i}">
      <div class="cp-swatch cp-md${isLight(c.hex) ? ' cp-light' : ''}"
           style="background:${c.hex}" data-color="${c.hex}"
           title="${(c.name || c.hex).replace(/"/g, '&quot;')}&#10;${c.hex}"
           tabindex="0" role="button"></div>
      <button class="cp-del" data-key="${key}" data-idx="${i}" title="Remove this color" aria-label="Remove">×</button>
    </div>`).join('');

  // ── Standard colors ──
  const standardSwatches = STANDARD_COLORS.map(c => makeSwatch(c.hex, c.name, 'cp-md')).join('');

  // ── Recent colors ──
  const recentHtml = recent.length ? `
    <div class="cp-section-label">Recent Colors</div>
    <div class="cp-row cp-recent-row">${recent.map(h => makeSwatch(h, h, 'cp-md')).join('')}</div>` : '';

  container.innerHTML = `
    <div class="cp-section-label">APi Group Colors</div>
    <div class="cp-brand-grid">
      ${baseRow}
      <div class="cp-shade-divider"></div>
      ${shadeRows}
    </div>

    <div class="cp-section-label">
      Custom Colors
      <span class="cp-custom-hint">hover to remove</span>
    </div>
    <div class="cp-row cp-custom-row" id="cr-${key}">
      ${customSwatches}
      <button class="cp-add-toggle" data-key="${key}" aria-label="Add custom color" title="Add a color">
        <svg width="11" height="11" viewBox="0 0 11 11" fill="none">
          <line x1="5.5" y1="1.5" x2="5.5" y2="9.5" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/>
          <line x1="1.5" y1="5.5" x2="9.5" y2="5.5" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/>
        </svg>
      </button>
    </div>

    <div class="cp-add-row hidden" id="car-${key}">
      <input type="color" class="cp-color-input" id="cpi-${key}" value="#2E75B6" />
      <input type="text"  class="cp-name-input"  id="cpn-${key}" placeholder="Color name (optional)" maxlength="24" />
      <button class="cp-confirm-btn" data-key="${key}">Add</button>
    </div>

    <div class="cp-section-label">Standard Colors</div>
    <div class="cp-row">${standardSwatches}</div>

    ${recentHtml}

    <button class="cp-no-fill-btn" data-key="${key}">
      <svg width="13" height="13" viewBox="0 0 13 13" fill="none">
        <rect x="1" y="1" width="11" height="11" rx="2" stroke="currentColor" stroke-width="1.2"/>
        <line x1="2" y1="11" x2="11" y2="2" stroke="currentColor" stroke-width="1.2"/>
      </svg>
      ${NO_FILL_LABELS[key] || 'No Color'}
    </button>
  `;

  // ── Bind events ──

  // All color swatches
  container.querySelectorAll('.cp-swatch[data-color]').forEach(s => {
    const handler = () => {
      const color = s.dataset.color;
      applyFn(color);
      trackRecent(key, color);
      renderPicker(containerId, key, applyFn);
    };
    s.addEventListener('click', handler);
    s.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); handler(); } });
  });

  // Delete custom color
  container.querySelectorAll('.cp-del').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      deleteCustomColor(btn.dataset.key, parseInt(btn.dataset.idx, 10));
      renderPicker(containerId, key, applyFn);
    });
  });

  // Toggle inline add row
  const toggleBtn = container.querySelector('.cp-add-toggle');
  if (toggleBtn) {
    toggleBtn.addEventListener('click', () => {
      const row = document.getElementById(`car-${key}`);
      const opening = row.classList.contains('hidden');
      row.classList.toggle('hidden');
      toggleBtn.classList.toggle('active', opening);
      if (opening) {
        setTimeout(() => document.getElementById(`cpn-${key}`)?.focus(), 50);
      }
    });
  }

  // Confirm add
  const confirmBtn = container.querySelector('.cp-confirm-btn');
  if (confirmBtn) {
    confirmBtn.addEventListener('click', () => {
      const hexVal  = document.getElementById(`cpi-${key}`).value;
      const nameVal = document.getElementById(`cpn-${key}`).value.trim();
      addCustomColor(key, hexVal, nameVal || hexVal);
      renderPicker(containerId, key, applyFn);
    });
  }

  // Enter key in name field
  const nameInput = document.getElementById(`cpn-${key}`);
  if (nameInput) {
    nameInput.addEventListener('keydown', e => {
      if (e.key === 'Enter') container.querySelector('.cp-confirm-btn')?.click();
    });
  }

  // No fill
  const noFillBtn = container.querySelector('.cp-no-fill-btn');
  if (noFillBtn) {
    noFillBtn.addEventListener('click', () => applyFn('none'));
  }
}

// ─────────────────────────────────────────────────────────────────
// Office ready
// ─────────────────────────────────────────────────────────────────

Office.onReady(() => {
  // Render all three pickers immediately
  renderPicker('picker-fill',   'fill',   applyFill);
  renderPicker('picker-border', 'border', applyBorderColor);
  renderPicker('picker-text',   'text',   applyTextColor);

  // Border weight chips
  document.querySelectorAll('[data-weight]').forEach(btn =>
    btn.addEventListener('click', () => applyBorderWeight(parseFloat(btn.dataset.weight))));

  // Border dash style chips
  document.querySelectorAll('[data-dash]').forEach(btn =>
    btn.addEventListener('click', () => applyBorderDash(btn.dataset.dash)));

  // Line ends
  document.getElementById('btn-apply-ends')
    ?.addEventListener('click', applyLineEnds);

  // Selection change
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    () => inspectSelection()
  );

  inspectSelection();
});

// ─────────────────────────────────────────────────────────────────
// Shape inspection
// ─────────────────────────────────────────────────────────────────

async function inspectSelection() {
  try {
    await PowerPoint.run(async ctx => {
      const sel = ctx.presentation.getSelectedShapes();
      sel.load('items/type,items/name');
      await ctx.sync();

      const items = sel.items;
      if (!items.length) { renderEmpty(); return; }

      const merged   = { fill: false, border: false, text: false, lineEnds: false };
      let   allSame  = true;
      const firstKey = toKey(items[0].type);
      const firstName = items[0].name || '';

      items.forEach(s => {
        const k = toKey(s.type);
        if (k !== firstKey) allSame = false;
        const caps = CAPABILITIES[k] || {};
        if (caps.fill)     merged.fill     = true;
        if (caps.border)   merged.border   = true;
        if (caps.text)     merged.text     = true;
        if (caps.lineEnds) merged.lineEnds = true;
      });

      const meta = allSame
        ? (TYPE_META[firstKey] || { icon: '?', label: cap(firstKey || 'Shape') })
        : { icon: '⊕', label: 'Mixed Selection' };

      const anySupported = Object.values(merged).some(Boolean);
      renderUI({ merged, meta, firstName, count: items.length, anySupported,
                 isLine: allSame && firstKey === 'line' });
    });
  } catch (_) { renderEmpty(); }
}

function toKey(raw) {
  if (!raw) return 'unsupported';
  const s = String(raw);
  return s.charAt(0).toLowerCase() + s.slice(1);
}

function cap(s) { return s.charAt(0).toUpperCase() + s.slice(1); }

// ─────────────────────────────────────────────────────────────────
// UI helpers
// ─────────────────────────────────────────────────────────────────

function renderUI({ merged, meta, firstName, count, anySupported, isLine }) {
  hide('empty-state');
  hide('unsupported-state');
  show('shape-banner');

  el('shape-icon').textContent       = meta.icon;
  el('shape-type-label').textContent = meta.label;
  el('shape-name').textContent       = firstName ? `"${firstName}"` : '';

  const badge = el('shape-count-badge');
  badge.textContent = `×${count}`;
  tog('shape-count-badge', count > 1);

  tog('section-fill',      merged.fill);
  tog('section-border',    merged.border);
  tog('section-line-ends', merged.lineEnds);
  tog('section-text',      merged.text);

  el('border-heading').textContent = isLine ? 'Line Style' : 'Border';

  if (!anySupported) show('unsupported-state');
}

function renderEmpty() {
  show('empty-state');
  hide('shape-banner');
  hide('unsupported-state');
  ['section-fill','section-border','section-line-ends','section-text'].forEach(hide);
  setStatus('');
}

function el(id)          { return document.getElementById(id); }
function show(id)        { el(id)?.classList.remove('hidden'); }
function hide(id)        { el(id)?.classList.add('hidden'); }
function tog(id, vis)    { el(id)?.classList.toggle('hidden', !vis); }
function setStatus(msg)  { el('status').textContent = msg; }

// ─────────────────────────────────────────────────────────────────
// Apply functions
// ─────────────────────────────────────────────────────────────────

async function getSelected(ctx) {
  const col = ctx.presentation.getSelectedShapes();
  col.load('items');
  await ctx.sync();
  return col.items;
}

async function applyFill(color) {
  await PowerPoint.run(async ctx => {
    const shapes = await getSelected(ctx);
    if (!shapes.length) return setStatus('No shape selected.');
    shapes.forEach(s => {
      if (color === 'none') s.fill.clear();
      else s.fill.setSolidColor(color);
    });
    await ctx.sync();
    setStatus(`Fill → ${color === 'none' ? 'cleared' : color}`);
  });
}

async function applyBorderColor(color) {
  await PowerPoint.run(async ctx => {
    const shapes = await getSelected(ctx);
    if (!shapes.length) return setStatus('No shape selected.');
    shapes.forEach(s => {
      if (color === 'none') { s.lineFormat.visible = false; }
      else { s.lineFormat.visible = true; s.lineFormat.color = color; }
    });
    await ctx.sync();
    setStatus(`Border → ${color === 'none' ? 'removed' : color}`);
  });
}

async function applyBorderWeight(pts) {
  await PowerPoint.run(async ctx => {
    const shapes = await getSelected(ctx);
    if (!shapes.length) return setStatus('No shape selected.');
    shapes.forEach(s => { s.lineFormat.weight = pts; s.lineFormat.visible = true; });
    await ctx.sync();
    setStatus(`Weight → ${pts}pt`);
  });
}

async function applyBorderDash(style) {
  const map = {
    solid: PowerPoint.LineDashStyle.solid,
    dash:  PowerPoint.LineDashStyle.dash,
    dot:   PowerPoint.LineDashStyle.dot,
  };
  await PowerPoint.run(async ctx => {
    const shapes = await getSelected(ctx);
    if (!shapes.length) return setStatus('No shape selected.');
    shapes.forEach(s => {
      s.lineFormat.dashStyle = map[style] ?? PowerPoint.LineDashStyle.solid;
      s.lineFormat.visible = true;
    });
    await ctx.sync();
    setStatus(`Style → ${style}`);
  });
}

async function applyTextColor(color) {
  await PowerPoint.run(async ctx => {
    const shapes = await getSelected(ctx);
    if (!shapes.length) return setStatus('No shape selected.');
    shapes.forEach(s => {
      try { s.textFrame.textRange.font.color = color; } catch (_) {}
    });
    await ctx.sync();
    setStatus(`Text → ${color}`);
  });
}

async function applyLineEnds() {
  const arrowMap = {
    none:      PowerPoint.ArrowheadStyle.none,
    arrow:     PowerPoint.ArrowheadStyle.arrow,
    openArrow: PowerPoint.ArrowheadStyle.openArrow,
    diamond:   PowerPoint.ArrowheadStyle.diamond,
    oval:      PowerPoint.ArrowheadStyle.oval,
  };
  const startVal = el('line-start').value;
  const endVal   = el('line-end').value;

  await PowerPoint.run(async ctx => {
    const shapes = await getSelected(ctx);
    if (!shapes.length) return setStatus('No shape selected.');
    shapes.forEach(s => {
      if (toKey(s.type) !== 'line') return;
      try {
        s.lineFormat.beginArrowheadStyle = arrowMap[startVal] ?? PowerPoint.ArrowheadStyle.none;
        s.lineFormat.endArrowheadStyle   = arrowMap[endVal]   ?? PowerPoint.ArrowheadStyle.none;
      } catch (_) {}
    });
    await ctx.sync();
    setStatus(`Ends → ${startVal} / ${endVal}`);
  });
}
