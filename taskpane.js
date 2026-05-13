'use strict';

// ─────────────────────────────────────────────────────────────────
// APi Group brand colors (theme columns)
// ─────────────────────────────────────────────────────────────────
const BRAND_COLORS = [
  { hex: '#151F37', name: 'Navy'       },
  { hex: '#2A3E6D', name: 'Blue'       },
  { hex: '#D50032', name: 'Red'        },
  { hex: '#008579', name: 'Teal'       },
  { hex: '#E2E3E2', name: 'Light Gray' },
  { hex: '#FFFFFF', name: 'White'      },
  { hex: '#000000', name: 'Black'      },
];

// Shade rows beneath base: [lighter 80%, 60%, 40%, darker 25%, 50%]
const SHADE_STEPS = [
  { fn: h => tint(h, 0.80),  label: 'Lighter 80%' },
  { fn: h => tint(h, 0.60),  label: 'Lighter 60%' },
  { fn: h => tint(h, 0.40),  label: 'Lighter 40%' },
  { fn: h => darken(h, 0.25),label: 'Darker 25%'  },
  { fn: h => darken(h, 0.50),label: 'Darker 50%'  },
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
    { hex: '#000000', name: 'Black'      },
  ],
  border: [
    { hex: '#151F37', name: 'Navy'       },
    { hex: '#2A3E6D', name: 'Blue'       },
    { hex: '#D50032', name: 'Red'        },
    { hex: '#000000', name: 'Black'      },
    { hex: '#FFFFFF', name: 'White'      },
    { hex: '#E2E3E2', name: 'Light Gray' },
    { hex: '#008579', name: 'Teal'       },
  ],
  text: [
    { hex: '#151F37', name: 'Navy'       },
    { hex: '#2A3E6D', name: 'Blue'       },
    { hex: '#FFFFFF', name: 'White'      },
    { hex: '#D50032', name: 'Red'        },
    { hex: '#000000', name: 'Black'      },
    { hex: '#008579', name: 'Teal'       },
    { hex: '#E2E3E2', name: 'Light Gray' },
  ],
};

const NO_LABEL = { fill: 'No Fill', border: 'No Border', text: 'No Color' };

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

function tint(hex, pct) {
  const { r, g, b } = hexToRgb(hex);
  return rgbToHex(r + (255 - r) * pct, g + (255 - g) * pct, b + (255 - b) * pct);
}

function darken(hex, pct) {
  const { r, g, b } = hexToRgb(hex);
  return rgbToHex(r * (1 - pct), g * (1 - pct), b * (1 - pct));
}

function isLight(hex) {
  const { r, g, b } = hexToRgb(hex);
  return (r * 299 + g * 587 + b * 114) / 1000 > 155;
}

function normalizeHex(raw) {
  const h = (raw || '').replace('#', '').trim();
  return /^[0-9A-Fa-f]{6}$/.test(h) ? '#' + h.toUpperCase() : null;
}

function esc(s) {
  return (s || '').replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// ─────────────────────────────────────────────────────────────────
// Storage
// ─────────────────────────────────────────────────────────────────

function getCustom(key) {
  try {
    const s = localStorage.getItem(`sct_${key}`);
    return s ? JSON.parse(s) : [...DEFAULT_PALETTES[key]];
  } catch { return [...DEFAULT_PALETTES[key]]; }
}

function saveCustom(key, arr) {
  localStorage.setItem(`sct_${key}`, JSON.stringify(arr));
}

function addCustom(key, hex, name) {
  const norm = normalizeHex(hex);
  if (!norm) return;
  const arr = getCustom(key);
  if (!arr.find(c => c.hex.toUpperCase() === norm)) {
    arr.push({ hex: norm, name: name || norm });
    saveCustom(key, arr);
  }
}

function delCustom(key, idx) {
  const arr = getCustom(key);
  arr.splice(idx, 1);
  saveCustom(key, arr);
}

function getRecent(key) {
  try { return JSON.parse(localStorage.getItem(`sct_recent_${key}`)) || []; }
  catch { return []; }
}

function pushRecent(key, hex) {
  if (!hex || hex === 'none') return;
  const norm = normalizeHex(hex) || hex.toUpperCase();
  let arr = getRecent(key).filter(h => h !== norm);
  arr.unshift(norm);
  localStorage.setItem(`sct_recent_${key}`, JSON.stringify(arr.slice(0, 10)));
}

// ─────────────────────────────────────────────────────────────────
// Edit panel state
// ─────────────────────────────────────────────────────────────────

let _editKey = null;
let _applyFn = null;

// ─────────────────────────────────────────────────────────────────
// Main panel — render 7 swatches + pencil edit button per section
// ─────────────────────────────────────────────────────────────────

function renderMainSwatches(key) {
  const row = document.getElementById(`swatches-${key}`);
  if (!row) return;

  const applyFnMap = { fill: applyFill, border: applyBorderColor, text: applyTextColor };
  const fn = applyFnMap[key];
  const colors = getCustom(key);

  row.innerHTML =
    colors.map(c => `
      <button class="main-swatch${isLight(c.hex) ? ' light' : ''}"
              style="background:${c.hex}"
              data-color="${c.hex}"
              title="${esc(c.name)}\n${c.hex}"
              aria-label="${esc(c.name)}">
      </button>`).join('') +
    `<button class="edit-palette-btn" data-key="${key}"
             title="Edit ${key} palette" aria-label="Edit ${key} colors">
       <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
         <path d="M8.5 1.5a1.414 1.414 0 012 2L4 10H2V8L8.5 1.5z"
               stroke="currentColor" stroke-width="1.3"
               stroke-linecap="round" stroke-linejoin="round" fill="none"/>
       </svg>
     </button>`;

  // Apply color on swatch click
  row.querySelectorAll('.main-swatch').forEach(btn => {
    btn.addEventListener('click', () => {
      fn(btn.dataset.color);
      pushRecent(key, btn.dataset.color);
    });
  });

  // Open edit panel
  row.querySelector('.edit-palette-btn').addEventListener('click', () => {
    openEditPanel(key, fn);
  });
}

// ─────────────────────────────────────────────────────────────────
// Edit panel — open / close / render
// ─────────────────────────────────────────────────────────────────

function openEditPanel(key, fn) {
  _editKey = key;
  _applyFn = fn;
  renderEditPanel();
  document.getElementById('main-content').classList.add('hidden');
  document.getElementById('edit-panel').classList.remove('hidden');
}

function closeEditPanel() {
  document.getElementById('edit-panel').classList.add('hidden');
  document.getElementById('main-content').classList.remove('hidden');
  renderMainSwatches(_editKey); // refresh main swatches in case palette changed
  _editKey = null;
  _applyFn = null;
}

function renderEditPanel() {
  const key = _editKey;
  const titleMap = { fill: 'Fill Color', border: 'Border Color', text: 'Text Color' };
  document.getElementById('edit-panel-title').textContent = titleMap[key] || 'Edit Colors';

  // ── APi Group theme colors: base row + 5 shade rows ──
  const baseRowHtml = BRAND_COLORS.map(c =>
    `<div class="theme-swatch base-swatch${isLight(c.hex) ? ' light' : ''}"
          style="background:${c.hex}" data-color="${c.hex}"
          title="${esc(c.name)}" role="button" tabindex="0"></div>`
  ).join('');

  const shadeRowsHtml = SHADE_STEPS.map(step =>
    `<div class="theme-row">${BRAND_COLORS.map(c => {
      const hex = step.fn(c.hex);
      return `<div class="theme-swatch${isLight(hex) ? ' light' : ''}"
                   style="background:${hex}" data-color="${hex}"
                   title="${esc(c.name)} — ${step.label}" role="button" tabindex="0"></div>`;
    }).join('')}</div>`
  ).join('');

  // ── Standard colors ──
  const standardHtml = STANDARD_COLORS.map(c =>
    `<div class="std-swatch${isLight(c.hex) ? ' light' : ''}"
          style="background:${c.hex}" data-color="${c.hex}"
          title="${esc(c.name)}" role="button" tabindex="0"></div>`
  ).join('');

  // ── Custom colors ──
  const customColors = getCustom(key);
  const customHtml = customColors.map((c, i) =>
    `<div class="custom-wrap">
       <div class="custom-swatch${isLight(c.hex) ? ' light' : ''}"
            style="background:${c.hex}" data-color="${c.hex}"
            title="${esc(c.name)}\n${c.hex}" role="button" tabindex="0"></div>
       <button class="del-btn" data-idx="${i}" aria-label="Remove ${esc(c.name)}">×</button>
     </div>`
  ).join('');

  // ── Recent colors ──
  const recent = getRecent(key);
  const recentHtml = recent.length ? `
    <div class="ep-label">Recent Colors</div>
    <div class="std-row recent-row">${recent.map(hex =>
      `<div class="std-swatch${isLight(hex) ? ' light' : ''}"
            style="background:${hex}" data-color="${hex}"
            title="${hex}" role="button" tabindex="0"></div>`
    ).join('')}</div>` : '';

  document.getElementById('edit-panel-body').innerHTML = `

    <!-- Theme colors -->
    <div class="ep-label">APi Group Theme Colors</div>
    <div class="theme-grid">
      <div class="theme-row base-row">${baseRowHtml}</div>
      ${shadeRowsHtml}
    </div>

    <!-- Standard colors -->
    <div class="ep-label">Standard Colors</div>
    <div class="std-row">${standardHtml}</div>

    <!-- Recent colors -->
    ${recentHtml}

    <!-- Custom colors -->
    <div class="ep-label">
      My Colors
      <span class="ep-hint">hover to remove</span>
    </div>
    <div class="custom-row">
      ${customHtml}
      <button class="add-toggle-btn" id="ep-add-toggle"
              aria-label="Add a custom color" title="Add color">
        <svg width="11" height="11" viewBox="0 0 11 11" fill="none">
          <line x1="5.5" y1="1" x2="5.5" y2="10" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/>
          <line x1="1" y1="5.5" x2="10" y2="5.5" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/>
        </svg>
      </button>
    </div>

    <!-- Inline add form -->
    <div class="add-form hidden" id="ep-add-form">
      <div class="add-form-inputs">
        <input type="color" id="ep-wheel" class="ep-wheel" value="#2A3E6D" />
        <div class="hex-field">
          <span class="hex-hash">#</span>
          <input type="text" id="ep-hex" class="ep-hex-input"
                 value="2A3E6D" maxlength="6" placeholder="RRGGBB" spellcheck="false" />
        </div>
        <div class="ep-preview-swatch" id="ep-preview" style="background:#2A3E6D"></div>
      </div>
      <input type="text" id="ep-name" class="ep-name-input"
             placeholder="Color name (optional)" maxlength="28" />
      <button class="ep-add-btn" id="ep-confirm-add">+ Add to My Colors</button>
    </div>

    <!-- More colors picker -->
    <div class="ep-label" style="margin-top:14px">More Colors</div>
    <div class="more-colors-box">
      <div class="mc-top">
        <input type="color" id="mc-wheel" class="mc-wheel" value="#2A3E6D" />
        <div class="mc-right">
          <div class="mc-preview" id="mc-preview" style="background:#2A3E6D"></div>
          <div class="hex-field">
            <span class="hex-hash">#</span>
            <input type="text" id="mc-hex" class="ep-hex-input"
                   value="2A3E6D" maxlength="6" placeholder="RRGGBB" spellcheck="false" />
          </div>
        </div>
      </div>
      <button class="mc-apply-btn" id="mc-apply">Apply This Color</button>
    </div>

    <!-- No color -->
    <button class="no-color-btn" id="ep-no-color">
      <svg width="13" height="13" viewBox="0 0 13 13" fill="none">
        <rect x="1" y="1" width="11" height="11" rx="2" stroke="currentColor" stroke-width="1.2"/>
        <line x1="2" y1="11" x2="11" y2="2" stroke="currentColor" stroke-width="1.2"/>
      </svg>
      ${NO_LABEL[key] || 'No Color'}
    </button>
  `;

  bindEditEvents(key);
}

function bindEditEvents(key) {
  const body = document.getElementById('edit-panel-body');

  // ── Theme + standard + recent: click to apply and close ──
  body.querySelectorAll('.theme-swatch, .std-swatch').forEach(s => {
    const handler = () => applyAndClose(s.dataset.color);
    s.addEventListener('click', handler);
    s.addEventListener('keydown', e => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); handler(); }
    });
  });

  // ── Custom swatches: click to apply and close ──
  body.querySelectorAll('.custom-swatch').forEach(s => {
    s.addEventListener('click', () => applyAndClose(s.dataset.color));
  });

  // ── Delete custom color ──
  body.querySelectorAll('.del-btn').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      delCustom(key, parseInt(btn.dataset.idx, 10));
      renderEditPanel();
    });
  });

  // ── Toggle inline add form ──
  document.getElementById('ep-add-toggle').addEventListener('click', () => {
    const form    = document.getElementById('ep-add-form');
    const opening = form.classList.contains('hidden');
    form.classList.toggle('hidden');
    document.getElementById('ep-add-toggle').classList.toggle('active', opening);
    if (opening) setTimeout(() => document.getElementById('ep-hex')?.focus(), 40);
  });

  // ── Add form: sync wheel ↔ hex ↔ preview ──
  const wheel   = document.getElementById('ep-wheel');
  const hexInp  = document.getElementById('ep-hex');
  const preview = document.getElementById('ep-preview');

  wheel.addEventListener('input', () => {
    hexInp.value = wheel.value.replace('#', '').toUpperCase();
    preview.style.background = wheel.value;
  });
  hexInp.addEventListener('input', () => {
    const norm = normalizeHex(hexInp.value);
    if (norm) { wheel.value = norm; preview.style.background = norm; }
  });

  // ── Confirm add to My Colors ──
  document.getElementById('ep-confirm-add').addEventListener('click', () => {
    const norm = normalizeHex(hexInp.value);
    if (!norm) return;
    const name = document.getElementById('ep-name').value.trim();
    addCustom(key, norm, name || norm);
    document.getElementById('ep-name').value = '';
    document.getElementById('ep-add-form').classList.add('hidden');
    document.getElementById('ep-add-toggle').classList.remove('active');
    renderEditPanel();
  });
  document.getElementById('ep-name').addEventListener('keydown', e => {
    if (e.key === 'Enter') document.getElementById('ep-confirm-add').click();
  });

  // ── More colors: sync wheel ↔ hex ↔ preview ──
  const mcWheel   = document.getElementById('mc-wheel');
  const mcHex     = document.getElementById('mc-hex');
  const mcPreview = document.getElementById('mc-preview');

  mcWheel.addEventListener('input', () => {
    mcHex.value = mcWheel.value.replace('#', '').toUpperCase();
    mcPreview.style.background = mcWheel.value;
  });
  mcHex.addEventListener('input', () => {
    const norm = normalizeHex(mcHex.value);
    if (norm) { mcWheel.value = norm; mcPreview.style.background = norm; }
  });
  document.getElementById('mc-apply').addEventListener('click', () => {
    const norm = normalizeHex(mcHex.value);
    if (norm) applyAndClose(norm);
  });

  // ── No color / no fill ──
  document.getElementById('ep-no-color').addEventListener('click', () => applyAndClose('none'));
}

// Apply color, track recent, return to main panel
function applyAndClose(color) {
  if (_applyFn) _applyFn(color);
  if (color !== 'none') pushRecent(_editKey, color);
  closeEditPanel();
}

// ─────────────────────────────────────────────────────────────────
// Office ready
// ─────────────────────────────────────────────────────────────────

Office.onReady(() => {
  // Render all three main-panel swatch rows
  renderMainSwatches('fill');
  renderMainSwatches('border');
  renderMainSwatches('text');

  // Border weight chips
  document.querySelectorAll('[data-weight]').forEach(btn =>
    btn.addEventListener('click', () => applyBorderWeight(parseFloat(btn.dataset.weight))));

  // Border dash chips
  document.querySelectorAll('[data-dash]').forEach(btn =>
    btn.addEventListener('click', () => applyBorderDash(btn.dataset.dash)));

  // Line ends
  document.getElementById('btn-apply-ends')
    ?.addEventListener('click', applyLineEnds);

  // Edit panel back button
  document.getElementById('edit-back-btn')
    .addEventListener('click', closeEditPanel);

  // Selection listener — skip while edit panel is open
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    () => { if (!_editKey) inspectSelection(); }
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

      const merged    = { fill: false, border: false, text: false, lineEnds: false };
      let   allSame   = true;
      const firstKey  = toKey(items[0].type);
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

      const anySupported = Object.values(merged).some(Boolean);
      const meta = allSame
        ? (TYPE_META[firstKey] || { icon: '?', label: cap(firstKey || 'Shape') })
        : { icon: '⊕', label: 'Mixed Selection' };

      renderUI({ merged, meta, firstName, count: items.length,
                 anySupported, isLine: allSame && firstKey === 'line' });
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

  el('shape-count-badge').textContent = `×${count}`;
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
  ['section-fill', 'section-border', 'section-line-ends', 'section-text'].forEach(hide);
  setStatus('');
}

function el(id)         { return document.getElementById(id); }
function show(id)       { el(id)?.classList.remove('hidden'); }
function hide(id)       { el(id)?.classList.add('hidden'); }
function tog(id, vis)   { el(id)?.classList.toggle('hidden', !vis); }
function setStatus(msg) { el('status').textContent = msg; }

// ─────────────────────────────────────────────────────────────────
// Shared: get selected shapes
// ─────────────────────────────────────────────────────────────────

async function getSelected(ctx) {
  const col = ctx.presentation.getSelectedShapes();
  col.load('items');
  await ctx.sync();
  return col.items;
}

// ─────────────────────────────────────────────────────────────────
// Apply functions
// ─────────────────────────────────────────────────────────────────

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
