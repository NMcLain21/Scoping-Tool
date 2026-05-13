'use strict';

// ─────────────────────────────────────────────────────────────────
// Brand / static color data
// ─────────────────────────────────────────────────────────────────
const BRAND_COLORS = [
  { hex: '#D50032', name: 'APi Corp Red'  },
  { hex: '#2A3E6D', name: 'Dark Navy'     },
  { hex: '#008579', name: 'Teal'          },
  { hex: '#D4D800', name: 'TBD'           },
  { hex: '#7030A0', name: 'APi Seg'       },
  { hex: '#00B0F0', name: 'Target HQ'    },
  { hex: '#92D050', name: 'Target OpCo'  },
];

const SHADE_STEPS = [
  { fn: h => tint(h, 0.80),   label: 'Lighter 80%' },
  { fn: h => tint(h, 0.60),   label: 'Lighter 60%' },
  { fn: h => tint(h, 0.40),   label: 'Lighter 40%' },
  { fn: h => darken(h, 0.25), label: 'Darker 25%'  },
  { fn: h => darken(h, 0.50), label: 'Darker 50%'  },
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
    { hex: '#D50032', name: 'Target Converts to Acquiring' },
    { hex: '#2A3E6D', name: 'APi Corp Entity'               },
    { hex: '#008579', name: 'Shared Services'               },
    { hex: '#595959', name: 'Leave As Is'                   },
    { hex: '#BFBFBF', name: 'To Be Determined'              },
    { hex: '#FFFFFF', name: 'No Fill'                       },
    { hex: '#000000', name: 'Divest / Sunset'               },
  ],
  border: [
    { hex: '#D4D800', name: 'TBD'                    },
    { hex: '#D50032', name: 'APi Corp'               },
    { hex: '#7030A0', name: 'APi Segment'            },
    { hex: '#00B0F0', name: 'Target HQ'              },
    { hex: '#92D050', name: 'Target OpCo'            },
    { hex: '#808080', name: 'No Reporting Structure' },
    { hex: '#151F37', name: 'Acquiring Entity'       },
  ],
  text: [
    { hex: '#FFFFFF', name: 'White'        },
    { hex: '#000000', name: 'Black'        },
    { hex: '#D50032', name: 'APi Corp Red' },
    { hex: '#2A3E6D', name: 'Dark Navy'    },
    { hex: '#008579', name: 'Teal'         },
    { hex: '#595959', name: 'Dark Gray'    },
    { hex: '#BFBFBF', name: 'Light Gray'   },
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
    .join('').toUpperCase();
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
// HSV ↔ RGB conversions
// ─────────────────────────────────────────────────────────────────
function hsvToRgb(h, s, v) {
  const c = v * s;
  const x = c * (1 - Math.abs((h / 60) % 2 - 1));
  const m = v - c;
  let r = 0, g = 0, b = 0;
  if      (h < 60)  { r = c; g = x; b = 0; }
  else if (h < 120) { r = x; g = c; b = 0; }
  else if (h < 180) { r = 0; g = c; b = x; }
  else if (h < 240) { r = 0; g = x; b = c; }
  else if (h < 300) { r = x; g = 0; b = c; }
  else              { r = c; g = 0; b = x; }
  return {
    r: Math.round((r + m) * 255),
    g: Math.round((g + m) * 255),
    b: Math.round((b + m) * 255),
  };
}

function rgbToHsv(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b), d = max - min;
  let h = 0;
  const s = max === 0 ? 0 : d / max;
  const v = max;
  if (d !== 0) {
    if      (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) * 60;
    else if (max === g) h = ((b - r) / d + 2) * 60;
    else                h = ((r - g) / d + 4) * 60;
  }
  return { h: h < 0 ? h + 360 : h, s, v };
}

function hexToHsv(hex) {
  const { r, g, b } = hexToRgb(hex);
  return rgbToHsv(r, g, b);
}

function hsvToHex(h, s, v) {
  const { r, g, b } = hsvToRgb(h, s, v);
  return rgbToHex(r, g, b);
}

// ─────────────────────────────────────────────────────────────────
// Storage
// ─────────────────────────────────────────────────────────────────
function getPalette(key) {
  try {
    const s = localStorage.getItem(`sct_${key}`);
    return s ? JSON.parse(s) : JSON.parse(JSON.stringify(DEFAULT_PALETTES[key]));
  } catch { return JSON.parse(JSON.stringify(DEFAULT_PALETTES[key])); }
}

function savePalette(key, arr) {
  localStorage.setItem(`sct_${key}`, JSON.stringify(arr));
}

function updateColor(key, idx, hex, name) {
  const arr = getPalette(key);
  if (arr[idx]) { arr[idx] = { hex, name: name || hex }; savePalette(key, arr); }
}

function addColor(key, hex, name) {
  const norm = normalizeHex(hex);
  if (!norm) return;
  const arr = getPalette(key);
  arr.push({ hex: norm, name: name || norm });
  savePalette(key, arr);
}

function deleteColor(key, idx) {
  const arr = getPalette(key);
  arr.splice(idx, 1);
  savePalette(key, arr);
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
let _editKey  = null;
let _applyFn  = null;
let _openForm = null; // { type: 'edit'|'add', idx: number|null }

// ─────────────────────────────────────────────────────────────────
// Performance state — cached from inspectSelection so apply
// functions skip the first load+sync round-trip entirely
// ─────────────────────────────────────────────────────────────────
let _cachedShapeCount = 0; // number of currently selected shapes
let _applyVer         = 0; // version counter — drops superseded clicks

// ─────────────────────────────────────────────────────────────────
// Main panel — 7 swatches + pencil button
// ─────────────────────────────────────────────────────────────────
function renderMainSwatches(key) {
  const row = document.getElementById(`swatches-${key}`);
  if (!row) return;

  const fnMap = { fill: applyFill, border: applyBorderColor, text: applyTextColor };
  const fn    = fnMap[key];
  const colors = getPalette(key);

  row.innerHTML =
    colors.map(c => {
      // Border palette — legend-style line tile, not a solid fill box
      if (key === 'border') {
        return `<button class="main-swatch border-legend-tile"
                 data-color="${c.hex}"
                 title="${esc(c.name)}\n${c.hex}"
                 aria-label="${esc(c.name)}">
           <span class="legend-line-sample" style="background:${c.hex}"></span>
           <span class="legend-line-name">${esc(c.name)}</span>
         </button>`;
      }
      // Fill / text — solid color tile with label
      return `<button class="main-swatch${isLight(c.hex) ? ' light' : ''}"
               style="background:${c.hex}"
               data-color="${c.hex}"
               title="${esc(c.name)}\n${c.hex}"
               aria-label="${esc(c.name)}">
         <span class="swatch-label${isLight(c.hex) ? ' dark-text' : ''}">${esc(c.name)}</span>
       </button>`;
    }).join('') +
    `<button class="edit-palette-btn" data-key="${key}"
             title="Edit ${key} palette" aria-label="Edit ${key} colors">
       <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
         <path d="M8.5 1.5a1.414 1.414 0 012 2L4 10H2V8L8.5 1.5z"
               stroke="currentColor" stroke-width="1.3"
               stroke-linecap="round" stroke-linejoin="round" fill="none"/>
       </svg>
     </button>`;

  row.querySelectorAll('.main-swatch').forEach(btn => {
    btn.addEventListener('click', () => {
      fn(btn.dataset.color);
      pushRecent(key, btn.dataset.color);
    });
  });

  row.querySelector('.edit-palette-btn').addEventListener('click', () => {
    openEditPanel(key, fn);
  });
}

// ─────────────────────────────────────────────────────────────────
// Edit panel — open / close
// ─────────────────────────────────────────────────────────────────
function openEditPanel(key, fn) {
  _editKey  = key;
  _applyFn  = fn;
  _openForm = null;
  renderEditPanel();
  document.getElementById('main-content').classList.add('hidden');
  document.getElementById('edit-panel').classList.remove('hidden');
}

function closeEditPanel() {
  document.getElementById('edit-panel').classList.add('hidden');
  document.getElementById('main-content').classList.remove('hidden');
  renderMainSwatches(_editKey);
  _editKey  = null;
  _applyFn  = null;
  _openForm = null;
}

// ─────────────────────────────────────────────────────────────────
// Edit panel — full render
// ─────────────────────────────────────────────────────────────────
function renderEditPanel() {
  const key = _editKey;
  const titleMap = { fill: 'Fill Color', border: 'Border Color', text: 'Text Color' };
  document.getElementById('edit-panel-title').textContent = titleMap[key] || 'Edit Colors';

  // My Colors list
  const palette = getPalette(key);
  const myColorsHtml = palette.map((c, i) => {
    const isEditing = _openForm && _openForm.type === 'edit' && _openForm.idx === i;
    return `
      <div class="my-color-item" data-idx="${i}">
        <div class="my-color-top">
          <div class="my-swatch-wrap">
            <div class="my-swatch${isLight(c.hex) ? ' light' : ''}"
                 style="background:${c.hex}"
                 title="${esc(c.name)}\n${c.hex}"
                 data-color="${c.hex}"
                 role="button" tabindex="0"
                 aria-label="Apply ${esc(c.name)}"></div>
          </div>
          <div class="my-color-info">
            <span class="my-color-name">${esc(c.name)}</span>
            <span class="my-color-hex">${c.hex}</span>
          </div>
          <div class="my-color-actions">
            <button class="my-edit-btn${isEditing ? ' active' : ''}"
                    data-idx="${i}" aria-label="Edit ${esc(c.name)}" title="Edit color">
              <svg width="11" height="11" viewBox="0 0 12 12" fill="none">
                <path d="M8.5 1.5a1.414 1.414 0 012 2L4 10H2V8L8.5 1.5z"
                      stroke="currentColor" stroke-width="1.4"
                      stroke-linecap="round" stroke-linejoin="round" fill="none"/>
              </svg>
            </button>
            <button class="my-del-btn" data-idx="${i}"
                    aria-label="Delete ${esc(c.name)}" title="Delete color">
              <svg width="10" height="10" viewBox="0 0 10 10" fill="none">
                <line x1="1.5" y1="1.5" x2="8.5" y2="8.5" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/>
                <line x1="8.5" y1="1.5" x2="1.5" y2="8.5" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/>
              </svg>
            </button>
          </div>
        </div>
        ${isEditing ? buildColorForm('edit', i, c.hex, c.name) : ''}
      </div>`;
  }).join('');

  const isAdding    = _openForm && _openForm.type === 'add';
  const addFormHtml = isAdding ? buildColorForm('add', null, '#2A3E6D', '') : '';

  document.getElementById('edit-panel-body').innerHTML = `
    <div class="ep-label">
      My Colors
      <span class="ep-hint">click swatch to apply</span>
    </div>
    <div id="my-colors-list">${myColorsHtml}</div>

    <button class="add-new-btn${isAdding ? ' active' : ''}" id="ep-add-toggle">
      <svg width="11" height="11" viewBox="0 0 11 11" fill="none">
        <line x1="5.5" y1="1" x2="5.5" y2="10" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/>
        <line x1="1" y1="5.5" x2="10" y2="5.5" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/>
      </svg>
      Add New Color
    </button>
    ${addFormHtml}

    <button class="no-color-btn" id="ep-no-color">
      <svg width="13" height="13" viewBox="0 0 13 13" fill="none">
        <rect x="1" y="1" width="11" height="11" rx="2" stroke="currentColor" stroke-width="1.2"/>
        <line x1="2" y1="11" x2="11" y2="2" stroke="currentColor" stroke-width="1.2"/>
      </svg>
      ${NO_LABEL[key] || 'No Color'}
    </button>

    <button class="reset-defaults-btn" id="ep-reset">
      ↺ Reset to defaults
    </button>
  `;

  bindEditEvents(key);
}

// ─────────────────────────────────────────────────────────────────
// Build color form — APi quick-pick + Standard quick-pick + HSV wheel
// Clicking quick-pick swatches SETS the picker value, not applies.
// Only Save / Add button commits the change.
// ─────────────────────────────────────────────────────────────────
function buildColorForm(type, idx, startHex, startName) {
  const norm   = normalizeHex(startHex) || '#2A3E6D';
  const hexVal = norm.replace('#', '');
  const sufx   = type === 'edit' ? `edit-${idx}` : 'add';

  // Quick-pick: APi Group base colors
  const apiSwatchHtml = BRAND_COLORS.map(c =>
    `<button class="qp-swatch${isLight(c.hex) ? ' light' : ''}"
             style="background:${c.hex}"
             data-setcolor="${c.hex}"
             title="${esc(c.name)}"
             type="button"></button>`
  ).join('');

  // Quick-pick: Standard colors
  const stdSwatchHtml = STANDARD_COLORS.map(c =>
    `<button class="qp-swatch${isLight(c.hex) ? ' light' : ''}"
             style="background:${c.hex}"
             data-setcolor="${c.hex}"
             title="${esc(c.name)}"
             type="button"></button>`
  ).join('');

  // Quick-pick: Recent colors — uses _editKey to pull the right history
  const recentColors = getRecent(_editKey);
  const recentQpHtml = recentColors.length
    ? `<div class="cform-qlabel" style="margin-top:7px;">Recent Colors</div>
       <div class="qp-row recent-qp">${recentColors.map(hex =>
         `<button class="qp-swatch${isLight(hex) ? ' light' : ''}"
                  style="background:${hex}"
                  data-setcolor="${hex}"
                  title="${hex}"
                  type="button"></button>`
       ).join('')}</div>`
    : '';

  return `
    <div class="color-form" id="cform-${sufx}" data-type="${type}" data-idx="${idx ?? ''}">

      <!-- Quick-pick: APi Group Colors -->
      <div class="cform-qlabel">APi Group Colors</div>
      <div class="qp-row api-qp">${apiSwatchHtml}</div>

      <!-- Quick-pick: Standard Colors -->
      <div class="cform-qlabel" style="margin-top:7px;">Standard Colors</div>
      <div class="qp-row std-qp">${stdSwatchHtml}</div>

      <!-- Quick-pick: Recent Colors (only shown when there are recents) -->
      ${recentQpHtml}

      <!-- Divider -->
      <div class="cform-divider"><span>Custom Color</span></div>

      <!-- HSV gradient square -->
      <div class="hsv-square" id="hsv-sq-${sufx}">
        <div class="hsv-layer hsv-sat"></div>
        <div class="hsv-layer hsv-val"></div>
        <div class="hsv-dot" id="hsv-dot-${sufx}"></div>
      </div>

      <!-- Hue slider -->
      <div class="hue-wrap">
        <input type="range" class="hue-slider" id="hue-${sufx}"
               min="0" max="359" step="1" value="0" />
      </div>

      <!-- Hex input + preview -->
      <div class="cform-row">
        <div class="hex-field">
          <span class="hex-hash">#</span>
          <input type="text" class="ep-hex-input" id="cfh-${sufx}"
                 value="${hexVal}" maxlength="6"
                 placeholder="RRGGBB" spellcheck="false" autocomplete="off" />
        </div>
        <div class="cform-preview" id="cfp-${sufx}" style="background:${norm}"></div>
      </div>

      <!-- Color name -->
      <input type="text" class="ep-name-input" id="cfn-${sufx}"
             placeholder="Color name (optional)" maxlength="28"
             value="${esc(startName)}" />

      <!-- Save / Cancel -->
      <div class="cform-btns">
        <button class="cform-save" type="button" data-type="${type}" data-idx="${idx ?? ''}">
          ${type === 'edit' ? 'Save Changes' : '+ Add to My Colors'}
        </button>
        <button class="cform-cancel" type="button">Cancel</button>
      </div>
    </div>`;
}

// ─────────────────────────────────────────────────────────────────
// Bind all events inside the edit panel
// ─────────────────────────────────────────────────────────────────
function bindEditEvents(key) {
  const body = document.getElementById('edit-panel-body');

  // ── My Colors: click swatch = APPLY + close ──
  body.querySelectorAll('.my-swatch').forEach(s => {
    s.addEventListener('click', () => applyAndClose(s.dataset.color));
    s.addEventListener('keydown', e => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); applyAndClose(s.dataset.color); }
    });
  });

  // ── My Colors: pencil button — toggle form ──
  body.querySelectorAll('.my-edit-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const idx = parseInt(btn.dataset.idx, 10);
      _openForm = (_openForm && _openForm.type === 'edit' && _openForm.idx === idx)
        ? null
        : { type: 'edit', idx };
      renderEditPanel();
    });
  });

  // ── My Colors: delete button ──
  body.querySelectorAll('.my-del-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      deleteColor(key, parseInt(btn.dataset.idx, 10));
      _openForm = null;
      renderEditPanel();
    });
  });

  // ── Add New Color toggle ──
  document.getElementById('ep-add-toggle').addEventListener('click', () => {
    _openForm = (_openForm && _openForm.type === 'add')
      ? null
      : { type: 'add', idx: null };
    renderEditPanel();
  });

  // ── No color ──
  document.getElementById('ep-no-color').addEventListener('click', () => applyAndClose('none'));

  // ── Reset to defaults ──
  document.getElementById('ep-reset').addEventListener('click', () => {
    if (!confirm(`Reset ${key} palette to defaults? Your custom colors will be lost.`)) return;
    localStorage.removeItem(`sct_${key}`);
    _openForm = null;
    renderEditPanel();
    renderMainSwatches(key);
  });

  // ── Wire all open color forms ──
  body.querySelectorAll('.color-form').forEach(form => {
    wireColorForm(form, key);
  });
}

// ─────────────────────────────────────────────────────────────────
// Wire a single color form — HSV picker + quick-select
// Quick-pick swatches SET the picker value (no auto-apply).
// ─────────────────────────────────────────────────────────────────
function wireColorForm(form, key) {
  const type    = form.dataset.type;
  const idx     = form.dataset.idx !== '' ? parseInt(form.dataset.idx, 10) : null;
  const sufx    = type === 'edit' ? `edit-${idx}` : 'add';

  const hexInp    = document.getElementById(`cfh-${sufx}`);
  const preview   = document.getElementById(`cfp-${sufx}`);
  const nameInp   = document.getElementById(`cfn-${sufx}`);
  const sq        = document.getElementById(`hsv-sq-${sufx}`);
  const dot       = document.getElementById(`hsv-dot-${sufx}`);
  const hueSlider = document.getElementById(`hue-${sufx}`);

  // Initialize HSV from starting hex
  const startNorm = normalizeHex('#' + hexInp.value) || '#2A3E6D';
  const startHsv  = hexToHsv(startNorm);
  let H = startHsv.h, S = startHsv.s, V = startHsv.v;

  // ── Sync all UI from current H, S, V ──
  function syncFromHsv() {
    const hex = hsvToHex(H, S, V);
    hexInp.value            = hex.replace('#', '').toUpperCase();
    preview.style.background = hex;
    sq.style.background     = `hsl(${Math.round(H)}, 100%, 50%)`;
    dot.style.left          = `${S * 100}%`;
    dot.style.top           = `${(1 - V) * 100}%`;
    hueSlider.value         = Math.round(H);
  }

  // ── Set color from any hex string — updates H, S, V and all UI ──
  function setFromHex(rawHex) {
    const candidate = rawHex.length === 6 && !rawHex.startsWith('#') ? '#' + rawHex : rawHex;
    const norm = normalizeHex(candidate);
    if (!norm) return;
    const hsv = hexToHsv(norm);
    H = hsv.h; S = hsv.s; V = hsv.v;
    syncFromHsv();
  }

  // Initial render
  syncFromHsv();

  // ── Hue slider ──
  hueSlider.addEventListener('input', () => {
    H = parseFloat(hueSlider.value);
    syncFromHsv();
  });

  // ── HSV square — click + drag ──
  let dragging = false;

  function handleSquareDrag(e) {
    if (!sq.isConnected) { dragging = false; return; }
    const rect    = sq.getBoundingClientRect();
    const clientX = e.touches ? e.touches[0].clientX : e.clientX;
    const clientY = e.touches ? e.touches[0].clientY : e.clientY;
    S = Math.max(0, Math.min(1, (clientX - rect.left)  / rect.width));
    V = Math.max(0, Math.min(1, 1 - (clientY - rect.top) / rect.height));
    syncFromHsv();
  }

  sq.addEventListener('mousedown', e => {
    e.preventDefault();
    dragging = true;
    handleSquareDrag(e);
  });

  const onMouseMove = e => { if (dragging) handleSquareDrag(e); };
  const onMouseUp   = () => { dragging = false; };
  document.addEventListener('mousemove', onMouseMove);
  document.addEventListener('mouseup',   onMouseUp);

  sq.addEventListener('touchstart', e => { e.preventDefault(); handleSquareDrag(e); }, { passive: false });
  sq.addEventListener('touchmove',  e => { e.preventDefault(); handleSquareDrag(e); }, { passive: false });

  // ── Hex input — manual entry ──
  hexInp.addEventListener('input', () => setFromHex(hexInp.value));

  // ── Quick-pick swatches (APi + Standard inside the form) ──
  // These SET the picker value — they do NOT apply or save.
  form.querySelectorAll('.qp-swatch').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      setFromHex(btn.dataset.setcolor);
      // Visual feedback: briefly highlight the selected swatch
      form.querySelectorAll('.qp-swatch').forEach(b => b.classList.remove('qp-selected'));
      btn.classList.add('qp-selected');
    });
  });

  // ── Save / Add button ──
  const cleanup = () => {
    document.removeEventListener('mousemove', onMouseMove);
    document.removeEventListener('mouseup',   onMouseUp);
  };

  form.querySelector('.cform-save').addEventListener('click', () => {
    const norm = normalizeHex('#' + hexInp.value);
    if (!norm) { hexInp.focus(); hexInp.select(); return; }
    const name = nameInp.value.trim();
    if (type === 'edit' && idx !== null) {
      updateColor(key, idx, norm, name || norm);
    } else {
      addColor(key, norm, name);
    }
    _openForm = null;
    cleanup();
    renderEditPanel();
  });

  // ── Cancel ──
  form.querySelector('.cform-cancel').addEventListener('click', () => {
    _openForm = null;
    cleanup();
    renderEditPanel();
  });

  // ── Enter in name field = save ──
  nameInp.addEventListener('keydown', e => {
    if (e.key === 'Enter') form.querySelector('.cform-save').click();
  });

  // Auto-select hex input on open
  setTimeout(() => { if (hexInp.isConnected) hexInp.select(); }, 50);
}

// ─────────────────────────────────────────────────────────────────
// Apply color and close edit panel
// ─────────────────────────────────────────────────────────────────
function applyAndClose(color) {
  if (_applyFn) _applyFn(color);
  if (color !== 'none') pushRecent(_editKey, color);
  closeEditPanel();
}

// ─────────────────────────────────────────────────────────────────
// Office ready
// ─────────────────────────────────────────────────────────────────
Office.onReady(() => {
  renderMainSwatches('fill');
  renderMainSwatches('border');
  renderMainSwatches('text');

  document.querySelectorAll('[data-weight]').forEach(btn =>
    btn.addEventListener('click', () => applyBorderWeight(parseFloat(btn.dataset.weight))));

  document.querySelectorAll('[data-dash]').forEach(btn =>
    btn.addEventListener('click', () => applyBorderDash(btn.dataset.dash)));

  document.getElementById('btn-apply-ends')
    ?.addEventListener('click', applyLineEnds);

  document.getElementById('edit-back-btn')
    .addEventListener('click', closeEditPanel);

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
      _cachedShapeCount = items.length; // cache for fast apply

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

      const anySupported = Object.values(merged).some(Boolean);
      const meta = allSame
        ? (TYPE_META[firstKey] || { icon: '?', label: cap(firstKey || 'Shape') })
        : { icon: '⊕', label: 'Mixed Selection' };

      renderUI({
        merged,
        meta,
        firstName,
        count: items.length,
        anySupported,
        isLine: allSame && firstKey === 'line',
      });
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
// Render helpers
// ─────────────────────────────────────────────────────────────────
function renderUI({ merged, meta, firstName, count, anySupported, isLine }) {
  hide('empty-state');
  hide('unsupported-state');
  show('shape-banner');

  el('shape-icon').textContent        = meta.icon;
  el('shape-type-label').textContent  = meta.label;
  el('shape-name').textContent        = firstName ? `"${firstName}"` : '';
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
  _cachedShapeCount = 0;
  show('empty-state');
  hide('shape-banner');
  hide('unsupported-state');
  ['section-fill', 'section-border', 'section-line-ends', 'section-text'].forEach(hide);
  setStatus('');
}

function el(id)       { return document.getElementById(id); }
function show(id)     { el(id)?.classList.remove('hidden'); }
function hide(id)     { el(id)?.classList.add('hidden'); }
function tog(id, vis) { el(id)?.classList.toggle('hidden', !vis); }
function setStatus(m) { if (el('status')) el('status').textContent = m; }

// ─────────────────────────────────────────────────────────────────
// Apply functions
// ─────────────────────────────────────────────────────────────────
// Fast helper — no load() needed for write operations.
// Uses cached count + getItemAt(i) to skip the first sync entirely.
function getSelectedFast(ctx) {
  const col    = ctx.presentation.getSelectedShapes();
  const shapes = [];
  for (let i = 0; i < _cachedShapeCount; i++) {
    shapes.push(col.getItemAt(i));
  }
  return shapes; // proxy objects — safe to write to before sync
}

async function applyFill(color) {
  const ver = ++_applyVer;
  if (!_cachedShapeCount) return setStatus('No shape selected.');
  await PowerPoint.run(async ctx => {
    if (ver !== _applyVer) return; // superseded by a newer click
    getSelectedFast(ctx).forEach(s => {
      if (color === 'none') s.fill.clear();
      else s.fill.setSolidColor(color);
    });
    await ctx.sync();
    if (ver === _applyVer) setStatus(`Fill → ${color === 'none' ? 'cleared' : color}`);
  });
}

async function applyBorderColor(color) {
  const ver = ++_applyVer;
  if (!_cachedShapeCount) return setStatus('No shape selected.');
  await PowerPoint.run(async ctx => {
    if (ver !== _applyVer) return;
    getSelectedFast(ctx).forEach(s => {
      if (color === 'none') s.lineFormat.visible = false;
      else { s.lineFormat.visible = true; s.lineFormat.color = color; }
    });
    await ctx.sync();
    if (ver === _applyVer) setStatus(`Border → ${color === 'none' ? 'removed' : color}`);
  });
}

async function applyBorderWeight(pts) {
  const ver = ++_applyVer;
  if (!_cachedShapeCount) return setStatus('No shape selected.');
  await PowerPoint.run(async ctx => {
    if (ver !== _applyVer) return;
    getSelectedFast(ctx).forEach(s => { s.lineFormat.weight = pts; s.lineFormat.visible = true; });
    await ctx.sync();
    if (ver === _applyVer) setStatus(`Weight → ${pts}pt`);
  });
}

async function applyBorderDash(style) {
  const ver = ++_applyVer;
  if (!_cachedShapeCount) return setStatus('No shape selected.');
  await PowerPoint.run(async ctx => {
    if (ver !== _applyVer) return;
    // Resolve enum INSIDE PowerPoint.run where it is guaranteed to be available.
    // LineDashStyle.dot does not exist in the API — roundDot is the correct dotted value.
    let dashVal;
    switch (style) {
      case 'dash':     dashVal = PowerPoint.LineDashStyle.dash;     break;
      case 'roundDot': dashVal = PowerPoint.LineDashStyle.roundDot; break;
      default:         dashVal = PowerPoint.LineDashStyle.solid;    break;
    }
    getSelectedFast(ctx).forEach(s => {
      s.lineFormat.dashStyle = dashVal;
      s.lineFormat.visible   = true;
    });
    await ctx.sync();
    if (ver === _applyVer) setStatus(`Style → ${style}`);
  });
}

async function applyTextColor(color) {
  const ver = ++_applyVer;
  if (!_cachedShapeCount) return setStatus('No shape selected.');
  await PowerPoint.run(async ctx => {
    if (ver !== _applyVer) return;
    getSelectedFast(ctx).forEach(s => {
      try { s.textFrame.textRange.font.color = color; } catch (_) {}
    });
    await ctx.sync();
    if (ver === _applyVer) setStatus(`Text → ${color}`);
  });
}

async function applyLineEnds() {
  const aMap = {
    none:      PowerPoint.ArrowheadStyle.none,
    arrow:     PowerPoint.ArrowheadStyle.arrow,
    openArrow: PowerPoint.ArrowheadStyle.openArrow,
    diamond:   PowerPoint.ArrowheadStyle.diamond,
    oval:      PowerPoint.ArrowheadStyle.oval,
  };
  const sv = el('line-start').value;
  const ev = el('line-end').value;
  const ver = ++_applyVer;
  if (!_cachedShapeCount) return setStatus('No shape selected.');
  await PowerPoint.run(async ctx => {
    if (ver !== _applyVer) return;
    getSelectedFast(ctx).forEach(s => {
      try {
        s.lineFormat.beginArrowheadStyle = aMap[sv] ?? PowerPoint.ArrowheadStyle.none;
        s.lineFormat.endArrowheadStyle   = aMap[ev] ?? PowerPoint.ArrowheadStyle.none;
      } catch (_) {}
    });
    await ctx.sync();
    if (ver === _applyVer) setStatus(`Ends → ${sv} / ${ev}`);
  });
}
