'use strict';

// ─────────────────────────────────────────────────────────────────
// Brand / static color data
// ─────────────────────────────────────────────────────────────────
const BRAND_COLORS = [
  { hex: '#D50032', name: 'APi Corp Red'  },
  { hex: '#2A3E6D', name: 'Dark Navy'     },
  { hex: '#008579', name: 'Teal'          },
  { hex: '#D4D800', name: 'TBD Yellow'    },
  { hex: '#7030A0', name: 'APi Seg'       },
  { hex: '#00B0F0', name: 'Target HQ'     },
  { hex: '#92D050', name: 'Target OpCo'   },
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

// Text color quick-picks used inside the DOI color form
const TEXT_QP_COLORS = [
  { hex: '#FFFFFF', name: 'White'      },
  { hex: '#000000', name: 'Black'      },
  { hex: '#D50032', name: 'APi Red'    },
  { hex: '#2A3E6D', name: 'Dark Navy'  },
  { hex: '#008579', name: 'Teal'       },
  { hex: '#595959', name: 'Dark Gray'  },
  { hex: '#BFBFBF', name: 'Light Gray' },
];

// Default fill entry text color — white on dark, black on light
const autoTextHex = hex => isLight(hex) ? '#000000' : '#FFFFFF';

const DEFAULT_PALETTES = {
  // Each fill entry carries textHex — the font color applied alongside the fill
  fill: [
    { hex: '#D50032', name: 'Target Converts to Acquiring', textHex: '#FFFFFF' },
    { hex: '#2A3E6D', name: 'APi Corp Entity',               textHex: '#FFFFFF' },
    { hex: '#008579', name: 'Shared Services',               textHex: '#FFFFFF' },
    { hex: '#595959', name: 'Leave As Is',                   textHex: '#FFFFFF' },
    { hex: '#BFBFBF', name: 'To Be Determined',              textHex: '#000000' },
    { hex: '#FFFFFF', name: 'No Fill',                       textHex: '#000000' },
    { hex: '#000000', name: 'Divest / Sunset',               textHex: '#FFFFFF' },
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
  const c = v * s, x = c * (1 - Math.abs((h / 60) % 2 - 1)), m = v - c;
  let r = 0, g = 0, b = 0;
  if      (h < 60)  { r = c; g = x; b = 0; }
  else if (h < 120) { r = x; g = c; b = 0; }
  else if (h < 180) { r = 0; g = c; b = x; }
  else if (h < 240) { r = 0; g = x; b = c; }
  else if (h < 300) { r = x; g = 0; b = c; }
  else              { r = c; g = 0; b = x; }
  return { r: Math.round((r + m) * 255), g: Math.round((g + m) * 255), b: Math.round((b + m) * 255) };
}

function rgbToHsv(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b), d = max - min;
  let h = 0;
  const s = max === 0 ? 0 : d / max, v = max;
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
    const s   = localStorage.getItem(`sct_${key}`);
    const arr = s ? JSON.parse(s) : JSON.parse(JSON.stringify(DEFAULT_PALETTES[key]));
    // Backwards-compat: ensure fill entries have textHex
    if (key === 'fill') {
      arr.forEach(c => { if (!c.textHex) c.textHex = autoTextHex(c.hex); });
    }
    return arr;
  } catch { return JSON.parse(JSON.stringify(DEFAULT_PALETTES[key])); }
}

function savePalette(key, arr) {
  localStorage.setItem(`sct_${key}`, JSON.stringify(arr));
}

function updateColor(key, idx, hex, name, textHex) {
  const arr = getPalette(key);
  if (!arr[idx]) return;
  arr[idx] = { hex, name: name || hex };
  if (key === 'fill') arr[idx].textHex = textHex || autoTextHex(hex);
  savePalette(key, arr);
}

function addColor(key, hex, name, textHex) {
  const norm = normalizeHex(hex);
  if (!norm) return;
  const arr   = getPalette(key);
  const entry = { hex: norm, name: name || norm };
  if (key === 'fill') entry.textHex = textHex || autoTextHex(norm);
  arr.push(entry);
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
  let arr    = getRecent(key).filter(h => h !== norm);
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
// Performance state
// ─────────────────────────────────────────────────────────────────
let _cachedShapeCount = 0;
let _applyVer         = 0;

// ─────────────────────────────────────────────────────────────────
// Main panel — 7 labeled tiles + pencil button
// ─────────────────────────────────────────────────────────────────
function renderMainSwatches(key) {
  const row    = document.getElementById(`swatches-${key}`);
  if (!row) return;

  const fnMap  = { fill: applyFill, border: applyBorderColor, text: applyTextColor };
  const fn     = fnMap[key];
  const colors = getPalette(key);

  row.innerHTML =
    colors.map(c => {
      // ── Border: legend line tile ──
      if (key === 'border') {
        return `<button class="main-swatch border-legend-tile"
                         data-color="${c.hex}"
                         title="${esc(c.name)}\n${c.hex}"
                         aria-label="${esc(c.name)}">
                  <span class="legend-line-sample" style="background:${c.hex}"></span>
                  <span class="legend-line-name">${esc(c.name)}</span>
                </button>`;
      }
      // ── Fill (DOI): solid tile — label in configured textHex, small text-color indicator ──
      if (key === 'fill') {
        const txHex = c.textHex || autoTextHex(c.hex);
        return `<button class="main-swatch doi-tile"
                         style="background:${c.hex};"
                         data-color="${c.hex}"
                         data-textcolor="${txHex}"
                         title="${esc(c.name)}\n${c.hex}"
                         aria-label="${esc(c.name)}">
                  <span class="swatch-label" style="color:${txHex};">${esc(c.name)}</span>
                  <span class="doi-text-dot" style="background:${txHex}; outline:2px solid ${c.hex}; outline-offset:1px;" title="Font: ${txHex}"></span>
                </button>`;
      }
      // ── Text: solid color tile ──
      const txtColor = isLight(c.hex) ? '#000000' : '#FFFFFF';
      return `<button class="main-swatch${isLight(c.hex) ? ' light' : ''}"
                       style="background:${c.hex}"
                       data-color="${c.hex}"
                       title="${esc(c.name)}\n${c.hex}"
                       aria-label="${esc(c.name)}">
                <span class="swatch-label" style="color:${txtColor};">${esc(c.name)}</span>
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

  // Wire swatch clicks
  row.querySelectorAll('.main-swatch').forEach(btn => {
    btn.addEventListener('click', () => {
      const color    = btn.dataset.color;
      const txtColor = btn.dataset.textcolor;
      fn(color);
      // DOI tiles auto-apply their paired font color
      if (key === 'fill' && txtColor) applyTextColor(txtColor);
      pushRecent(key, color);
    });
  });

  // Wire edit button
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
  const key      = _editKey;
  const titleMap = { fill: 'Degree of Integration', border: 'Border Color', text: 'Text Color' };
  document.getElementById('edit-panel-title').textContent = titleMap[key] || 'Edit Colors';

  const palette = getPalette(key);

  const myColorsHtml = palette.map((c, i) => {
    const isEditing = _openForm && _openForm.type === 'edit' && _openForm.idx === i;
    const txHex     = c.textHex || autoTextHex(c.hex);
    // For DOI (fill) swatches in the edit list, show the text-color dot too
    const swatchInner = (key === 'fill')
      ? `<span class="doi-text-dot" style="background:${txHex}; outline:2px solid ${c.hex}; outline-offset:1px;"></span>`
      : '';

    return `
      <div class="my-color-item" data-idx="${i}">
        <div class="my-color-top">
          <div class="my-swatch-wrap">
            <div class="my-swatch${isLight(c.hex) ? ' light' : ''}"
                 style="background:${c.hex}"
                 data-color="${c.hex}"
                 data-textcolor="${txHex}"
                 role="button" tabindex="0"
                 aria-label="Apply ${esc(c.name)}">
              ${swatchInner}
            </div>
          </div>
          <div class="my-color-info">
            <span class="my-color-name">${esc(c.name)}</span>
            <span class="my-color-hex">${c.hex}${key === 'fill' ? ` · T: ${txHex}` : ''}</span>
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
                    aria-label="Delete ${esc(c.name)}" title="Delete">
              <svg width="10" height="10" viewBox="0 0 10 10" fill="none">
                <line x1="1.5" y1="1.5" x2="8.5" y2="8.5" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/>
                <line x1="8.5" y1="1.5" x2="1.5" y2="8.5" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/>
              </svg>
            </button>
          </div>
        </div>
        ${isEditing ? buildColorForm('edit', i, c.hex, c.name, txHex) : ''}
      </div>`;
  }).join('');

  const isAdding    = _openForm && _openForm.type === 'add';
  const addFormHtml = isAdding ? buildColorForm('add', null, '#2A3E6D', '', '#FFFFFF') : '';

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

    <button class="reset-defaults-btn" id="ep-reset">↺ Reset to defaults</button>
  `;

  bindEditEvents(key);
}

// ─────────────────────────────────────────────────────────────────
// Build color form
// For fill (DOI): includes a text color section below the fill picker.
// Quick-pick swatches SET the picker value — they don't save/apply.
// ─────────────────────────────────────────────────────────────────
function buildColorForm(type, idx, startHex, startName, startTextHex) {
  const norm      = normalizeHex(startHex) || '#2A3E6D';
  const hexVal    = norm.replace('#', '');
  const sufx      = type === 'edit' ? `edit-${idx}` : 'add';
  const txNorm    = normalizeHex(startTextHex) || '#FFFFFF';
  const txHexVal  = txNorm.replace('#', '');
  const isDOI     = _editKey === 'fill';

  // Quick-pick rows
  const apiQp = BRAND_COLORS.map(c =>
    `<button class="qp-swatch${isLight(c.hex) ? ' light' : ''}"
             style="background:${c.hex}" data-setcolor="${c.hex}"
             title="${esc(c.name)}" type="button"></button>`
  ).join('');

  const stdQp = STANDARD_COLORS.map(c =>
    `<button class="qp-swatch${isLight(c.hex) ? ' light' : ''}"
             style="background:${c.hex}" data-setcolor="${c.hex}"
             title="${esc(c.name)}" type="button"></button>`
  ).join('');

  const recent    = getRecent(_editKey);
  const recentQp  = recent.length
    ? `<div class="cform-qlabel" style="margin-top:7px;">Recent Colors</div>
       <div class="qp-row">${recent.map(hex =>
         `<button class="qp-swatch${isLight(hex) ? ' light' : ''}"
                  style="background:${hex}" data-setcolor="${hex}"
                  title="${hex}" type="button"></button>`
       ).join('')}</div>`
    : '';

  // Text color quick-picks (DOI only)
  const textQp = isDOI
    ? TEXT_QP_COLORS.map(c =>
        `<button class="tc-qp-swatch${isLight(c.hex) ? ' light' : ''}${c.hex.toUpperCase() === txNorm ? ' tc-selected' : ''}"
                 style="background:${c.hex}" data-settextcolor="${c.hex}"
                 title="${esc(c.name)}" type="button"></button>`
      ).join('')
    : '';

  const textSection = isDOI ? `
    <!-- ── Text Color section (DOI only) ── -->
    <div class="cform-divider"><span>Font Color</span></div>

    <div class="tc-qp-row" id="tc-qp-${sufx}">${textQp}</div>

    <div class="cform-row" style="margin-top:6px;">
      <div class="hex-field">
        <span class="hex-hash">#</span>
        <input type="text" class="ep-hex-input" id="tc-hex-${sufx}"
               value="${txHexVal}" maxlength="6"
               placeholder="RRGGBB" spellcheck="false" autocomplete="off" />
      </div>
      <div class="cform-preview" id="tc-prev-${sufx}" style="background:${txNorm}"></div>
    </div>
  ` : '';

  return `
    <div class="color-form" id="cform-${sufx}" data-type="${type}" data-idx="${idx ?? ''}">

      <div class="cform-qlabel">APi Group Colors</div>
      <div class="qp-row api-qp">${apiQp}</div>

      <div class="cform-qlabel" style="margin-top:7px;">Standard Colors</div>
      <div class="qp-row std-qp">${stdQp}</div>

      ${recentQp}

      <div class="cform-divider"><span>Fill Color</span></div>

      <div class="hsv-square" id="hsv-sq-${sufx}">
        <div class="hsv-layer hsv-sat"></div>
        <div class="hsv-layer hsv-val"></div>
        <div class="hsv-dot" id="hsv-dot-${sufx}"></div>
      </div>

      <div class="hue-wrap">
        <input type="range" class="hue-slider" id="hue-${sufx}"
               min="0" max="359" step="1" value="0" />
      </div>

      <div class="cform-row">
        <div class="hex-field">
          <span class="hex-hash">#</span>
          <input type="text" class="ep-hex-input" id="cfh-${sufx}"
                 value="${hexVal}" maxlength="6"
                 placeholder="RRGGBB" spellcheck="false" autocomplete="off" />
        </div>
        <div class="cform-preview" id="cfp-${sufx}" style="background:${norm}"></div>
      </div>

      ${textSection}

      <input type="text" class="ep-name-input" id="cfn-${sufx}"
             placeholder="Color name (optional)" maxlength="28"
             value="${esc(startName)}" />

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

  // My Colors: click swatch → APPLY (fill applies text too) + close
  body.querySelectorAll('.my-swatch').forEach(s => {
    const applyClick = () => {
      const textColor = (key === 'fill') ? (s.dataset.textcolor || null) : null;
      applyAndClose(s.dataset.color, textColor);
    };
    s.addEventListener('click', applyClick);
    s.addEventListener('keydown', e => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); applyClick(); }
    });
  });

  // My Colors: pencil — toggle form
  body.querySelectorAll('.my-edit-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const idx = parseInt(btn.dataset.idx, 10);
      _openForm = (_openForm && _openForm.type === 'edit' && _openForm.idx === idx)
        ? null : { type: 'edit', idx };
      renderEditPanel();
    });
  });

  // My Colors: delete
  body.querySelectorAll('.my-del-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      deleteColor(key, parseInt(btn.dataset.idx, 10));
      _openForm = null;
      renderEditPanel();
    });
  });

  // Add New Color toggle
  document.getElementById('ep-add-toggle').addEventListener('click', () => {
    _openForm = (_openForm && _openForm.type === 'add')
      ? null : { type: 'add', idx: null };
    renderEditPanel();
  });

  // No color
  document.getElementById('ep-no-color').addEventListener('click', () => applyAndClose('none', null));

  // Reset to defaults
  document.getElementById('ep-reset').addEventListener('click', () => {
    if (!confirm(`Reset ${key} palette to defaults? Your custom colors will be lost.`)) return;
    localStorage.removeItem(`sct_${key}`);
    _openForm = null;
    renderEditPanel();
    renderMainSwatches(key);
  });

  // Wire all open color forms
  body.querySelectorAll('.color-form').forEach(form => wireColorForm(form, key));
}

// ─────────────────────────────────────────────────────────────────
// Wire a single color form — HSV fill picker + optional text picker
// ─────────────────────────────────────────────────────────────────
function wireColorForm(form, key) {
  const type      = form.dataset.type;
  const idx       = form.dataset.idx !== '' ? parseInt(form.dataset.idx, 10) : null;
  const sufx      = type === 'edit' ? `edit-${idx}` : 'add';
  const isDOI     = key === 'fill';

  // Fill picker refs
  const hexInp    = document.getElementById(`cfh-${sufx}`);
  const preview   = document.getElementById(`cfp-${sufx}`);
  const nameInp   = document.getElementById(`cfn-${sufx}`);
  const sq        = document.getElementById(`hsv-sq-${sufx}`);
  const dot       = document.getElementById(`hsv-dot-${sufx}`);
  const hueSlider = document.getElementById(`hue-${sufx}`);

  // Text color refs (DOI only)
  const tcHexInp  = isDOI ? document.getElementById(`tc-hex-${sufx}`) : null;
  const tcPreview = isDOI ? document.getElementById(`tc-prev-${sufx}`) : null;

  // ── Fill HSV state ──
  const startNorm = normalizeHex('#' + hexInp.value) || '#2A3E6D';
  const startHsv  = hexToHsv(startNorm);
  let H = startHsv.h, S = startHsv.s, V = startHsv.v;

  function syncFromHsv() {
    const hex = hsvToHex(H, S, V);
    hexInp.value             = hex.replace('#', '').toUpperCase();
    preview.style.background = hex;
    sq.style.background      = `hsl(${Math.round(H)}, 100%, 50%)`;
    dot.style.left           = `${S * 100}%`;
    dot.style.top            = `${(1 - V) * 100}%`;
    hueSlider.value          = Math.round(H);
  }

  function setFillFromHex(rawHex) {
    const c  = rawHex.length === 6 && !rawHex.startsWith('#') ? '#' + rawHex : rawHex;
    const n  = normalizeHex(c);
    if (!n) return;
    const hsv = hexToHsv(n);
    H = hsv.h; S = hsv.s; V = hsv.v;
    syncFromHsv();
  }

  syncFromHsv();

  hueSlider.addEventListener('input', () => { H = parseFloat(hueSlider.value); syncFromHsv(); });

  let dragging = false;
  function handleSquareDrag(e) {
    if (!sq.isConnected) { dragging = false; return; }
    const rect = sq.getBoundingClientRect();
    const cx   = e.touches ? e.touches[0].clientX : e.clientX;
    const cy   = e.touches ? e.touches[0].clientY : e.clientY;
    S = Math.max(0, Math.min(1, (cx - rect.left) / rect.width));
    V = Math.max(0, Math.min(1, 1 - (cy - rect.top) / rect.height));
    syncFromHsv();
  }

  sq.addEventListener('mousedown', e => { e.preventDefault(); dragging = true; handleSquareDrag(e); });
  const onMove = e => { if (dragging) handleSquareDrag(e); };
  const onUp   = () => { dragging = false; };
  document.addEventListener('mousemove', onMove);
  document.addEventListener('mouseup',   onUp);
  sq.addEventListener('touchstart', e => { e.preventDefault(); handleSquareDrag(e); }, { passive: false });
  sq.addEventListener('touchmove',  e => { e.preventDefault(); handleSquareDrag(e); }, { passive: false });

  hexInp.addEventListener('input', () => setFillFromHex(hexInp.value));

  // Fill quick-picks (APi + Standard + Recent)
  form.querySelectorAll('.qp-swatch').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      setFillFromHex(btn.dataset.setcolor);
      form.querySelectorAll('.qp-swatch').forEach(b => b.classList.remove('qp-selected'));
      btn.classList.add('qp-selected');
    });
  });

  // ── Text color state (DOI only) ──
  function syncTextPreview(hex) {
    if (!tcPreview || !tcHexInp) return;
    const n = normalizeHex(hex.startsWith('#') ? hex : '#' + hex);
    if (!n) return;
    tcPreview.style.background = n;
    tcHexInp.value             = n.replace('#', '').toUpperCase();
    // Update active state on text QP swatches
    form.querySelectorAll('.tc-qp-swatch').forEach(b => {
      b.classList.toggle('tc-selected', b.dataset.settextcolor.toUpperCase() === n);
    });
  }

  if (isDOI) {
    // Text color quick-picks
    form.querySelectorAll('.tc-qp-swatch').forEach(btn => {
      btn.addEventListener('click', e => {
        e.stopPropagation();
        syncTextPreview(btn.dataset.settextcolor);
      });
    });

    // Text hex manual input
    tcHexInp.addEventListener('input', () => {
      const val = tcHexInp.value.trim();
      if (val.length === 6) syncTextPreview(val);
    });
  }

  // ── Save / Add ──
  const cleanup = () => {
    document.removeEventListener('mousemove', onMove);
    document.removeEventListener('mouseup',   onUp);
  };

  form.querySelector('.cform-save').addEventListener('click', () => {
    const norm = normalizeHex('#' + hexInp.value);
    if (!norm) { hexInp.focus(); hexInp.select(); return; }
    const name    = nameInp.value.trim();
    const txtNorm = isDOI
      ? (normalizeHex('#' + (tcHexInp ? tcHexInp.value : '')) || '#FFFFFF')
      : null;

    if (type === 'edit' && idx !== null) {
      updateColor(key, idx, norm, name || norm, txtNorm);
    } else {
      addColor(key, norm, name, txtNorm);
    }
    _openForm = null;
    cleanup();
    renderEditPanel();
  });

  form.querySelector('.cform-cancel').addEventListener('click', () => {
    _openForm = null;
    cleanup();
    renderEditPanel();
  });

  nameInp.addEventListener('keydown', e => {
    if (e.key === 'Enter') form.querySelector('.cform-save').click();
  });

  setTimeout(() => { if (hexInp.isConnected) hexInp.select(); }, 50);
}

// ─────────────────────────────────────────────────────────────────
// Apply color and close edit panel
// textColor is only passed for DOI (fill) entries
// ─────────────────────────────────────────────────────────────────
function applyAndClose(color, textColor) {
  if (_applyFn) _applyFn(color);
  // DOI: also apply paired text color
  if (textColor && _editKey === 'fill') applyTextColor(textColor);
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
      _cachedShapeCount = items.length;

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

      renderUI({ merged, meta, firstName, count: items.length, anySupported, isLine: allSame && firstKey === 'line' });
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
  show('empty-state'); hide('shape-banner'); hide('unsupported-state');
  ['section-fill', 'section-border', 'section-line-ends', 'section-text'].forEach(hide);
  setStatus('');
}

function el(id)       { return document.getElementById(id); }
function show(id)     { el(id)?.classList.remove('hidden'); }
function hide(id)     { el(id)?.classList.add('hidden'); }
function tog(id, vis) { el(id)?.classList.toggle('hidden', !vis); }
function setStatus(m) { if (el('status')) el('status').textContent = m; }

// ─────────────────────────────────────────────────────────────────
// Apply functions — single sync, version-gated
// ─────────────────────────────────────────────────────────────────
function getSelectedFast(ctx) {
  const col = ctx.presentation.getSelectedShapes();
  const shapes = [];
  for (let i = 0; i < _cachedShapeCount; i++) shapes.push(col.getItemAt(i));
  return shapes;
}

async function applyFill(color) {
  const ver = ++_applyVer;
  if (!_cachedShapeCount) return setStatus('No shape selected.');
  await PowerPoint.run(async ctx => {
    if (ver !== _applyVer) return;
    getSelectedFast(ctx).forEach(s => {
      if (color === 'none') s.fill.clear(); else s.fill.setSolidColor(color);
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
  const sv  = el('line-start').value;
  const ev  = el('line-end').value;
  const ver = ++_applyVer;
  if (!_cachedShapeCount) return setStatus('No shape selected.');
  await PowerPoint.run(async ctx => {
    if (ver !== _applyVer) return;
    getSelectedFast(ctx).forEach(s => {
      try {
        s.lineFormat.beginArrowheadStyle = PowerPoint.ArrowheadStyle[sv] ?? PowerPoint.ArrowheadStyle.none;
        s.lineFormat.endArrowheadStyle   = PowerPoint.ArrowheadStyle[ev] ?? PowerPoint.ArrowheadStyle.none;
      } catch (_) {}
    });
    await ctx.sync();
    if (ver === _applyVer) setStatus(`Ends → ${sv} / ${ev}`);
  });
}
