'use strict';

// ─────────────────────────────────────────────────────────────────
// Brand / static color data
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
  return { r: parseInt(h.slice(0,2),16), g: parseInt(h.slice(2,4),16), b: parseInt(h.slice(4,6),16) };
}
function rgbToHex(r,g,b) {
  return '#' + [r,g,b].map(x => Math.round(Math.min(255,Math.max(0,x))).toString(16).padStart(2,'0')).join('').toUpperCase();
}
function tint(hex, pct) {
  const {r,g,b} = hexToRgb(hex);
  return rgbToHex(r+(255-r)*pct, g+(255-g)*pct, b+(255-b)*pct);
}
function darken(hex, pct) {
  const {r,g,b} = hexToRgb(hex);
  return rgbToHex(r*(1-pct), g*(1-pct), b*(1-pct));
}
function isLight(hex) {
  const {r,g,b} = hexToRgb(hex);
  return (r*299 + g*587 + b*114)/1000 > 155;
}
function normalizeHex(raw) {
  const h = (raw||'').replace('#','').trim();
  return /^[0-9A-Fa-f]{6}$/.test(h) ? '#'+h.toUpperCase() : null;
}
function esc(s) {
  return (s||'').replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
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
  localStorage.setItem(`sct_recent_${key}`, JSON.stringify(arr.slice(0,10)));
}

// ─────────────────────────────────────────────────────────────────
// Edit panel state
// ─────────────────────────────────────────────────────────────────
let _editKey  = null;
let _applyFn  = null;

// Which inline form is currently open: null | { type:'edit'|'add', idx:number|null }
let _openForm = null;

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
    colors.map(c => `
      <button class="main-swatch${isLight(c.hex) ? ' light' : ''}"
              style="background:${c.hex}"
              data-color="${c.hex}"
              title="${esc(c.name)}\n${c.hex}"
              aria-label="${esc(c.name)}"></button>`
    ).join('') +
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

  // Theme color grid
  const baseRowHtml = BRAND_COLORS.map(c =>
    `<div class="theme-swatch base-swatch${isLight(c.hex) ? ' light':''}"
          style="background:${c.hex}" data-color="${c.hex}"
          title="${esc(c.name)}" role="button" tabindex="0"></div>`
  ).join('');

  const shadeRowsHtml = SHADE_STEPS.map(step =>
    `<div class="theme-row">${BRAND_COLORS.map(c => {
      const hex = step.fn(c.hex);
      return `<div class="theme-swatch${isLight(hex)?' light':''}"
                   style="background:${hex}" data-color="${hex}"
                   title="${esc(c.name)} — ${step.label}" role="button" tabindex="0"></div>`;
    }).join('')}</div>`
  ).join('');

  // Standard colors
  const standardHtml = STANDARD_COLORS.map(c =>
    `<div class="std-swatch${isLight(c.hex)?' light':''}"
          style="background:${c.hex}" data-color="${c.hex}"
          title="${esc(c.name)}" role="button" tabindex="0"></div>`
  ).join('');

  // Recent colors
  const recent = getRecent(key);
  const recentHtml = recent.length ? `
    <div class="ep-label">Recent Colors</div>
    <div class="std-row">${recent.map(hex =>
      `<div class="std-swatch${isLight(hex)?' light':''}"
            style="background:${hex}" data-color="${hex}"
            title="${hex}" role="button" tabindex="0"></div>`
    ).join('')}</div>` : '';

  // My colors — each has edit + delete overlays
  const palette     = getPalette(key);
  const myColorsHtml = palette.map((c, i) => {
    const isEditing = _openForm && _openForm.type === 'edit' && _openForm.idx === i;
    return `
      <div class="my-color-item" data-idx="${i}">
        <div class="my-color-top">
          <div class="my-swatch-wrap">
            <div class="my-swatch${isLight(c.hex)?' light':''}"
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
            <button class="my-edit-btn${isEditing?' active':''}"
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

  // Add new color row + form
  const isAdding    = _openForm && _openForm.type === 'add';
  const addFormHtml = isAdding ? buildColorForm('add', null, '#2A3E6D', '') : '';

  document.getElementById('edit-panel-body').innerHTML = `

    <div class="ep-label">APi Group Theme Colors</div>
    <div class="theme-grid">
      <div class="theme-row base-row">${baseRowHtml}</div>
      ${shadeRowsHtml}
    </div>

    <div class="ep-label">Standard Colors</div>
    <div class="std-row">${standardHtml}</div>

    ${recentHtml}

    <div class="ep-label">
      My Colors
      <span class="ep-hint">click swatch to apply</span>
    </div>
    <div id="my-colors-list">${myColorsHtml}</div>

    <button class="add-new-btn${isAdding?' active':''}" id="ep-add-toggle">
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
  `;

  bindEditEvents(key);
}

// ─────────────────────────────────────────────────────────────────
// Build a color form (used for both edit and add)
// type: 'edit' | 'add'
// idx:  number (for edit) | null (for add)
// ─────────────────────────────────────────────────────────────────
function buildColorForm(type, idx, startHex, startName) {
  const norm    = normalizeHex(startHex) || '#2A3E6D';
  const hexVal  = norm.replace('#','');
  const idSufx  = type === 'edit' ? `edit-${idx}` : 'add';

  return `
    <div class="color-form" id="cform-${idSufx}" data-type="${type}" data-idx="${idx ?? ''}">
      <div class="cform-row">
        <input type="color" class="cform-wheel" id="cfw-${idSufx}" value="${norm}" />
        <div class="hex-field">
          <span class="hex-hash">#</span>
          <input type="text" class="cform-hex ep-hex-input"
                 id="cfh-${idSufx}" value="${hexVal}"
                 maxlength="6" placeholder="RRGGBB" spellcheck="false" />
        </div>
        <div class="cform-preview" id="cfp-${idSufx}" style="background:${norm}"></div>
      </div>
      <input type="text" class="ep-name-input cform-name" id="cfn-${idSufx}"
             placeholder="Color name (optional)" maxlength="28"
             value="${esc(startName)}" />
      <div class="cform-btns">
        <button class="cform-save" data-type="${type}" data-idx="${idx ?? ''}">
          ${type === 'edit' ? 'Save Changes' : '+ Add to My Colors'}
        </button>
        <button class="cform-cancel">Cancel</button>
      </div>
    </div>`;
}

// ─────────────────────────────────────────────────────────────────
// Bind all events inside the edit panel
// ─────────────────────────────────────────────────────────────────
function bindEditEvents(key) {
  const body = document.getElementById('edit-panel-body');

  // ── Theme swatches: apply + close ──
  body.querySelectorAll('.theme-swatch').forEach(s => {
    const apply = () => applyAndClose(s.dataset.color);
    s.addEventListener('click', apply);
    s.addEventListener('keydown', e => { if (e.key==='Enter'||e.key===' '){e.preventDefault();apply();} });
  });

  // ── Standard + recent: apply + close ──
  body.querySelectorAll('.std-swatch').forEach(s => {
    s.addEventListener('click', () => applyAndClose(s.dataset.color));
  });

  // ── My colors: click swatch = apply + close ──
  body.querySelectorAll('.my-swatch').forEach(s => {
    s.addEventListener('click', () => applyAndClose(s.dataset.color));
    s.addEventListener('keydown', e => { if (e.key==='Enter'||e.key===' '){e.preventDefault();applyAndClose(s.dataset.color);} });
  });

  // ── My colors: edit button ──
  body.querySelectorAll('.my-edit-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const idx = parseInt(btn.dataset.idx, 10);
      // Toggle: if this form is already open, close it; otherwise open it
      if (_openForm && _openForm.type === 'edit' && _openForm.idx === idx) {
        _openForm = null;
      } else {
        _openForm = { type: 'edit', idx };
      }
      renderEditPanel();
    });
  });

  // ── My colors: delete button ──
  body.querySelectorAll('.my-del-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      deleteColor(key, parseInt(btn.dataset.idx, 10));
      _openForm = null;
      renderEditPanel();
    });
  });

  // ── Add new color toggle ──
  document.getElementById('ep-add-toggle').addEventListener('click', () => {
    _openForm = (_openForm && _openForm.type === 'add') ? null : { type: 'add', idx: null };
    renderEditPanel();
  });

  // ── No color button ──
  document.getElementById('ep-no-color').addEventListener('click', () => applyAndClose('none'));

  // ── Wire up all open color forms ──
  body.querySelectorAll('.color-form').forEach(form => {
    wireColorForm(form, key);
  });
}

// ─────────────────────────────────────────────────────────────────
// Wire a single color form's inputs + buttons
// ─────────────────────────────────────────────────────────────────
function wireColorForm(form, key) {
  const type   = form.dataset.type;
  const idx    = form.dataset.idx !== '' ? parseInt(form.dataset.idx, 10) : null;
  const sufx   = type === 'edit' ? `edit-${idx}` : 'add';

  const wheel   = document.getElementById(`cfw-${sufx}`);
  const hexInp  = document.getElementById(`cfh-${sufx}`);
  const preview = document.getElementById(`cfp-${sufx}`);
  const nameInp = document.getElementById(`cfn-${sufx}`);

  // Sync wheel → hex + preview
  wheel.addEventListener('input', () => {
    hexInp.value = wheel.value.replace('#','').toUpperCase();
    preview.style.background = wheel.value;
  });

  // Sync hex → wheel + preview
  hexInp.addEventListener('input', () => {
    const norm = normalizeHex(hexInp.value);
    if (norm) { wheel.value = norm; preview.style.background = norm; }
  });

  // Save / Add button
  form.querySelector('.cform-save').addEventListener('click', () => {
    const norm = normalizeHex(hexInp.value);
    if (!norm) { hexInp.focus(); return; }
    const name = nameInp.value.trim();

    if (type === 'edit' && idx !== null) {
      updateColor(key, idx, norm, name || norm);
    } else {
      addColor(key, norm, name);
    }
    _openForm = null;
    renderEditPanel();
  });

  // Cancel
  form.querySelector('.cform-cancel').addEventListener('click', () => {
    _openForm = null;
    renderEditPanel();
  });

  // Enter key in name field = save
  nameInp.addEventListener('keydown', e => {
    if (e.key === 'Enter') form.querySelector('.cform-save').click();
  });

  // Auto-focus hex input
  setTimeout(() => hexInp.focus(), 40);
}

// ─────────────────────────────────────────────────────────────────
// Apply and close
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

      const merged   = { fill:false, border:false, text:false, lineEnds:false };
      let   allSame  = true;
      const firstKey = toKey(items[0].type);
      const firstName= items[0].name || '';

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
        ? (TYPE_META[firstKey] || { icon:'?', label: cap(firstKey||'Shape') })
        : { icon:'⊕', label:'Mixed Selection' };

      renderUI({ merged, meta, firstName, count:items.length, anySupported, isLine: allSame && firstKey==='line' });
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
  ['section-fill','section-border','section-line-ends','section-text'].forEach(hide);
  setStatus('');
}

function el(id)       { return document.getElementById(id); }
function show(id)     { el(id)?.classList.remove('hidden'); }
function hide(id)     { el(id)?.classList.add('hidden'); }
function tog(id, vis) { el(id)?.classList.toggle('hidden', !vis); }
function setStatus(m) { el('status').textContent = m; }

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
    shapes.forEach(s => { if (color==='none') s.fill.clear(); else s.fill.setSolidColor(color); });
    await ctx.sync();
    setStatus(`Fill → ${color==='none'?'cleared':color}`);
  });
}

async function applyBorderColor(color) {
  await PowerPoint.run(async ctx => {
    const shapes = await getSelected(ctx);
    if (!shapes.length) return setStatus('No shape selected.');
    shapes.forEach(s => {
      if (color==='none') { s.lineFormat.visible=false; }
      else { s.lineFormat.visible=true; s.lineFormat.color=color; }
    });
    await ctx.sync();
    setStatus(`Border → ${color==='none'?'removed':color}`);
  });
}

async function applyBorderWeight(pts) {
  await PowerPoint.run(async ctx => {
    const shapes = await getSelected(ctx);
    if (!shapes.length) return setStatus('No shape selected.');
    shapes.forEach(s => { s.lineFormat.weight=pts; s.lineFormat.visible=true; });
    await ctx.sync();
    setStatus(`Weight → ${pts}pt`);
  });
}

async function applyBorderDash(style) {
  const map = { solid:PowerPoint.LineDashStyle.solid, dash:PowerPoint.LineDashStyle.dash, dot:PowerPoint.LineDashStyle.dot };
  await PowerPoint.run(async ctx => {
    const shapes = await getSelected(ctx);
    if (!shapes.length) return setStatus('No shape selected.');
    shapes.forEach(s => { s.lineFormat.dashStyle=map[style]??PowerPoint.LineDashStyle.solid; s.lineFormat.visible=true; });
    await ctx.sync();
    setStatus(`Style → ${style}`);
  });
}

async function applyTextColor(color) {
  await PowerPoint.run(async ctx => {
    const shapes = await getSelected(ctx);
    if (!shapes.length) return setStatus('No shape selected.');
    shapes.forEach(s => { try { s.textFrame.textRange.font.color=color; } catch(_){} });
    await ctx.sync();
    setStatus(`Text → ${color}`);
  });
}

async function applyLineEnds() {
  const aMap = { none:PowerPoint.ArrowheadStyle.none, arrow:PowerPoint.ArrowheadStyle.arrow,
                 openArrow:PowerPoint.ArrowheadStyle.openArrow, diamond:PowerPoint.ArrowheadStyle.diamond,
                 oval:PowerPoint.ArrowheadStyle.oval };
  const sv = el('line-start').value, ev = el('line-end').value;
  await PowerPoint.run(async ctx => {
    const shapes = await getSelected(ctx);
    if (!shapes.length) return setStatus('No shape selected.');
    shapes.forEach(s => {
      if (toKey(s.type)!=='line') return;
      try { s.lineFormat.beginArrowheadStyle=aMap[sv]??PowerPoint.ArrowheadStyle.none;
            s.lineFormat.endArrowheadStyle  =aMap[ev]??PowerPoint.ArrowheadStyle.none; } catch(_){}
    });
    await ctx.sync();
    setStatus(`Ends → ${sv} / ${ev}`);
  });
}
