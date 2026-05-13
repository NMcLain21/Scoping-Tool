'use strict';

// ─────────────────────────────────────────────────────────────────
// Brand / static color data
// ─────────────────────────────────────────────────────────────────
const BRAND_COLORS = [
  { hex: '#D50032', name: 'APi Corp Red' },
  { hex: '#2A3E6D', name: 'Dark Navy'    },
  { hex: '#008579', name: 'Teal'         },
  { hex: '#D4D800', name: 'TBD Yellow'   },
  { hex: '#7030A0', name: 'APi Segment'  },
  { hex: '#00B0F0', name: 'Target HQ'    },
  { hex: '#92D050', name: 'Target OpCo'  },
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

const autoTextHex = hex => isLight(hex) ? '#000000' : '#FFFFFF';

const DEFAULT_PALETTES = {
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
    { hex: '#BFBFBF', name: 'Light Gray'  },
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
function rgbToHex(r, g, b) {
  return '#' + [r,g,b].map(x => Math.round(Math.min(255,Math.max(0,x))).toString(16).padStart(2,'0')).join('').toUpperCase();
}
function isLight(hex) {
  const { r, g, b } = hexToRgb(hex);
  return (r * 299 + g * 587 + b * 114) / 1000 > 155;
}
function normalizeHex(raw) {
  const h = (raw || '').replace('#','').trim();
  return /^[0-9A-Fa-f]{6}$/.test(h) ? '#' + h.toUpperCase() : null;
}
function esc(s) {
  return (s||'').replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

// Tint: mix hex toward white by factor (0–1)
function tintHex(hex, factor) {
  const { r, g, b } = hexToRgb(hex);
  return rgbToHex(
    Math.round(r + (255 - r) * factor),
    Math.round(g + (255 - g) * factor),
    Math.round(b + (255 - b) * factor)
  );
}

// Shade: darken hex by factor (0–1)
function shadeHex(hex, factor) {
  const { r, g, b } = hexToRgb(hex);
  return rgbToHex(
    Math.round(r * (1 - factor)),
    Math.round(g * (1 - factor)),
    Math.round(b * (1 - factor))
  );
}

// ─────────────────────────────────────────────────────────────────
// HSV utilities
// ─────────────────────────────────────────────────────────────────
function hsvToRgb(h, s, v) {
  const c = v*s, x = c*(1-Math.abs((h/60)%2-1)), m = v-c;
  let r=0, g=0, b=0;
  if      (h<60)  { r=c; g=x; b=0; }
  else if (h<120) { r=x; g=c; b=0; }
  else if (h<180) { r=0; g=c; b=x; }
  else if (h<240) { r=0; g=x; b=c; }
  else if (h<300) { r=x; g=0; b=c; }
  else            { r=c; g=0; b=x; }
  return { r: Math.round((r+m)*255), g: Math.round((g+m)*255), b: Math.round((b+m)*255) };
}
function rgbToHsv(r, g, b) {
  r/=255; g/=255; b/=255;
  const max=Math.max(r,g,b), min=Math.min(r,g,b), d=max-min;
  let h=0;
  const s=max===0?0:d/max, v=max;
  if (d!==0) {
    if      (max===r) h=((g-b)/d+(g<b?6:0))*60;
    else if (max===g) h=((b-r)/d+2)*60;
    else              h=((r-g)/d+4)*60;
  }
  return { h: h<0?h+360:h, s, v };
}
function hexToHsv(hex) { const {r,g,b}=hexToRgb(hex); return rgbToHsv(r,g,b); }
function hsvToHex(h,s,v) { const {r,g,b}=hsvToRgb(h,s,v); return rgbToHex(r,g,b); }

// ─────────────────────────────────────────────────────────────────
// Storage
// ─────────────────────────────────────────────────────────────────
function getPalette(key) {
  try {
    const s   = localStorage.getItem(`sct_${key}`);
    const arr = s ? JSON.parse(s) : JSON.parse(JSON.stringify(DEFAULT_PALETTES[key]));
    if (key==='fill') arr.forEach(c => { if (!c.textHex) c.textHex = autoTextHex(c.hex); });
    return arr;
  } catch { return JSON.parse(JSON.stringify(DEFAULT_PALETTES[key])); }
}
function savePalette(key, arr) { localStorage.setItem(`sct_${key}`, JSON.stringify(arr)); }

function updateColor(key, idx, hex, name, textHex) {
  const arr = getPalette(key);
  if (!arr[idx]) return;
  arr[idx] = { hex, name: name||hex };
  if (key==='fill') arr[idx].textHex = textHex || autoTextHex(hex);
  savePalette(key, arr);
}
function addColor(key, hex, name, textHex) {
  const norm = normalizeHex(hex);
  if (!norm) return;
  const arr   = getPalette(key);
  const entry = { hex: norm, name: name||norm };
  if (key==='fill') entry.textHex = textHex || autoTextHex(norm);
  arr.push(entry);
  savePalette(key, arr);
}
function deleteColor(key, idx) {
  const arr = getPalette(key); arr.splice(idx,1); savePalette(key,arr);
}
function getRecent(key) {
  try { return JSON.parse(localStorage.getItem(`sct_recent_${key}`))||[]; } catch { return []; }
}
function pushRecent(key, hex) {
  if (!hex||hex==='none') return;
  const norm = normalizeHex(hex)||hex.toUpperCase();
  let arr    = getRecent(key).filter(h=>h!==norm);
  arr.unshift(norm);
  localStorage.setItem(`sct_recent_${key}`, JSON.stringify(arr.slice(0,10)));
}

// ─────────────────────────────────────────────────────────────────
// Edit panel state
// ─────────────────────────────────────────────────────────────────
let _editKey  = null;
let _applyFn  = null;
let _openForm = null;

// ─────────────────────────────────────────────────────────────────
// Performance state
// ─────────────────────────────────────────────────────────────────
let _cachedShapeCount = 0;
let _applyVer         = 0;

// ─────────────────────────────────────────────────────────────────
// Build the PowerPoint-style 7×6 theme grid
// ─────────────────────────────────────────────────────────────────
function buildThemeGrid(dataAttr) {
  const rows = [
    BRAND_COLORS.map(c => c.hex),                          // Row 0: base
    BRAND_COLORS.map(c => tintHex(c.hex, 0.8)),           // Row 1: 80% lighter
    BRAND_COLORS.map(c => tintHex(c.hex, 0.6)),           // Row 2: 60% lighter
    BRAND_COLORS.map(c => tintHex(c.hex, 0.4)),           // Row 3: 40% lighter
    BRAND_COLORS.map(c => shadeHex(c.hex, 0.25)),         // Row 4: 25% darker
    BRAND_COLORS.map(c => shadeHex(c.hex, 0.5)),          // Row 5: 50% darker
  ];
  return `<div class="ppt-theme-grid">${
    rows.map((row, ri) =>
      `<div class="ppt-theme-row${ri===0?' ppt-base-row':''}">${
        row.map(hex =>
          `<button class="ppt-theme-swatch${isLight(hex)?' light':''}"
                   style="background:${hex}" ${dataAttr}="${hex}"
                   title="${hex}" type="button"></button>`
        ).join('')
      }</div>`
    ).join('')
  }</div>`;
}

// ─────────────────────────────────────────────────────────────────
// Build a single PowerPoint-style picker panel (collapsible sections)
// ─────────────────────────────────────────────────────────────────
function buildPptPicker(pickerType, sufx, startHex, dataAttr, recentKey) {
  const norm   = normalizeHex(startHex) || '#2A3E6D';
  const hexVal = norm.replace('#','');
  const recent = getRecent(recentKey);
  const moreId = `mc-${pickerType}-${sufx}`;

  const stdRow = STANDARD_COLORS.map(c =>
    `<button class="ppt-theme-swatch${isLight(c.hex)?' light':''}"
             style="background:${c.hex}" ${dataAttr}="${c.hex}"
             title="${esc(c.name)}" type="button"></button>`
  ).join('');

  const recentHtml = recent.length ? `
    <div class="acc-section">
      <button class="acc-hdr" data-acc="recent-${pickerType}-${sufx}" type="button">
        <span>Recent Colors</span>
        <svg class="acc-chevron" width="10" height="10" viewBox="0 0 10 10" fill="none">
          <path d="M2 3.5l3 3 3-3" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
      </button>
      <div class="acc-body hidden" id="acc-recent-${pickerType}-${sufx}">
        <div class="ppt-std-row" style="padding:4px 0 2px;">${
          recent.slice(0,10).map(hex =>
            `<button class="ppt-theme-swatch${isLight(hex)?' light':''}"
                     style="background:${hex}" ${dataAttr}="${hex}"
                     title="${hex}" type="button"></button>`
          ).join('')
        }</div>
      </div>
    </div>` : '';

  return `
    <div class="ppt-picker" data-picker="${pickerType}">

      <div class="acc-section">
        <button class="acc-hdr acc-open" data-acc="theme-${pickerType}-${sufx}" type="button">
          <span>Theme Colors</span>
          <svg class="acc-chevron acc-chevron-open" width="10" height="10" viewBox="0 0 10 10" fill="none">
            <path d="M2 3.5l3 3 3-3" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/>
          </svg>
        </button>
        <div class="acc-body" id="acc-theme-${pickerType}-${sufx}">
          <div style="padding:4px 0 2px;">${buildThemeGrid(dataAttr)}</div>
        </div>
      </div>

      <div class="acc-section">
        <button class="acc-hdr" data-acc="std-${pickerType}-${sufx}" type="button">
          <span>Standard Colors</span>
          <svg class="acc-chevron" width="10" height="10" viewBox="0 0 10 10" fill="none">
            <path d="M2 3.5l3 3 3-3" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/>
          </svg>
        </button>
        <div class="acc-body hidden" id="acc-std-${pickerType}-${sufx}">
          <div class="ppt-std-row" style="padding:4px 0 2px;">${stdRow}</div>
        </div>
      </div>

      ${recentHtml}

      <div class="acc-section">
        <button class="acc-hdr" data-acc="${moreId}" type="button"
                id="more-btn-${pickerType}-${sufx}">
          <span>Custom Color</span>
          <svg class="acc-chevron" width="10" height="10" viewBox="0 0 10 10" fill="none">
            <path d="M2 3.5l3 3 3-3" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/>
          </svg>
        </button>
        <div class="acc-body hidden" id="${moreId}">
          <div style="padding:4px 0 2px;">
            <div class="hsv-square" id="hsv-sq-${pickerType}-${sufx}">
              <div class="hsv-layer hsv-sat"></div>
              <div class="hsv-layer hsv-val"></div>
              <div class="hsv-dot" id="hsv-dot-${pickerType}-${sufx}"></div>
            </div>
            <div class="hue-wrap">
              <input type="range" class="hue-slider" id="hue-${pickerType}-${sufx}"
                     min="0" max="359" step="1" value="0"/>
            </div>
            <div class="cform-row">
              <div class="hex-field">
                <span class="hex-hash">#</span>
                <input type="text" class="ep-hex-input" id="hex-${pickerType}-${sufx}"
                       value="${hexVal}" maxlength="6"
                       placeholder="RRGGBB" spellcheck="false" autocomplete="off"/>
              </div>
              <div class="cform-preview${isLight(norm)?' light':''}"
                   id="prev-${pickerType}-${sufx}" style="background:${norm};"></div>
            </div>
          </div>
        </div>
      </div>

    </div>`;
}

// ─────────────────────────────────────────────────────────────────
// Build color form — flat PowerPoint-style pickers
// ─────────────────────────────────────────────────────────────────
function buildColorForm(type, idx, startHex, startName, startTextHex) {
  const norm = normalizeHex(startHex) || '#2A3E6D';
  const sufx = type==='edit' ? `edit-${idx}` : 'add';

  const fillPicker = buildPptPicker('fill', sufx, norm, 'data-setfill', _editKey);

  return `
    <div class="color-form" id="cform-${sufx}" data-type="${type}" data-idx="${idx??''}">

      <input type="text" class="ep-name-input" id="cfn-${sufx}"
             placeholder="Color name (optional)" maxlength="36"
             value="${esc(startName)}" />

      <div class="ppt-sect-hdr">
        <svg class="ppt-fill-svg" id="fill-svg-${sufx}" width="16" height="16" viewBox="0 0 16 16" fill="none">
          <path d="M3 12.5c0-.9.8-1.7 1.7-1.7s1.7.8 1.7 1.7S4.7 15 4.7 15 3 13.4 3 12.5z"
                fill="${norm}" stroke="${isLight(norm)?'#aaa':norm}" stroke-width="0.8"/>
          <path d="M5.5 10.7l5.8-5.8a1.1 1.1 0 011.6 0l.5.5a1.1 1.1 0 010 1.6L7.5 12"
                stroke="#666" stroke-width="1.2" stroke-linecap="round" fill="none"/>
        </svg>
        <span class="ppt-sect-label">Shape Fill</span>
        <div class="ppt-curr-swatch${isLight(norm)?' light':''}"
             id="fill-swatch-${sufx}" style="background:${norm};"></div>
      </div>
      ${fillPicker}

      <div class="cform-btns">
        <button class="cform-save" type="button" data-type="${type}" data-idx="${idx??''}">
          ${type==='edit' ? 'Save Changes' : '+ Add to My Colors'}
        </button>
        <button class="cform-cancel" type="button">Cancel</button>
      </div>
    </div>`;
}

// ─────────────────────────────────────────────────────────────────
// Wire a single picker's interactions
// Returns a cleanup function to remove global mouse listeners
// ─────────────────────────────────────────────────────────────────
function wirePicker(form, pickerType, sufx, onColorChange) {
  const moreId    = `mc-${pickerType}-${sufx}`;
  const morePanel = document.getElementById(moreId);
  const moreBtn   = document.getElementById(`more-btn-${pickerType}-${sufx}`);
  const hexInp    = document.getElementById(`hex-${pickerType}-${sufx}`);
  const preview   = document.getElementById(`prev-${pickerType}-${sufx}`);
  const sq        = document.getElementById(`hsv-sq-${pickerType}-${sufx}`);
  const dot       = document.getElementById(`hsv-dot-${pickerType}-${sufx}`);
  const hueSlider = document.getElementById(`hue-${pickerType}-${sufx}`);
  const dataAttr  = pickerType==='fill' ? 'data-setfill' : 'data-settextcolor';

  // HSV state
  const startNorm = normalizeHex('#' + (hexInp?.value||'')) || '#2A3E6D';
  const startHsv  = hexToHsv(startNorm);
  let H = startHsv.h, S = startHsv.s, V = startHsv.v;

  function syncFromHsv() {
    const hex = hsvToHex(H, S, V);
    if (hexInp)  hexInp.value             = hex.replace('#','').toUpperCase();
    if (preview) preview.style.background = hex;
    if (preview) preview.classList.toggle('light', isLight(hex));
    if (sq) {
      sq.style.background = `hsl(${Math.round(H)},100%,50%)`;
      if (dot) { dot.style.left = `${S*100}%`; dot.style.top = `${(1-V)*100}%`; }
    }
    if (hueSlider) hueSlider.value = Math.round(H);
    onColorChange(hex);
  }

  function setFromHex(rawHex) {
    const c = rawHex.startsWith('#') ? rawHex : '#'+rawHex;
    const n = normalizeHex(c);
    if (!n) return;
    const hsv = hexToHsv(n);
    H = hsv.h; S = hsv.s; V = hsv.v;
    if (hexInp)  hexInp.value             = n.replace('#','').toUpperCase();
    if (preview) preview.style.background = n;
    if (preview) preview.classList.toggle('light', isLight(n));
    if (sq) {
      sq.style.background = `hsl(${Math.round(H)},100%,50%)`;
      if (dot) { dot.style.left = `${S*100}%`; dot.style.top = `${(1-V)*100}%`; }
    }
    if (hueSlider) hueSlider.value = Math.round(H);
    onColorChange(n);
  }

  // Quick-pick swatches (theme grid + standard + recent)
  const pickerEl = form.querySelector(`.ppt-picker[data-picker="${pickerType}"]`);
  pickerEl?.querySelectorAll(`[${dataAttr}]`).forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      setFromHex(btn.getAttribute(dataAttr));
      pickerEl.querySelectorAll(`[${dataAttr}]`).forEach(b => b.classList.remove('qp-selected'));
      btn.classList.add('qp-selected');
    });
  });

  // Accordion headers inside this picker
  const pickerContainer = form.querySelector(`.ppt-picker[data-picker="${pickerType}"]`);
  pickerContainer?.querySelectorAll('.acc-hdr').forEach(hdr => {
    hdr.addEventListener('click', () => {
      const bodyId = hdr.dataset.acc;
      const body   = document.getElementById(bodyId);
      if (!body) return;
      const opening = body.classList.contains('hidden');
      body.classList.toggle('hidden', !opening);
      hdr.classList.toggle('acc-open', opening);
      hdr.querySelector('.acc-chevron')?.classList.toggle('acc-chevron-open', opening);
      // Initialize HSV square on first open of Custom Color panel
      if (opening && bodyId === moreId) syncFromHsv();
    });
  });

  // Hue slider
  hueSlider?.addEventListener('input', () => { H = parseFloat(hueSlider.value); syncFromHsv(); });

  // Square drag
  let dragging = false;
  function handleDrag(e) {
    if (!sq?.isConnected) { dragging = false; return; }
    const rect = sq.getBoundingClientRect();
    const cx   = e.touches ? e.touches[0].clientX : e.clientX;
    const cy   = e.touches ? e.touches[0].clientY : e.clientY;
    S = Math.max(0, Math.min(1, (cx - rect.left) / rect.width));
    V = Math.max(0, Math.min(1, 1 - (cy - rect.top) / rect.height));
    syncFromHsv();
  }
  if (sq) {
    sq.addEventListener('mousedown',  e => { e.preventDefault(); dragging=true; handleDrag(e); });
    sq.addEventListener('touchstart', e => { e.preventDefault(); handleDrag(e); }, { passive: false });
    sq.addEventListener('touchmove',  e => { e.preventDefault(); handleDrag(e); }, { passive: false });
  }
  const onMove = e => { if (dragging) handleDrag(e); };
  const onUp   = ()  => { dragging = false; };
  document.addEventListener('mousemove', onMove);
  document.addEventListener('mouseup',   onUp);

  // Hex input
  hexInp?.addEventListener('input', () => { if (hexInp.value.length===6) setFromHex(hexInp.value); });

  // Initialize square on load (More panel is hidden; still init state)
  if (sq) syncFromHsv();

  return () => {
    document.removeEventListener('mousemove', onMove);
    document.removeEventListener('mouseup',   onUp);
  };
}

// ─────────────────────────────────────────────────────────────────
// Wire the whole color form
// ─────────────────────────────────────────────────────────────────
function wireColorForm(form, key) {
  const type = form.dataset.type;
  const idx  = form.dataset.idx !== '' ? parseInt(form.dataset.idx, 10) : null;
  const sufx = type==='edit' ? `edit-${idx}` : 'add';

  // Track current fill selection
  let currentFill = normalizeHex('#'+(document.getElementById(`hex-fill-${sufx}`)?.value||'')) || '#2A3E6D';

  // Fill color change callback — update header swatch + icon
  function onFillChange(hex) {
    currentFill = hex;
    const sw = document.getElementById(`fill-swatch-${sufx}`);
    if (sw) { sw.style.background = hex; sw.classList.toggle('light', isLight(hex)); }
    const sv = document.getElementById(`fill-svg-${sufx}`);
    if (sv) {
      const p = sv.querySelector('path');
      if (p) { p.setAttribute('fill', hex); p.setAttribute('stroke', isLight(hex)?'#aaa':hex); }
    }
  }

  // Wire fill picker
  const cleanFill = wirePicker(form, 'fill', sufx, onFillChange);

  const nameInp = document.getElementById(`cfn-${sufx}`);

  // Save
  form.querySelector('.cform-save')?.addEventListener('click', () => {
    const norm = normalizeHex(currentFill);
    if (!norm) return;
    const name    = nameInp?.value.trim() || '';
    const txtNorm = key==='fill' ? autoTextHex(norm) : null;
    if (type==='edit' && idx !== null) updateColor(key, idx, norm, name||norm, txtNorm);
    else addColor(key, norm, name, txtNorm);
    _openForm = null;
    cleanFill?.();
    renderEditPanel();
  });

  // Cancel
  form.querySelector('.cform-cancel')?.addEventListener('click', () => {
    _openForm = null;
    cleanFill?.();
    renderEditPanel();
  });

  nameInp?.addEventListener('keydown', e => {
    if (e.key==='Enter') form.querySelector('.cform-save')?.click();
  });

  setTimeout(() => { if (nameInp?.isConnected) nameInp.select(); }, 50);
}

// ─────────────────────────────────────────────────────────────────
// Main panel — labeled tiles + pencil button
// ─────────────────────────────────────────────────────────────────
function renderMainSwatches(key) {
  const row = document.getElementById(`swatches-${key}`);
  if (!row) return;
  const colors = getPalette(key);

  row.innerHTML =
    colors.map(c => {
      if (key === 'border') {
        return `<button class="main-swatch border-legend-tile"
                         data-color="${c.hex}" title="${esc(c.name)}" aria-label="${esc(c.name)}">
                  <span class="legend-line-sample" style="background:${c.hex}"></span>
                  <span class="legend-line-name">${esc(c.name)}</span>
                </button>`;
      }
      if (key === 'fill') {
        const txHex = c.textHex || autoTextHex(c.hex);
        return `<button class="main-swatch doi-tile"
                         style="background:${c.hex};"
                         data-color="${c.hex}" data-textcolor="${txHex}"
                         title="${esc(c.name)}" aria-label="${esc(c.name)}">
                  <span class="swatch-label" style="color:${txHex};">${esc(c.name)}</span>
                </button>`;
      }
      const txtColor = isLight(c.hex) ? '#000000' : '#FFFFFF';
      return `<button class="main-swatch${isLight(c.hex)?' light':''}"
                       style="background:${c.hex}"
                       data-color="${c.hex}" title="${esc(c.name)}" aria-label="${esc(c.name)}">
                <span class="swatch-label" style="color:${txtColor};">${esc(c.name)}</span>
              </button>`;
    }).join('') +
    `<button class="edit-palette-btn" data-key="${key}" title="Edit ${key} palette">
       <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
         <path d="M8.5 1.5a1.414 1.414 0 012 2L4 10H2V8L8.5 1.5z"
               stroke="currentColor" stroke-width="1.3"
               stroke-linecap="round" stroke-linejoin="round" fill="none"/>
       </svg>
     </button>`;

  row.querySelectorAll('.main-swatch').forEach(btn => {
    btn.addEventListener('click', () => {
      const color    = btn.dataset.color;
      const txtColor = btn.dataset.textcolor;
      if (key==='fill' && txtColor) applyDOITile(color, txtColor);
      else ({ border: applyBorderColor, text: applyTextColor })[key]?.(color);
      pushRecent(key, color);
    });
  });

  row.querySelector('.edit-palette-btn').addEventListener('click', () => {
    const fnMap = { fill: applyDOITile, border: applyBorderColor, text: applyTextColor };
    openEditPanel(key, fnMap[key]);
  });
}

// ─────────────────────────────────────────────────────────────────
// Edit panel — open / close
// ─────────────────────────────────────────────────────────────────
function openEditPanel(key, fn) {
  _editKey = key; _applyFn = fn; _openForm = null;
  renderEditPanel();
  document.getElementById('main-content').classList.add('hidden');
  document.getElementById('edit-panel').classList.remove('hidden');
}

function closeEditPanel() {
  document.getElementById('edit-panel').classList.add('hidden');
  document.getElementById('main-content').classList.remove('hidden');
  renderMainSwatches(_editKey);
  _editKey = null; _applyFn = null; _openForm = null;
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
    const isEditing = _openForm && _openForm.type==='edit' && _openForm.idx===i;
    const txHex     = c.textHex || autoTextHex(c.hex);
    return `
      <div class="my-color-item" data-idx="${i}">
        <div class="my-color-top">
          <div class="my-swatch-wrap">
            <div class="my-swatch${isLight(c.hex)?' light':''}"
                 style="background:${c.hex}"
                 data-color="${c.hex}" data-textcolor="${txHex}"
                 role="button" tabindex="0" aria-label="Apply ${esc(c.name)}"></div>
          </div>
          <div class="my-color-info">
            <span class="my-color-name">${esc(c.name)}</span>
            <span class="my-color-hex">${c.hex}${key==='fill'?` · T: ${txHex}`:''}</span>
          </div>
          <div class="my-color-actions">
            <button class="my-edit-btn${isEditing?' active':''}" data-idx="${i}" title="Edit">
              <svg width="11" height="11" viewBox="0 0 12 12" fill="none">
                <path d="M8.5 1.5a1.414 1.414 0 012 2L4 10H2V8L8.5 1.5z"
                      stroke="currentColor" stroke-width="1.4"
                      stroke-linecap="round" stroke-linejoin="round" fill="none"/>
              </svg>
            </button>
            <button class="my-del-btn" data-idx="${i}" title="Delete">
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

  const isAdding = _openForm && _openForm.type==='add';

  document.getElementById('edit-panel-body').innerHTML = `
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
    ${isAdding ? buildColorForm('add', null, '#2A3E6D', '', '#FFFFFF') : ''}

    <button class="no-color-btn" id="ep-no-color">
      <svg width="13" height="13" viewBox="0 0 13 13" fill="none">
        <rect x="1" y="1" width="11" height="11" rx="2" stroke="currentColor" stroke-width="1.2"/>
        <line x1="2" y1="11" x2="11" y2="2" stroke="currentColor" stroke-width="1.2"/>
      </svg>
      ${NO_LABEL[key]||'No Color'}
    </button>

    <button class="reset-defaults-btn" id="ep-reset">↺ Reset to defaults</button>
  `;

  bindEditEvents(key);
}

// ─────────────────────────────────────────────────────────────────
// Bind events inside the edit panel
// ─────────────────────────────────────────────────────────────────
function bindEditEvents(key) {
  const body = document.getElementById('edit-panel-body');

  // My Colors: click swatch → apply + close
  body.querySelectorAll('.my-swatch').forEach(s => {
    const go = () => {
      const color = s.dataset.color, tc = s.dataset.textcolor;
      if (key==='fill'&&tc) applyDOITile(color,tc); else _applyFn?.(color);
      if (color!=='none') pushRecent(key, color);
      closeEditPanel();
    };
    s.addEventListener('click', go);
    s.addEventListener('keydown', e => { if (e.key==='Enter'||e.key===' '){e.preventDefault();go();} });
  });

  // My Colors: pencil
  body.querySelectorAll('.my-edit-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const idx = parseInt(btn.dataset.idx, 10);
      _openForm = (_openForm&&_openForm.type==='edit'&&_openForm.idx===idx)
        ? null : { type:'edit', idx };
      renderEditPanel();
    });
  });

  // My Colors: delete
  body.querySelectorAll('.my-del-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      deleteColor(key, parseInt(btn.dataset.idx, 10));
      _openForm = null; renderEditPanel();
    });
  });

  // Add New Color toggle
  document.getElementById('ep-add-toggle').addEventListener('click', () => {
    _openForm = (_openForm&&_openForm.type==='add') ? null : { type:'add', idx:null };
    renderEditPanel();
  });

  // No Color
  document.getElementById('ep-no-color').addEventListener('click', () => {
    if (key==='fill') applyDOITile('none',null); else _applyFn?.('none');
    closeEditPanel();
  });

  // Reset to defaults
  document.getElementById('ep-reset').addEventListener('click', () => {
    if (!confirm(`Reset ${key} palette to defaults? Your custom colors will be lost.`)) return;
    localStorage.removeItem(`sct_${key}`);
    _openForm = null; renderEditPanel(); renderMainSwatches(key);
  });

  // Wire all open color forms
  body.querySelectorAll('.color-form').forEach(form => wireColorForm(form, key));
}

// ─────────────────────────────────────────────────────────────────
// Apply functions
// ─────────────────────────────────────────────────────────────────
function getSelectedFast(ctx) {
  const col = ctx.presentation.getSelectedShapes();
  const out = [];
  for (let i=0; i<_cachedShapeCount; i++) out.push(col.getItemAt(i));
  return out;
}

async function applyDOITile(fillColor, textColor) {
  const ver = ++_applyVer;
  if (!_cachedShapeCount) return setStatus('No shape selected.');
  await PowerPoint.run(async ctx => {
    if (ver !== _applyVer) return;
    getSelectedFast(ctx).forEach(s => {
      if (fillColor==='none') s.fill.clear(); else s.fill.setSolidColor(fillColor);
      if (textColor && textColor!=='none') {
        try { s.textFrame.textRange.font.color = textColor; } catch(_) {}
      }
    });
    await ctx.sync();
    if (ver===_applyVer) setStatus(`DOI → ${fillColor==='none'?'cleared':fillColor}`);
  });
}

async function applyBorderColor(color) {
  const ver = ++_applyVer;
  if (!_cachedShapeCount) return setStatus('No shape selected.');
  await PowerPoint.run(async ctx => {
    if (ver!==_applyVer) return;
    getSelectedFast(ctx).forEach(s => {
      if (color==='none') s.lineFormat.visible=false;
      else { s.lineFormat.visible=true; s.lineFormat.color=color; }
    });
    await ctx.sync();
    if (ver===_applyVer) setStatus(`Border → ${color==='none'?'removed':color}`);
  });
}

async function applyBorderWeight(pts) {
  const ver = ++_applyVer;
  if (!_cachedShapeCount) return setStatus('No shape selected.');
  await PowerPoint.run(async ctx => {
    if (ver!==_applyVer) return;
    getSelectedFast(ctx).forEach(s => { s.lineFormat.weight=pts; s.lineFormat.visible=true; });
    await ctx.sync();
    if (ver===_applyVer) setStatus(`Weight → ${pts}pt`);
  });
}

async function applyBorderDash(style) {
  const ver = ++_applyVer;
  if (!_cachedShapeCount) return setStatus('No shape selected.');
  await PowerPoint.run(async ctx => {
    if (ver!==_applyVer) return;
    let dashVal;
    switch(style) {
      case 'dash':     dashVal = PowerPoint.LineDashStyle.dash;     break;
      case 'roundDot': dashVal = PowerPoint.LineDashStyle.roundDot; break;
      default:         dashVal = PowerPoint.LineDashStyle.solid;    break;
    }
    getSelectedFast(ctx).forEach(s => { s.lineFormat.dashStyle=dashVal; s.lineFormat.visible=true; });
    await ctx.sync();
    if (ver===_applyVer) setStatus(`Style → ${style}`);
  });
}

async function applyTextColor(color) {
  const ver = ++_applyVer;
  if (!_cachedShapeCount) return setStatus('No shape selected.');
  await PowerPoint.run(async ctx => {
    if (ver!==_applyVer) return;
    getSelectedFast(ctx).forEach(s => {
      try { s.textFrame.textRange.font.color = color; } catch(_) {}
    });
    await ctx.sync();
    if (ver===_applyVer) setStatus(`Text → ${color}`);
  });
}

async function applyLineEnds() {
  const sv=el('line-start').value, ev=el('line-end').value;
  const ver=++_applyVer;
  if (!_cachedShapeCount) return setStatus('No shape selected.');
  await PowerPoint.run(async ctx => {
    if (ver!==_applyVer) return;
    getSelectedFast(ctx).forEach(s => {
      try {
        s.lineFormat.beginArrowheadStyle = PowerPoint.ArrowheadStyle[sv]??PowerPoint.ArrowheadStyle.none;
        s.lineFormat.endArrowheadStyle   = PowerPoint.ArrowheadStyle[ev]??PowerPoint.ArrowheadStyle.none;
      } catch(_) {}
    });
    await ctx.sync();
    if (ver===_applyVer) setStatus(`Ends → ${sv} / ${ev}`);
  });
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
  document.getElementById('btn-apply-ends')?.addEventListener('click', applyLineEnds);
  document.getElementById('edit-back-btn').addEventListener('click', closeEditPanel);

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
      const merged={fill:false,border:false,text:false,lineEnds:false};
      let allSame=true;
      const firstKey=toKey(items[0].type), firstName=items[0].name||'';
      items.forEach(s => {
        const k=toKey(s.type);
        if (k!==firstKey) allSame=false;
        const caps=CAPABILITIES[k]||{};
        if (caps.fill)     merged.fill=true;
        if (caps.border)   merged.border=true;
        if (caps.text)     merged.text=true;
        if (caps.lineEnds) merged.lineEnds=true;
      });
      const meta=allSame
        ? (TYPE_META[firstKey]||{icon:'?',label:cap(firstKey||'Shape')})
        : {icon:'⊕',label:'Mixed Selection'};
      renderUI({merged,meta,firstName,count:items.length,anySupported:Object.values(merged).some(Boolean),isLine:allSame&&firstKey==='line'});
    });
  } catch(_) { renderEmpty(); }
}

function toKey(raw) {
  if (!raw) return 'unsupported';
  const s=String(raw); return s.charAt(0).toLowerCase()+s.slice(1);
}
function cap(s) { return s.charAt(0).toUpperCase()+s.slice(1); }

function renderUI({merged,meta,firstName,count,anySupported,isLine}) {
  hide('empty-state'); hide('unsupported-state'); show('shape-banner');
  el('shape-icon').textContent       = meta.icon;
  el('shape-type-label').textContent = meta.label;
  el('shape-name').textContent       = firstName?`"${firstName}"` : '';
  el('shape-count-badge').textContent= `×${count}`;
  tog('shape-count-badge', count>1);
  tog('section-fill',      merged.fill);
  tog('section-border',    merged.border);
  tog('section-line-ends', merged.lineEnds);
  tog('section-text',      merged.text);
  el('border-heading').textContent = isLine ? 'Line Style' : 'Border';
  if (!anySupported) show('unsupported-state');
}

function renderEmpty() {
  _cachedShapeCount=0;
  show('empty-state'); hide('shape-banner'); hide('unsupported-state');
  ['section-fill','section-border','section-line-ends','section-text'].forEach(hide);
  setStatus('');
}

function el(id)        { return document.getElementById(id); }
function show(id)      { el(id)?.classList.remove('hidden'); }
function hide(id)      { el(id)?.classList.add('hidden'); }
function tog(id,vis)   { el(id)?.classList.toggle('hidden',!vis); }
function setStatus(m)  { if(el('status')) el('status').textContent=m; }
