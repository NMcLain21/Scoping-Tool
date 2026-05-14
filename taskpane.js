'use strict';

// ─────────────────────────────────────────────────────────────────
// Brand / static color data
// ─────────────────────────────────────────────────────────────────
const BRAND_COLORS = [
  { hex: '#D50032', name: 'APi Corp Red'  },
  { hex: '#2A3E6D', name: 'Dark Navy'     },
  { hex: '#008579', name: 'Teal'          },
  { hex: '#D4D800', name: 'TBD Yellow'    },
  { hex: '#7030A0', name: 'APi Segment'   },
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
function tintHex(hex, factor) {
  const { r, g, b } = hexToRgb(hex);
  return rgbToHex(Math.round(r+(255-r)*factor),Math.round(g+(255-g)*factor),Math.round(b+(255-b)*factor));
}
function shadeHex(hex, factor) {
  const { r, g, b } = hexToRgb(hex);
  return rgbToHex(Math.round(r*(1-factor)),Math.round(g*(1-factor)),Math.round(b*(1-factor)));
}

// ─────────────────────────────────────────────────────────────────
// HSV utilities
// ─────────────────────────────────────────────────────────────────
function hsvToRgb(h,s,v) {
  const c=v*s, x=c*(1-Math.abs((h/60)%2-1)), m=v-c;
  let r=0,g=0,b=0;
  if(h<60){r=c;g=x;b=0;}else if(h<120){r=x;g=c;b=0;}
  else if(h<180){r=0;g=c;b=x;}else if(h<240){r=0;g=x;b=c;}
  else if(h<300){r=x;g=0;b=c;}else{r=c;g=0;b=x;}
  return{r:Math.round((r+m)*255),g:Math.round((g+m)*255),b:Math.round((b+m)*255)};
}
function rgbToHsv(r,g,b) {
  r/=255;g/=255;b/=255;
  const max=Math.max(r,g,b),min=Math.min(r,g,b),d=max-min;
  let h=0;const s=max===0?0:d/max,v=max;
  if(d!==0){if(max===r)h=((g-b)/d+(g<b?6:0))*60;else if(max===g)h=((b-r)/d+2)*60;else h=((r-g)/d+4)*60;}
  return{h:h<0?h+360:h,s,v};
}
function hexToHsv(hex){const{r,g,b}=hexToRgb(hex);return rgbToHsv(r,g,b);}
function hsvToHex(h,s,v){const{r,g,b}=hsvToRgb(h,s,v);return rgbToHex(r,g,b);}

// ─────────────────────────────────────────────────────────────────
// Storage
// ─────────────────────────────────────────────────────────────────
function getPalette(key) {
  try {
    const s=localStorage.getItem(`sct_${key}`);
    const arr=s?JSON.parse(s):JSON.parse(JSON.stringify(DEFAULT_PALETTES[key]));
    if(key==='fill') arr.forEach(c=>{if(!c.textHex)c.textHex=autoTextHex(c.hex);});
    return arr;
  } catch { return JSON.parse(JSON.stringify(DEFAULT_PALETTES[key])); }
}
function savePalette(key,arr){localStorage.setItem(`sct_${key}`,JSON.stringify(arr));}
function updateColor(key,idx,hex,name,textHex) {
  const arr=getPalette(key);
  if(!arr[idx])return;
  arr[idx]={hex,name:name||hex};
  if(key==='fill')arr[idx].textHex=textHex||autoTextHex(hex);
  savePalette(key,arr);
}
function addColor(key,hex,name,textHex) {
  const norm=normalizeHex(hex);if(!norm)return;
  const arr=getPalette(key);
  const entry={hex:norm,name:name||norm};
  if(key==='fill')entry.textHex=textHex||autoTextHex(norm);
  arr.push(entry);savePalette(key,arr);
}
function deleteColor(key,idx){const arr=getPalette(key);arr.splice(idx,1);savePalette(key,arr);}
function getRecent(key){try{return JSON.parse(localStorage.getItem(`sct_recent_${key}`))||[];}catch{return[];}}
function pushRecent(key,hex) {
  if(!hex||hex==='none')return;
  const norm=normalizeHex(hex)||hex.toUpperCase();
  let arr=getRecent(key).filter(h=>h!==norm);
  arr.unshift(norm);
  localStorage.setItem(`sct_recent_${key}`,JSON.stringify(arr.slice(0,10)));
}

// ─────────────────────────────────────────────────────────────────
// ══ MODE STATE ══
// _currentMode: 'style' | 'paint'
// STAGED: what the paint mode will apply
// ─────────────────────────────────────────────────────────────────
let _currentMode = 'style';

const STAGED = {
  fill:   { dirty: false, hex: null, textHex: null, name: '' },
  border: { dirty: false, hex: null, name: '' },
  text:   { dirty: false, hex: null },
  locked: false,
};

// Persist staged state across mode switches (sessionStorage so it
// survives Style→Paint→Style→Paint round-trips but not page reloads)
function saveStagedState() {
  try { sessionStorage.setItem('sct_staged', JSON.stringify(STAGED)); } catch(_) {}
}
function loadStagedState() {
  try {
    const s = sessionStorage.getItem('sct_staged');
    if (!s) return;
    const parsed = JSON.parse(s);
    Object.assign(STAGED, parsed);
  } catch(_) {}
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
// Theme colors from document
// ─────────────────────────────────────────────────────────────────
let _themeColors = null;

async function loadThemeColors() {
  try {
    await PowerPoint.run(async ctx => {
      ctx.presentation.slideMasters.load('items');
      await ctx.sync();
      const master = ctx.presentation.slideMasters.getItemAt(0);
      const theme  = master.getTheme();
      theme.load('themeColorScheme');
      await ctx.sync();
      const cs = theme.themeColorScheme;
      const raw = [
        { raw: cs.dark1,   name: 'Text / Dark 1'  },
        { raw: cs.light1,  name: 'Background 1'   },
        { raw: cs.dark2,   name: 'Text / Dark 2'  },
        { raw: cs.light2,  name: 'Background 2'   },
        { raw: cs.accent1, name: 'Accent 1'        },
        { raw: cs.accent2, name: 'Accent 2'        },
        { raw: cs.accent3, name: 'Accent 3'        },
        { raw: cs.accent4, name: 'Accent 4'        },
        { raw: cs.accent5, name: 'Accent 5'        },
        { raw: cs.accent6, name: 'Accent 6'        },
      ];
      const parsed = raw.map(c=>({hex:normalizeHex(c.raw),name:c.name})).filter(c=>c.hex);
      if (parsed.length >= 4) {
        _themeColors = parsed;
        if (_editKey) renderEditPanel();
      }
    });
  } catch(e) {
    console.log('Theme colors unavailable:', e.message);
  }
}

// ─────────────────────────────────────────────────────────────────
// Theme grid builder
// ─────────────────────────────────────────────────────────────────
function buildThemeGrid(dataAttr) {
  const baseColors = _themeColors || BRAND_COLORS;
  const rows = [
    baseColors.map(c=>c.hex),
    baseColors.map(c=>tintHex(c.hex,0.8)),
    baseColors.map(c=>tintHex(c.hex,0.6)),
    baseColors.map(c=>tintHex(c.hex,0.4)),
    baseColors.map(c=>shadeHex(c.hex,0.25)),
    baseColors.map(c=>shadeHex(c.hex,0.5)),
  ];
  return `<div class="ppt-theme-grid">${
    rows.map((row,ri)=>
      `<div class="ppt-theme-row${ri===0?' ppt-base-row':''}">${
        row.map(hex=>
          `<button class="ppt-theme-swatch${isLight(hex)?' light':''}"
                   style="background:${hex}" ${dataAttr}="${hex}"
                   title="${hex}" type="button"></button>`
        ).join('')
      }</div>`
    ).join('')
  }</div>`;
}

// ─────────────────────────────────────────────────────────────────
// Build PPT-style picker
// ─────────────────────────────────────────────────────────────────
function buildPptPicker(pickerType,sufx,startHex,dataAttr,recentKey) {
  const norm=normalizeHex(startHex)||'#2A3E6D';
  const hexVal=norm.replace('#','');
  const recent=getRecent(recentKey);
  const moreId=`mc-${pickerType}-${sufx}`;

  const stdRow=STANDARD_COLORS.map(c=>
    `<button class="ppt-theme-swatch${isLight(c.hex)?' light':''}"
             style="background:${c.hex}" ${dataAttr}="${c.hex}"
             title="${esc(c.name)}" type="button"></button>`
  ).join('');

  const recentHtml=recent.length?`
    <div class="acc-section">
      <button class="acc-hdr" data-acc="acc-recent-${pickerType}-${sufx}" type="button">
        <span>Recent Colors</span>
        <svg class="acc-chevron" width="10" height="10" viewBox="0 0 10 10" fill="none">
          <path d="M2 3.5l3 3 3-3" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
      </button>
      <div class="acc-body hidden" id="acc-recent-${pickerType}-${sufx}">
        <div class="ppt-std-row" style="padding:4px 0 2px;">${
          recent.slice(0,10).map(hex=>
            `<button class="ppt-theme-swatch${isLight(hex)?' light':''}"
                     style="background:${hex}" ${dataAttr}="${hex}"
                     title="${hex}" type="button"></button>`
          ).join('')
        }</div>
      </div>
    </div>`:'';

  return `
    <div class="ppt-picker" data-picker="${pickerType}">
      <div class="acc-section">
        <button class="acc-hdr acc-open" data-acc="acc-theme-${pickerType}-${sufx}" type="button">
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
        <button class="acc-hdr" data-acc="acc-std-${pickerType}-${sufx}" type="button">
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
        <button class="acc-hdr" data-acc="${moreId}" type="button">
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
// Build color form
// ─────────────────────────────────────────────────────────────────
function buildColorForm(type,idx,startHex,startName,startTextHex) {
  const norm=normalizeHex(startHex)||'#2A3E6D';
  const txNorm=normalizeHex(startTextHex)||autoTextHex(norm);
  const sufx=type==='edit'?`edit-${idx}`:'add';
  const isFill=_editKey==='fill';
  const fillPicker=buildPptPicker('fill',sufx,norm,'data-setfill',_editKey);
  const fontPicker=isFill?buildPptPicker('textcolor',sufx,txNorm,'data-settextcolor','text'):'';

  return `
    <div class="color-form" id="cform-${sufx}" data-type="${type}" data-idx="${idx??''}">
      <input type="text" class="ep-name-input" id="cfn-${sufx}"
             placeholder="Color name (optional)" maxlength="36"
             value="${esc(startName)}" />
      <!-- Shape Fill — collapsible -->
      <div class="ppt-sect-hdr ppt-sect-toggle acc-open" data-target="fill-picker-${sufx}"
           id="fill-hdr-${sufx}" style="cursor:pointer;">
        <svg class="ppt-fill-svg" id="fill-svg-${sufx}" width="16" height="16" viewBox="0 0 16 16" fill="none">
          <path d="M3 12.5c0-.9.8-1.7 1.7-1.7s1.7.8 1.7 1.7S4.7 15 4.7 15 3 13.4 3 12.5z"
                fill="${norm}" stroke="${isLight(norm)?'#aaa':norm}" stroke-width="0.8"/>
          <path d="M5.5 10.7l5.8-5.8a1.1 1.1 0 011.6 0l.5.5a1.1 1.1 0 010 1.6L7.5 12"
                stroke="#666" stroke-width="1.2" stroke-linecap="round" fill="none"/>
        </svg>
        <span class="ppt-sect-label">Shape Fill</span>
        <div class="ppt-curr-swatch${isLight(norm)?' light':''}"
             id="fill-swatch-${sufx}" style="background:${norm};"></div>
        <svg class="acc-chevron acc-chevron-open" width="10" height="10" viewBox="0 0 10 10" fill="none">
          <path d="M2 3.5l3 3 3-3" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
      </div>
      <div id="fill-picker-${sufx}" class="ppt-sect-body">
        ${fillPicker}
      </div>
      ${isFill?`
      <!-- Font Color — collapsible -->
      <div class="ppt-sect-hdr ppt-sect-toggle" data-target="font-picker-${sufx}"
           id="font-hdr-${sufx}" style="margin-top:8px;cursor:pointer;">
        <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
          <text x="2" y="13" font-family="Segoe UI,sans-serif" font-size="12" font-weight="700"
                fill="${txNorm}" id="font-svg-${sufx}">A</text>
          <rect x="2" y="13.5" width="12" height="2.5" rx="1"
                fill="${txNorm}" id="font-bar-${sufx}"/>
        </svg>
        <span class="ppt-sect-label">Font Color</span>
        <div class="ppt-curr-swatch${isLight(txNorm)?' light':''}"
             id="font-swatch-${sufx}" style="background:${txNorm};"></div>
        <svg class="acc-chevron" width="10" height="10" viewBox="0 0 10 10" fill="none">
          <path d="M2 3.5l3 3 3-3" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
      </div>
      <div id="font-picker-${sufx}" class="ppt-sect-body hidden">
        ${fontPicker}
      </div>`:''}
      <div class="cform-btns">
        <button class="cform-save" type="button" data-type="${type}" data-idx="${idx??''}">
          ${type==='edit'?'Save Changes':'+ Add to My Colors'}
        </button>
        <button class="cform-cancel" type="button">Cancel</button>
      </div>
    </div>`;
}

// ─────────────────────────────────────────────────────────────────
// Wire a single picker
// ─────────────────────────────────────────────────────────────────
function wirePicker(form,pickerType,sufx,onColorChange) {
  const moreId=`mc-${pickerType}-${sufx}`;
  const hexInp=document.getElementById(`hex-${pickerType}-${sufx}`);
  const preview=document.getElementById(`prev-${pickerType}-${sufx}`);
  const sq=document.getElementById(`hsv-sq-${pickerType}-${sufx}`);
  const dot=document.getElementById(`hsv-dot-${pickerType}-${sufx}`);
  const hueSlider=document.getElementById(`hue-${pickerType}-${sufx}`);
  const dataAttr=pickerType==='fill'?'data-setfill':'data-settextcolor';

  const startNorm=normalizeHex('#'+(hexInp?.value||''))||'#2A3E6D';
  const startHsv=hexToHsv(startNorm);
  let H=startHsv.h,S=startHsv.s,V=startHsv.v;

  function syncFromHsv() {
    const hex=hsvToHex(H,S,V);
    if(hexInp)hexInp.value=hex.replace('#','').toUpperCase();
    if(preview){preview.style.background=hex;preview.classList.toggle('light',isLight(hex));}
    if(sq){sq.style.background=`hsl(${Math.round(H)},100%,50%)`;if(dot){dot.style.left=`${S*100}%`;dot.style.top=`${(1-V)*100}%`;}}
    if(hueSlider)hueSlider.value=Math.round(H);
    onColorChange(hex);
  }
  function setFromHex(rawHex) {
    const c=rawHex.startsWith('#')?rawHex:'#'+rawHex;
    const n=normalizeHex(c);if(!n)return;
    const hsv=hexToHsv(n);H=hsv.h;S=hsv.s;V=hsv.v;
    if(hexInp)hexInp.value=n.replace('#','').toUpperCase();
    if(preview){preview.style.background=n;preview.classList.toggle('light',isLight(n));}
    if(sq){sq.style.background=`hsl(${Math.round(H)},100%,50%)`;if(dot){dot.style.left=`${S*100}%`;dot.style.top=`${(1-V)*100}%`;}}
    if(hueSlider)hueSlider.value=Math.round(H);
    onColorChange(n);
  }

  const pickerEl=form.querySelector(`.ppt-picker[data-picker="${pickerType}"]`);
  pickerEl?.querySelectorAll(`[${dataAttr}]`).forEach(btn=>{
    btn.addEventListener('click',e=>{
      e.stopPropagation();
      setFromHex(btn.getAttribute(dataAttr));
      pickerEl.querySelectorAll(`[${dataAttr}]`).forEach(b=>b.classList.remove('qp-selected'));
      btn.classList.add('qp-selected');
    });
  });

  const pickerContainer=form.querySelector(`.ppt-picker[data-picker="${pickerType}"]`);
  pickerContainer?.querySelectorAll('.acc-hdr').forEach(hdr=>{
    hdr.addEventListener('click',()=>{
      const bodyId=hdr.dataset.acc;
      const body=document.getElementById(bodyId);if(!body)return;
      const opening=body.classList.contains('hidden');
      body.classList.toggle('hidden',!opening);
      hdr.classList.toggle('acc-open',opening);
      hdr.querySelector('.acc-chevron')?.classList.toggle('acc-chevron-open',opening);
      if(opening&&bodyId===moreId)syncFromHsv();
    });
  });

  hueSlider?.addEventListener('input',()=>{H=parseFloat(hueSlider.value);syncFromHsv();});

  let dragging=false;
  function handleDrag(e) {
    if(!sq?.isConnected){dragging=false;return;}
    const rect=sq.getBoundingClientRect();
    const cx=e.touches?e.touches[0].clientX:e.clientX;
    const cy=e.touches?e.touches[0].clientY:e.clientY;
    S=Math.max(0,Math.min(1,(cx-rect.left)/rect.width));
    V=Math.max(0,Math.min(1,1-(cy-rect.top)/rect.height));
    syncFromHsv();
  }
  if(sq){
    sq.addEventListener('mousedown',e=>{e.preventDefault();dragging=true;handleDrag(e);});
    sq.addEventListener('touchstart',e=>{e.preventDefault();handleDrag(e);},{passive:false});
    sq.addEventListener('touchmove',e=>{e.preventDefault();handleDrag(e);},{passive:false});
  }
  const onMove=e=>{if(dragging)handleDrag(e);};
  const onUp=()=>{dragging=false;};
  document.addEventListener('mousemove',onMove);
  document.addEventListener('mouseup',onUp);
  hexInp?.addEventListener('input',()=>{if(hexInp.value.length===6)setFromHex(hexInp.value);});
  if(sq)syncFromHsv();
  return ()=>{document.removeEventListener('mousemove',onMove);document.removeEventListener('mouseup',onUp);};
}

// ─────────────────────────────────────────────────────────────────
// Wire the whole color form
// ─────────────────────────────────────────────────────────────────
function wireColorForm(form,key) {
  const type=form.dataset.type;
  const idx=form.dataset.idx!==''?parseInt(form.dataset.idx,10):null;
  const sufx=type==='edit'?`edit-${idx}`:'add';
  const isFill=key==='fill';
  let currentFill=normalizeHex('#'+(document.getElementById(`hex-fill-${sufx}`)?.value||''))||'#2A3E6D';
  let currentText=normalizeHex('#'+(document.getElementById(`hex-textcolor-${sufx}`)?.value||''))||autoTextHex(currentFill);

  function onFillChange(hex) {
    currentFill=hex;
    const sw=document.getElementById(`fill-swatch-${sufx}`);
    if(sw){sw.style.background=hex;sw.classList.toggle('light',isLight(hex));}
    const sv=document.getElementById(`fill-svg-${sufx}`);
    if(sv){const p=sv.querySelector('path');if(p){p.setAttribute('fill',hex);p.setAttribute('stroke',isLight(hex)?'#aaa':hex);}}
  }
  function onFontChange(hex) {
    currentText=hex;
    const sw=document.getElementById(`font-swatch-${sufx}`);
    if(sw){sw.style.background=hex;sw.classList.toggle('light',isLight(hex));}
    const txt=document.getElementById(`font-svg-${sufx}`);if(txt)txt.setAttribute('fill',hex);
    const bar=document.getElementById(`font-bar-${sufx}`);if(bar)bar.setAttribute('fill',hex);
  }

  const cleanFill=wirePicker(form,'fill',sufx,onFillChange);
  const cleanFont=isFill?wirePicker(form,'textcolor',sufx,onFontChange):null;
  const nameInp=document.getElementById(`cfn-${sufx}`);

  form.querySelectorAll('.ppt-sect-toggle').forEach(hdr=>{
    hdr.addEventListener('click',()=>{
      const targetId=hdr.dataset.target;
      const body=document.getElementById(targetId);if(!body)return;
      const opening=body.classList.contains('hidden');
      body.classList.toggle('hidden',!opening);
      hdr.classList.toggle('acc-open',opening);
      hdr.querySelector('.acc-chevron')?.classList.toggle('acc-chevron-open',opening);
    });
  });

  form.querySelector('.cform-save')?.addEventListener('click',()=>{
    const norm=normalizeHex(currentFill);if(!norm)return;
    const name=nameInp?.value.trim()||'';
    const txtNorm=isFill?(normalizeHex(currentText)||autoTextHex(norm)):null;
    if(type==='edit'&&idx!==null)updateColor(key,idx,norm,name||norm,txtNorm);
    else addColor(key,norm,name,txtNorm);
    _openForm=null;cleanFill?.();cleanFont?.();renderEditPanel();
  });
  form.querySelector('.cform-cancel')?.addEventListener('click',()=>{
    _openForm=null;cleanFill?.();cleanFont?.();renderEditPanel();
  });
  nameInp?.addEventListener('keydown',e=>{if(e.key==='Enter')form.querySelector('.cform-save')?.click();});
  setTimeout(()=>{if(nameInp?.isConnected)nameInp.select();},50);
}

// ─────────────────────────────────────────────────────────────────
// Main panel — labeled tiles
// ─────────────────────────────────────────────────────────────────
function renderMainSwatches(key) {
  const row=document.getElementById(`swatches-${key}`);if(!row)return;
  const colors=getPalette(key);
  row.innerHTML=colors.map(c=>{
    if(key==='border'){
      return `<button class="main-swatch border-legend-tile"
                       data-color="${c.hex}" title="${esc(c.name)}" aria-label="${esc(c.name)}">
                <span class="legend-line-sample" style="background:${c.hex}"></span>
                <span class="legend-line-name">${esc(c.name)}</span>
              </button>`;
    }
    if(key==='fill'){
      const txHex=c.textHex||autoTextHex(c.hex);
      return `<button class="main-swatch doi-tile"
                       style="background:${c.hex};"
                       data-color="${c.hex}" data-textcolor="${txHex}"
                       title="${esc(c.name)}" aria-label="${esc(c.name)}">
                <span class="swatch-label" style="color:${txHex};">${esc(c.name)}</span>
              </button>`;
    }
    const txtColor=isLight(c.hex)?'#000000':'#FFFFFF';
    return `<button class="main-swatch${isLight(c.hex)?' light':''}"
                     style="background:${c.hex}"
                     data-color="${c.hex}" title="${esc(c.name)}" aria-label="${esc(c.name)}">
              <span class="swatch-label" style="color:${txtColor};">${esc(c.name)}</span>
            </button>`;
  }).join('')+
  `<button class="edit-palette-btn" data-key="${key}" title="Edit ${key} palette">
     <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
       <path d="M8.5 1.5a1.414 1.414 0 012 2L4 10H2V8L8.5 1.5z"
             stroke="currentColor" stroke-width="1.3"
             stroke-linecap="round" stroke-linejoin="round" fill="none"/>
     </svg>
   </button>`;

  row.querySelectorAll('.main-swatch').forEach(btn=>{
    btn.addEventListener('click',()=>{
      const color=btn.dataset.color;
      const txtColor=btn.dataset.textcolor;
      if(key==='fill'&&txtColor)applyDOITile(color,txtColor);
      else({border:applyBorderColor,text:applyTextColor})[key]?.(color);
      pushRecent(key,color);
    });
  });
  row.querySelector('.edit-palette-btn').addEventListener('click',()=>{
    const fnMap={fill:applyDOITile,border:applyBorderColor,text:applyTextColor};
    openEditPanel(key,fnMap[key]);
  });
}

// ─────────────────────────────────────────────────────────────────
// Trash icon
// ─────────────────────────────────────────────────────────────────
function trashIcon() {
  return `<svg width="11" height="12" viewBox="0 0 11 12" fill="none">
    <path d="M1 3h9M4 3V2a.5.5 0 01.5-.5h2A.5.5 0 017 2v1M9 3l-.5 7a.5.5 0 01-.5.5H3a.5.5 0 01-.5-.5L2 3"
          stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/>
  </svg>`;
}

// ─────────────────────────────────────────────────────────────────
// Edit panel
// ─────────────────────────────────────────────────────────────────
function openEditPanel(key,fn) {
  _editKey=key;_applyFn=fn;_openForm=null;
  renderEditPanel();
  document.getElementById('main-content').classList.add('hidden');
  document.getElementById('paint-panel').classList.add('hidden');
  document.getElementById('edit-panel').classList.remove('hidden');
}
function closeEditPanel() {
  document.getElementById('edit-panel').classList.add('hidden');
  if(_currentMode==='paint'){
    document.getElementById('paint-panel').classList.remove('hidden');
  } else {
    document.getElementById('main-content').classList.remove('hidden');
  }
  renderMainSwatches(_editKey);
  _editKey=null;_applyFn=null;_openForm=null;
}

function renderEditPanel() {
  const key=_editKey;
  const titleMap={fill:'Degree of Integration',border:'Post Int Process Owner',text:'Text Color'};
  document.getElementById('edit-panel-title').textContent=titleMap[key]||'Edit Colors';
  const palette=getPalette(key);

  const myColorsHtml=palette.map((c,i)=>{
    const isEditing=_openForm&&_openForm.type==='edit'&&_openForm.idx===i;
    const txHex=c.textHex||autoTextHex(c.hex);
    return `
      <div class="my-color-item" draggable="true" data-idx="${i}">
        <div class="my-color-top">
          <div class="drag-handle" title="Drag to reorder">
            <svg width="10" height="14" viewBox="0 0 10 14" fill="none">
              <circle cx="3" cy="3"  r="1.2" fill="currentColor"/>
              <circle cx="7" cy="3"  r="1.2" fill="currentColor"/>
              <circle cx="3" cy="7"  r="1.2" fill="currentColor"/>
              <circle cx="7" cy="7"  r="1.2" fill="currentColor"/>
              <circle cx="3" cy="11" r="1.2" fill="currentColor"/>
              <circle cx="7" cy="11" r="1.2" fill="currentColor"/>
            </svg>
          </div>
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
              ${trashIcon()}
            </button>
          </div>
        </div>
        ${isEditing?buildColorForm('edit',i,c.hex,c.name,txHex):''}
      </div>`;
  }).join('');

  const isAdding=_openForm&&_openForm.type==='add';
  document.getElementById('edit-panel-body').innerHTML=`
    <div class="ep-label">My Colors <span class="ep-hint">drag to reorder · click swatch to apply</span></div>
    <div id="my-colors-list">${myColorsHtml}</div>
    <button class="add-new-btn${isAdding?' active':''}" id="ep-add-toggle">
      <svg width="11" height="11" viewBox="0 0 11 11" fill="none">
        <line x1="5.5" y1="1" x2="5.5" y2="10" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/>
        <line x1="1" y1="5.5" x2="10" y2="5.5" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/>
      </svg>
      Add New Color
    </button>
    ${isAdding?buildColorForm('add',null,'#2A3E6D','','#FFFFFF'):''}
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

function bindEditEvents(key) {
  const body=document.getElementById('edit-panel-body');

  body.querySelectorAll('.my-swatch').forEach(s=>{
    const go=()=>{
      const color=s.dataset.color,tc=s.dataset.textcolor;
      if(key==='fill'&&tc)applyDOITile(color,tc);else _applyFn?.(color);
      if(color!=='none')pushRecent(key,color);
      closeEditPanel();
    };
    s.addEventListener('click',go);
    s.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();go();}});
  });

  body.querySelectorAll('.my-edit-btn').forEach(btn=>{
    btn.addEventListener('click',()=>{
      const idx=parseInt(btn.dataset.idx,10);
      _openForm=(_openForm&&_openForm.type==='edit'&&_openForm.idx===idx)?null:{type:'edit',idx};
      renderEditPanel();
    });
  });

  body.querySelectorAll('.my-del-btn').forEach(btn=>{
    btn.addEventListener('click',()=>{
      deleteColor(key,parseInt(btn.dataset.idx,10));
      _openForm=null;renderEditPanel();
    });
  });

  document.getElementById('ep-add-toggle').addEventListener('click',()=>{
    _openForm=(_openForm&&_openForm.type==='add')?null:{type:'add',idx:null};
    renderEditPanel();
  });

  document.getElementById('ep-no-color').addEventListener('click',()=>{
    if(key==='fill')applyDOITile('none',null);else _applyFn?.('none');
    closeEditPanel();
  });

  document.getElementById('ep-reset').addEventListener('click',()=>{
    if(!confirm(`Reset ${key} palette to defaults? Your custom colors will be lost.`))return;
    localStorage.removeItem(`sct_${key}`);
    _openForm=null;renderEditPanel();renderMainSwatches(key);
  });

  body.querySelectorAll('.color-form').forEach(form=>wireColorForm(form,key));

  // Drag-and-drop
  const list=document.getElementById('my-colors-list');
  let dragSrcIdx=null;
  list.querySelectorAll('.my-color-item[draggable]').forEach(item=>{
    item.addEventListener('dragstart',e=>{
      dragSrcIdx=parseInt(item.dataset.idx,10);
      item.classList.add('dragging');
      e.dataTransfer.effectAllowed='move';
      e.dataTransfer.setData('text/plain',dragSrcIdx);
    });
    item.addEventListener('dragend',()=>{
      item.classList.remove('dragging');
      list.querySelectorAll('.my-color-item').forEach(i=>i.classList.remove('drag-over'));
    });
    item.addEventListener('dragover',e=>{
      e.preventDefault();e.dataTransfer.dropEffect='move';
      list.querySelectorAll('.my-color-item').forEach(i=>i.classList.remove('drag-over'));
      item.classList.add('drag-over');
    });
    item.addEventListener('drop',e=>{
      e.preventDefault();
      const destIdx=parseInt(item.dataset.idx,10);
      if(dragSrcIdx===null||dragSrcIdx===destIdx)return;
      const arr=getPalette(key);
      const [moved]=arr.splice(dragSrcIdx,1);
      arr.splice(destIdx,0,moved);
      savePalette(key,arr);
      _openForm=null;renderEditPanel();
    });
  });
}

// ─────────────────────────────────────────────────────────────────
// ══ MODE SWITCHING ══
// ─────────────────────────────────────────────────────────────────
function switchMode(mode) {
  if(_currentMode === mode) return;
  _currentMode = mode;

  // Save/restore staged state on every switch
  saveStagedState();

  const header = document.getElementById('app-header');
  const stylBtn = document.getElementById('btn-mode-style');
  const paintBtn = document.getElementById('btn-mode-paint');
  const mainContent = document.getElementById('main-content');
  const paintPanel = document.getElementById('paint-panel');
  const emptyState = document.getElementById('empty-state');
  const shapeBanner = document.getElementById('shape-banner');

  if(mode === 'paint') {
    // Update toggle
    stylBtn.classList.remove('mode-active');
    paintBtn.classList.add('mode-active');
    // Header tint
    header.classList.add('paint-mode-header');
    // Show/hide panels
    mainContent.classList.add('hidden');
    emptyState.classList.add('hidden');
    shapeBanner.classList.add('hidden');
    paintPanel.classList.remove('hidden');
    // Render paint panel
    renderPaintPanel();
    setStatus('Paint mode — capture a shape or select attributes');
  } else {
    // Update toggle
    paintBtn.classList.remove('mode-active');
    stylBtn.classList.add('mode-active');
    // Remove header tint
    header.classList.remove('paint-mode-header');
    // Show/hide panels
    paintPanel.classList.add('hidden');
    mainContent.classList.remove('hidden');
    // Re-inspect selection
    inspectSelection();
    setStatus('');
  }
}

// ─────────────────────────────────────────────────────────────────
// ══ PAINT PANEL RENDER ══
// ─────────────────────────────────────────────────────────────────
function renderPaintPanel() {
  // Fill / DOI preview
  updateStagedPreview('fill');
  updateStagedPreview('border');
  updateStagedPreview('text');

  // Checkboxes — only enabled when dirty
  const fillChk  = document.getElementById('staged-fill-chk');
  const borderChk = document.getElementById('staged-border-chk');
  const textChk  = document.getElementById('staged-text-chk');

  if(fillChk) {
    fillChk.disabled = !STAGED.fill.dirty;
    fillChk.checked  = STAGED.fill.dirty;
  }
  if(borderChk) {
    borderChk.disabled = !STAGED.border.dirty;
    borderChk.checked  = STAGED.border.dirty;
  }
  if(textChk) {
    textChk.disabled = !STAGED.text.dirty;
    textChk.checked  = STAGED.text.dirty;
  }

  // Lock toggle
  updateLockToggle();

  // Capture hint — update based on state
  const anyDirty = STAGED.fill.dirty || STAGED.border.dirty || STAGED.text.dirty;
  const hint = document.getElementById('capture-hint');
  if(hint) {
    hint.textContent = anyDirty
      ? 'Format captured. Select shapes and click "Apply to Selection" or enable auto-paint.'
      : 'Select a shape, then click Capture to read its formatting.';
  }
}

function updateStagedPreview(attr) {
  if(attr === 'fill') {
    const prev = document.getElementById('staged-fill-preview');
    if(!prev) return;
    if(STAGED.fill.dirty && STAGED.fill.hex) {
      const txHex = STAGED.fill.textHex || autoTextHex(STAGED.fill.hex);
      prev.innerHTML = `<span class="staged-fill-tile"
        style="background:${STAGED.fill.hex};color:${txHex};">
        ${esc(STAGED.fill.name || STAGED.fill.hex)}
      </span>`;
    } else {
      prev.innerHTML = '<span class="staged-not-set">Not captured</span>';
    }
  } else if(attr === 'border') {
    const prev = document.getElementById('staged-border-preview');
    if(!prev) return;
    if(STAGED.border.dirty && STAGED.border.hex) {
      prev.innerHTML = `<span class="staged-line-preview"
        style="background:${STAGED.border.hex};"></span>
        <span class="staged-border-name">${esc(STAGED.border.name || STAGED.border.hex)}</span>`;
    } else {
      prev.innerHTML = '<span class="staged-not-set">Not captured</span>';
    }
  } else if(attr === 'text') {
    const prev = document.getElementById('staged-text-preview');
    if(!prev) return;
    if(STAGED.text.dirty && STAGED.text.hex) {
      prev.innerHTML = `<span class="staged-color-dot${isLight(STAGED.text.hex)?' light':''}"
        style="background:${STAGED.text.hex};"></span>
        <span class="staged-hex-label">${STAGED.text.hex}</span>`;
    } else {
      prev.innerHTML = '<span class="staged-not-set">Not captured</span>';
    }
  }
}

function updateLockToggle() {
  const btn = document.getElementById('btn-lock');
  const lbl = document.getElementById('lock-state-label');
  if(!btn) return;
  btn.setAttribute('aria-checked', STAGED.locked ? 'true' : 'false');
  btn.classList.toggle('lock-on', STAGED.locked);
  if(lbl) lbl.textContent = STAGED.locked ? 'On' : 'Off';
}

// ─────────────────────────────────────────────────────────────────
// ══ CAPTURE FROM SHAPE ══
// ─────────────────────────────────────────────────────────────────
async function captureFromShape() {
  try {
    await PowerPoint.run(async ctx => {
      const sel = ctx.presentation.getSelectedShapes();
      sel.load('items/type,items/name,items/fill,items/lineFormat,items/textFrame');
      await ctx.sync();

      if(!sel.items.length) {
        setStatus('No shape selected — select a shape first.');
        return;
      }

      const s = sel.items[0];
      // Load the properties we need
      s.fill.load('type,foreColor');
      s.lineFormat.load('color,weight,visible');
      try { s.textFrame.textRange.font.load('color'); } catch(_) {}
      await ctx.sync();

      // ── Capture fill ──
      try {
        const fillType = s.fill.type;
        if(fillType === PowerPoint.ShapeFillType.solid || fillType === 'Solid') {
          const rawFill = s.fill.foreColor;
          const normFill = normalizeHex(rawFill.startsWith('#') ? rawFill : '#'+rawFill);
          if(normFill) {
            STAGED.fill.dirty   = true;
            STAGED.fill.hex     = normFill;
            STAGED.fill.textHex = autoTextHex(normFill);
            // Try to match a palette name
            const match = getPalette('fill').find(c=>c.hex.toUpperCase()===normFill.toUpperCase());
            STAGED.fill.name = match ? match.name : normFill;
            if(match && match.textHex) STAGED.fill.textHex = match.textHex;
          }
        }
      } catch(_) {}

      // ── Capture border ──
      try {
        if(s.lineFormat.visible !== false) {
          const rawBorder = s.lineFormat.color;
          const normBorder = normalizeHex(rawBorder.startsWith('#') ? rawBorder : '#'+rawBorder);
          if(normBorder) {
            STAGED.border.dirty = true;
            STAGED.border.hex   = normBorder;
            const match = getPalette('border').find(c=>c.hex.toUpperCase()===normBorder.toUpperCase());
            STAGED.border.name = match ? match.name : normBorder;
          }
        }
      } catch(_) {}

      // ── Capture text color ──
      try {
        const rawText = s.textFrame.textRange.font.color;
        const normText = normalizeHex(rawText.startsWith('#') ? rawText : '#'+rawText);
        if(normText) {
          STAGED.text.dirty = true;
          STAGED.text.hex   = normText;
        }
      } catch(_) {}

      _cachedShapeCount = sel.items.length;
      saveStagedState();
      renderPaintPanel();
      setStatus(`Captured from "${s.name || 'shape'}"`);
    });
  } catch(e) {
    setStatus('Capture failed — make sure a shape is selected.');
  }
}

// ─────────────────────────────────────────────────────────────────
// ══ APPLY STAGED FORMAT ══
// ─────────────────────────────────────────────────────────────────
async function applyStagedFormat() {
  const anyChecked =
    (STAGED.fill.dirty   && document.getElementById('staged-fill-chk')?.checked)   ||
    (STAGED.border.dirty && document.getElementById('staged-border-chk')?.checked) ||
    (STAGED.text.dirty   && document.getElementById('staged-text-chk')?.checked);

  if(!anyChecked) { setStatus('Nothing staged — capture a shape first.'); return; }
  if(!_cachedShapeCount) { setStatus('No shape selected.'); return; }

  const ver = ++_applyVer;
  await PowerPoint.run(async ctx => {
    if(ver !== _applyVer) return;
    const shapes = getSelectedFast(ctx);
    const applyFill   = STAGED.fill.dirty   && document.getElementById('staged-fill-chk')?.checked;
    const applyBorder = STAGED.border.dirty && document.getElementById('staged-border-chk')?.checked;
    const applyText   = STAGED.text.dirty   && document.getElementById('staged-text-chk')?.checked;

    shapes.forEach(s => {
      if(applyFill) {
        if(STAGED.fill.hex === 'none') s.fill.clear();
        else s.fill.setSolidColor(STAGED.fill.hex);
        if(STAGED.fill.textHex && !applyText) {
          try { s.textFrame.textRange.font.color = STAGED.fill.textHex; } catch(_) {}
        }
      }
      if(applyBorder) {
        if(STAGED.border.hex === 'none') s.lineFormat.visible = false;
        else { s.lineFormat.visible = true; s.lineFormat.color = STAGED.border.hex; }
      }
      if(applyText) {
        try { s.textFrame.textRange.font.color = STAGED.text.hex; } catch(_) {}
      }
    });

    await ctx.sync();
    if(ver === _applyVer) {
      const parts = [];
      if(applyFill)   parts.push('fill');
      if(applyBorder) parts.push('border');
      if(applyText)   parts.push('text');
      setStatus(`Painted: ${parts.join(' + ')}`);
    }
  });
}

// ─────────────────────────────────────────────────────────────────
// Update shape count for paint mode (no UI update)
// ─────────────────────────────────────────────────────────────────
async function updateShapeCount() {
  try {
    await PowerPoint.run(async ctx => {
      const sel = ctx.presentation.getSelectedShapes();
      sel.load('items');
      await ctx.sync();
      _cachedShapeCount = sel.items.length;
    });
  } catch(_) { _cachedShapeCount = 0; }
}

// ─────────────────────────────────────────────────────────────────
// Apply functions (Style mode)
// ─────────────────────────────────────────────────────────────────
function getSelectedFast(ctx) {
  const col=ctx.presentation.getSelectedShapes();
  const out=[];
  for(let i=0;i<_cachedShapeCount;i++)out.push(col.getItemAt(i));
  return out;
}

async function applyDOITile(fillColor,textColor) {
  const ver=++_applyVer;
  if(!_cachedShapeCount)return setStatus('No shape selected.');
  await PowerPoint.run(async ctx=>{
    if(ver!==_applyVer)return;
    getSelectedFast(ctx).forEach(s=>{
      if(fillColor==='none')s.fill.clear();else s.fill.setSolidColor(fillColor);
      if(textColor&&textColor!=='none'){try{s.textFrame.textRange.font.color=textColor;}catch(_){}}
    });
    await ctx.sync();
    if(ver===_applyVer)setStatus(`DOI → ${fillColor==='none'?'cleared':fillColor}`);
  });
}

async function applyBorderColor(color) {
  const ver=++_applyVer;
  if(!_cachedShapeCount)return setStatus('No shape selected.');
  await PowerPoint.run(async ctx=>{
    if(ver!==_applyVer)return;
    getSelectedFast(ctx).forEach(s=>{
      if(color==='none')s.lineFormat.visible=false;
      else{s.lineFormat.visible=true;s.lineFormat.color=color;}
    });
    await ctx.sync();
    if(ver===_applyVer)setStatus(`Border → ${color==='none'?'removed':color}`);
  });
}

async function applyBorderWeight(pts) {
  const ver=++_applyVer;
  if(!_cachedShapeCount)return setStatus('No shape selected.');
  await PowerPoint.run(async ctx=>{
    if(ver!==_applyVer)return;
    getSelectedFast(ctx).forEach(s=>{s.lineFormat.weight=pts;s.lineFormat.visible=true;});
    await ctx.sync();
    if(ver===_applyVer)setStatus(`Weight → ${pts}pt`);
  });
}

async function applyBorderDash(style) {
  const ver=++_applyVer;
  if(!_cachedShapeCount)return setStatus('No shape selected.');
  await PowerPoint.run(async ctx=>{
    if(ver!==_applyVer)return;
    let dashVal;
    switch(style){
      case 'dash':     dashVal=PowerPoint.LineDashStyle.dash;     break;
      case 'roundDot': dashVal=PowerPoint.LineDashStyle.roundDot; break;
      default:         dashVal=PowerPoint.LineDashStyle.solid;    break;
    }
    getSelectedFast(ctx).forEach(s=>{s.lineFormat.dashStyle=dashVal;s.lineFormat.visible=true;});
    await ctx.sync();
    if(ver===_applyVer)setStatus(`Style → ${style}`);
  });
}

async function applyTextColor(color) {
  const ver=++_applyVer;
  if(!_cachedShapeCount)return setStatus('No shape selected.');
  await PowerPoint.run(async ctx=>{
    if(ver!==_applyVer)return;
    getSelectedFast(ctx).forEach(s=>{try{s.textFrame.textRange.font.color=color;}catch(_){}});
    await ctx.sync();
    if(ver===_applyVer)setStatus(`Text → ${color}`);
  });
}

async function applyLineEnds() {
  const sv=el('line-start').value,ev=el('line-end').value;
  const ver=++_applyVer;
  if(!_cachedShapeCount)return setStatus('No shape selected.');
  await PowerPoint.run(async ctx=>{
    if(ver!==_applyVer)return;
    getSelectedFast(ctx).forEach(s=>{
      try{
        s.lineFormat.beginArrowheadStyle=PowerPoint.ArrowheadStyle[sv]??PowerPoint.ArrowheadStyle.none;
        s.lineFormat.endArrowheadStyle=PowerPoint.ArrowheadStyle[ev]??PowerPoint.ArrowheadStyle.none;
      }catch(_){}
    });
    await ctx.sync();
    if(ver===_applyVer)setStatus(`Ends → ${sv} / ${ev}`);
  });
}

// ─────────────────────────────────────────────────────────────────
// Insert New Capability
// ─────────────────────────────────────────────────────────────────
async function insertCapabilityShape() {
  await PowerPoint.run(async (context) => {
    const CM_TO_PT = 28.3465;
    const widthPt  = 5.6  * CM_TO_PT;
    const heightPt = 1.46 * CM_TO_PT;
    const leftPt   = (720 - widthPt)  / 2;
    const topPt    = (540 - heightPt) / 2;

    let slide;
    try {
      const sel = context.presentation.getSelectedSlides();
      sel.load('items');
      await context.sync();
      slide = sel.items.length
        ? context.presentation.getSelectedSlides().getItemAt(0)
        : context.presentation.slides.getItemAt(0);
    } catch (_) {
      slide = context.presentation.slides.getItemAt(0);
    }

    const shape = slide.shapes.addGeometricShape(
      PowerPoint.GeometricShapeType.roundedRectangle,
      { left: leftPt, top: topPt, width: widthPt, height: heightPt }
    );

    shape.name = 'Capability';
    const tf = shape.textFrame;
    tf.textRange.text = 'New Capability\n(xx)';
    tf.textRange.font.name = 'Aptos';
    tf.textRange.font.size = 12;
    try {
      tf.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
      tf.textRange.paragraphFormat.alignment =
        PowerPoint.ParagraphHorizontalAlignment.center;
    } catch (_) {}

    await context.sync();
    setStatus('Inserted New Capability shape');
  }).catch(err => setStatus('Insert failed: ' + err.message));
}

// ─────────────────────────────────────────────────────────────────
// Office ready
// ─────────────────────────────────────────────────────────────────
Office.onReady(async () => {
  // Load persisted staged state
  loadStagedState();

  // Render main swatches
  renderMainSwatches('fill');
  renderMainSwatches('border');

  // Style mode controls
  document.querySelectorAll('[data-weight]').forEach(btn=>
    btn.addEventListener('click',()=>applyBorderWeight(parseFloat(btn.dataset.weight))));
  document.querySelectorAll('[data-dash]').forEach(btn=>
    btn.addEventListener('click',()=>applyBorderDash(btn.dataset.dash)));
  document.getElementById('btn-apply-ends')?.addEventListener('click',applyLineEnds);
  document.getElementById('btn-insert-capability')?.addEventListener('click',insertCapabilityShape);
  document.getElementById('edit-back-btn').addEventListener('click',closeEditPanel);

  // ── Mode toggle ──
  document.getElementById('btn-mode-style').addEventListener('click',()=>switchMode('style'));
  document.getElementById('btn-mode-paint').addEventListener('click',()=>switchMode('paint'));

  // ── Paint panel controls ──
  document.getElementById('btn-capture').addEventListener('click',captureFromShape);

  document.getElementById('btn-apply-staged').addEventListener('click',applyStagedFormat);

  document.getElementById('btn-lock').addEventListener('click',()=>{
    STAGED.locked = !STAGED.locked;
    saveStagedState();
    updateLockToggle();
    setStatus(STAGED.locked
      ? 'Lock ON — format will auto-apply on every shape click'
      : 'Lock OFF — click "Apply to Selection" manually');
  });

  // Staged checkboxes — re-save when toggled
  ['fill','border','text'].forEach(attr=>{
    document.getElementById(`staged-${attr}-chk`)?.addEventListener('change',()=>saveStagedState());
  });

  // ── Selection changed — bifurcated by mode ──
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    async () => {
      if(_editKey) return; // Edit panel is open — don't interrupt
      if(_currentMode === 'paint') {
        await updateShapeCount();
        if(STAGED.locked) await applyStagedFormat();
      } else {
        inspectSelection();
      }
    }
  );

  inspectSelection();
  loadThemeColors();
});

// ─────────────────────────────────────────────────────────────────
// Shape inspection (Style mode)
// ─────────────────────────────────────────────────────────────────
async function inspectSelection() {
  try {
    await PowerPoint.run(async ctx=>{
      const sel=ctx.presentation.getSelectedShapes();
      sel.load('items/type,items/name');
      await ctx.sync();
      const items=sel.items;
      if(!items.length){renderEmpty();return;}
      _cachedShapeCount=items.length;
      const merged={fill:false,border:false,text:false,lineEnds:false};
      let allSame=true;
      const firstKey=toKey(items[0].type),firstName=items[0].name||'';
      items.forEach(s=>{
        const k=toKey(s.type);if(k!==firstKey)allSame=false;
        const caps=CAPABILITIES[k]||{};
        if(caps.fill)merged.fill=true;
        if(caps.border)merged.border=true;
        if(caps.text)merged.text=true;
        if(caps.lineEnds)merged.lineEnds=true;
      });
      const meta=allSame?(TYPE_META[firstKey]||{icon:'?',label:cap(firstKey||'Shape')}):{icon:'⊕',label:'Mixed Selection'};
      renderUI({merged,meta,firstName,count:items.length,anySupported:Object.values(merged).some(Boolean),isLine:allSame&&firstKey==='line'});
    });
  }catch(_){renderEmpty();}
}

function toKey(raw){if(!raw)return'unsupported';const s=String(raw);return s.charAt(0).toLowerCase()+s.slice(1);}
function cap(s){return s.charAt(0).toUpperCase()+s.slice(1);}

function renderUI({merged,meta,firstName,count,anySupported,isLine}) {
  hide('empty-state');hide('unsupported-state');show('shape-banner');
  el('shape-icon').textContent=meta.icon;
  el('shape-type-label').textContent=meta.label;
  el('shape-name').textContent=firstName?`"${firstName}"` :'';
  el('shape-count-badge').textContent=`×${count}`;
  tog('shape-count-badge',count>1);
  tog('section-fill',merged.fill);
  tog('section-border',merged.border);
  tog('section-line-ends',merged.lineEnds);
  el('border-heading').textContent=isLine?'Line Style':'Post Int Process Owner';
  if(!anySupported)show('unsupported-state');
}

function renderEmpty() {
  _cachedShapeCount=0;
  show('empty-state');hide('shape-banner');hide('unsupported-state');
  ['section-fill','section-border','section-line-ends'].forEach(hide);
  setStatus('');
}

function el(id){return document.getElementById(id);}
function show(id){el(id)?.classList.remove('hidden');}
function hide(id){el(id)?.classList.add('hidden');}
function tog(id,vis){el(id)?.classList.toggle('hidden',!vis);}
function setStatus(m){if(el('status'))el('status').textContent=m;}
