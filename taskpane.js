// ─────────────────────────────────────────────────────────────────
// CAPABILITY MAP
// ─────────────────────────────────────────────────────────────────
const CAPABILITIES = {
  geometricShape:{fill:true,border:true,text:true,lineEnds:false},
  textBox:       {fill:true,border:true,text:true,lineEnds:false},
  placeholder:   {fill:true,border:true,text:true,lineEnds:false},
  callout:       {fill:true,border:true,text:true,lineEnds:false},
  freeform:      {fill:true,border:true,text:true,lineEnds:false},
  group:         {fill:true,border:true,text:true,lineEnds:false},
  line:          {fill:false,border:true,text:false,lineEnds:true},
  image:         {fill:false,border:true,text:false,lineEnds:false},
  table:         {fill:false,border:true,text:true,lineEnds:false},
};
const TYPE_META = {
  geometricShape:{icon:'◻',label:'Geometric Shape'},
  textBox:       {icon:'T', label:'Text Box'},
  placeholder:   {icon:'⊞',label:'Placeholder'},
  callout:       {icon:'💬',label:'Callout'},
  freeform:      {icon:'✏',label:'Freeform'},
  group:         {icon:'⊡',label:'Group'},
  line:          {icon:'╱',label:'Line'},
  image:         {icon:'⬜',label:'Image'},
  table:         {icon:'⊟',label:'Table'},
};
const NO_LABEL = {fill:'No Fill',border:'No Border',text:'No Color'};

// ─────────────────────────────────────────────────────────────────
// DEFAULT PALETTES
// ─────────────────────────────────────────────────────────────────
const DEFAULTS = {
  fill:[
    {hex:'#D50032',name:'Target Converts to Acquiring',textHex:'#FFFFFF'},
    {hex:'#2A3E6D',name:'APi Corp Entity',            textHex:'#FFFFFF'},
    {hex:'#008579',name:'Shared Services',             textHex:'#FFFFFF'},
    {hex:'#595959',name:'Leave As Is',                 textHex:'#FFFFFF'},
    {hex:'#BFBFBF',name:'To Be Determined',            textHex:'#000000'},
    {hex:'#FFFFFF',name:'No Fill',                     textHex:'#000000'},
    {hex:'#000000',name:'Divest / Sunset',             textHex:'#FFFFFF'},
  ],
  border:[
    {hex:'#D4D800',name:'TBD'},
    {hex:'#D50032',name:'APi Corp'},
    {hex:'#7030A0',name:'APi Segment'},
    {hex:'#00B0F0',name:'Target HQ'},
    {hex:'#92D050',name:'Target OpCo'},
    {hex:'#808080',name:'No Reporting Structure'},
    {hex:'#2A3E6D',name:'Acquiring Entity'},
  ],
  text:[
    {hex:'#FFFFFF',name:'White'},
    {hex:'#000000',name:'Black'},
    {hex:'#2A3E6D',name:'Dark Navy'},
    {hex:'#D50032',name:'APi Red'},
    {hex:'#008579',name:'Teal'},
    {hex:'#595959',name:'Dark Gray'},
    {hex:'#BFBFBF',name:'Light Gray'},
  ],
};

const BRAND_COLORS = ['#151F37','#D50032','#2A3E6D','#008579','#595959','#BFBFBF','#FFFFFF'];
const STANDARD_COLORS = [
  '#C00000','#FF0000','#FFC000','#FFFF00','#92D050',
  '#00B050','#00B0F0','#0070C0','#002060','#7030A0'
];

// ─────────────────────────────────────────────────────────────────
// STATE
// ─────────────────────────────────────────────────────────────────
let _currentMode      = 'style';
let _cachedShapeCount = 0;
let _applyVer         = 0;
let _editKey          = null;
let _applyFn          = null;
let _openForm         = null;
let _themeColors      = null;

// STAGED — persisted across mode switches
const STAGED = {
  fill:   {dirty:false, hex:null, textHex:null, name:null},
  border: {dirty:false, hex:null, name:null},
  locked: false,
};

function saveStagedState() {
  try { localStorage.setItem('sct_staged', JSON.stringify(STAGED)); } catch(_) {}
}
function loadStagedState() {
  try {
    const s = JSON.parse(localStorage.getItem('sct_staged')||'{}');
    if(s.fill)   Object.assign(STAGED.fill,   s.fill);
    if(s.border) Object.assign(STAGED.border, s.border);
    if(typeof s.locked === 'boolean') STAGED.locked = s.locked;
  } catch(_) {}
}

// ─────────────────────────────────────────────────────────────────
// PALETTE STORAGE
// ─────────────────────────────────────────────────────────────────
function getPalette(key)         { try{const s=localStorage.getItem(`sct_${key}`);return s?JSON.parse(s):DEFAULTS[key]||[];}catch(_){return DEFAULTS[key]||[];} }
function savePalette(key,arr)    { try{localStorage.setItem(`sct_${key}`,JSON.stringify(arr));}catch(_){} }
function deleteColor(key,idx)    { const a=getPalette(key);a.splice(idx,1);savePalette(key,a); }
function pushRecent(key,hex)     {
  if(!hex||hex==='none')return;
  const k=`sct_recent_${key}`;
  try{
    let a=JSON.parse(localStorage.getItem(k)||'[]');
    a=a.filter(x=>x.toUpperCase()!==hex.toUpperCase());
    a.unshift(hex);if(a.length>10)a.length=10;
    localStorage.setItem(k,JSON.stringify(a));
  }catch(_){}
}
function getRecents(key)         { try{return JSON.parse(localStorage.getItem(`sct_recent_${key}`)||'[]');}catch(_){return[];} }

// ─────────────────────────────────────────────────────────────────
// COLOR UTILITIES
// ─────────────────────────────────────────────────────────────────
function isLight(hex) {
  const r=parseInt(hex.slice(1,3),16),g=parseInt(hex.slice(3,5),16),b=parseInt(hex.slice(5,7),16);
  return (r*299+g*587+b*114)/1000>155;
}
function autoTextHex(hex){ return isLight(hex)?'#000000':'#FFFFFF'; }
function normalizeHex(h) {
  if(!h)return null;
  const c=h.replace('#','').toUpperCase();
  if(c.length===6&&/^[0-9A-F]+$/.test(c))return '#'+c;
  if(c.length===3&&/^[0-9A-F]+$/.test(c))return '#'+c[0]+c[0]+c[1]+c[1]+c[2]+c[2];
  return null;
}
function hexToHsv(hex) {
  let r=parseInt(hex.slice(1,3),16)/255, g=parseInt(hex.slice(3,5),16)/255, b=parseInt(hex.slice(5,7),16)/255;
  const max=Math.max(r,g,b),min=Math.min(r,g,b),d=max-min;
  let h=0,s=max?d/max:0,v=max;
  if(d){if(max===r)h=((g-b)/d)%6;else if(max===g)h=(b-r)/d+2;else h=(r-g)/d+4;h=Math.round(h*60);if(h<0)h+=360;}
  return{h,s:Math.round(s*100),v:Math.round(v*100)};
}
function hsvToHex(h,s,v) {
  s/=100;v/=100;
  const f=(n,k=(n+h/60)%6)=>v-v*s*Math.max(Math.min(k,4-k,1),0);
  const toB=x=>Math.round(x*255).toString(16).padStart(2,'0');
  return '#'+toB(f(5))+toB(f(3))+toB(f(1));
}
function tint(hex,pct){ // pct 0-1: blend toward white
  const r=parseInt(hex.slice(1,3),16),g=parseInt(hex.slice(3,5),16),b=parseInt(hex.slice(5,7),16);
  const t=x=>Math.round(x+(255-x)*pct).toString(16).padStart(2,'0');
  return '#'+t(r)+t(g)+t(b);
}
function shade(hex,pct){ // pct 0-1: blend toward black
  const r=parseInt(hex.slice(1,3),16),g=parseInt(hex.slice(3,5),16),b=parseInt(hex.slice(5,7),16);
  const s=x=>Math.round(x*(1-pct)).toString(16).padStart(2,'0');
  return '#'+s(r)+s(g)+s(b);
}
function esc(s){ const d=document.createElement('div');d.textContent=s||'';return d.innerHTML; }

// ─────────────────────────────────────────────────────────────────
// THEME COLORS
// ─────────────────────────────────────────────────────────────────
async function loadThemeColors() {
  try {
    await PowerPoint.run(async ctx => {
      const master = ctx.presentation.slideMasters.getItemAt(0);
      const theme  = master.getTheme();
      theme.load('themeColorScheme');
      await ctx.sync();
      const cs = theme.themeColorScheme;
      _themeColors = [
        cs.dark1, cs.light1, cs.dark2, cs.light2,
        cs.accent1, cs.accent2, cs.accent3, cs.accent4, cs.accent5, cs.accent6
      ].map(c => normalizeHex('#'+c.replace('#','')) || '#CCCCCC');
    });
  } catch(_) { _themeColors = null; }
}
function getThemeCols() { return _themeColors || BRAND_COLORS; }

// ─────────────────────────────────────────────────────────────────
// MAIN SWATCH RENDER — Style mode
// ─────────────────────────────────────────────────────────────────
function renderMainSwatches(key) {
  const row = document.getElementById(`swatches-${key}`);
  if(!row) return;
  const palette = getPalette(key);

  if(key === 'border') {
    row.innerHTML = palette.map((c,i) => `
      <button class="main-swatch border-box-tile" data-color="${c.hex}" data-key="border"
              style="border-color:${c.hex};"
              title="${esc(c.name)}" aria-label="${esc(c.name)}">
        <span class="border-box-name" style="color:${c.hex};">${esc(c.name)}</span>
      </button>`).join('') +
      `<button class="edit-palette-btn" data-key="border" title="Edit border palette">
         <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
           <path d="M8.5 1.5a1.414 1.414 0 012 2L4 10H2V8L8.5 1.5z" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round" fill="none"/>
         </svg>
       </button>`;
  } else {
    row.innerHTML = palette.map((c,i) => {
      const txtColor = c.textHex || autoTextHex(c.hex);
      return `<button class="main-swatch${isLight(c.hex)?' light':''}"
                      style="background:${c.hex}"
                      data-color="${c.hex}"
                      data-textcolor="${txtColor}"
                      title="${esc(c.name)}" aria-label="${esc(c.name)}">
                <span class="swatch-label" style="color:${txtColor};">${esc(c.name)}</span>
              </button>`;
    }).join('') +
    `<button class="edit-palette-btn" data-key="${key}" title="Edit ${key} palette">
       <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
         <path d="M8.5 1.5a1.414 1.414 0 012 2L4 10H2V8L8.5 1.5z" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round" fill="none"/>
       </svg>
     </button>`;
  }

  // Wire clicks
  row.querySelectorAll('.main-swatch').forEach(btn => {
    btn.addEventListener('click', () => {
      const color    = btn.dataset.color;
      const txtColor = btn.dataset.textcolor;
      applyDOITile(color, txtColor);
      pushRecent(key, color);
    });
  });
  row.querySelectorAll('.main-swatch').forEach(btn => {
    btn.addEventListener('click', () => {
      applyBorderColor(btn.dataset.color);
      pushRecent('border', btn.dataset.color);
    });
  });
  row.querySelectorAll('.edit-palette-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const k = btn.dataset.key;
      const fnMap = {fill:applyDOITile, border:applyBorderColor, text:applyTextColor};
      openEditPanel(k, fnMap[k]);
    });
  });
}

// ─────────────────────────────────────────────────────────────────
// PAINT SWATCH RENDER — stages on click instead of applying
// ─────────────────────────────────────────────────────────────────
function renderPaintSwatches() {
  renderPaintFillSwatches();
  renderPaintBorderSwatches();
}

function renderPaintFillSwatches() {
  const row = document.getElementById('paint-swatches-fill');
  if(!row) return;
  const palette = getPalette('fill');
  const stagedHex = (STAGED.fill.dirty && STAGED.fill.hex) ? STAGED.fill.hex.toUpperCase() : null;

  row.innerHTML = palette.map(c => {
    const txtColor = c.textHex || autoTextHex(c.hex);
    const isActive = stagedHex && c.hex.toUpperCase() === stagedHex;
    return `<button class="main-swatch${isLight(c.hex)?' light':''}${isActive?' paint-staged':''}"
                    style="background:${c.hex};"
                    data-color="${c.hex}"
                    data-textcolor="${txtColor}"
                    data-name="${esc(c.name)}"
                    title="${esc(c.name)}" aria-label="${esc(c.name)}">
              <span class="swatch-label" style="color:${txtColor};">${esc(c.name)}</span>
              ${isActive?'<span class="staged-check">✓</span>':''}
            </button>`;
  }).join('');

  row.querySelectorAll('.main-swatch').forEach(btn => {
    btn.addEventListener('click', () => {
      const hex      = btn.dataset.color;
      const textHex  = btn.dataset.textcolor;
      const name     = btn.dataset.name;
      const already  = STAGED.fill.dirty && STAGED.fill.hex && STAGED.fill.hex.toUpperCase() === hex.toUpperCase();
      if(already) {
        // deselect
        STAGED.fill.dirty   = false;
        STAGED.fill.hex     = null;
        STAGED.fill.textHex = null;
        STAGED.fill.name    = null;
      } else {
        STAGED.fill.dirty   = true;
        STAGED.fill.hex     = hex;
        STAGED.fill.textHex = textHex;
        STAGED.fill.name    = name;
      }
      saveStagedState();
      renderPaintFillSwatches();
      updatePaintBadge('fill');
      updateApplyBtn();
    });
  });
  updatePaintBadge('fill');
}

function renderPaintBorderSwatches() {
  const row = document.getElementById('paint-swatches-border');
  if(!row) return;
  const palette   = getPalette('border');
  const stagedHex = (STAGED.border.dirty && STAGED.border.hex) ? STAGED.border.hex.toUpperCase() : null;

  row.innerHTML = palette.map(c => {
    const isActive = stagedHex && c.hex.toUpperCase() === stagedHex;
    return `<button class="main-swatch border-box-tile${isActive?' paint-staged':''}"
                    style="border-color:${c.hex};"
                    data-color="${c.hex}"
                    data-name="${esc(c.name)}"
                    title="${esc(c.name)}" aria-label="${esc(c.name)}">
              <span class="border-box-name" style="color:${c.hex};">${esc(c.name)}</span>
              ${isActive?'<span class="staged-check" style="color:'+c.hex+';">✓</span>':''}
            </button>`;
  }).join('');

  row.querySelectorAll('.main-swatch').forEach(btn => {
    btn.addEventListener('click', () => {
      const hex  = btn.dataset.color;
      const name = btn.dataset.name;
      const already = STAGED.border.dirty && STAGED.border.hex && STAGED.border.hex.toUpperCase() === hex.toUpperCase();
      if(already) {
        STAGED.border.dirty = false;
        STAGED.border.hex   = null;
        STAGED.border.name  = null;
      } else {
        STAGED.border.dirty = true;
        STAGED.border.hex   = hex;
        STAGED.border.name  = name;
      }
      saveStagedState();
      renderPaintBorderSwatches();
      updatePaintBadge('border');
      updateApplyBtn();
    });
  });
  updatePaintBadge('border');
}

function updatePaintBadge(attr) {
  const badge = document.getElementById(`paint-${attr}-badge`);
  if(!badge) return;
  if(attr === 'fill' && STAGED.fill.dirty && STAGED.fill.hex) {
    const txHex = STAGED.fill.textHex || autoTextHex(STAGED.fill.hex);
    badge.style.cssText = `display:inline-flex;background:${STAGED.fill.hex};color:${txHex};`;
    badge.textContent   = STAGED.fill.name || STAGED.fill.hex;
    badge.classList.add('has-value');
  } else if(attr === 'border' && STAGED.border.dirty && STAGED.border.hex) {
    badge.innerHTML = `<span style="display:inline-block;width:28px;height:5px;background:${STAGED.border.hex};border-radius:2px;vertical-align:middle;margin-right:5px;"></span>${esc(STAGED.border.name||STAGED.border.hex)}`;
    badge.style.cssText = '';
    badge.classList.add('has-value');
  } else {
    badge.textContent = '';
    badge.style.cssText = '';
    badge.classList.remove('has-value');
  }
}

function updateApplyBtn() {
  const btn = document.getElementById('btn-apply-staged');
  if(!btn) return;
  const anyStaged = STAGED.fill.dirty || STAGED.border.dirty;
  btn.disabled = !anyStaged;
  btn.classList.toggle('apply-ready', anyStaged);
}

function updateLockToggle() {
  const btn = document.getElementById('btn-lock');
  const lbl = document.getElementById('lock-state-lbl');
  if(!btn) return;
  btn.setAttribute('aria-checked', STAGED.locked ? 'true' : 'false');
  btn.classList.toggle('lock-on', STAGED.locked);
  if(lbl) lbl.textContent = STAGED.locked ? 'On' : 'Off';
}

function updatePaintInstruction() {
  const txt = document.getElementById('paint-instruction-text');
  if(!txt) return;
  const anyStaged = STAGED.fill.dirty || STAGED.border.dirty;
  if(STAGED.locked) {
    txt.textContent = anyStaged
      ? 'Auto-paint ON — click any shape to apply staged format.'
      : 'Auto-paint ON — select attributes above first.';
  } else {
    txt.textContent = anyStaged
      ? 'Attributes staged — select shapes then click "Apply".'
      : 'Select attributes above, then click shapes to apply.';
  }
}

// ─────────────────────────────────────────────────────────────────
// TRASH ICON
// ─────────────────────────────────────────────────────────────────
function trashIcon() {
  return `<svg width="11" height="12" viewBox="0 0 11 12" fill="none">
    <path d="M1 3h9M4 3V2a.5.5 0 01.5-.5h2A.5.5 0 017 2v1M9 3l-.5 7a.5.5 0 01-.5.5H3a.5.5 0 01-.5-.5L2 3"
          stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/>
  </svg>`;
}

// ─────────────────────────────────────────────────────────────────
// EDIT PANEL
// ─────────────────────────────────────────────────────────────────
function openEditPanel(key, fn) {
  _editKey=key; _applyFn=fn; _openForm=null;
  renderEditPanel();
  document.getElementById('main-content').classList.add('hidden');
  document.getElementById('paint-panel').classList.add('hidden');
  document.getElementById('edit-panel').classList.remove('hidden');
}
function closeEditPanel() {
  document.getElementById('edit-panel').classList.add('hidden');
  if(_currentMode==='paint') {
    document.getElementById('paint-panel').classList.remove('hidden');
    renderPaintSwatches();
  } else {
    document.getElementById('main-content').classList.remove('hidden');
  }
  renderMainSwatches(_editKey);
  _editKey=null; _applyFn=null; _openForm=null;
}

function renderEditPanel() {
  const key = _editKey;
  const titleMap = {fill:'Degree of Integration',border:'Post Int Process Owner',text:'Text Color'};
  document.getElementById('edit-panel-title').textContent = titleMap[key]||'Edit Colors';
  const palette = getPalette(key);

  const myColorsHtml = palette.map((c,i) => {
    const isEditing = _openForm && _openForm.type==='edit' && _openForm.idx===i;
    const txHex = c.textHex || autoTextHex(c.hex);
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
            <button class="my-del-btn" data-idx="${i}" title="Delete">${trashIcon()}</button>
          </div>
        </div>
        ${isEditing ? buildColorForm('edit',i,c.hex,c.name,txHex) : ''}
      </div>`;
  }).join('');

  const isAdding = _openForm && _openForm.type==='add';
  document.getElementById('edit-panel-body').innerHTML = `
    <div class="ep-label">My Colors <span class="ep-hint">drag to reorder · click swatch to apply</span></div>
    <div id="my-colors-list">${myColorsHtml}</div>
    <button class="add-new-btn${isAdding?' active':''}" id="ep-add-toggle">
      <svg width="11" height="11" viewBox="0 0 11 11" fill="none">
        <line x1="5.5" y1="1" x2="5.5" y2="10" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/>
        <line x1="1" y1="5.5" x2="10" y2="5.5" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/>
      </svg>
      Add New Color
    </button>
    ${isAdding ? buildColorForm('add',null,'#2A3E6D','','#FFFFFF') : ''}
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
  const body = document.getElementById('edit-panel-body');

  body.querySelectorAll('.my-swatch').forEach(s => {
    const go = () => {
      const color=s.dataset.color, tc=s.dataset.textcolor;
      if(key==='fill' && tc) applyDOITile(color,tc);
      else _applyFn?.(color);
      if(color!=='none') pushRecent(key,color);
      closeEditPanel();
    };
    s.addEventListener('click', go);
    s.addEventListener('keydown', e => { if(e.key==='Enter'||e.key===' '){e.preventDefault();go();} });
  });

  body.querySelectorAll('.my-edit-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const idx = parseInt(btn.dataset.idx,10);
      _openForm = (_openForm && _openForm.type==='edit' && _openForm.idx===idx) ? null : {type:'edit',idx};
      renderEditPanel();
    });
  });

  body.querySelectorAll('.my-del-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      deleteColor(key, parseInt(btn.dataset.idx,10));
      _openForm=null; renderEditPanel();
    });
  });

  document.getElementById('ep-add-toggle').addEventListener('click', () => {
    _openForm = (_openForm && _openForm.type==='add') ? null : {type:'add',idx:null};
    renderEditPanel();
  });

  document.getElementById('ep-no-color').addEventListener('click', () => {
    if(key==='fill') applyDOITile('none',null);
    else _applyFn?.('none');
    closeEditPanel();
  });

  document.getElementById('ep-reset').addEventListener('click', () => {
    if(!confirm(`Reset ${key} palette to defaults? Your custom colors will be lost.`)) return;
    localStorage.removeItem(`sct_${key}`);
    _openForm=null; renderEditPanel(); renderMainSwatches(key);
  });

  body.querySelectorAll('.color-form').forEach(form => wireColorForm(form, key));

  // Drag-and-drop
  const list = document.getElementById('my-colors-list');
  let dragSrcIdx = null;
  list.querySelectorAll('.my-color-item[draggable]').forEach(item => {
    item.addEventListener('dragstart', e => {
      dragSrcIdx = parseInt(item.dataset.idx,10);
      item.classList.add('dragging');
      e.dataTransfer.effectAllowed='move';
    });
    item.addEventListener('dragend', () => {
      item.classList.remove('dragging');
      list.querySelectorAll('.my-color-item').forEach(i=>i.classList.remove('drag-over'));
    });
    item.addEventListener('dragover', e => {
      e.preventDefault(); e.dataTransfer.dropEffect='move';
      list.querySelectorAll('.my-color-item').forEach(i=>i.classList.remove('drag-over'));
      item.classList.add('drag-over');
    });
    item.addEventListener('drop', e => {
      e.preventDefault();
      const destIdx = parseInt(item.dataset.idx,10);
      if(dragSrcIdx===null || dragSrcIdx===destIdx) return;
      const arr = getPalette(key);
      const [moved] = arr.splice(dragSrcIdx,1);
      arr.splice(destIdx,0,moved);
      savePalette(key,arr);
      _openForm=null; renderEditPanel();
    });
  });
}

// ─────────────────────────────────────────────────────────────────
// COLOR FORM (inside edit panel)
// ─────────────────────────────────────────────────────────────────
function buildColorForm(mode, idx, initFillHex, initName, initTextHex) {
  const fid = `cf-${mode}-${idx??'new'}`;
  const isDOI = _editKey === 'fill';
  const themeCols = getThemeCols();

  function themeGrid(prefix) {
    return themeCols.map((base,ci) => {
      const variants = [base, tint(base,.8), tint(base,.6), tint(base,.4), shade(base,.25), shade(base,.5)];
      return variants.map((v,ri) =>
        `<button class="ppt-swatch ppt-theme-swatch" data-hex="${v}" data-form="${fid}"
                 data-picker="${prefix}" title="${v}"
                 style="background:${v};width:18px;height:${ri===0?20:14}px;"></button>`
      ).join('');
    }).join('');
  }

  function stdRow(prefix) {
    return STANDARD_COLORS.map(c =>
      `<button class="ppt-swatch ppt-std-swatch" data-hex="${c}" data-form="${fid}"
               data-picker="${prefix}" title="${c}"
               style="background:${c};width:18px;height:18px;"></button>`
    ).join('');
  }

  function recentRow(prefix) {
    const recs = getRecents(_editKey);
    if(!recs.length) return '';
    return `
      <div class="acc-hdr" data-acc="acc-recent-${prefix}-${fid}">
        <span>Recent Colors</span><svg class="acc-chev" width="10" height="10" viewBox="0 0 10 10"><polyline points="2,3 5,7 8,3" fill="none" stroke="currentColor" stroke-width="1.5"/></svg>
      </div>
      <div class="acc-body" id="acc-recent-${prefix}-${fid}">
        <div class="ppt-row">${recs.map(c=>`<button class="ppt-swatch ppt-std-swatch" data-hex="${c}" data-form="${fid}" data-picker="${prefix}" title="${c}" style="background:${c};width:18px;height:18px;"></button>`).join('')}</div>
      </div>`;
  }

  function picker(prefix, label, iconHtml, initHex) {
    return `
      <div class="ppt-picker-hdr ppt-sect-hdr" data-sect-key="${prefix}-${fid}">
        <span class="ppt-picker-icon">${iconHtml}</span>
        <span class="ppt-picker-label">${label}</span>
        <span class="ppt-picker-preview" id="prev-${prefix}-${fid}" style="background:${initHex||'#FFFFFF'};${isLight(initHex||'#FFFFFF')?'border:1px solid #ddd;':''}"></span>
        <svg class="ppt-sect-chev" width="10" height="10" viewBox="0 0 10 10"><polyline points="2,3 5,7 8,3" fill="none" stroke="currentColor" stroke-width="1.5"/></svg>
      </div>
      <div class="ppt-sect-body" id="sect-${prefix}-${fid}" style="${prefix==='fill'?'':'display:none;'}">
        <div class="acc-hdr acc-open" data-acc="acc-theme-${prefix}-${fid}">
          <span>Theme Colors</span><svg class="acc-chev" width="10" height="10" viewBox="0 0 10 10"><polyline points="2,3 5,7 8,3" fill="none" stroke="currentColor" stroke-width="1.5"/></svg>
        </div>
        <div class="acc-body acc-open" id="acc-theme-${prefix}-${fid}">
          <div class="ppt-theme-grid">${themeGrid(prefix)}</div>
        </div>

        <div class="acc-hdr" data-acc="acc-std-${prefix}-${fid}">
          <span>Standard Colors</span><svg class="acc-chev" width="10" height="10" viewBox="0 0 10 10"><polyline points="2,3 5,7 8,3" fill="none" stroke="currentColor" stroke-width="1.5"/></svg>
        </div>
        <div class="acc-body" id="acc-std-${prefix}-${fid}">
          <div class="ppt-row">${stdRow(prefix)}</div>
        </div>

        ${recentRow(prefix)}

        <div class="acc-hdr" data-acc="acc-more-${prefix}-${fid}">
          <span>More Colors</span><svg class="acc-chev" width="10" height="10" viewBox="0 0 10 10"><polyline points="2,3 5,7 8,3" fill="none" stroke="currentColor" stroke-width="1.5"/></svg>
        </div>
        <div class="acc-body" id="acc-more-${prefix}-${fid}">
          <canvas class="hsv-square" id="hsv-sq-${prefix}-${fid}" width="180" height="150"></canvas>
          <div class="hue-slider-wrap"><input type="range" class="hue-slider" id="hue-sl-${prefix}-${fid}" min="0" max="360" value="${hexToHsv(initHex||'#2A3E6D').h}"/></div>
          <div class="hex-input-row">
            <span class="hex-hash">#</span>
            <input type="text" class="hex-input" id="hex-in-${prefix}-${fid}" maxlength="6" value="${(initHex||'#2A3E6D').replace('#','')}"/>
            <span class="hex-preview" id="hex-prev-${prefix}-${fid}" style="background:${initHex||'#2A3E6D'};"></span>
          </div>
        </div>
      </div>`;
  }

  const fillIcon   = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 11H5m14 0a2 2 0 012 2v6a2 2 0 01-2 2H5a2 2 0 01-2-2v-6a2 2 0 012-2"/><path d="M12 11V3m0 0L8 7m4-4l4 4"/></svg>`;
  const textAIcon  = `<span style="font-weight:700;font-size:13px;border-bottom:3px solid currentColor;padding-bottom:1px;">A</span>`;

  return `
    <div class="color-form" data-form-id="${fid}" data-mode="${mode}" data-idx="${idx??''}" data-key="${_editKey}">
      <div class="cf-name-row">
        <label class="cf-name-lbl">Name</label>
        <input type="text" class="cf-name-input" id="cfname-${fid}" value="${esc(initName||'')}"/>
      </div>
      ${picker('fill', isDOI?'Shape Fill':'Color', fillIcon, initFillHex)}
      ${isDOI ? picker('text', 'Font Color', textAIcon, initTextHex) : ''}
      <div class="cf-actions">
        <button class="cf-save-btn" data-form="${fid}">${mode==='edit'?'Save Changes':'+ Add to My Colors'}</button>
        <button class="cf-cancel-btn" data-form="${fid}">Cancel</button>
      </div>
    </div>`;
}

function wireColorForm(form, key) {
  const fid     = form.dataset.formId;
  const isDOI   = key === 'fill';
  const prefixes = isDOI ? ['fill','text'] : ['fill'];

  // Collapsible section headers (Shape Fill / Font Color)
  form.querySelectorAll('.ppt-sect-hdr').forEach(hdr => {
    hdr.addEventListener('click', () => {
      const sectKey = hdr.dataset.sectKey;
      const [prefix, ...rest] = sectKey.split('-');
      const bodyId = `sect-${sectKey}`;
      const body = document.getElementById(bodyId);
      if(!body) return;
      const open = body.style.display !== 'none';
      body.style.display = open ? 'none' : '';
      hdr.classList.toggle('ppt-sect-open', !open);
    });
  });

  // Accordion toggles
  form.querySelectorAll('.acc-hdr').forEach(hdr => {
    hdr.addEventListener('click', () => {
      const bodyId = hdr.dataset.acc;
      const body   = document.getElementById(bodyId);
      if(!body) return;
      const open = body.classList.contains('acc-open');
      body.classList.toggle('acc-open', !open);
      hdr.classList.toggle('acc-open', !open);
      if(!open) initHsvIfNeeded(hdr.dataset.acc, fid);
    });
  });

  // Color swatch clicks
  form.querySelectorAll('.ppt-swatch').forEach(sw => {
    sw.addEventListener('click', () => {
      const hex    = sw.dataset.hex;
      const picker = sw.dataset.picker;
      setPickerColor(fid, picker, hex);
      markActive(form, picker, hex);
    });
  });

  // Hue sliders + HSV
  prefixes.forEach(prefix => {
    const hueSl  = document.getElementById(`hue-sl-${prefix}-${fid}`);
    const hexIn  = document.getElementById(`hex-in-${prefix}-${fid}`);
    const hexPrv = document.getElementById(`hex-prev-${prefix}-${fid}`);
    if(hueSl) {
      hueSl.addEventListener('input', () => {
        const sq  = document.getElementById(`hsv-sq-${prefix}-${fid}`);
        if(sq && sq._sData) {
          sq._sData.h = parseInt(hueSl.value,10);
          drawHsvSquare(sq);
          const hex = hsvToHex(sq._sData.h, sq._sData.s, sq._sData.v);
          setPickerColor(fid, prefix, hex);
        }
      });
    }
    if(hexIn) {
      hexIn.addEventListener('input', () => {
        const raw = hexIn.value.replace('#','');
        if(raw.length===6 && /^[0-9A-Fa-f]+$/.test(raw)) {
          const hex = '#'+raw.toUpperCase();
          if(hexPrv) hexPrv.style.background = hex;
          syncPickerPreview(fid, prefix, hex);
        }
      });
    }
  });

  // Save / Cancel
  form.querySelector('.cf-save-btn')?.addEventListener('click', () => {
    const nameVal = document.getElementById(`cfname-${fid}`)?.value.trim() || 'Custom';
    const hexIn   = document.getElementById(`hex-in-fill-${fid}`);
    const raw     = hexIn?.value.replace('#','').toUpperCase();
    const hex     = raw && raw.length===6 ? '#'+raw : (STAGED.fill.hex||'#2A3E6D');
    const norm    = normalizeHex(hex);
    if(!norm) { setStatus('Invalid hex color'); return; }

    let entry;
    if(key==='fill') {
      const txIn  = document.getElementById(`hex-in-text-${fid}`);
      const txRaw = txIn?.value.replace('#','').toUpperCase();
      const txHex = (txRaw&&txRaw.length===6) ? '#'+txRaw : autoTextHex(norm);
      entry = {hex:norm, name:nameVal, textHex:txHex};
    } else {
      entry = {hex:norm, name:nameVal};
    }

    const arr = getPalette(key);
    const mode = form.dataset.mode;
    const idx  = form.dataset.idx;
    if(mode==='edit' && idx!=='') arr[parseInt(idx,10)] = entry;
    else arr.push(entry);
    savePalette(key, arr);
    pushRecent(key, norm);
    _openForm=null; renderEditPanel(); renderMainSwatches(key);
    setStatus(`${mode==='edit'?'Updated':'Added'}: ${nameVal}`);
  });

  form.querySelector('.cf-cancel-btn')?.addEventListener('click', () => {
    _openForm=null; renderEditPanel();
  });

  // Init visible HSV squares
  prefixes.forEach(prefix => {
    const sq = document.getElementById(`hsv-sq-${prefix}-${fid}`);
    if(sq) {
      const hexIn = document.getElementById(`hex-in-${prefix}-${fid}`);
      const initHex = hexIn ? '#'+(hexIn.value||'2A3E6D').replace('#','') : '#2A3E6D';
      const {h,s,v} = hexToHsv(initHex);
      sq._sData = {h, s, v};
      drawHsvSquare(sq);
      wireHsvSquare(sq, fid, prefix);
    }
  });
}

function initHsvIfNeeded(bodyId, fid) {
  const body = document.getElementById(bodyId);
  if(!body) return;
  body.querySelectorAll('.hsv-square').forEach(sq => {
    if(sq._sData) return;
    const prefix = sq.id.replace(`hsv-sq-`,'').replace(`-${fid}`,'');
    const hexIn = document.getElementById(`hex-in-${prefix}-${fid}`);
    const initHex = hexIn ? '#'+(hexIn.value||'2A3E6D').replace('#','') : '#2A3E6D';
    const {h,s,v} = hexToHsv(initHex);
    sq._sData = {h, s, v};
    drawHsvSquare(sq);
    wireHsvSquare(sq, fid, prefix);
  });
}

function drawHsvSquare(sq) {
  const ctx = sq.getContext('2d');
  const w=sq.width, h=sq.height;
  const {h:hue} = sq._sData;
  const baseHex = hsvToHex(hue,100,100);
  const gradH = ctx.createLinearGradient(0,0,w,0);
  gradH.addColorStop(0,'#FFFFFF'); gradH.addColorStop(1,baseHex);
  ctx.fillStyle=gradH; ctx.fillRect(0,0,w,h);
  const gradV = ctx.createLinearGradient(0,0,0,h);
  gradV.addColorStop(0,'rgba(0,0,0,0)'); gradV.addColorStop(1,'rgba(0,0,0,1)');
  ctx.fillStyle=gradV; ctx.fillRect(0,0,w,h);
  const {s,v} = sq._sData;
  const cx=Math.round(s/100*w), cy=Math.round((1-v/100)*h);
  ctx.beginPath(); ctx.arc(cx,cy,6,0,2*Math.PI);
  ctx.strokeStyle='#fff'; ctx.lineWidth=2; ctx.stroke();
  ctx.beginPath(); ctx.arc(cx,cy,5,0,2*Math.PI);
  ctx.strokeStyle='rgba(0,0,0,.4)'; ctx.lineWidth=1; ctx.stroke();
}

function wireHsvSquare(sq, fid, prefix) {
  let down=false;
  const pick = e => {
    const r=sq.getBoundingClientRect();
    sq._sData.s = Math.round(Math.max(0,Math.min(1,(e.clientX-r.left)/r.width))*100);
    sq._sData.v = Math.round(Math.max(0,Math.min(1,1-(e.clientY-r.top)/r.height))*100);
    drawHsvSquare(sq);
    const hex = hsvToHex(sq._sData.h, sq._sData.s, sq._sData.v);
    setPickerColor(fid, prefix, hex);
  };
  sq.addEventListener('mousedown', e=>{ down=true; pick(e); });
  window.addEventListener('mousemove', e=>{ if(down) pick(e); });
  window.addEventListener('mouseup', ()=>{ down=false; });
}

function setPickerColor(fid, prefix, hex) {
  const hexIn  = document.getElementById(`hex-in-${prefix}-${fid}`);
  const hexPrv = document.getElementById(`hex-prev-${prefix}-${fid}`);
  const topPrv = document.getElementById(`prev-${prefix}-${fid}`);
  if(hexIn) hexIn.value = hex.replace('#','');
  if(hexPrv) hexPrv.style.background = hex;
  if(topPrv) { topPrv.style.background=hex; topPrv.style.border=isLight(hex)?'1px solid #ddd':''; }
  const hueSl = document.getElementById(`hue-sl-${prefix}-${fid}`);
  const sq    = document.getElementById(`hsv-sq-${prefix}-${fid}`);
  if(sq && sq._sData) {
    const {h,s,v} = hexToHsv(hex);
    sq._sData = {h,s,v};
    drawHsvSquare(sq);
    if(hueSl) hueSl.value = h;
  }
}

function syncPickerPreview(fid, prefix, hex) {
  const topPrv = document.getElementById(`prev-${prefix}-${fid}`);
  if(topPrv) { topPrv.style.background=hex; topPrv.style.border=isLight(hex)?'1px solid #ddd':''; }
}

function markActive(form, prefix, hex) {
  form.querySelectorAll(`[data-picker="${prefix}"]`).forEach(s=>s.classList.remove('ppt-swatch-active'));
  form.querySelectorAll(`[data-picker="${prefix}"][data-hex="${hex}"]`).forEach(s=>s.classList.add('ppt-swatch-active'));
}

// ─────────────────────────────────────────────────────────────────
// MODE SWITCHING
// ─────────────────────────────────────────────────────────────────
function switchMode(mode) {
  if(_currentMode===mode) return;
  _currentMode = mode;
  const header    = document.getElementById('app-header');
  const stylBtn   = document.getElementById('btn-mode-style');
  const paintBtn  = document.getElementById('btn-mode-paint');
  const mainCont  = document.getElementById('main-content');
  const paintPnl  = document.getElementById('paint-panel');
  const emptySt   = document.getElementById('empty-state');
  const shapeBnr  = document.getElementById('shape-banner');

  if(mode==='paint') {
    stylBtn.classList.remove('mode-active');
    paintBtn.classList.add('mode-active');
    header.classList.add('paint-mode-header');
    mainCont.classList.add('hidden');
    emptySt.classList.add('hidden');
    shapeBnr.classList.add('hidden');
    paintPnl.classList.remove('hidden');
    renderPaintSwatches();
    updateApplyBtn();
    updateLockToggle();
    updatePaintInstruction();
    setStatus('Paint mode — select attributes then click shapes');
  } else {
    paintBtn.classList.remove('mode-active');
    stylBtn.classList.add('mode-active');
    header.classList.remove('paint-mode-header');
    paintPnl.classList.add('hidden');
    mainCont.classList.remove('hidden');
    inspectSelection();
    setStatus('');
  }
}

// ─────────────────────────────────────────────────────────────────
// CAPTURE FROM SHAPE
// ─────────────────────────────────────────────────────────────────
async function captureFromShape() {
  try {
    await PowerPoint.run(async ctx => {
      const sel = ctx.presentation.getSelectedShapes();
      sel.load('items/type,items/name');
      await ctx.sync();
      if(!sel.items.length) { setStatus('Select a shape first.'); return; }
      const s = sel.items[0];
      s.fill.load('type,foreColor');
      s.lineFormat.load('color,visible');
      try { s.textFrame.textRange.font.load('color'); } catch(_) {}
      await ctx.sync();

      // Capture fill
      try {
        if(s.fill.type==='Solid'||s.fill.type===PowerPoint.ShapeFillType.solid) {
          const raw = s.fill.foreColor;
          const norm = normalizeHex(raw.startsWith('#')?raw:'#'+raw);
          if(norm) {
            const match = getPalette('fill').find(c=>c.hex.toUpperCase()===norm.toUpperCase());
            STAGED.fill = {
              dirty:true, hex:norm,
              textHex: match?.textHex || autoTextHex(norm),
              name: match?.name || norm
            };
          }
        }
      } catch(_) {}

      // Capture border
      try {
        if(s.lineFormat.visible!==false) {
          const raw = s.lineFormat.color;
          const norm = normalizeHex(raw.startsWith('#')?raw:'#'+raw);
          if(norm) {
            const match = getPalette('border').find(c=>c.hex.toUpperCase()===norm.toUpperCase());
            STAGED.border = {dirty:true, hex:norm, name:match?.name||norm};
          }
        }
      } catch(_) {}

      _cachedShapeCount = sel.items.length;
      saveStagedState();
      renderPaintSwatches();
      updateApplyBtn();
      updatePaintInstruction();
      setStatus(`Captured from "${sel.items[0].name||'shape'}"`);
    });
  } catch(e) { setStatus('Capture failed — select a shape first.'); }
}

// ─────────────────────────────────────────────────────────────────
// APPLY STAGED FORMAT
// ─────────────────────────────────────────────────────────────────
async function applyStagedFormat() {
  const doFill   = STAGED.fill.dirty   && STAGED.fill.hex;
  const doBorder = STAGED.border.dirty && STAGED.border.hex;
  if(!doFill && !doBorder) { setStatus('Select at least one attribute above.'); return; }

  const ver = ++_applyVer;
  try {
    await PowerPoint.run(async ctx => {
      if(ver!==_applyVer) return;
      // ── One run: load count + write properties in two syncs ──────────────
      const col = ctx.presentation.getSelectedShapes();
      const cnt = col.getCount();
      await ctx.sync();                              // sync 1 — get count
      if(ver!==_applyVer) return;
      _cachedShapeCount = cnt.value;
      if(!_cachedShapeCount) { setStatus('No shape selected.'); return; }

      for(let i=0; i<_cachedShapeCount; i++) {
        const s = col.getItemAt(i);
        if(doFill) {
          if(STAGED.fill.hex==='none') s.fill.clear();
          else {
            s.fill.setSolidColor(STAGED.fill.hex);
            const tx = STAGED.fill.textHex || autoTextHex(STAGED.fill.hex);
            try { s.textFrame.textRange.font.color=tx; } catch(_) {}
          }
        }
        if(doBorder) {
          if(STAGED.border.hex==='none') s.lineFormat.visible=false;
          else { s.lineFormat.visible=true; s.lineFormat.color=STAGED.border.hex; }
        }
      }
      await ctx.sync();                              // sync 2 — commit writes
      if(ver===_applyVer) {
        const parts=[];
        if(doFill)   parts.push(STAGED.fill.name   || STAGED.fill.hex);
        if(doBorder) parts.push(STAGED.border.name || STAGED.border.hex);
        setStatus('Painted — ' + parts.join(' + '));
      }
    });
  } catch(e) { setStatus('Apply failed — select a shape first.'); }
}

async function updateShapeCount() {
  try {
    await PowerPoint.run(async ctx=>{
      const col = ctx.presentation.getSelectedShapes();
      const cnt = col.getCount();
      await ctx.sync();
      _cachedShapeCount = cnt.value;
    });
  } catch(_) { _cachedShapeCount=0; }
}

// ─────────────────────────────────────────────────────────────────
// APPLY FUNCTIONS — Style mode
// ─────────────────────────────────────────────────────────────────
function getSelectedFast(ctx) {
  const col=ctx.presentation.getSelectedShapes();
  const out=[];
  for(let i=0;i<_cachedShapeCount;i++) out.push(col.getItemAt(i));
  return out;
}

async function applyDOITile(fillColor, textColor) {
  const ver=++_applyVer;
  if(!_cachedShapeCount) return setStatus('No shape selected.');
  await PowerPoint.run(async ctx=>{
    if(ver!==_applyVer) return;
    getSelectedFast(ctx).forEach(s=>{
      if(fillColor==='none') s.fill.clear();
      else s.fill.setSolidColor(fillColor);
      if(textColor&&textColor!=='none') try{s.textFrame.textRange.font.color=textColor;}catch(_){}
    });
    await ctx.sync();
    if(ver===_applyVer) setStatus(`DOI → ${fillColor==='none'?'cleared':fillColor}`);
  });
}

async function applyBorderColor(color) {
  const ver=++_applyVer;
  if(!_cachedShapeCount) return setStatus('No shape selected.');
  await PowerPoint.run(async ctx=>{
    if(ver!==_applyVer) return;
    getSelectedFast(ctx).forEach(s=>{
      if(color==='none') s.lineFormat.visible=false;
      else { s.lineFormat.visible=true; s.lineFormat.color=color; }
    });
    await ctx.sync();
    if(ver===_applyVer) setStatus(`Border → ${color==='none'?'removed':color}`);
  });
}

async function applyBorderWeight(pts) {
  const ver=++_applyVer;
  if(!_cachedShapeCount) return setStatus('No shape selected.');
  await PowerPoint.run(async ctx=>{
    if(ver!==_applyVer) return;
    getSelectedFast(ctx).forEach(s=>{ s.lineFormat.weight=pts; s.lineFormat.visible=true; });
    await ctx.sync();
    if(ver===_applyVer) setStatus(`Weight → ${pts}pt`);
  });
}

async function applyBorderDash(style) {
  const ver=++_applyVer;
  if(!_cachedShapeCount) return setStatus('No shape selected.');
  await PowerPoint.run(async ctx=>{
    if(ver!==_applyVer) return;
    let dashVal;
    switch(style) {
      case 'dash':     dashVal=PowerPoint.LineDashStyle.dash;     break;
      case 'roundDot': dashVal=PowerPoint.LineDashStyle.roundDot; break;
      default:         dashVal=PowerPoint.LineDashStyle.solid;    break;
    }
    getSelectedFast(ctx).forEach(s=>{ s.lineFormat.dashStyle=dashVal; s.lineFormat.visible=true; });
    await ctx.sync();
    if(ver===_applyVer) setStatus(`Style → ${style}`);
  });
}

async function applyTextColor(color) {
  const ver=++_applyVer;
  if(!_cachedShapeCount) return setStatus('No shape selected.');
  await PowerPoint.run(async ctx=>{
    if(ver!==_applyVer) return;
    getSelectedFast(ctx).forEach(s=>{ try{s.textFrame.textRange.font.color=color;}catch(_){} });
    await ctx.sync();
    if(ver===_applyVer) setStatus(`Text → ${color}`);
  });
}

async function applyLineEnds() {
  const sv=el('line-start').value, ev=el('line-end').value;
  const ver=++_applyVer;
  if(!_cachedShapeCount) return setStatus('No shape selected.');
  await PowerPoint.run(async ctx=>{
    if(ver!==_applyVer) return;
    getSelectedFast(ctx).forEach(s=>{
      try {
        s.lineFormat.beginArrowheadStyle=PowerPoint.ArrowheadStyle[sv]??PowerPoint.ArrowheadStyle.none;
        s.lineFormat.endArrowheadStyle  =PowerPoint.ArrowheadStyle[ev]??PowerPoint.ArrowheadStyle.none;
      } catch(_) {}
    });
    await ctx.sync();
    if(ver===_applyVer) setStatus(`Ends → ${sv} / ${ev}`);
  });
}

// ─────────────────────────────────────────────────────────────────
// INSERT NEW CAPABILITY
// ─────────────────────────────────────────────────────────────────
async function insertCapabilityShape() {
  await PowerPoint.run(async ctx => {
    // Convert cm to points (1 cm = 28.3465 pt)
    const CM = 28.3465;
    const w    = 5.6  * CM;   // 158.74 pt
    const h    = 1.46 * CM;   // 41.39 pt

    // Get the active slide
    let slide;
    try {
      const sel = ctx.presentation.getSelectedSlides();
      sel.load('items');
      await ctx.sync();
      slide = sel.items.length ? sel.items[0] : ctx.presentation.slides.getItemAt(0);
    } catch (_) {
      slide = ctx.presentation.slides.getItemAt(0);
    }

    // Add shape — set geometry AFTER creation to avoid Online options-ignore bug
    const shape = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.roundedRectangle);
    shape.width  = w;
    shape.height = h;
    shape.left   = (720 - w) / 2;
    shape.top    = (540 - h) / 2;
    shape.name   = 'Capability';

    // Default white fill, light border
    shape.fill.setSolidColor('#FFFFFF');
    shape.lineFormat.color   = '#2A3E6D';
    shape.lineFormat.weight  = 1.5;
    shape.lineFormat.visible = true;

    // Text — two paragraphs so each line is independently aligned
    const tf = shape.textFrame;
    tf.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;

    // Clear default and set line 1
    const range = tf.textRange;
    range.text = 'New Capability';
    range.font.name = 'Aptos';
    range.font.size = 12;
    range.font.color = '#151F37';

    // Add line 2 as a second paragraph
    const para2 = tf.textRange.paragraphs.add();
    para2.text = '(xx)';
    para2.font.name = 'Aptos';
    para2.font.size = 12;
    para2.font.color = '#151F37';

    // Vertical + horizontal alignment
    try {
      tf.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
      range.paragraphFormat.alignment          = PowerPoint.ParagraphHorizontalAlignment.center;
      para2.paragraphFormat.alignment          = PowerPoint.ParagraphHorizontalAlignment.center;
    } catch (_) {}

    await ctx.sync();
    setStatus('Inserted New Capability shape');
  }).catch(e => setStatus('Insert failed: ' + e.message));
}

// ─────────────────────────────────────────────────────────────────
// OFFICE READY
// ─────────────────────────────────────────────────────────────────
Office.onReady(async () => {
  loadStagedState();
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

  // Mode toggle
  document.getElementById('btn-mode-style').addEventListener('click',()=>switchMode('style'));
  document.getElementById('btn-mode-paint').addEventListener('click',()=>switchMode('paint'));

  // Paint panel
  document.getElementById('btn-capture').addEventListener('click',captureFromShape);

  document.getElementById('btn-apply-staged').addEventListener('click',applyStagedFormat);

  document.getElementById('btn-lock').addEventListener('click',()=>{
    STAGED.locked=!STAGED.locked;
    saveStagedState();
    updateLockToggle();
    updatePaintInstruction();
    setStatus(STAGED.locked?'Auto-paint ON — click shapes to apply':'Auto-paint OFF');
  });

  // Clear staged buttons
  document.getElementById('paint-fill-clear')?.addEventListener('click',()=>{
    STAGED.fill={dirty:false,hex:null,textHex:null,name:null};
    saveStagedState(); renderPaintFillSwatches(); updateApplyBtn(); updatePaintInstruction();
  });
  document.getElementById('paint-border-clear')?.addEventListener('click',()=>{
    STAGED.border={dirty:false,hex:null,name:null};
    saveStagedState(); renderPaintBorderSwatches(); updateApplyBtn(); updatePaintInstruction();
  });

  // Selection change — bifurcated by mode
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    async () => {
      if(_editKey) return;
      if(_currentMode==='paint') {
        const anyStaged = STAGED.fill.dirty || STAGED.border.dirty;
        if(STAGED.locked && anyStaged) {
          // Combined: count + apply in one PowerPoint.run — fastest path
          await applyStagedFormat();
        } else {
          // Lightweight count-only update (no apply)
          await updateShapeCount();
        }
        updateApplyBtn();
      } else {
        inspectSelection();
      }
    }
  );

  inspectSelection();
  loadThemeColors();
});

// ─────────────────────────────────────────────────────────────────
// SHAPE INSPECTION — Style mode
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
      const firstKey=toKey(items[0].type), firstName=items[0].name||'';
      items.forEach(s=>{
        const k=toKey(s.type); if(k!==firstKey)allSame=false;
        const caps=CAPABILITIES[k]||{};
        if(caps.fill)    merged.fill=true;
        if(caps.border)  merged.border=true;
        if(caps.text)    merged.text=true;
        if(caps.lineEnds)merged.lineEnds=true;
      });
      const meta=allSame?(TYPE_META[firstKey]||{icon:'?',label:cap(firstKey||'Shape')}):{icon:'⊕',label:'Mixed Selection'};
      renderUI({merged,meta,firstName,count:items.length,anySupported:Object.values(merged).some(Boolean),isLine:allSame&&firstKey==='line'});
    });
  } catch(_){renderEmpty();}
}

function toKey(raw){if(!raw)return'unsupported';const s=String(raw);return s.charAt(0).toLowerCase()+s.slice(1);}
function cap(s){return s.charAt(0).toUpperCase()+s.slice(1);}

function renderUI({merged,meta,firstName,count,anySupported,isLine}) {
  hide('empty-state'); hide('unsupported-state'); show('shape-banner');
  el('shape-icon').textContent      = meta.icon;
  el('shape-type-label').textContent= meta.label;
  el('shape-name').textContent      = firstName?`"${firstName}"` :'';
  el('shape-count-badge').textContent=`×${count}`;
  tog('shape-count-badge', count>1);
  tog('section-fill',      merged.fill);
  tog('section-border',    merged.border);
  tog('section-line-ends', merged.lineEnds);
  el('border-heading').textContent=isLine?'Line Style':'Post Int Process Owner';
  if(!anySupported) show('unsupported-state');
}

function renderEmpty() {
  _cachedShapeCount=0;
  show('empty-state'); hide('shape-banner'); hide('unsupported-state');
  ['section-fill','section-border','section-line-ends'].forEach(hide);
  setStatus('');
}

function el(id)       { return document.getElementById(id); }
function show(id)     { el(id)?.classList.remove('hidden'); }
function hide(id)     { el(id)?.classList.add('hidden'); }
function tog(id,vis)  { el(id)?.classList.toggle('hidden',!vis); }
function setStatus(m) { if(el('status')) el('status').textContent=m; }
