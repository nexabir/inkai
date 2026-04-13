/* ═══════════════════════════════════════════════════════════════
   InkAI — Full Excel/Google Sheets Editor (Complete Rewrite)
   Features: Multi-select, 30+ formulas, bold/italic/color/align,
   undo/redo, copy/paste, column resize, fill handle, sort,
   named cells, dropdowns, autocomplete, number formats,
   freeze header, row/col insert/delete, Ctrl+Z/Y/C/V/X/A/F
═══════════════════════════════════════════════════════════════ */

window.ExcelEditor = (() => {
  'use strict';

  /* ── Constants ── */
  const DEF_ROWS = 100, DEF_COLS = 26;
  const COL_W = 100, ROW_H = 25, HDR_W = 46;
  const MAX_HIST = 60;

  /* ── State ── */
  let _cells = {}, _fmts = {}, _named = {}, _valid = {};
  let _colW = {}, _rowH = {};
  let _rows = DEF_ROWS, _cols = DEF_COLS;

  // Selection
  let _anchor = 'A1', _focus = 'A1';
  let _mouseSelecting = false;
  let _colSelecting = false, _rowSelecting = false;

  // Edit mode
  let _editing = false, _editRef = null;

  // Formula bar cell‑ref insertion
  let _fbarActive = false, _fbarRange = false;
  let _fbarRangeStart = null, _fbarRangeEnd = null;
  let _fbarPreVal = null, _fbarPreCursor = 0;

  // Undo/redo
  let _hist = [], _histIdx = -1;

  // Clipboard
  let _clip = null; // {cells, fmts, range, cut}

  // Column resize
  let _resCol = null, _resX0 = 0, _resW0 = 0;

  // Fill handle
  let _fillActive = false, _fillSrc = null;

  // Find
  let _findOpen = false;

  let _saveCallback = null, _saveTimer = null;
  let _pendingCtxRef = null;

  /* ────────────────────────────────────────
     FORMULA ENGINE  (30+ functions)
  ──────────────────────────────────────── */
  class FE {
    constructor(cells, named) { this.c = cells; this.n = named; this.stk = new Set(); }

    getVal(ref) {
      ref = ref.toUpperCase();
      if (this.stk.has(ref)) return '#CIRCULAR!';
      const cell = this.c[ref];
      if (!cell || (cell.value === '' && !cell.formula)) return 0;
      if (cell.formula && cell.formula.startsWith('=')) {
        this.stk.add(ref);
        try { const r = this.eval(cell.formula); this.stk.delete(ref); return r; }
        catch { this.stk.delete(ref); return '#ERROR!'; }
      }
      const n = parseFloat(cell.value);
      return isNaN(n) ? (cell.value || '') : n;
    }

    eval(f) { return this.expr(f.substring(1).trim()); }

    expr(e) {
      e = (e || '').trim();
      const m = e.match(/^([A-Z]+)\(([\s\S]*)\)$/);
      if (m) return this.fn(m[1], m[2]);
      return this.arith(e);
    }

    arith(e) {
      let s = e.replace(/([A-Z]+\d+)/g, r => {
        const v = this.getVal(r);
        return typeof v === 'number' ? v : (isNaN(parseFloat(v)) ? `"${v}"` : parseFloat(v));
      });
      for (const nm in this.n) {
        s = s.replace(new RegExp(`\\b${nm}\\b`, 'g'), () => {
          const vv = this.range(this.n[nm]);
          return parseFloat(vv[0]) || 0;
        });
      }
      try { return Function('"use strict";return(' + s + ')')(); }
      catch { return '#VALUE!'; }
    }

    range(r) {
      r = (r || '').trim().toUpperCase();
      if (this.n[r]) r = this.n[r].toUpperCase();
      if (!r.includes(':')) return [this.getVal(r)];
      const [s, e2] = r.split(':');
      const sc = cti(s.match(/[A-Z]+/)[0]), sr = parseInt(s.match(/\d+/)[0]);
      const ec = cti(e2.match(/[A-Z]+/)[0]), er = parseInt(e2.match(/\d+/)[0]);
      const vs = [];
      for (let rr = sr; rr <= er; rr++) for (let cc = sc; cc <= ec; cc++) vs.push(this.getVal(itc(cc) + rr));
      return vs;
    }

    args(s) {
      const a = []; let d = 0, cur = '';
      for (const ch of s) {
        if (ch === '(') d++;
        else if (ch === ')') d--;
        if (ch === ',' && d === 0) { a.push(cur.trim()); cur = ''; continue; }
        cur += ch;
      }
      if (cur.trim()) a.push(cur.trim());
      return a;
    }

    resolveArg(a) {
      a = a.trim();
      if (a.includes(':') || this.n[a]) return this.range(a);
      if (/^[A-Z]+\d+$/i.test(a)) return [this.getVal(a.toUpperCase())];
      if (/^"(.*)"$/.test(a)) return [a.slice(1, -1)];
      if (!isNaN(a)) return [parseFloat(a)];
      try { return [this.expr(a)]; } catch { return [a]; }
    }

    nums(args) {
      return args.flatMap(a => this.resolveArg(a)).map(v => parseFloat(v)).filter(n => !isNaN(n));
    }

    cond(s) {
      let e = s.replace(/([A-Z]+\d+)/g, r => { const v = this.getVal(r.toUpperCase()); return typeof v === 'number' ? v : `"${v}"`; });
      try { return !!Function('"use strict";return(' + e + ')')(); } catch { return false; }
    }

    match(val, crit) {
      const s = String(crit);
      if (s.startsWith('>=')) return parseFloat(val) >= parseFloat(s.slice(2));
      if (s.startsWith('<=')) return parseFloat(val) <= parseFloat(s.slice(2));
      if (s.startsWith('<>')) return String(val).toLowerCase() !== s.slice(2).toLowerCase();
      if (s.startsWith('>'))  return parseFloat(val) > parseFloat(s.slice(1));
      if (s.startsWith('<'))  return parseFloat(val) < parseFloat(s.slice(1));
      if (s.includes('*') || s.includes('?')) {
        return new RegExp('^' + s.replace(/\*/g,'.*').replace(/\?/g,'.') + '$', 'i').test(String(val));
      }
      return String(val).toLowerCase() === s.toLowerCase();
    }

    fn(nm, raw) {
      const a = this.args(raw);
      switch (nm.toUpperCase()) {
        /* Math */
        case 'SUM':      return this.nums(a).reduce((x,y)=>x+y,0);
        case 'AVERAGE':  { const n=this.nums(a); return n.length ? n.reduce((x,y)=>x+y,0)/n.length : '#DIV/0!'; }
        case 'COUNT':    return this.nums(a).length;
        case 'COUNTA':   return a.flatMap(x=>this.resolveArg(x)).filter(v=>v!==''&&v!==null&&v!==undefined).length;
        case 'MIN':      { const n=this.nums(a); return n.length ? Math.min(...n) : 0; }
        case 'MAX':      { const n=this.nums(a); return n.length ? Math.max(...n) : 0; }
        case 'SUBTRACT': { const n=this.nums(a); return n.length>=2? n[0]-n[1]:'#VALUE!'; }
        case 'MULTIPLY': { const n=this.nums(a); return n.length>=2? n[0]*n[1]:'#VALUE!'; }
        case 'DIVIDE':   { const n=this.nums(a); return n.length>=2? (n[1]===0?'#DIV/0!':n[0]/n[1]):'#VALUE!'; }
        case 'ROUND':    { const n=this.nums(a); return n.length>=2 ? +n[0].toFixed(n[1]) : (n.length?Math.round(n[0]):0); }
        case 'ROUNDUP':  { const n=this.nums(a); const d=n[1]||0; return n.length>=1?Math.ceil(n[0]*10**d)/10**d:0; }
        case 'ROUNDDOWN':{ const n=this.nums(a); const d=n[1]||0; return n.length>=1?Math.floor(n[0]*10**d)/10**d:0; }
        case 'INT':      { const n=this.nums(a); return n.length?Math.floor(n[0]):0; }
        case 'ABS':      { const n=this.nums(a); return n.length?Math.abs(n[0]):0; }
        case 'SQRT':     { const n=this.nums(a); return n.length?(n[0]<0?'#NUM!':Math.sqrt(n[0])):0; }
        case 'POWER':    { const n=this.nums(a); return n.length>=2?Math.pow(n[0],n[1]):0; }
        case 'MOD':      { const n=this.nums(a); return n.length>=2?(n[1]===0?'#DIV/0!':n[0]%n[1]):0; }
        case 'SIGN':     { const n=this.nums(a); return n.length?Math.sign(n[0]):0; }
        case 'PI':       return Math.PI;
        case 'EXP':      { const n=this.nums(a); return n.length?Math.exp(n[0]):1; }
        case 'LOG':      { const n=this.nums(a); return n.length?((n[1]?Math.log(n[0])/Math.log(n[1]):Math.log10(n[0]))):0; }
        case 'LN':       { const n=this.nums(a); return n.length?Math.log(n[0]):0; }
        case 'RAND':     return Math.random();
        case 'RANDBETWEEN': { const n=this.nums(a); return n.length>=2?Math.floor(Math.random()*(n[1]-n[0]+1))+n[0]:0; }
        case 'LARGE':    { const n=this.nums([a[0]]); n.sort((x,y)=>y-x); const k=(this.nums([a[1]])[0]||1)-1; return n[k]??'#NUM!'; }
        case 'SMALL':    { const n=this.nums([a[0]]); n.sort((x,y)=>x-y); const k=(this.nums([a[1]])[0]||1)-1; return n[k]??'#NUM!'; }
        case 'RANK': {
          const val=this.nums([a[0]])[0]; const rng=this.nums([a[1]]); const ord=(this.nums([a[2]])[0]||0);
          const sorted=[...rng].sort((x,y)=>ord?x-y:y-x);
          return sorted.indexOf(val)+1;
        }
        /* Logic */
        case 'IF': {
          if (a.length<2) return '#VALUE!';
          return this.cond(a[0]) ? (a[1]?this.resolveArg(a[1])[0]:'') : (a[2]?this.resolveArg(a[2])[0]:'');
        }
        case 'IFS': {
          for (let i=0;i<a.length-1;i+=2) if (this.cond(a[i])) return this.resolveArg(a[i+1])[0];
          return '#N/A';
        }
        case 'AND':  return a.every(x=>this.cond(x));
        case 'OR':   return a.some(x=>this.cond(x));
        case 'NOT':  return !this.cond(a[0]);
        case 'XOR':  return a.filter(x=>this.cond(x)).length % 2 === 1;
        case 'IFERROR': { try { const v=this.expr(a[0]); if(typeof v==='string'&&v.startsWith('#'))return a[1]?this.resolveArg(a[1])[0]:''; return v; } catch { return a[1]?this.resolveArg(a[1])[0]:''; } }
        case 'ISBLANK':  { const v=this.resolveArg(a[0])[0]; return v===''||v===null||v===undefined; }
        case 'ISNUMBER': { const v=this.resolveArg(a[0])[0]; return !isNaN(parseFloat(v))&&isFinite(v); }
        case 'ISTEXT':   { const v=this.resolveArg(a[0])[0]; return typeof v==='string'&&isNaN(parseFloat(v)); }
        case 'ISERROR':  { const v=this.resolveArg(a[0])[0]; return typeof v==='string'&&v.startsWith('#'); }
        /* Conditional aggregation */
        case 'SUMIF': {
          const rv=this.range(a[0]); const crit=this.resolveArg(a[1])[0];
          const sv=a[2]?this.range(a[2]):rv;
          return rv.reduce((s,v,i)=>s+(this.match(v,crit)?(parseFloat(sv[i])||0):0),0);
        }
        case 'SUMIFS': {
          const sv=this.range(a[0]);
          let ok=sv.map((_,i)=>i);
          for (let i=1;i<a.length-1;i+=2) {
            const cr=this.range(a[i]); const ct=this.resolveArg(a[i+1])[0];
            ok=ok.filter(idx=>this.match(cr[idx],ct));
          }
          return ok.reduce((s,i)=>s+(parseFloat(sv[i])||0),0);
        }
        case 'COUNTIF':  { const rv=this.range(a[0]); const ct=this.resolveArg(a[1])[0]; return rv.filter(v=>this.match(v,ct)).length; }
        case 'COUNTIFS': {
          const fr=this.range(a[0]); let ok=fr.map((_,i)=>i);
          for (let i=0;i<a.length-1;i+=2) { const cr=this.range(a[i]); const ct=this.resolveArg(a[i+1])[0]; ok=ok.filter(idx=>this.match(cr[idx],ct)); }
          return ok.length;
        }
        /* Text */
        case 'LEN':         { const v=String(this.resolveArg(a[0])[0]||''); return v.length; }
        case 'LEFT':        { const v=String(this.resolveArg(a[0])[0]||''); const n=this.nums([a[1]])[0]||1; return v.substring(0,n); }
        case 'RIGHT':       { const v=String(this.resolveArg(a[0])[0]||''); const n=this.nums([a[1]])[0]||1; return v.slice(-n); }
        case 'MID':         { const v=String(this.resolveArg(a[0])[0]||''); const s=this.nums([a[1]])[0]||1; const n=this.nums([a[2]])[0]||1; return v.substring(s-1,s-1+n); }
        case 'UPPER':       return String(this.resolveArg(a[0])[0]||'').toUpperCase();
        case 'LOWER':       return String(this.resolveArg(a[0])[0]||'').toLowerCase();
        case 'TRIM':        return String(this.resolveArg(a[0])[0]||'').trim();
        case 'PROPER':      { const v=String(this.resolveArg(a[0])[0]||''); return v.replace(/\b\w/g,c=>c.toUpperCase()); }
        case 'REPT':        { const v=String(this.resolveArg(a[0])[0]||''); const n=this.nums([a[1]])[0]||0; return v.repeat(n); }
        case 'SUBSTITUTE':  { const v=String(this.resolveArg(a[0])[0]||''); const f=String(this.resolveArg(a[1])[0]||''); const r=String(this.resolveArg(a[2])[0]||''); return v.split(f).join(r); }
        case 'FIND':        { const find=String(this.resolveArg(a[0])[0]||''); const within=String(this.resolveArg(a[1])[0]||''); const idx=within.indexOf(find); return idx>=0?idx+1:'#VALUE!'; }
        case 'CONCATENATE':
        case 'CONCAT':      return a.flatMap(x=>this.resolveArg(x)).map(v=>String(v??'')).join('');
        case 'TEXTJOIN':    { const dlm=String(this.resolveArg(a[0])[0]||''); const skip=this.cond(a[1]); const vals=a.slice(2).flatMap(x=>this.resolveArg(x)).map(String); return (skip?vals.filter(v=>v!==''):vals).join(dlm); }
        case 'TEXT':        {
          const v=this.resolveArg(a[0])[0]; const fmt=String(this.resolveArg(a[1])[0]||'');
          if (fmt.includes('%')) return (parseFloat(v)*100).toFixed(2)+'%';
          if (fmt.includes('$')) return '$'+parseFloat(v).toFixed(2);
          if (fmt.match(/0\.0+/)) { const d=(fmt.match(/\.0*/)?.[0].length-1)||0; return parseFloat(v).toFixed(d); }
          return String(v??'');
        }
        /* Date/Time */
        case 'TODAY':  return new Date().toLocaleDateString();
        case 'NOW':    return new Date().toLocaleString();
        case 'YEAR':   { const d=new Date(this.resolveArg(a[0])[0]); return isNaN(d)?'#VALUE!':d.getFullYear(); }
        case 'MONTH':  { const d=new Date(this.resolveArg(a[0])[0]); return isNaN(d)?'#VALUE!':d.getMonth()+1; }
        case 'DAY':    { const d=new Date(this.resolveArg(a[0])[0]); return isNaN(d)?'#VALUE!':d.getDate(); }
        case 'HOUR':   { const d=new Date(this.resolveArg(a[0])[0]); return isNaN(d)?'#VALUE!':d.getHours(); }
        case 'MINUTE': { const d=new Date(this.resolveArg(a[0])[0]); return isNaN(d)?'#VALUE!':d.getMinutes(); }
        case 'DAYS':   { const d1=new Date(this.resolveArg(a[0])[0]); const d2=new Date(this.resolveArg(a[1])[0]); return Math.round((d1-d2)/(1000*60*60*24)); }
        /* Lookup */
        case 'VLOOKUP': {
          const lv=this.resolveArg(a[0])[0]; const rng=a[1]; const ci=(this.nums([a[2]])[0]||1)-1;
          const exact=a[3]?this.cond(a[3]):true;
          const [s3,e3]=rng.split(':');
          const sc=cti(s3.match(/[A-Z]+/)[0]); const sr=parseInt(s3.match(/\d+/)[0]);
          const er=parseInt(e3.match(/\d+/)[0]);
          for (let r=sr;r<=er;r++) {
            const tv=this.getVal(itc(sc)+r);
            if (String(tv).toLowerCase()===String(lv).toLowerCase()) return this.getVal(itc(sc+ci)+r);
          }
          return '#N/A';
        }
        default: return '#NAME?';
      }
    }
  }

  /* ── Cell helpers ── */
  function cti(col) { let i=0; for (const ch of col.toUpperCase()) i=i*26+(ch.charCodeAt(0)-64); return i; }
  function itc(idx) { let c=''; while(idx>0){c=String.fromCharCode(64+(((idx-1)%26)+1))+c;idx=Math.floor((idx-1)/26);}return c; }
  function cid(c, r) { return itc(c)+r; }
  function parseRef(id) {
    const m=String(id).match(/^([A-Z]+)(\d+)$/i);
    return m?{col:cti(m[1].toUpperCase()),row:parseInt(m[2])}:null;
  }
  function esc(s) { return String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
  function escAttr(s) { return esc(s).replace(/'/g,'&#39;'); }

  /* ── Evaluate cell ── */
  function evalCell(ref) {
    const cell = _cells[ref];
    if (!cell) return '';
    if (cell.formula && cell.formula.startsWith('=')) {
      try {
        const engine = new FE(_cells, _named);
        const r = engine.eval(cell.formula);
        if (typeof r === 'number') return Number.isInteger(r) ? r : +r.toFixed(10);
        return r;
      } catch { return '#ERROR!'; }
    }
    return cell.value ?? '';
  }

  /* ── Number formatting ── */
  function fmtVal(val, fmt, dec) {
    if (val === '' || val === null || val === undefined) return '';
    if (typeof val === 'string' && val.startsWith('#')) return val;
    dec = dec ?? 2;
    switch ((fmt||'general').toLowerCase()) {
      case 'number':   { const n=parseFloat(val); return isNaN(n)?val:n.toLocaleString(undefined,{minimumFractionDigits:dec,maximumFractionDigits:dec}); }
      case 'currency': { const n=parseFloat(val); return isNaN(n)?val:n.toLocaleString(undefined,{style:'currency',currency:'USD',minimumFractionDigits:dec,maximumFractionDigits:dec}); }
      case 'percent':  { const n=parseFloat(val); return isNaN(n)?val:(n*100).toFixed(dec)+'%'; }
      case 'integer':  { const n=parseFloat(val); return isNaN(n)?val:Math.round(n).toLocaleString(); }
      case 'date':     { const d=new Date(val); return isNaN(d.getTime())?val:d.toLocaleDateString(); }
      case 'text':     return String(val);
      default:         return String(val);
    }
  }

  /* ── Selection ── */
  function getSel() {
    const a=parseRef(_anchor)||{col:1,row:1}, f=parseRef(_focus)||a;
    return { minC:Math.min(a.col,f.col), maxC:Math.max(a.col,f.col), minR:Math.min(a.row,f.row), maxR:Math.max(a.row,f.row) };
  }

  function inSel(ref) {
    const s=getSel(), p=parseRef(ref);
    return p&&p.col>=s.minC&&p.col<=s.maxC&&p.row>=s.minR&&p.row<=s.maxR;
  }

  function setSelection(anchor, focus, skipScroll) {
    _anchor = anchor || 'A1';
    _focus  = focus  || _anchor;
    refreshSelectionUI();
    updateFormulaBar();
    updateNameBox();
    updateFmtToolbar();
    if (!skipScroll) scrollToCell(_anchor);
  }

  function refreshSelectionUI() {
    const s = getSel();
    // Reset all headers
    document.querySelectorAll('.xl-col-hdr').forEach(el => el.classList.remove('hdr-active','hdr-sel'));
    document.querySelectorAll('.xl-row-hdr').forEach(el => el.classList.remove('hdr-active','hdr-sel'));
    // Reset all cells
    document.querySelectorAll('.xl-cell').forEach(el => {
      el.classList.remove('xl-in-sel','xl-active');
    });
    // Highlight selection
    for (let r = s.minR; r <= s.maxR; r++) {
      for (let c = s.minC; c <= s.maxC; c++) {
        const ref = cid(c, r);
        const el = document.querySelector(`.xl-cell[data-ref="${ref}"]`);
        if (el) el.classList.add('xl-in-sel');
      }
      const rh = document.querySelector(`.xl-row-hdr[data-row="${r}"]`);
      if (rh) rh.classList.add('hdr-sel');
    }
    for (let c = s.minC; c <= s.maxC; c++) {
      const ch = document.querySelector(`.xl-col-hdr[data-col="${c}"]`);
      if (ch) ch.classList.add('hdr-sel');
    }
    // Active cell
    const act = document.querySelector(`.xl-cell[data-ref="${_anchor}"]`);
    if (act) act.classList.add('xl-active');
    // Fill handle
    placeFillHandle();
  }

  /* ── Fill handle ── */
  function placeFillHandle() {
    document.querySelector('.xl-fill-handle')?.remove();
    const s = getSel();
    const brRef = cid(s.maxC, s.maxR);
    const brEl  = document.querySelector(`.xl-cell[data-ref="${brRef}"]`);
    if (!brEl) return;
    const handle = document.createElement('div');
    handle.className = 'xl-fill-handle';
    handle.title = 'Fill Handle — drag to copy/fill';
    brEl.style.position = 'relative';
    brEl.appendChild(handle);
    handle.addEventListener('mousedown', startFill);
  }

  function startFill(e) {
    e.preventDefault(); e.stopPropagation();
    _fillActive = true;
    _fillSrc = getSel();
    document.addEventListener('mouseover', duringFill);
    document.addEventListener('mouseup', endFill, { once: true });
    document.body.style.cursor = 'crosshair';
  }

  function duringFill(e) {
    if (!_fillActive) return;
    const el = e.target.closest('.xl-cell[data-ref]');
    if (!el) return;
    const ref = el.dataset.ref;
    const p = parseRef(ref); if (!p) return;
    const s = _fillSrc;
    // Determine direction: down or right
    const downDist  = p.row - s.maxR;
    const rightDist = p.col - s.maxC;
    // Highlight range to fill
    document.querySelectorAll('.xl-filling').forEach(x => x.classList.remove('xl-filling'));
    if (downDist > 0) {
      for (let r = s.maxR+1; r <= p.row; r++)
        for (let c = s.minC; c <= s.maxC; c++)
          document.querySelector(`.xl-cell[data-ref="${cid(c,r)}"]`)?.classList.add('xl-filling');
    } else if (rightDist > 0) {
      for (let c = s.maxC+1; c <= p.col; c++)
        for (let r = s.minR; r <= s.maxR; r++)
          document.querySelector(`.xl-cell[data-ref="${cid(c,r)}"]`)?.classList.add('xl-filling');
    }
  }

  function endFill(e) {
    if (!_fillActive) return;
    _fillActive = false;
    document.removeEventListener('mouseover', duringFill);
    document.body.style.cursor = '';
    // Find destination
    const el = document.elementFromPoint(e.clientX, e.clientY)?.closest?.('.xl-cell[data-ref]');
    document.querySelectorAll('.xl-filling').forEach(x => x.classList.remove('xl-filling'));
    if (!el) return;
    const p = parseRef(el.dataset.ref); if (!p) return;
    const s = _fillSrc;
    const downDist = p.row - s.maxR, rightDist = p.col - s.maxC;
    if (downDist <= 0 && rightDist <= 0) return;
    pushHistory();
    if (downDist > 0) {
      const srcH = s.maxR - s.minR + 1;
      for (let r = s.maxR+1; r <= p.row; r++) {
        for (let c = s.minC; c <= s.maxC; c++) {
          const srcRef = cid(c, s.minR + (r - s.maxR - 1) % srcH);
          const dstRef = cid(c, r);
          _cells[dstRef] = {...(_cells[srcRef] || {})};
          _fmts[dstRef]  = {...(_fmts[srcRef]  || {})};
        }
      }
    } else if (rightDist > 0) {
      const srcW = s.maxC - s.minC + 1;
      for (let c = s.maxC+1; c <= p.col; c++) {
        for (let r = s.minR; r <= s.maxR; r++) {
          const srcRef = cid(s.minC + (c - s.maxC - 1) % srcW, r);
          const dstRef = cid(c, r);
          _cells[dstRef] = {...(_cells[srcRef] || {})};
          _fmts[dstRef]  = {...(_fmts[srcRef]  || {})};
        }
      }
    }
    rebuildGrid();
    scheduleSave();
  }

  /* ── Undo / Redo ── */
  function snap() {
    return JSON.stringify({ cells:_cells, fmts:_fmts, named:_named, valid:_valid, colW:_colW, rowH:_rowH, rows:_rows, cols:_cols });
  }

  function pushHistory() {
    _hist = _hist.slice(0, _histIdx + 1);
    _hist.push(snap());
    if (_hist.length > MAX_HIST) _hist.shift();
    _histIdx = _hist.length - 1;
  }

  function restoreSnap(s) {
    const d = JSON.parse(s);
    _cells=d.cells; _fmts=d.fmts; _named=d.named; _valid=d.valid;
    _colW=d.colW; _rowH=d.rowH; _rows=d.rows; _cols=d.cols;
    rebuildGrid();
  }

  function undo() {
    if (_histIdx <= 0) return;
    _histIdx--;
    restoreSnap(_hist[_histIdx]);
  }

  function redo() {
    if (_histIdx >= _hist.length - 1) return;
    _histIdx++;
    restoreSnap(_hist[_histIdx]);
  }

  /* ── Clipboard ── */
  function copy(cut = false) {
    const s = getSel();
    const cells = {}, fmts = {};
    for (let r = s.minR; r <= s.maxR; r++)
      for (let c = s.minC; c <= s.maxC; c++) {
        const ref = cid(c, r);
        cells[cid(c - s.minC + 1, r - s.minR + 1)] = {...(_cells[ref]||{})};
        fmts [cid(c - s.minC + 1, r - s.minR + 1)] = {...(_fmts[ref] ||{})};
      }
    _clip = { cells, fmts, srcSel: s, h: s.maxR-s.minR+1, w: s.maxC-s.minC+1, cut };
    // System clipboard (plain text)
    let text = '';
    for (let r = s.minR; r <= s.maxR; r++) {
      const row = [];
      for (let c = s.minC; c <= s.maxC; c++) row.push(fmtVal(evalCell(cid(c,r)), _fmts[cid(c,r)]?.numberFormat));
      text += row.join('\t') + '\n';
    }
    try { navigator.clipboard.writeText(text.trim()); } catch {}
    // Flash visual
    document.querySelectorAll('.xl-in-sel').forEach(el => el.classList.add('xl-cut-flash'));
    setTimeout(() => document.querySelectorAll('.xl-cut-flash').forEach(el => el.classList.remove('xl-cut-flash')), 600);
  }

  async function paste() {
    const dst = parseRef(_anchor); if (!dst) return;
    let pasted = false;

    // Try system clipboard first (handles paste from Excel/Sheets)
    try {
      const txt = await navigator.clipboard.readText();
      if (txt && txt.trim()) {
        pushHistory();
        const rows2 = txt.split('\n').filter(r=>r.trim()!='');
        rows2.forEach((rowStr, ri) => {
          rowStr.split('\t').forEach((val, ci) => {
            const ref = cid(dst.col + ci, dst.row + ri);
            if(!_cells[ref]) _cells[ref] = {};
            const v = val.trim();
            if (v.startsWith('=')) { _cells[ref].formula = v; _cells[ref].value = ''; }
            else { _cells[ref].value = v; _cells[ref].formula = ''; }
          });
        });
        pasted = true;
      }
    } catch {}

    // Internal clipboard (preserves formats)
    if (!pasted && _clip) {
      pushHistory();
      const { cells, fmts, h, w, srcSel, cut } = _clip;
      for (let r = 1; r <= h; r++) for (let c = 1; c <= w; c++) {
        const srcKey = cid(c, r);
        const dstRef = cid(dst.col + c - 1, dst.row + r - 1);
        if (cells[srcKey]) _cells[dstRef] = {...cells[srcKey]};
        if (fmts[srcKey])  _fmts[dstRef]  = {...fmts[srcKey]};
      }
      if (cut && srcSel) {
        for (let r=srcSel.minR;r<=srcSel.maxR;r++)
          for (let c=srcSel.minC;c<=srcSel.maxC;c++) { delete _cells[cid(c,r)]; delete _fmts[cid(c,r)]; }
        _clip = null;
      }
      pasted = true;
    }

    if (pasted) { rebuildGrid(); scheduleSave(); }
  }

  /* ── Column resize ── */
  function startColResize(e) {
    e.preventDefault(); e.stopPropagation();
    _resCol = parseInt(e.target.dataset.col);
    _resX0  = e.clientX;
    _resW0  = _colW[_resCol] || COL_W;
    document.body.style.cursor = 'col-resize';
    document.addEventListener('mousemove', duringColResize);
    document.addEventListener('mouseup', endColResize, { once: true });
  }

  function duringColResize(e) {
    if (!_resCol) return;
    const nw = Math.max(30, _resW0 + (e.clientX - _resX0));
    _colW[_resCol] = nw;
    // Live update
    document.querySelector(`.xl-col-hdr[data-col="${_resCol}"]`)?.style && (document.querySelector(`.xl-col-hdr[data-col="${_resCol}"]`).style.width = nw + 'px');
    document.querySelectorAll(`.xl-cell[data-ref^="${itc(_resCol)}"]`).forEach(el => el.style.width = nw + 'px');
  }

  function endColResize() {
    _resCol = null;
    document.removeEventListener('mousemove', duringColResize);
    document.body.style.cursor = '';
    scheduleSave();
  }

  /* ── Sort ── */
  function sortCol(asc) {
    const s = getSel();
    const col = s.minC;
    pushHistory();
    // Collect all rows with data, sort by this col
    const rowNums = [];
    for (let r = 1; r <= _rows; r++) rowNums.push(r);
    rowNums.sort((a, b) => {
      const va = evalCell(cid(col, a)), vb = evalCell(cid(col, b));
      const na = parseFloat(va), nb = parseFloat(vb);
      const isNA = !isNaN(na), isNB = !isNaN(nb);
      if (isNA && isNB) return asc ? na-nb : nb-na;
      return asc ? String(va).localeCompare(String(vb)) : String(vb).localeCompare(String(va));
    });
    const newCells = {}, newFmts = {};
    rowNums.forEach((origRow, newRowIdx) => {
      const nr = newRowIdx + 1;
      for (let c = 1; c <= _cols; c++) {
        const src = cid(c, origRow), dst = cid(c, nr);
        if (_cells[src]) newCells[dst] = {..._cells[src]};
        if (_fmts[src])  newFmts[dst]  = {..._fmts[src]};
      }
    });
    _cells = newCells; _fmts = newFmts;
    rebuildGrid(); scheduleSave();
  }

  /* ── Row / Col insert / delete ── */
  function insertRow() {
    const s = getSel(); pushHistory();
    const nr = {};
    for (const ref in _cells) {
      const p = parseRef(ref); if (!p) continue;
      nr[cid(p.col, p.row >= s.minR ? p.row + 1 : p.row)] = _cells[ref];
    }
    const nf = {};
    for (const ref in _fmts) {
      const p = parseRef(ref); if (!p) continue;
      nf[cid(p.col, p.row >= s.minR ? p.row + 1 : p.row)] = _fmts[ref];
    }
    _cells = nr; _fmts = nf; _rows++;
    rebuildGrid(); scheduleSave();
  }

  function deleteRow() {
    const s = getSel(); pushHistory();
    const nr = {}, nf = {};
    for (const ref in _cells) {
      const p = parseRef(ref); if (!p) continue;
      if (p.row >= s.minR && p.row <= s.maxR) continue;
      nr[cid(p.col, p.row > s.maxR ? p.row - (s.maxR - s.minR + 1) : p.row)] = _cells[ref];
    }
    for (const ref in _fmts) {
      const p = parseRef(ref); if (!p) continue;
      if (p.row >= s.minR && p.row <= s.maxR) continue;
      nf[cid(p.col, p.row > s.maxR ? p.row - (s.maxR - s.minR + 1) : p.row)] = _fmts[ref];
    }
    _cells = nr; _fmts = nf; _rows = Math.max(5, _rows - (s.maxR - s.minR + 1));
    rebuildGrid(); scheduleSave();
  }

  function insertCol() {
    const s = getSel(); pushHistory();
    const nr = {}, nf = {};
    for (const ref in _cells) {
      const p = parseRef(ref); if (!p) continue;
      nr[cid(p.col >= s.minC ? p.col + 1 : p.col, p.row)] = _cells[ref];
    }
    for (const ref in _fmts) {
      const p = parseRef(ref); if (!p) continue;
      nf[cid(p.col >= s.minC ? p.col + 1 : p.col, p.row)] = _fmts[ref];
    }
    _cells = nr; _fmts = nf; _cols = Math.min(702, _cols + 1);
    rebuildGrid(); scheduleSave();
  }

  function deleteCol() {
    const s = getSel(); pushHistory();
    const nr = {}, nf = {};
    for (const ref in _cells) {
      const p = parseRef(ref); if (!p) continue;
      if (p.col >= s.minC && p.col <= s.maxC) continue;
      nr[cid(p.col > s.maxC ? p.col - (s.maxC - s.minC + 1) : p.col, p.row)] = _cells[ref];
    }
    for (const ref in _fmts) {
      const p = parseRef(ref); if (!p) continue;
      if (p.col >= s.minC && p.col <= s.maxC) continue;
      nf[cid(p.col > s.maxC ? p.col - (s.maxC - s.minC + 1) : p.col, p.row)] = _fmts[ref];
    }
    _cells = nr; _fmts = nf; _cols = Math.max(3, _cols - (s.maxC - s.minC + 1));
    rebuildGrid(); scheduleSave();
  }

  /* ── Cell Editing ── */
  function startEdit(ref, char) {
    if (_editing) commitEdit();
    _editing = true; _editRef = ref;
    const el = document.querySelector(`.xl-cell[data-ref="${ref}"]`);
    if (!el) return;
    const inp = el.querySelector('.xl-ci');
    if (!inp) return;
    inp.readOnly = false;
    const cell = _cells[ref] || {};
    if (char && char !== 'F2') {
      inp.value = char === 'Delete' ? '' : (cell.formula ? cell.formula : (char || ''));
    } else {
      inp.value = cell.formula || cell.value || '';
    }
    inp.focus();
    inp.select();
    el.classList.add('xl-editing');
    // Sync formula bar
    const fb = document.getElementById('formula-bar');
    if (fb) fb.value = inp.value;
  }

  function commitEdit() {
    if (!_editing || !_editRef) return;
    _editing = false;
    const ref = _editRef; _editRef = null;
    const el = document.querySelector(`.xl-cell[data-ref="${ref}"]`);
    if (!el) return;
    const inp = el.querySelector('.xl-ci');
    if (!inp) return;
    inp.readOnly = true;
    el.classList.remove('xl-editing');
    const val = inp.value.trim();
    pushHistory();
    if (!_cells[ref]) _cells[ref] = {};
    if (val.startsWith('=')) { _cells[ref].formula = val; _cells[ref].value = ''; }
    else { _cells[ref].value = val; _cells[ref].formula = ''; }
    refreshCell(ref);
    scheduleSave();
  }

  function cancelEdit() {
    if (!_editing || !_editRef) return;
    _editing = false;
    const ref = _editRef; _editRef = null;
    const el = document.querySelector(`.xl-cell[data-ref="${ref}"]`);
    if (!el) return;
    el.classList.remove('xl-editing');
    const inp = el.querySelector('.xl-ci');
    if (inp) { inp.readOnly = true; inp.value = fmtVal(evalCell(ref), _fmts[ref]?.numberFormat, _fmts[ref]?.decimals); }
  }

  function refreshCell(ref) {
    const el = document.querySelector(`.xl-cell[data-ref="${ref}"]`);
    if (!el) return;
    const inp = el.querySelector('.xl-ci');
    if (!inp) return;
    const val = evalCell(ref);
    const fmt = _fmts[ref] || {};
    inp.value = fmtVal(val, fmt.numberFormat, fmt.decimals);
    inp.readOnly = true;
    applyFmtToInp(inp, fmt);
    el.style.background = fmt.fillColor || '';
    el.classList.toggle('is-error', typeof val === 'string' && val.startsWith('#'));
    el.classList.toggle('is-formula', !!(_cells[ref]?.formula?.startsWith('=')));
  }

  function applyFmtToInp(inp, fmt) {
    inp.style.fontWeight      = fmt.bold ? 'bold' : '';
    inp.style.fontStyle       = fmt.italic ? 'italic' : '';
    inp.style.textDecoration  = [fmt.underline?'underline':'', fmt.strikethrough?'line-through':''].filter(Boolean).join(' ') || '';
    inp.style.fontFamily      = fmt.fontFamily || '';
    inp.style.fontSize        = fmt.fontSize ? fmt.fontSize + 'px' : '';
    inp.style.color           = fmt.textColor || '';
    inp.style.textAlign       = fmt.align || '';
    inp.style.whiteSpace      = fmt.wrapText ? 'pre-wrap' : '';
  }

  /* ── Formatting ── */
  function applyFmt(key, value) {
    const s = getSel();
    pushHistory();
    for (let r = s.minR; r <= s.maxR; r++) for (let c = s.minC; c <= s.maxC; c++) {
      const ref = cid(c, r);
      if (!_fmts[ref]) _fmts[ref] = {};
      if (key === 'toggle_bold')        { _fmts[ref].bold = !_fmts[ref].bold; }
      else if (key === 'toggle_italic') { _fmts[ref].italic = !_fmts[ref].italic; }
      else if (key === 'toggle_under')  { _fmts[ref].underline = !_fmts[ref].underline; }
      else if (key === 'toggle_strike') { _fmts[ref].strikethrough = !_fmts[ref].strikethrough; }
      else if (key === 'toggle_wrap')   { _fmts[ref].wrapText = !_fmts[ref].wrapText; }
      else _fmts[ref][key] = value;
      refreshCell(ref);
    }
    updateFmtToolbar();
    scheduleSave();
  }

  function updateFmtToolbar() {
    const fmt = _fmts[_anchor] || {};
    document.getElementById('xl-bold')?.classList.toggle('active', !!fmt.bold);
    document.getElementById('xl-italic')?.classList.toggle('active', !!fmt.italic);
    document.getElementById('xl-underline')?.classList.toggle('active', !!fmt.underline);
    document.getElementById('xl-strike')?.classList.toggle('active', !!fmt.strikethrough);
    document.getElementById('xl-wrap')?.classList.toggle('active', !!fmt.wrapText);
    document.getElementById('xl-al-left')?.classList.toggle('active', fmt.align === 'left');
    document.getElementById('xl-al-center')?.classList.toggle('active', fmt.align === 'center');
    document.getElementById('xl-al-right')?.classList.toggle('active', fmt.align === 'right');
    if (document.getElementById('xl-font')) document.getElementById('xl-font').value = fmt.fontFamily || '';
    if (document.getElementById('xl-fsize')) document.getElementById('xl-fsize').value = fmt.fontSize || 11;
    if (document.getElementById('xl-numfmt')) document.getElementById('xl-numfmt').value = fmt.numberFormat || 'general';
    if (document.getElementById('xl-txt-color')) document.getElementById('xl-txt-color').value = fmt.textColor || (document.documentElement.dataset.theme === 'dark' ? '#c0caf5' : '#1a1a1a');
    if (document.getElementById('xl-fill-color')) document.getElementById('xl-fill-color').value = fmt.fillColor || (document.documentElement.dataset.theme === 'dark' ? '#292e42' : '#ffffff');
  }

  /* ── Formula bar ── */
  function updateFormulaBar() {
    const fb = document.getElementById('formula-bar');
    if (!fb) return;
    const cell = _cells[_anchor] || {};
    fb.value = cell.formula || cell.value || '';
  }

  function updateNameBox() {
    const nb = document.getElementById('name-box');
    if (!nb) return;
    const s = getSel();
    if (_anchor === _focus) {
      let name = _anchor;
      for (const nm in _named) if (_named[nm].toUpperCase() === _anchor) { name = nm; break; }
      nb.value = name;
    } else {
      nb.value = `${cid(s.minC,s.minR)}:${cid(s.maxC,s.maxR)}`;
    }
  }

  /* ── Grid rebuild ── */
  function rebuildGrid() {
    const inner = document.getElementById('xl-grid-inner');
    if (!inner) { buildGrid(); return; }
    buildGrid();
  }

  function buildGrid() {
    const wrap = document.getElementById('excel-grid-wrap');
    if (!wrap) return;

    let html = '<div class="xl-grid" id="xl-grid-inner">';

    // Header row
    html += '<div class="xl-hdr-row">';
    html += `<div class="xl-corner" id="xl-corner" title="Select All (Ctrl+A)"></div>`;
    for (let c = 1; c <= _cols; c++) {
      const w = _colW[c] || COL_W;
      html += `<div class="xl-col-hdr" data-col="${c}" style="width:${w}px;min-width:${w}px">${itc(c)}<div class="xl-col-resizer" data-col="${c}"></div></div>`;
    }
    html += '</div>';

    // Data rows
    for (let r = 1; r <= _rows; r++) {
      const rh = _rowH[r] || ROW_H;
      html += `<div class="xl-data-row" data-row="${r}">`;
      html += `<div class="xl-row-hdr" data-row="${r}" style="height:${rh}px">${r}</div>`;
      for (let c = 1; c <= _cols; c++) {
        const ref = cid(c, r);
        const w = _colW[c] || COL_W;
        const cell = _cells[ref] || {};
        const fmt  = _fmts[ref]  || {};
        const val  = evalCell(ref);
        const disp = fmtVal(val, fmt.numberFormat, fmt.decimals);
        const bgStyle = fmt.fillColor ? `background:${fmt.fillColor};` : '';
        const isFormula = !!(cell.formula?.startsWith('='));
        const isError   = typeof val === 'string' && val.startsWith('#');
        const hasDrop   = !!_valid[ref];
        let cls = 'xl-cell';
        if (isFormula) cls += ' is-formula';
        if (isError)   cls += ' is-error';
        if (hasDrop)   cls += ' has-dropdown';
        // Input style
        const inpSt = buildInpStyle(fmt);
        html += `<div class="${cls}" data-ref="${ref}" style="width:${w}px;height:${rh}px;${bgStyle}">` +
          `<input class="xl-ci" value="${escAttr(disp)}" readonly tabindex="-1" data-ref="${ref}" style="${inpSt}" />` +
          (hasDrop ? `<span class="xl-dd-arr" data-ref="${ref}">▼</span>` : '') +
          `</div>`;
      }
      html += '</div>';
    }
    html += '</div>'; // .xl-grid

    wrap.innerHTML = html;

    // Bind events
    const gi = document.getElementById('xl-grid-inner');
    gi.addEventListener('mousedown', onGridMousedown);
    gi.addEventListener('mouseover', onGridMouseover);
    gi.addEventListener('dblclick',  onGridDblclick);
    gi.addEventListener('contextmenu', onContextMenu);
    document.getElementById('xl-corner')?.addEventListener('click', selectAll);
    document.querySelectorAll('.xl-col-resizer').forEach(el => el.addEventListener('mousedown', startColResize));
    document.querySelectorAll('.xl-col-hdr').forEach(el => el.addEventListener('mousedown', onColHdrDn));
    document.querySelectorAll('.xl-row-hdr').forEach(el => el.addEventListener('mousedown', onRowHdrDn));
    document.querySelectorAll('.xl-dd-arr').forEach(el => el.addEventListener('mousedown', e => { e.stopPropagation(); openDropdown(el.dataset.ref); }));

    refreshSelectionUI();
  }

  function buildInpStyle(fmt) {
    const parts = [];
    if (fmt.bold)          parts.push('font-weight:bold');
    if (fmt.italic)        parts.push('font-style:italic');
    const td = [fmt.underline?'underline':'', fmt.strikethrough?'line-through':''].filter(Boolean).join(' ');
    if (td) parts.push('text-decoration:'+td);
    if (fmt.fontFamily)    parts.push('font-family:'+fmt.fontFamily);
    if (fmt.fontSize)      parts.push('font-size:'+fmt.fontSize+'px');
    if (fmt.textColor)     parts.push('color:'+fmt.textColor);
    if (fmt.align)         parts.push('text-align:'+fmt.align);
    if (fmt.wrapText)      parts.push('white-space:pre-wrap');
    return parts.join(';');
  }

  /* ── Grid mouse events ── */
  function onGridMousedown(e) {
    if (e.target.classList.contains('xl-col-resizer') ||
        e.target.classList.contains('xl-fill-handle') ||
        e.target.classList.contains('xl-dd-arr')) return;

    const cellEl = e.target.closest('.xl-cell[data-ref]');
    if (!cellEl) return;
    const ref = cellEl.dataset.ref;

    // Formula bar range-selection mode
    if (_fbarActive) {
      e.preventDefault();
      if (!_fbarRange) {
        _fbarRange = true;
        _fbarRangeStart = ref;
        _fbarRangeEnd   = ref;
        _fbarPreVal     = document.getElementById('formula-bar').value;
        _fbarPreCursor  = document.getElementById('formula-bar').selectionStart;
      }
      _fbarRangeEnd = ref;
      applyRangeToFbar();
      highlightFbarRange(_fbarRangeStart, _fbarRangeEnd);
      return;
    }

    if (_editing && _editRef !== ref) commitEdit();

    if (e.shiftKey) {
      _focus = ref;
      setSelection(_anchor, _focus);
    } else {
      _anchor = ref; _focus = ref;
      _mouseSelecting = true;
      setSelection(_anchor, _focus);
    }
  }

  function onGridMouseover(e) {
    if (_fbarRange) {
      const cellEl = e.target.closest('.xl-cell[data-ref]');
      if (!cellEl) return;
      _fbarRangeEnd = cellEl.dataset.ref;
      applyRangeToFbar();
      highlightFbarRange(_fbarRangeStart, _fbarRangeEnd);
      return;
    }
    if (_mouseSelecting) {
      const cellEl = e.target.closest('.xl-cell[data-ref]');
      if (cellEl) { _focus = cellEl.dataset.ref; refreshSelectionUI(); }
    }
  }

  function onGridDblclick(e) {
    const cellEl = e.target.closest('.xl-cell[data-ref]');
    if (cellEl) startEdit(cellEl.dataset.ref, 'F2');
  }

  function onColHdrDn(e) {
    if (e.target.classList.contains('xl-col-resizer')) return;
    const c = parseInt(e.currentTarget.dataset.col);
    _anchor = cid(c, 1); _focus = cid(c, _rows);
    _colSelecting = true;
    setSelection(_anchor, _focus);
  }

  function onRowHdrDn(e) {
    const r = parseInt(e.currentTarget.dataset.row);
    _anchor = cid(1, r); _focus = cid(_cols, r);
    setSelection(_anchor, _focus);
  }

  function selectAll() {
    _anchor = cid(1, 1); _focus = cid(_cols, _rows);
    setSelection(_anchor, _focus);
  }

  /* ── Context menu ── */
  function onContextMenu(e) {
    e.preventDefault();
    const cellEl = e.target.closest('.xl-cell[data-ref]');
    if (cellEl) { _pendingCtxRef = cellEl.dataset.ref; if (!inSel(cellEl.dataset.ref)) setSelection(cellEl.dataset.ref); }
    const menu = document.getElementById('context-menu');
    if (!menu) return;
    menu.style.top  = Math.min(e.clientY, window.innerHeight-220)+'px';
    menu.style.left = Math.min(e.clientX, window.innerWidth-200)+'px';
    menu.classList.remove('hidden');
  }

  /* ── Dropdown ── */
  function openDropdown(ref) {
    document.querySelector('.xl-dd-list')?.remove();
    const opts = _valid[ref]; if (!opts?.length) return;
    const el = document.querySelector(`.xl-cell[data-ref="${ref}"]`);
    if (!el) return;
    const rect = el.getBoundingClientRect();
    const list = document.createElement('div');
    list.className = 'xl-dd-list';
    list.style.cssText = `top:${rect.bottom}px;left:${rect.left}px;min-width:${rect.width}px`;
    opts.forEach(opt => {
      const btn = document.createElement('button');
      btn.textContent = opt;
      btn.addEventListener('click', () => {
        pushHistory();
        if (!_cells[ref]) _cells[ref] = {};
        _cells[ref].value = opt; _cells[ref].formula = '';
        refreshCell(ref); scheduleSave(); list.remove();
      });
      list.appendChild(btn);
    });
    document.body.appendChild(list);
    setTimeout(() => document.addEventListener('click', () => list.remove(), { once: true }), 50);
  }

  /* ── Scroll to cell ── */
  function scrollToCell(ref) {
    const el = document.querySelector(`.xl-cell[data-ref="${ref}"]`);
    if (el) el.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }

  /* ── Formula bar range insertion ── */
  function buildFbarRange(start, end) {
    if (!start) return '';
    if (!end || start === end) return start;
    const a = parseRef(start), b = parseRef(end); if (!a||!b) return start;
    const mc=Math.min(a.col,b.col), xc=Math.max(a.col,b.col);
    const mr=Math.min(a.row,b.row), xr=Math.max(a.row,b.row);
    const tl=cid(mc,mr), br=cid(xc,xr);
    return tl===br?tl:tl+':'+br;
  }

  function applyRangeToFbar() {
    const fb = document.getElementById('formula-bar'); if (!fb) return;
    const rng = buildFbarRange(_fbarRangeStart, _fbarRangeEnd);
    fb.value = _fbarPreVal.substring(0, _fbarPreCursor) + rng + _fbarPreVal.substring(_fbarPreCursor);
    const nb = document.getElementById('name-box'); if (nb) nb.value = rng;
  }

  function highlightFbarRange(start, end) {
    document.querySelectorAll('.xl-cell.formula-selecting').forEach(el => el.classList.remove('formula-selecting'));
    const a=parseRef(start), b=end?parseRef(end):a; if (!a||!b) return;
    const mc=Math.min(a.col,b.col), xc=Math.max(a.col,b.col);
    const mr=Math.min(a.row,b.row), xr=Math.max(a.row,b.row);
    for (let r=mr;r<=xr;r++) for (let c=mc;c<=xc;c++) document.querySelector(`.xl-cell[data-ref="${cid(c,r)}"]`)?.classList.add('formula-selecting');
  }

  /* ── Formula autocomplete ── */
  const FN_LIST = ['SUM','AVERAGE','COUNT','COUNTA','MIN','MAX','IF','IFS','SUMIF','SUMIFS','COUNTIF','COUNTIFS',
    'SUBTRACT','MULTIPLY','DIVIDE','ROUND','ROUNDUP','ROUNDDOWN','INT','ABS','SQRT','POWER','MOD','SIGN','PI',
    'LEN','LEFT','RIGHT','MID','UPPER','LOWER','TRIM','PROPER','CONCATENATE','CONCAT','TEXTJOIN','SUBSTITUTE',
    'FIND','TEXT','REPT','TODAY','NOW','YEAR','MONTH','DAY','HOUR','MINUTE','DAYS',
    'IFERROR','ISBLANK','ISNUMBER','ISTEXT','ISERROR','AND','OR','NOT','XOR',
    'VLOOKUP','LARGE','SMALL','RANK','EXP','LOG','LN','RAND','RANDBETWEEN'];

  function showAutocomplete(input, val) {
    const ac = document.getElementById('xl-autocomplete'); if (!ac) return;
    const m = val.match(/=([A-Z]*)$/i);
    if (!m || !m[1]) { ac.classList.add('hidden'); return; }
    const prefix = m[1].toUpperCase();
    const matches = FN_LIST.filter(f => f.startsWith(prefix));
    if (!matches.length) { ac.classList.add('hidden'); return; }
    ac.innerHTML = '';
    matches.slice(0, 8).forEach(fn => {
      const item = document.createElement('div');
      item.className = 'xl-ac-item';
      item.textContent = fn;
      item.addEventListener('mousedown', e => {
        e.preventDefault();
        const base = val.replace(/=[A-Z]*$/i, '');
        document.getElementById('formula-bar').value = base + '=' + fn + '(';
        ac.classList.add('hidden');
        document.getElementById('formula-bar').focus();
      });
      ac.appendChild(item);
    });
    const fb = document.getElementById('formula-bar');
    const rect = fb?.getBoundingClientRect();
    if (rect) { ac.style.left = rect.left + 'px'; ac.style.top = (rect.bottom + 2) + 'px'; }
    ac.classList.remove('hidden');
  }

  /* ── Named ranges ── */
  function renderNamedList() {
    const list = document.getElementById('named-ranges-list'); if (!list) return;
    list.innerHTML = '<div class="named-list-title">Existing Names:</div>';
    Object.entries(_named).forEach(([name, range]) => {
      const tag = document.createElement('div');
      tag.className = 'named-range-tag';
      tag.innerHTML = `<strong>${esc(name)}</strong>: ${esc(range)} <button data-name="${esc(name)}" title="Delete">✕</button>`;
      tag.querySelector('button').addEventListener('click', () => { delete _named[name]; scheduleSave(); renderNamedList(); rebuildGrid(); });
      list.appendChild(tag);
    });
  }

  /* ── Find & Replace ── */
  function openFind() {
    let dialog = document.getElementById('xl-find-dialog');
    if (!dialog) {
      dialog = document.createElement('div');
      dialog.id = 'xl-find-dialog';
      dialog.className = 'xl-find-dialog';
      dialog.innerHTML = `
        <div class="xl-find-header"><strong>Find & Replace</strong><button id="xl-find-close">✕</button></div>
        <div class="xl-find-body">
          <label>Find: <input id="xl-find-inp" class="form-input" placeholder="Search…" /></label>
          <label>Replace: <input id="xl-rep-inp" class="form-input" placeholder="Replace with…" /></label>
          <div class="xl-find-actions">
            <button class="btn btn-ghost" id="xl-find-next">Find Next</button>
            <button class="btn btn-primary" id="xl-rep-all">Replace All</button>
          </div>
          <div id="xl-find-status" style="font-size:0.75rem;color:var(--txt3);margin-top:6px"></div>
        </div>`;
      document.body.appendChild(dialog);
      document.getElementById('xl-find-close').addEventListener('click',  () => dialog.remove());
      document.getElementById('xl-find-next').addEventListener('click', findNext);
      document.getElementById('xl-rep-all').addEventListener('click', replaceAll);
    }
    document.getElementById('xl-find-inp').focus();
  }

  let _findPos = null;
  function findNext() {
    const q = document.getElementById('xl-find-inp')?.value?.toLowerCase(); if (!q) return;
    const allRefs = [];
    for (let r=1;r<=_rows;r++) for (let c=1;c<=_cols;c++) allRefs.push(cid(c,r));
    if (!_findPos) _findPos = 0;
    let found = false;
    for (let i=_findPos; i<allRefs.length; i++) {
      const v = String(evalCell(allRefs[i])||'').toLowerCase();
      if (v.includes(q)) { setSelection(allRefs[i]); _findPos = i+1; found=true; break; }
    }
    if (!found) { _findPos = 0; document.getElementById('xl-find-status').textContent = 'Not found / wrapped to start'; }
  }

  function replaceAll() {
    const q = document.getElementById('xl-find-inp')?.value; if (!q) return;
    const rep = document.getElementById('xl-rep-inp')?.value || '';
    pushHistory();
    let count = 0;
    for (const ref in _cells) {
      const v = _cells[ref].value || '';
      if (v.includes(q)) { _cells[ref].value = v.split(q).join(rep); count++; refreshCell(ref); }
    }
    document.getElementById('xl-find-status').textContent = `Replaced ${count} occurrence(s)`;
    scheduleSave();
  }

  /* ── Keyboard ── */
  function handleKeydown(e) {
    const tag = document.activeElement?.tagName;
    const inFbar = document.activeElement?.id === 'formula-bar';
    const inCell = document.activeElement?.classList.contains('xl-ci');

    if (inFbar) {
      // Escape formula bar
      if (e.key === 'Escape') {
        document.querySelectorAll('.xl-cell.formula-selecting').forEach(el => el.classList.remove('formula-selecting'));
        _fbarActive = false; _fbarRange = false;
        updateFormulaBar();
        document.getElementById('formula-bar').blur();
        return;
      }
      // Enter commits formula bar
      if (e.key === 'Enter' && _anchor) {
        const val = document.getElementById('formula-bar').value.trim();
        pushHistory();
        if (!_cells[_anchor]) _cells[_anchor] = {};
        if (val.startsWith('=')) { _cells[_anchor].formula = val; _cells[_anchor].value = ''; }
        else { _cells[_anchor].value = val; _cells[_anchor].formula = ''; }
        refreshCell(_anchor);
        refreshAll();
        document.getElementById('formula-bar').blur();
        _fbarActive = false; _fbarRange = false;
        const p = parseRef(_anchor);
        if (p) setSelection(cid(p.col, p.row+1));
        scheduleSave();
        e.preventDefault();
      }
      return;
    }

    if (inCell && _editing) {
      if (e.key === 'Escape')  { cancelEdit(); document.getElementById('excel-grid-wrap')?.focus(); return; }
      if (e.key === 'Enter')   { e.preventDefault(); commitEdit(); const p=parseRef(_editRef||_anchor); if(p) setSelection(cid(p.col,p.row+1)); return; }
      if (e.key === 'Tab')     { e.preventDefault(); commitEdit(); const p=parseRef(_editRef||_anchor); if(p) setSelection(cid(p.col+1,p.row)); return; }
      // Sync formula bar while typing
      setTimeout(() => { const fb=document.getElementById('formula-bar'); const inp=document.activeElement; if(fb&&inp?.classList.contains('xl-ci')) fb.value=inp.value; showAutocomplete(inp, inp.value); }, 0);
      return;
    }

    // Grid-level shortcuts
    const p = parseRef(_anchor);
    if (!p) return;

    // Navigation
    if (e.key === 'ArrowUp')    { e.preventDefault(); const nr=cid(p.col,Math.max(1,p.row-1)); setSelection(e.shiftKey?_anchor:nr, e.shiftKey?nr:nr); if(!e.shiftKey) _anchor=nr; _focus=nr; refreshSelectionUI(); updateFormulaBar(); scrollToCell(_focus); return; }
    if (e.key === 'ArrowDown')  { e.preventDefault(); const nr=cid(p.col,Math.min(_rows,p.row+1)); _anchor=e.shiftKey?_anchor:nr; _focus=nr; refreshSelectionUI(); updateFormulaBar(); scrollToCell(_focus); return; }
    if (e.key === 'ArrowLeft')  { e.preventDefault(); const nr=cid(Math.max(1,p.col-1),p.row); _anchor=e.shiftKey?_anchor:nr; _focus=nr; refreshSelectionUI(); updateFormulaBar(); scrollToCell(_focus); return; }
    if (e.key === 'ArrowRight') { e.preventDefault(); const nr=cid(Math.min(_cols,p.col+1),p.row); _anchor=e.shiftKey?_anchor:nr; _focus=nr; refreshSelectionUI(); updateFormulaBar(); scrollToCell(_focus); return; }
    if (e.key === 'Tab')        { e.preventDefault(); const nr=cid(Math.min(_cols,p.col+1),p.row); setSelection(nr); return; }
    if (e.key === 'Enter')      { e.preventDefault(); const nr=cid(p.col,Math.min(_rows,p.row+1)); setSelection(nr); return; }
    if (e.key === 'Home')       { e.preventDefault(); setSelection(e.ctrlKey?cid(1,1):cid(1,p.row)); return; }
    if (e.key === 'End')        { e.preventDefault(); setSelection(e.ctrlKey?cid(_cols,_rows):cid(_cols,p.row)); return; }
    if (e.key === 'PageDown')   { e.preventDefault(); setSelection(cid(p.col,Math.min(_rows,p.row+20))); return; }
    if (e.key === 'PageUp')     { e.preventDefault(); setSelection(cid(p.col,Math.max(1,p.row-20))); return; }

    // Delete/Backspace
    if (e.key === 'Delete' || e.key === 'Backspace') {
      const s=getSel(); pushHistory();
      for (let r=s.minR;r<=s.maxR;r++) for (let c=s.minC;c<=s.maxC;c++) { delete _cells[cid(c,r)]; refreshCell(cid(c,r)); }
      updateFormulaBar(); scheduleSave(); return;
    }

    // F2 = edit
    if (e.key === 'F2') { startEdit(_anchor, 'F2'); return; }

    // Ctrl shortcuts
    if (e.ctrlKey || e.metaKey) {
      if (e.key === 'z') { e.preventDefault(); undo(); return; }
      if (e.key === 'y') { e.preventDefault(); redo(); return; }
      if (e.key === 'c') { e.preventDefault(); copy(false); return; }
      if (e.key === 'x') { e.preventDefault(); copy(true); return; }
      if (e.key === 'v') { e.preventDefault(); paste(); return; }
      if (e.key === 'a') { e.preventDefault(); selectAll(); return; }
      if (e.key === 'f') { e.preventDefault(); openFind(); return; }
      if (e.key === 'b') { e.preventDefault(); applyFmt('toggle_bold'); return; }
      if (e.key === 'i') { e.preventDefault(); applyFmt('toggle_italic'); return; }
      if (e.key === 'u') { e.preventDefault(); applyFmt('toggle_under'); return; }
      if (e.key === 'Home') { e.preventDefault(); setSelection('A1'); return; }
      if (e.key === 'End')  { e.preventDefault(); setSelection(cid(_cols,_rows)); return; }
    }

    // Any printable char starts editing
    if (!e.ctrlKey && !e.metaKey && !e.altKey && e.key.length === 1) {
      startEdit(_anchor, e.key);
    }
  }

  /* ── Refresh all cells ── */
  function refreshAll() {
    for (const ref in _cells) refreshCell(ref);
  }

  function scheduleSave() {
    clearTimeout(_saveTimer);
    _saveTimer = setTimeout(() => { if (_saveCallback) _saveCallback(); }, 800);
  }

  /* ── Init ── */
  function init(saveCallback) {
    _saveCallback = saveCallback;

    // Formula bar
    const fb = document.getElementById('formula-bar');
    fb.addEventListener('focus', () => { _fbarActive = true; });
    fb.addEventListener('blur', () => { setTimeout(() => { if (!_fbarRange) { _fbarActive=false; _fbarRange=false; } }, 150); });
    fb.addEventListener('input', e => { showAutocomplete(e.target, e.target.value); });
    fb.addEventListener('keydown', e => {
      if (e.key === 'Escape') {
        document.querySelectorAll('.xl-cell.formula-selecting').forEach(el => el.classList.remove('formula-selecting'));
        _fbarActive = false; _fbarRange = false; updateFormulaBar(); fb.blur(); return;
      }
    });

    // Name box
    const nb = document.getElementById('name-box');
    nb.addEventListener('keydown', e => {
      if (e.key === 'Enter') {
        const v = nb.value.trim().toUpperCase();
        // Check named range
        if (_named[v]) { setSelection(_named[v].split(':')[0], _named[v].split(':')[1]||_named[v]); }
        else if (/^[A-Z]+\d+$/.test(v)) setSelection(v);
        else if (/^[A-Z]+\d+:[A-Z]+\d+$/.test(v)) { const [a2,b2]=v.split(':'); setSelection(a2,b2); }
        nb.blur();
      }
    });

    // Mouse up
    document.addEventListener('mouseup', () => {
      _mouseSelecting = false; _colSelecting = false; _rowSelecting = false;
      if (_fbarRange) {
        _fbarRange = false;
        document.querySelectorAll('.xl-cell.formula-selecting').forEach(el => el.classList.remove('formula-selecting'));
        const inserted = buildFbarRange(_fbarRangeStart, _fbarRangeEnd);
        const newPos = _fbarPreCursor + inserted.length;
        fb.focus(); fb.setSelectionRange(newPos, newPos);
        _fbarRangeStart = null; _fbarRangeEnd = null;
      }
    });

    // Grid keyboard
    const wrap = document.getElementById('excel-grid-wrap');
    wrap.setAttribute('tabindex', '0');
    wrap.addEventListener('keydown', handleKeydown);

    // Formatting buttons
    bindFmt('xl-bold',      () => applyFmt('toggle_bold'));
    bindFmt('xl-italic',    () => applyFmt('toggle_italic'));
    bindFmt('xl-underline', () => applyFmt('toggle_under'));
    bindFmt('xl-strike',    () => applyFmt('toggle_strike'));
    bindFmt('xl-wrap',      () => applyFmt('toggle_wrap'));
    bindFmt('xl-al-left',   () => applyFmt('align', 'left'));
    bindFmt('xl-al-center', () => applyFmt('align', 'center'));
    bindFmt('xl-al-right',  () => applyFmt('align', 'right'));

    const xlFont = document.getElementById('xl-font');
    xlFont?.addEventListener('change', e => applyFmt('fontFamily', e.target.value));

    const xlFsize = document.getElementById('xl-fsize');
    xlFsize?.addEventListener('change', e => applyFmt('fontSize', parseInt(e.target.value)));

    const xlNumFmt = document.getElementById('xl-numfmt');
    xlNumFmt?.addEventListener('change', e => applyFmt('numberFormat', e.target.value));

    document.getElementById('xl-txt-color')?.addEventListener('input', e => applyFmt('textColor', e.target.value));
    document.getElementById('xl-fill-color')?.addEventListener('input', e => applyFmt('fillColor', e.target.value));

    document.getElementById('xl-dec-more')?.addEventListener('click', () => {
      const fmt = _fmts[_anchor] || {};
      applyFmt('decimals', Math.min(10, (fmt.decimals ?? 2) + 1));
    });
    document.getElementById('xl-dec-less')?.addEventListener('click', () => {
      const fmt = _fmts[_anchor] || {};
      applyFmt('decimals', Math.max(0, (fmt.decimals ?? 2) - 1));
    });

    // Undo/Redo
    document.getElementById('xl-undo')?.addEventListener('click', undo);
    document.getElementById('xl-redo')?.addEventListener('click', redo);

    // Sort
    document.getElementById('xl-sort-asc')?.addEventListener('click', () => sortCol(true));
    document.getElementById('xl-sort-desc')?.addEventListener('click', () => sortCol(false));

    // Row/Col ops
    document.getElementById('xl-ins-row')?.addEventListener('click', insertRow);
    document.getElementById('xl-ins-col')?.addEventListener('click', insertCol);
    document.getElementById('xl-del-row')?.addEventListener('click', deleteRow);
    document.getElementById('xl-del-col')?.addEventListener('click', deleteCol);

    // Formula insert dropdown
    document.getElementById('xl-formula-insert')?.addEventListener('change', e => {
      const f = e.target.value; if (!f) return;
      const bar = document.getElementById('formula-bar');
      bar.value = f; bar.focus();
      const pi = f.indexOf('(') + 1;
      bar.setSelectionRange(pi, pi);
      e.target.value = '';
    });

    // Name cell
    document.getElementById('xl-name-cell')?.addEventListener('click', () => {
      document.getElementById('cell-range-input').value = _anchor || '';
      renderNamedList();
      window._AppModal.open('named-cell-modal');
    });

    // Dropdown cell
    document.getElementById('xl-dropdown-cell')?.addEventListener('click', () => {
      if (!_anchor) return;
      const existing = (_valid[_anchor] || []).join('\n');
      document.getElementById('validation-options-input').value = existing;
      window._AppModal.open('validation-modal');
    });

    // Find button
    document.getElementById('xl-find-btn')?.addEventListener('click', openFind);

    // Save named cell
    document.getElementById('save-cell-name-btn')?.addEventListener('click', () => {
      const name  = document.getElementById('cell-name-input').value.trim();
      const range = document.getElementById('cell-range-input').value.trim().toUpperCase();
      if (name && range) { _named[name] = range; scheduleSave(); renderNamedList(); rebuildGrid(); document.getElementById('cell-name-input').value=''; }
    });

    // Save validation
    document.getElementById('save-validation-btn')?.addEventListener('click', () => {
      if (!_anchor) return;
      const opts = document.getElementById('validation-options-input').value.split('\n').map(s=>s.trim()).filter(Boolean);
      if (opts.length) _valid[_anchor] = opts; else delete _valid[_anchor];
      scheduleSave(); rebuildGrid(); window._AppModal.close();
    });

    // Context menu
    document.getElementById('ctx-name')?.addEventListener('click', () => {
      document.getElementById('context-menu').classList.add('hidden');
      document.getElementById('cell-range-input').value = _pendingCtxRef||_anchor||'';
      renderNamedList();
      window._AppModal.open('named-cell-modal');
    });
    document.getElementById('ctx-dropdown')?.addEventListener('click', () => {
      document.getElementById('context-menu').classList.add('hidden');
      if (_pendingCtxRef) { _anchor=_pendingCtxRef; const e=(_valid[_pendingCtxRef]||[]).join('\n'); document.getElementById('validation-options-input').value=e; window._AppModal.open('validation-modal'); }
    });
    document.getElementById('ctx-clear')?.addEventListener('click', () => {
      document.getElementById('context-menu').classList.add('hidden');
      const s=getSel(); pushHistory();
      for(let r=s.minR;r<=s.maxR;r++) for(let c=s.minC;c<=s.maxC;c++) { delete _cells[cid(c,r)]; refreshCell(cid(c,r)); }
      scheduleSave();
    });
    document.getElementById('ctx-insert-row')?.addEventListener('click', () => { document.getElementById('context-menu').classList.add('hidden'); insertRow(); });
    document.getElementById('ctx-insert-col')?.addEventListener('click', () => { document.getElementById('context-menu').classList.add('hidden'); insertCol(); });
    document.addEventListener('click', e => { if (!e.target.closest('#context-menu')) document.getElementById('context-menu')?.classList.add('hidden'); });
    document.addEventListener('click', e => { if (!e.target.closest('#xl-autocomplete,#formula-bar')) document.getElementById('xl-autocomplete')?.classList.add('hidden'); });

    pushHistory(); // initial snapshot
  }

  function bindFmt(id, fn) {
    document.getElementById(id)?.addEventListener('click', fn);
  }

  /* ── Load / Save ── */
  function load(pageData) {
    _cells = pageData.cells       ? JSON.parse(JSON.stringify(pageData.cells))       : {};
    _fmts  = pageData.fmts        ? JSON.parse(JSON.stringify(pageData.fmts))        : {};
    _named = pageData.namedRanges ? JSON.parse(JSON.stringify(pageData.namedRanges)) : {};
    _valid = pageData.validations ? JSON.parse(JSON.stringify(pageData.validations)) : {};
    _colW  = pageData.colWidths   ? JSON.parse(JSON.stringify(pageData.colWidths))   : {};
    _rowH  = pageData.rowHeights  ? JSON.parse(JSON.stringify(pageData.rowHeights))  : {};
    _rows  = pageData.rowCount    || DEF_ROWS;
    _cols  = pageData.colCount    || DEF_COLS;
    _anchor = 'A1'; _focus = 'A1';
    _hist = []; _histIdx = -1;
    buildGrid();
    pushHistory();
  }

  function save() {
    return {
      cells:       JSON.parse(JSON.stringify(_cells)),
      fmts:        JSON.parse(JSON.stringify(_fmts)),
      namedRanges: JSON.parse(JSON.stringify(_named)),
      validations: JSON.parse(JSON.stringify(_valid)),
      colWidths:   JSON.parse(JSON.stringify(_colW)),
      rowHeights:  JSON.parse(JSON.stringify(_rowH)),
      rowCount:    _rows,
      colCount:    _cols,
    };
  }

  return { init, load, save };
})();
