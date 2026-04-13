/* ═══════════════════════════════════════
   InkAI — Export Engine
   PNG / PDF — current page, selected, all
═══════════════════════════════════════ */

window.ExportEngine = (() => {

  let _getState = null; // callback to get app state
  let _renderPageForExport = null; // callback to render a page temporarily

  function init(getStateCb, renderPageCb) {
    _getState = getStateCb;
    _renderPageForExport = renderPageCb;

    // Export button opens modal
    document.getElementById('export-btn').addEventListener('click', openExportModal);

    // Scope radio shows/hides page checkboxes
    document.querySelectorAll('input[name="exp-scope"]').forEach(radio => {
      radio.addEventListener('change', () => {
        const box = document.getElementById('page-checkboxes');
        box.classList.toggle('hidden', radio.value !== 'selected');
      });
    });

    // Trigger export
    document.getElementById('do-export-btn').addEventListener('click', doExport);
  }

  function openExportModal() {
    const state = _getState();
    if (!state.activeNotebook) { alert('Please open a notebook first.'); return; }

    // Build page checkbox list
    const box = document.getElementById('page-checkboxes');
    box.innerHTML = '';
    state.activeNotebook.pages.forEach(page => {
      const label = document.createElement('label');
      label.innerHTML = `
        <input type="checkbox" value="${page.id}" checked>
        <span>${page.title}</span>
        <span class="page-format-badge">${page.format}</span>
      `;
      box.appendChild(label);
    });

    window._AppModal.open('export-modal');
  }

  async function doExport() {
    const state = _getState();
    if (!state.activeNotebook) return;

    const scope = document.querySelector('input[name="exp-scope"]:checked').value;
    const fmt = document.querySelector('input[name="exp-fmt"]:checked').value;

    let pagesToExport = [];

    if (scope === 'current') {
      const pg = state.activeNotebook.pages.find(p => p.id === state.activePageId);
      if (pg) pagesToExport = [pg];
    } else if (scope === 'selected') {
      const checked = [...document.querySelectorAll('#page-checkboxes input:checked')].map(i => i.value);
      pagesToExport = state.activeNotebook.pages.filter(p => checked.includes(p.id));
    } else {
      pagesToExport = [...state.activeNotebook.pages];
    }

    if (!pagesToExport.length) { alert('No pages selected.'); return; }

    window._AppModal.close();

    const btn = document.getElementById('do-export-btn');
    btn.textContent = 'Exporting…';
    btn.disabled = true;

    try {
      if (fmt === 'png') {
        await exportAsPNG(pagesToExport, state);
      } else {
        await exportAsPDF(pagesToExport, state);
      }
    } catch (err) {
      console.error('Export error:', err);
      alert('Export failed: ' + err.message);
    }

    btn.innerHTML = '<i class="fa-solid fa-download"></i> Export';
    btn.disabled = false;
  }

  // ──────────────────────────────────────
  // PNG EXPORT
  // ──────────────────────────────────────
  async function exportAsPNG(pages, state) {
    for (const page of pages) {
      const el = await buildExportElement(page, state);
      document.body.appendChild(el);
      await wait(100);

      const canvas = await html2canvas(el, {
        backgroundColor: '#ffffff',
        scale: 2,
        useCORS: true,
        allowTaint: true,
        logging: false,
      });

      document.body.removeChild(el);

      const link = document.createElement('a');
      link.download = `InkAI_${sanitizeFilename(page.title)}.png`;
      link.href = canvas.toDataURL('image/png');
      link.click();

      if (pages.length > 1) await wait(400);
    }
  }

  // ──────────────────────────────────────
  // PDF EXPORT
  // ──────────────────────────────────────
  async function exportAsPDF(pages, state) {
    const { jsPDF } = window.jspdf;
    let pdf = null;
    let firstPage = true;

    for (const page of pages) {
      const el = await buildExportElement(page, state);
      document.body.appendChild(el);
      await wait(120);

      const canvas = await html2canvas(el, {
        backgroundColor: '#ffffff',
        scale: 2,
        useCORS: true,
        allowTaint: true,
        logging: false,
      });

      document.body.removeChild(el);

      const imgW = canvas.width;
      const imgH = canvas.height;

      // A4: 210 x 297 mm
      const pageW = 210;
      const pageH = Math.min(297, (imgH / imgW) * pageW);
      const imgData = canvas.toDataURL('image/jpeg', 0.92);

      if (!pdf) {
        pdf = new jsPDF({ orientation: imgW > imgH ? 'landscape' : 'portrait', unit: 'mm', format: [pageW, pageH] });
      } else {
        pdf.addPage([pageW, pageH], imgW > imgH ? 'landscape' : 'portrait');
      }

      pdf.addImage(imgData, 'JPEG', 0, 0, pageW, pageH);

      if (pages.length > 1) await wait(200);
    }

    if (pdf) {
      const nbName = state.activeNotebook.title || 'InkAI';
      pdf.save(`InkAI_${sanitizeFilename(nbName)}.pdf`);
    }
  }

  // ──────────────────────────────────────
  // BUILD EXPORT ELEMENT
  // ──────────────────────────────────────
  async function buildExportElement(page, state) {
    const wrap = document.createElement('div');
    wrap.style.cssText = `
      position: fixed;
      left: -9999px;
      top: 0;
      width: 900px;
      min-height: 600px;
      background: #ffffff;
      color: #1a1a1a;
      font-family: Inter, sans-serif;
      padding: 0;
      z-index: -1;
      overflow: hidden;
    `;

    // Header
    const header = document.createElement('div');
    header.style.cssText = `
      background: #8B0000;
      color: #fff;
      padding: 14px 24px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      font-size: 13px;
    `;
    const nbTitle = state.activeNotebook ? state.activeNotebook.title : 'Notebook';
    header.innerHTML = `
      <div style="display:flex;align-items:center;gap:10px">
        <svg viewBox="0 0 24 24" width="20" height="20" style="flex-shrink:0">
          <ellipse cx="12" cy="16" rx="6" ry="8" fill="rgba(255,255,255,0.8)"/>
          <circle cx="12" cy="6" r="3" fill="rgba(255,255,255,0.8)"/>
        </svg>
        <strong style="font-size:15px">InkAI</strong>
      </div>
      <div style="text-align:right">
        <div style="font-weight:600">${escHtml(nbTitle)} › ${escHtml(page.title)}</div>
        <div style="opacity:0.8;font-size:11px">${new Date().toLocaleString()} · ${page.format.toUpperCase()} mode</div>
      </div>
    `;
    wrap.appendChild(header);

    // Content
    const content = document.createElement('div');
    content.style.cssText = 'padding: 24px; background:#fff; color:#1a1a1a; min-height:500px;';

    if (page.format === 'word') {
      content.innerHTML = page.content || '<p style="color:#aaa">Empty page</p>';
      // Fix images
      content.querySelectorAll('img').forEach(img => { img.style.maxWidth = '100%'; });
    } else if (page.format === 'excel') {
      content.appendChild(buildExcelTable(page));
    } else if (page.format === 'design') {
      if (page.imageData) {
        const img = document.createElement('img');
        img.src = page.imageData;
        img.style.cssText = 'max-width:100%;display:block;border:1px solid #eee;border-radius:8px;';
        content.appendChild(img);
      } else {
        content.innerHTML = '<p style="color:#aaa">Empty canvas</p>';
      }
    }

    wrap.appendChild(content);
    return wrap;
  }

  function buildExcelTable(page) {
    const cells = page.cells || {};
    const namedRanges = page.namedRanges || {};
    const wrap = document.createElement('div');
    wrap.style.cssText = 'overflow:auto;';

    // Find used range
    let maxRow = 1, maxCol = 1;
    for (const ref in cells) {
      const m = ref.match(/^([A-Z]+)(\d+)$/);
      if (!m) continue;
      const c = colToIdx(m[1]);
      const r = parseInt(m[2]);
      if (r > maxRow) maxRow = r;
      if (c > maxCol) maxCol = c;
    }
    maxRow = Math.min(maxRow + 2, 100);
    maxCol = Math.min(maxCol + 2, 26);

    const table = document.createElement('table');
    table.style.cssText = 'border-collapse:collapse;font-size:12px;width:100%;';

    // Header
    const thead = document.createElement('thead');
    const hrow = document.createElement('tr');
    const th0 = document.createElement('th');
    th0.style.cssText = 'width:40px;background:#f5f5f5;border:1px solid #ddd;padding:4px;';
    hrow.appendChild(th0);
    for (let c = 1; c <= maxCol; c++) {
      const th = document.createElement('th');
      th.style.cssText = 'background:#f5f5f5;border:1px solid #ddd;padding:4px 8px;text-align:center;font-size:11px;color:#555;';
      th.textContent = idxToCol(c);
      hrow.appendChild(th);
    }
    thead.appendChild(hrow);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    for (let r = 1; r <= maxRow; r++) {
      const row = document.createElement('tr');
      const rh = document.createElement('td');
      rh.style.cssText = 'background:#f5f5f5;border:1px solid #ddd;padding:3px 6px;text-align:center;font-size:10px;color:#888;';
      rh.textContent = r;
      row.appendChild(rh);
      for (let c = 1; c <= maxCol; c++) {
        const ref = idxToCol(c) + r;
        const cell = cells[ref] || {};
        const td = document.createElement('td');
        td.style.cssText = 'border:1px solid #ddd;padding:3px 8px;min-width:80px;';

        // Evaluate formula for export
        let displayVal = '';
        if (cell.formula && cell.formula.startsWith('=')) {
          try {
            // Simple engine for export
            displayVal = evalForExport(cell.formula, cells, namedRanges);
          } catch { displayVal = '#ERROR!'; }
        } else {
          displayVal = cell.value || '';
        }

        if (typeof displayVal === 'string' && displayVal.startsWith('#')) {
          td.style.color = '#cc0000';
        } else if (cell.formula && cell.formula.startsWith('=')) {
          td.style.color = '#0055cc';
        }
        td.textContent = displayVal;
        row.appendChild(td);
      }
      tbody.appendChild(row);
    }
    table.appendChild(tbody);
    wrap.appendChild(table);
    return wrap;
  }

  function evalForExport(formula, cells, namedRanges) {
    // Minimal inline evaluator for export
    const expr = formula.substring(1).trim();
    const funcM = expr.match(/^([A-Z]+)\(([\s\S]*)\)$/);
    if (!funcM) return expr; // arithmetic fallback

    const args = funcM[2].split(',').map(s => s.trim());
    const getVal = (ref) => {
      const c = cells[ref.toUpperCase()];
      if (!c) return 0;
      if (c.formula && c.formula.startsWith('=')) {
        try { return evalForExport(c.formula, cells, namedRanges); } catch { return 0; }
      }
      const n = parseFloat(c.value);
      return isNaN(n) ? (c.value || 0) : n;
    };
    const resolveRange = (r) => {
      if (namedRanges[r]) r = namedRanges[r];
      if (!r.includes(':')) return [getVal(r)];
      const [s, e] = r.split(':');
      const sc = colToIdx(s.match(/[A-Z]+/)[0]), sr = parseInt(s.match(/\d+/)[0]);
      const ec = colToIdx(e.match(/[A-Z]+/)[0]), er = parseInt(e.match(/\d+/)[0]);
      const vals = [];
      for (let row = sr; row <= er; row++)
        for (let col = sc; col <= ec; col++) vals.push(getVal(idxToCol(col) + row));
      return vals;
    };
    const nums = (argsList) => argsList.flatMap(a => resolveRange(a)).map(v => parseFloat(v)).filter(n => !isNaN(n));

    switch (funcM[1]) {
      case 'SUM': return nums(args).reduce((a,b)=>a+b,0);
      case 'AVERAGE': { const n=nums(args); return n.length?n.reduce((a,b)=>a+b,0)/n.length:0; }
      case 'COUNT': return nums(args).length;
      case 'MIN': { const n=nums(args); return n.length?Math.min(...n):0; }
      case 'MAX': { const n=nums(args); return n.length?Math.max(...n):0; }
      default: return formula;
    }
  }

  function colToIdx(col) {
    let idx = 0;
    for (let i = 0; i < col.length; i++) idx = idx * 26 + (col.toUpperCase().charCodeAt(i) - 64);
    return idx;
  }
  function idxToCol(idx) {
    let col = '';
    while (idx > 0) { col = String.fromCharCode(64 + (((idx - 1) % 26) + 1)) + col; idx = Math.floor((idx - 1) / 26); }
    return col;
  }

  // ── Utilities ──
  function wait(ms) { return new Promise(r => setTimeout(r, ms)); }
  function sanitizeFilename(name) { return (name || 'note').replace(/[^a-z0-9_\-]/gi, '_').substring(0, 40); }
  function escHtml(str) { return (str || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

  return { init };

})();
