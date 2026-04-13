/* ═══════════════════════════════════════
   InkAI — Design (Canvas Paint) Editor
═══════════════════════════════════════ */

window.DesignEditor = (() => {

  let _canvas = null;
  let _ctx = null;
  let _tool = 'pen';
  let _color = '#8B0000';
  let _fillColor = '#ffffff';
  let _brushSize = 5;
  let _useFill = false;
  let _drawing = false;
  let _startX = 0;
  let _startY = 0;
  let _snapshot = null; // for shape preview
  let _history = [];
  let _redoStack = [];
  let _saveCallback = null;
  let _saveTimer = null;
  let _initialized = false;
  let _textInput = null;

  const MAX_HISTORY = 50;

  function init(saveCallback) {
    _saveCallback = saveCallback;
    _canvas = document.getElementById('design-canvas');
    _ctx = _canvas.getContext('2d');

    bindToolbar();
    bindCanvasEvents();
    _initialized = true;
  }

  function bindToolbar() {
    // Tool buttons
    document.querySelectorAll('#design-toolbar .tool-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('#design-toolbar .tool-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        _tool = btn.dataset.tool;
        _canvas.style.cursor = _tool === 'text' ? 'text' : (_tool === 'fill' ? 'cell' : 'crosshair');
        // Remove text input if switching away from text
        if (_tool !== 'text' && _textInput) { _textInput.remove(); _textInput = null; }
      });
    });

    // Colors
    document.getElementById('design-color').addEventListener('input', e => { _color = e.target.value; });
    document.getElementById('fill-color').addEventListener('input', e => { _fillColor = e.target.value; });

    // Brush size
    const sizeSlider = document.getElementById('brush-size');
    const sizeLabel = document.getElementById('brush-size-label');
    sizeSlider.addEventListener('input', e => {
      _brushSize = parseInt(e.target.value);
      sizeLabel.textContent = _brushSize + 'px';
    });

    // Fill toggle
    const fillToggle = document.getElementById('fill-toggle');
    fillToggle.addEventListener('click', () => {
      _useFill = !_useFill;
      fillToggle.textContent = _useFill ? '■ Filled' : '⬜ Outline';
    });

    // Undo / Redo
    document.getElementById('design-undo').addEventListener('click', undo);
    document.getElementById('design-redo').addEventListener('click', redo);

    // Clear
    document.getElementById('design-clear').addEventListener('click', () => {
      if (!confirm('Clear the entire canvas?')) return;
      saveHistory();
      _ctx.clearRect(0, 0, _canvas.width, _canvas.height);
      fillBackground();
      scheduleSave();
    });
  }

  // ── Canvas sizing ──
  function resizeCanvas(width, height) {
    // Save image before resize
    let imgData = null;
    if (_canvas.width > 0 && _canvas.height > 0) {
      try { imgData = _ctx.getImageData(0, 0, _canvas.width, _canvas.height); } catch {}
    }
    _canvas.width = width || Math.max(800, document.getElementById('canvas-wrap').offsetWidth - 4);
    _canvas.height = height || Math.max(600, document.getElementById('design-editor').offsetHeight - 44);
    fillBackground();
    if (imgData) { try { _ctx.putImageData(imgData, 0, 0); } catch {} }
  }

  function fillBackground() {
    // White canvas background
    const prevStyle = _ctx.fillStyle;
    _ctx.fillStyle = '#ffffff';
    _ctx.fillRect(0, 0, _canvas.width, _canvas.height);
    _ctx.fillStyle = prevStyle;
  }

  // ── History ──
  function saveHistory() {
    if (_history.length >= MAX_HISTORY) _history.shift();
    _history.push(_ctx.getImageData(0, 0, _canvas.width, _canvas.height));
    _redoStack = [];
  }

  function undo() {
    if (!_history.length) return;
    _redoStack.push(_ctx.getImageData(0, 0, _canvas.width, _canvas.height));
    const state = _history.pop();
    _ctx.putImageData(state, 0, 0);
    scheduleSave();
  }

  function redo() {
    if (!_redoStack.length) return;
    _history.push(_ctx.getImageData(0, 0, _canvas.width, _canvas.height));
    const state = _redoStack.pop();
    _ctx.putImageData(state, 0, 0);
    scheduleSave();
  }

  // ── Events ──
  function getPos(e) {
    const rect = _canvas.getBoundingClientRect();
    const scaleX = _canvas.width / rect.width;
    const scaleY = _canvas.height / rect.height;
    if (e.touches) {
      return {
        x: (e.touches[0].clientX - rect.left) * scaleX,
        y: (e.touches[0].clientY - rect.top) * scaleY,
      };
    }
    return {
      x: (e.clientX - rect.left) * scaleX,
      y: (e.clientY - rect.top) * scaleY,
    };
  }

  function bindCanvasEvents() {
    _canvas.addEventListener('mousedown', startDraw);
    _canvas.addEventListener('mousemove', duringDraw);
    _canvas.addEventListener('mouseup', endDraw);
    _canvas.addEventListener('mouseleave', endDraw);

    _canvas.addEventListener('touchstart', e => { e.preventDefault(); startDraw(e); }, { passive: false });
    _canvas.addEventListener('touchmove', e => { e.preventDefault(); duringDraw(e); }, { passive: false });
    _canvas.addEventListener('touchend', e => { e.preventDefault(); endDraw(e); }, { passive: false });
  }

  function startDraw(e) {
    if (_tool === 'text') { handleTextTool(e); return; }
    if (_tool === 'fill') { handleFill(e); return; }
    const pos = getPos(e);
    _startX = pos.x;
    _startY = pos.y;
    _drawing = true;
    saveHistory();

    _ctx.strokeStyle = _color;
    _ctx.fillStyle = _useFill ? _fillColor : _color;
    _ctx.lineWidth = _brushSize;
    _ctx.lineCap = 'round';
    _ctx.lineJoin = 'round';

    if (_tool === 'pen' || _tool === 'brush') {
      _ctx.beginPath();
      _ctx.moveTo(_startX, _startY);
    }

    if (_tool === 'line' || _tool === 'rect' || _tool === 'circle') {
      _snapshot = _ctx.getImageData(0, 0, _canvas.width, _canvas.height);
    }
  }

  function duringDraw(e) {
    if (!_drawing) return;
    const pos = getPos(e);

    if (_tool === 'pen' || _tool === 'brush') {
      _ctx.strokeStyle = _color;
      _ctx.lineWidth = _tool === 'brush' ? _brushSize * 2 : _brushSize;
      _ctx.globalAlpha = _tool === 'brush' ? 0.6 : 1;
      _ctx.lineTo(pos.x, pos.y);
      _ctx.stroke();
      _ctx.globalAlpha = 1;
      return;
    }

    if (_tool === 'eraser') {
      _ctx.globalCompositeOperation = 'destination-out';
      _ctx.beginPath();
      _ctx.arc(pos.x, pos.y, _brushSize * 2, 0, Math.PI * 2);
      _ctx.fill();
      _ctx.globalCompositeOperation = 'source-over';
      return;
    }

    // Shape preview: restore snapshot then draw shape
    if (_snapshot) _ctx.putImageData(_snapshot, 0, 0);

    _ctx.strokeStyle = _color;
    _ctx.fillStyle = _useFill ? _fillColor : _color;
    _ctx.lineWidth = _brushSize;

    if (_tool === 'line') {
      _ctx.beginPath();
      _ctx.moveTo(_startX, _startY);
      _ctx.lineTo(pos.x, pos.y);
      _ctx.stroke();
    }
    if (_tool === 'rect') {
      const w = pos.x - _startX;
      const h = pos.y - _startY;
      if (_useFill) _ctx.fillRect(_startX, _startY, w, h);
      _ctx.strokeRect(_startX, _startY, w, h);
    }
    if (_tool === 'circle') {
      const rx = Math.abs(pos.x - _startX) / 2;
      const ry = Math.abs(pos.y - _startY) / 2;
      const cx = _startX + (pos.x - _startX) / 2;
      const cy = _startY + (pos.y - _startY) / 2;
      _ctx.beginPath();
      _ctx.ellipse(cx, cy, rx, ry, 0, 0, Math.PI * 2);
      if (_useFill) _ctx.fill();
      _ctx.stroke();
    }
  }

  function endDraw(e) {
    if (!_drawing) return;
    _drawing = false;
    _snapshot = null;
    _ctx.globalAlpha = 1;
    _ctx.globalCompositeOperation = 'source-over';
    scheduleSave();
  }

  function handleFill(e) {
    const pos = getPos(e);
    saveHistory();
    floodFill(Math.round(pos.x), Math.round(pos.y), hexToRgba(_color));
    scheduleSave();
  }

  function handleTextTool(e) {
    if (_textInput) { _textInput.remove(); _textInput = null; }
    const pos = getPos(e);
    const rect = _canvas.getBoundingClientRect();
    const scaleX = rect.width / _canvas.width;
    const scaleY = rect.height / _canvas.height;

    const inp = document.createElement('textarea');
    inp.style.cssText = `
      position:fixed;
      left:${e.clientX}px;
      top:${e.clientY}px;
      background:transparent;
      border:1px dashed var(--accent);
      color:${_color};
      font-size:${Math.max(12, _brushSize * 3)}px;
      font-family:var(--font);
      outline:none;
      resize:both;
      min-width:100px;
      min-height:40px;
      z-index:9999;
      padding:4px;
    `;
    document.body.appendChild(inp);
    inp.focus();
    _textInput = inp;

    inp.addEventListener('blur', () => {
      const text = inp.value;
      if (text) {
        saveHistory();
        _ctx.fillStyle = _color;
        _ctx.font = `${Math.max(12, _brushSize * 3)}px Inter, sans-serif`;
        // Multi-line text
        const lines = text.split('\n');
        lines.forEach((line, i) => {
          _ctx.fillText(line, pos.x, pos.y + i * Math.max(16, _brushSize * 3.5));
        });
        scheduleSave();
      }
      inp.remove();
      _textInput = null;
    });
  }

  // ── Flood fill ──
  function hexToRgba(hex) {
    const r = parseInt(hex.slice(1,3), 16);
    const g = parseInt(hex.slice(3,5), 16);
    const b = parseInt(hex.slice(5,7), 16);
    return [r, g, b, 255];
  }

  function floodFill(startX, startY, fillColor) {
    const imgData = _ctx.getImageData(0, 0, _canvas.width, _canvas.height);
    const data = imgData.data;
    const w = _canvas.width, h = _canvas.height;

    const getIdx = (x, y) => (y * w + x) * 4;
    const idx = getIdx(startX, startY);
    const targetColor = [data[idx], data[idx+1], data[idx+2], data[idx+3]];

    if (targetColor.every((v, i) => v === fillColor[i])) return;

    const colorsMatch = (x, y) => {
      const i = getIdx(x, y);
      return Math.abs(data[i]-targetColor[0]) < 30 &&
             Math.abs(data[i+1]-targetColor[1]) < 30 &&
             Math.abs(data[i+2]-targetColor[2]) < 30 &&
             Math.abs(data[i+3]-targetColor[3]) < 30;
    };

    const setColor = (x, y) => {
      const i = getIdx(x, y);
      data[i] = fillColor[0]; data[i+1] = fillColor[1];
      data[i+2] = fillColor[2]; data[i+3] = fillColor[3];
    };

    const stack = [[startX, startY]];
    const visited = new Set();
    while (stack.length) {
      const [x, y] = stack.pop();
      if (x < 0 || x >= w || y < 0 || y >= h) continue;
      const key = y * w + x;
      if (visited.has(key)) continue;
      visited.add(key);
      if (!colorsMatch(x, y)) continue;
      setColor(x, y);
      stack.push([x+1,y],[x-1,y],[x,y+1],[x,y-1]);
    }
    _ctx.putImageData(imgData, 0, 0);
  }

  // ──────────────────────────────────────
  // SAVE / LOAD
  // ──────────────────────────────────────
  function scheduleSave() {
    clearTimeout(_saveTimer);
    _saveTimer = setTimeout(() => { if (_saveCallback) _saveCallback(); }, 1000);
  }

  function load(pageData) {
    // Must wait for canvas to be visible
    requestAnimationFrame(() => {
      resizeCanvas();
      _history = [];
      _redoStack = [];

      if (pageData.imageData) {
        const img = new Image();
        img.onload = () => {
          fillBackground();
          _ctx.drawImage(img, 0, 0);
        };
        img.src = pageData.imageData;
      }
    });
  }

  function save() {
    if (!_canvas || !_canvas.width) return {};
    return { imageData: _canvas.toDataURL('image/png') };
  }

  return { init, load, save, resizeCanvas };

})();
