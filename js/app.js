/* ═══════════════════════════════════════
   InkAI — Core App
   State management, routing, modals, theme
═══════════════════════════════════════ */

(() => {
  'use strict';

  // ──────────────────────────────────────
  // STATE
  // ──────────────────────────────────────
  let state = {
    theme: 'dark',
    notebooks: [],
    activeNotebookId: null,
    activePageId: null,
  };

  // Transient (not persisted beyond session)
  let _pendingNewPageNotebookId = null;
  let _selectedPageFormat = 'word';
  let _sidebarCollapsed = false;

  // ──────────────────────────────────────
  // STORAGE
  // ──────────────────────────────────────
  const STORAGE_KEY = 'inkai_v1';

  function loadState() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const saved = JSON.parse(raw);
        state = { ...state, ...saved };
      }
    } catch (e) { console.warn('InkAI: failed to load state', e); }
  }

  function saveState() {
    try {
      // Save current editor data into the active page
      flushActiveEditor();
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    } catch (e) { console.warn('InkAI: failed to save state', e); }
  }

  function flushActiveEditor() {
    if (!state.activeNotebookId || !state.activePageId) return;
    const page = getActivePage();
    if (!page) return;

    if (page.format === 'word') {
      const data = window.WordEditor.save();
      page.content = data.content || '';
    } else if (page.format === 'excel') {
      const data = window.ExcelEditor.save();
      // Merge all fields (cells, fmts, namedRanges, validations, colWidths, rowHeights, rowCount, colCount)
      Object.assign(page, data);
    } else if (page.format === 'design') {
      const data = window.DesignEditor.save();
      if (data.imageData) page.imageData = data.imageData;
    }
  }

  // ──────────────────────────────────────
  // HELPERS
  // ──────────────────────────────────────
  function uuid() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => {
      const r = Math.random() * 16 | 0;
      return (c === 'x' ? r : (r & 0x3 | 0x8)).toString(16);
    });
  }

  function getActiveNotebook() {
    return state.notebooks.find(n => n.id === state.activeNotebookId) || null;
  }

  function getActivePage() {
    const nb = getActiveNotebook();
    if (!nb) return null;
    return nb.pages.find(p => p.id === state.activePageId) || null;
  }

  function getAppState() {
    return {
      activeNotebook: getActiveNotebook(),
      activePageId: state.activePageId,
    };
  }

  // ──────────────────────────────────────
  // MODALS
  // ──────────────────────────────────────
  window._AppModal = {
    open(id) {
      document.getElementById('modal-overlay').classList.remove('hidden');
      document.querySelectorAll('.modal').forEach(m => m.classList.add('hidden'));
      document.getElementById(id).classList.remove('hidden');
    },
    close() {
      document.getElementById('modal-overlay').classList.add('hidden');
      document.querySelectorAll('.modal').forEach(m => m.classList.add('hidden'));
    }
  };

  function initModals() {
    // Close on overlay click
    document.getElementById('modal-overlay').addEventListener('click', e => {
      if (e.target === document.getElementById('modal-overlay')) window._AppModal.close();
    });

    // All modal-close and modal-cancel buttons
    document.querySelectorAll('.modal-close, .modal-cancel').forEach(btn => {
      btn.addEventListener('click', () => window._AppModal.close());
    });

    // ── New Notebook ──
    document.getElementById('new-notebook-btn').addEventListener('click', () => {
      document.getElementById('notebook-name-input').value = '';
      window._AppModal.open('notebook-modal');
      setTimeout(() => document.getElementById('notebook-name-input').focus(), 100);
    });
    document.getElementById('welcome-start-btn').addEventListener('click', () => {
      document.getElementById('notebook-name-input').value = '';
      window._AppModal.open('notebook-modal');
      setTimeout(() => document.getElementById('notebook-name-input').focus(), 100);
    });
    document.getElementById('create-notebook-confirm').addEventListener('click', createNotebook);
    document.getElementById('notebook-name-input').addEventListener('keydown', e => {
      if (e.key === 'Enter') createNotebook();
    });

    // ── New Page (format selector) ──
    document.querySelectorAll('.format-card').forEach(card => {
      card.addEventListener('click', () => {
        document.querySelectorAll('.format-card').forEach(c => c.classList.remove('active'));
        card.classList.add('active');
        _selectedPageFormat = card.dataset.format;
      });
    });
    document.getElementById('create-page-confirm').addEventListener('click', createPage);
    document.getElementById('page-name-input').addEventListener('keydown', e => {
      if (e.key === 'Enter') createPage();
    });
  }

  // ──────────────────────────────────────
  // SIDEBAR
  // ──────────────────────────────────────
  function initSidebar() {
    const sidebar = document.getElementById('sidebar');

    document.getElementById('sidebar-collapse').addEventListener('click', () => {
      _sidebarCollapsed = !_sidebarCollapsed;
      sidebar.classList.toggle('collapsed', _sidebarCollapsed);
      document.getElementById('sidebar-open').classList.toggle('hidden', !_sidebarCollapsed);
    });

    document.getElementById('sidebar-open').addEventListener('click', () => {
      _sidebarCollapsed = false;
      sidebar.classList.remove('collapsed');
      document.getElementById('sidebar-open').classList.add('hidden');
    });

    // Mobile menu
    document.getElementById('mobile-menu-btn').addEventListener('click', () => {
      sidebar.classList.toggle('mobile-open');
    });

    // Close mobile sidebar when clicking outside
    document.addEventListener('click', e => {
      if (window.innerWidth <= 768 && !sidebar.contains(e.target) &&
          !document.getElementById('mobile-menu-btn').contains(e.target)) {
        sidebar.classList.remove('mobile-open');
      }
    });
  }

  function renderSidebar() {
    const list = document.getElementById('notebooks-list');
    list.innerHTML = '';

    state.notebooks.forEach(nb => {
      const item = document.createElement('div');
      item.className = 'notebook-item';

      const isOpen = nb.id === state.activeNotebookId;

      // Notebook header
      const header = document.createElement('div');
      header.className = 'notebook-header' + (isOpen ? ' open' : '');
      header.innerHTML = `
        <i class="fa-solid fa-chevron-right nb-icon"></i>
        <i class="fa-solid fa-book-open" style="color:var(--accent);font-size:0.75rem;flex-shrink:0"></i>
        <span class="nb-name" title="${escHtml(nb.title)}">${escHtml(nb.title)}</span>
        <div class="nb-actions">
          <button class="nb-action-btn" title="Add Page" data-action="add-page" data-nb="${nb.id}">
            <i class="fa-solid fa-plus"></i>
          </button>
          <button class="nb-action-btn danger" title="Delete Notebook" data-action="delete-nb" data-nb="${nb.id}">
            <i class="fa-solid fa-trash"></i>
          </button>
        </div>
      `;

      header.addEventListener('click', e => {
        if (e.target.closest('[data-action]')) return;
        if (state.activeNotebookId === nb.id) {
          header.classList.toggle('open');
          pagesList.classList.toggle('hidden');
        } else {
          state.activeNotebookId = nb.id;
          state.activePageId = null;
          renderSidebar();
          showWelcome();
        }
      });

      // Action buttons
      header.querySelectorAll('[data-action]').forEach(btn => {
        btn.addEventListener('click', e => {
          e.stopPropagation();
          const action = btn.dataset.action;
          if (action === 'add-page') {
            _pendingNewPageNotebookId = nb.id;
            openAddPageModal();
          } else if (action === 'delete-nb') {
            if (confirm(`Delete notebook "${nb.title}"? This cannot be undone.`)) {
              deleteNotebook(nb.id);
            }
          }
        });
      });

      item.appendChild(header);

      // Pages list
      const pagesList = document.createElement('div');
      pagesList.className = 'pages-list' + (isOpen ? '' : ' hidden');

      nb.pages.forEach(page => {
        const pageEl = createPageEl(nb, page);
        pagesList.appendChild(pageEl);
      });

      // Add page button
      const addBtn = document.createElement('button');
      addBtn.className = 'add-page-btn';
      addBtn.innerHTML = '<i class="fa-solid fa-plus"></i> Add Page';
      addBtn.addEventListener('click', () => {
        _pendingNewPageNotebookId = nb.id;
        openAddPageModal();
      });
      pagesList.appendChild(addBtn);

      item.appendChild(pagesList);
      list.appendChild(item);
    });

    updateBreadcrumb();
  }

  function createPageEl(nb, page) {
    const el = document.createElement('div');
    el.className = 'page-item' + (page.id === state.activePageId ? ' active' : '');
    el.dataset.pageId = page.id;

    const icon = { word: 'fa-file-lines', excel: 'fa-table', design: 'fa-paint-brush' }[page.format] || 'fa-file';
    el.innerHTML = `
      <i class="fa-solid ${icon}" style="color:var(--accent);font-size:0.75rem;flex-shrink:0"></i>
      <span style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escHtml(page.title)}</span>
      <span class="page-format-badge">${page.format}</span>
      <div class="pg-actions">
        <button class="nb-action-btn danger" title="Delete Page" data-action="delete-page" data-pg="${page.id}" data-nb="${nb.id}">
          <i class="fa-solid fa-xmark"></i>
        </button>
      </div>
    `;

    el.addEventListener('click', e => {
      if (e.target.closest('[data-action]')) return;
      openPage(nb.id, page.id);
    });

    el.querySelectorAll('[data-action]').forEach(btn => {
      btn.addEventListener('click', e => {
        e.stopPropagation();
        if (btn.dataset.action === 'delete-page') {
          if (confirm(`Delete page "${page.title}"?`)) {
            deletePage(nb.id, page.id);
          }
        }
      });
    });

    return el;
  }

  function updateBreadcrumb() {
    const bc = document.getElementById('breadcrumb');
    const nb = getActiveNotebook();
    const page = getActivePage();

    let html = '<span class="bc-note">InkAI</span>';
    if (nb) {
      html += ' <span class="bc-sep"><i class="fa-solid fa-chevron-right" style="font-size:0.65rem"></i></span> ';
      html += `<span class="bc-nb">${escHtml(nb.title)}</span>`;
    }
    if (page) {
      html += ' <span class="bc-sep"><i class="fa-solid fa-chevron-right" style="font-size:0.65rem"></i></span> ';
      html += `<span class="bc-page">${escHtml(page.title)}</span>`;
    }
    bc.innerHTML = html;
  }

  // ──────────────────────────────────────
  // NOTEBOOKS & PAGES CRUD
  // ──────────────────────────────────────
  function createNotebook() {
    const nameInput = document.getElementById('notebook-name-input');
    const name = nameInput.value.trim() || 'New Notebook';
    const nb = { id: uuid(), title: name, pages: [], createdAt: Date.now() };
    state.notebooks.push(nb);
    state.activeNotebookId = nb.id;
    state.activePageId = null;
    window._AppModal.close();
    renderSidebar();
    showWelcome();
    saveState();

    // Prompt to add first page
    _pendingNewPageNotebookId = nb.id;
    setTimeout(() => openAddPageModal(), 200);
  }

  function deleteNotebook(id) {
    state.notebooks = state.notebooks.filter(n => n.id !== id);
    if (state.activeNotebookId === id) {
      state.activeNotebookId = state.notebooks[0]?.id || null;
      state.activePageId = null;
    }
    renderSidebar();
    if (!getActivePage()) showWelcome();
    saveState();
  }

  function openAddPageModal() {
    document.getElementById('page-name-input').value = '';
    _selectedPageFormat = 'word';
    document.querySelectorAll('.format-card').forEach(c => c.classList.remove('active'));
    document.querySelector('.format-card[data-format="word"]').classList.add('active');
    window._AppModal.open('page-modal');
    setTimeout(() => document.getElementById('page-name-input').focus(), 100);
  }

  function createPage() {
    const nbId = _pendingNewPageNotebookId || state.activeNotebookId;
    if (!nbId) return;
    const nb = state.notebooks.find(n => n.id === nbId);
    if (!nb) return;

    const nameInput = document.getElementById('page-name-input');
    const name = nameInput.value.trim() || `Page ${nb.pages.length + 1}`;
    const page = {
      id: uuid(),
      title: name,
      format: _selectedPageFormat,
      content: '',
      cells: {},
      namedRanges: {},
      validations: {},
      rowCount: 50,
      colCount: 26,
      imageData: null,
      createdAt: Date.now(),
    };
    nb.pages.push(page);

    window._AppModal.close();
    state.activeNotebookId = nbId;
    state.activePageId = page.id;
    renderSidebar();
    openPage(nbId, page.id);
    saveState();
  }

  function deletePage(nbId, pageId) {
    const nb = state.notebooks.find(n => n.id === nbId);
    if (!nb) return;
    nb.pages = nb.pages.filter(p => p.id !== pageId);
    if (state.activePageId === pageId) {
      state.activePageId = nb.pages[nb.pages.length - 1]?.id || null;
    }
    renderSidebar();
    if (state.activePageId) openPage(nbId, state.activePageId);
    else showWelcome();
    saveState();
  }

  // ──────────────────────────────────────
  // PAGE SWITCHING
  // ──────────────────────────────────────
  function openPage(nbId, pageId) {
    // Flush current editor
    flushActiveEditor();

    state.activeNotebookId = nbId;
    state.activePageId = pageId;

    const page = getActivePage();
    if (!page) { showWelcome(); return; }

    // Hide welcome
    document.getElementById('welcome-screen').classList.add('hidden');

    // Show correct editor, hide others
    const editors = { word: 'word-editor', excel: 'excel-editor', design: 'design-editor' };
    const toolbars = { word: 'word-toolbar', excel: 'excel-toolbar', design: 'design-toolbar' };

    Object.values(editors).forEach(id => document.getElementById(id).classList.add('hidden'));
    Object.values(toolbars).forEach(id => document.getElementById(id).classList.add('hidden'));
    document.getElementById('ai-bar').classList.add('hidden');
    document.getElementById('ai-popover').classList.add('hidden');

    const editorId = editors[page.format];
    const toolbarId = toolbars[page.format];

    document.getElementById(editorId).classList.remove('hidden');
    document.getElementById(toolbarId).classList.remove('hidden');

    // Load content
    if (page.format === 'word') {
      window.WordEditor.load({ content: page.content || '' });
    } else if (page.format === 'excel') {
      window.ExcelEditor.load(page);
    } else if (page.format === 'design') {
      window.DesignEditor.load(page);
      setTimeout(() => window.DesignEditor.resizeCanvas(), 50);
    }

    updateBreadcrumb();
    renderSidebar();

    // Close mobile sidebar
    document.getElementById('sidebar').classList.remove('mobile-open');
  }

  function showWelcome() {
    document.getElementById('welcome-screen').classList.remove('hidden');
    document.querySelectorAll('.editor').forEach(el => el.classList.add('hidden'));
    ['word-toolbar', 'excel-toolbar', 'design-toolbar', 'ai-bar'].forEach(id => {
      document.getElementById(id).classList.add('hidden');
    });
    updateBreadcrumb();
  }

  // ──────────────────────────────────────
  // THEME
  // ──────────────────────────────────────
  function toggleTheme() {
    state.theme = state.theme === 'dark' ? 'light' : 'dark';
    applyTheme(state.theme);
    saveState();
  }

  function initTheme() {
    applyTheme(state.theme);

    // Sidebar theme toggle (still works)
    document.getElementById('theme-toggle').addEventListener('click', toggleTheme);

    // Topbar theme pill — always visible
    document.getElementById('topbar-theme-btn').addEventListener('click', toggleTheme);
  }

  function applyTheme(theme) {
    document.documentElement.dataset.theme = theme;
    const isDark = theme === 'dark';

    // Sidebar toggle icons
    document.getElementById('theme-icon-dark').style.display = isDark ? 'inline' : 'none';
    document.getElementById('theme-icon-light').style.display = isDark ? 'none' : 'inline';
    document.getElementById('theme-label').textContent = isDark ? 'Dark Mode' : 'Light Mode';

    // Topbar pill label
    const tpLabel = document.getElementById('tp-label');
    if (tpLabel) tpLabel.textContent = isDark ? 'Dark' : 'Light';

    // Update word editor text color default
    const tc = document.getElementById('text-color');
    if (tc) tc.value = isDark ? '#c0caf5' : '#1a1a1a';
  }

  // ──────────────────────────────────────
  // RESIZE
  // ──────────────────────────────────────
  function initResize() {
    window.addEventListener('resize', () => {
      const page = getActivePage();
      if (page && page.format === 'design') {
        window.DesignEditor.resizeCanvas();
      }
    });
  }

  // ──────────────────────────────────────
  // KEYBOARD SHORTCUTS
  // ──────────────────────────────────────
  function initKeyboardShortcuts() {
    document.addEventListener('keydown', e => {
      // Ctrl+S — save
      if ((e.ctrlKey || e.metaKey) && e.key === 's') {
        e.preventDefault();
        saveState();
        flashSaveIndicator();
      }
      // Ctrl+N — new page
      if ((e.ctrlKey || e.metaKey) && e.key === 'n' && !e.shiftKey) {
        e.preventDefault();
        if (state.activeNotebookId) {
          _pendingNewPageNotebookId = state.activeNotebookId;
          openAddPageModal();
        }
      }
    });
  }

  function flashSaveIndicator() {
    const bc = document.getElementById('breadcrumb');
    const saved = document.createElement('span');
    saved.style.cssText = 'color:var(--accent);font-size:0.72rem;margin-left:8px;animation:fadeIn 0.2s ease';
    saved.innerHTML = '<i class="fa-solid fa-check"></i> Saved';
    bc.appendChild(saved);
    setTimeout(() => saved.remove(), 1800);
  }

  // ──────────────────────────────────────
  // AUTO-SAVE
  // ──────────────────────────────────────
  function initAutoSave() {
    // Save on editor changes
    const editorSaveCb = () => saveState();

    // Expose to editors
    window._EditorSaveCallback = editorSaveCb;

    // Periodic auto-save every 30 seconds
    setInterval(saveState, 30000);

    // Save before page unload
    window.addEventListener('beforeunload', saveState);
  }

  // ──────────────────────────────────────
  // INIT
  // ──────────────────────────────────────
  function init() {
    loadState();

    // Init sub-systems
    window.WordEditor.init(saveState);
    window.ExcelEditor.init(saveState);
    window.DesignEditor.init(saveState);
    window.ExportEngine.init(getAppState, null);

    initModals();
    initSidebar();
    initTheme();
    initResize();
    initKeyboardShortcuts();
    initAutoSave();

    // Render sidebar
    renderSidebar();

    // Restore active page if any
    if (state.activeNotebookId && state.activePageId) {
      openPage(state.activeNotebookId, state.activePageId);
    } else {
      showWelcome();
    }

    console.log('%cInkAI loaded ✓', 'color:#cc1122;font-weight:700;font-size:14px');
  }

  // ──────────────────────────────────────
  // UTILITY
  // ──────────────────────────────────────
  function escHtml(str) {
    return (str || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }

  // Boot
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

})();
