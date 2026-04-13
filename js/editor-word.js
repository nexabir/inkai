/* ═══════════════════════════════════════
   InkAI — Word Editor
═══════════════════════════════════════ */

window.WordEditor = (() => {

  let _content = null;
  let _saveCallback = null;
  let _aiTimeout = null;

  function init(saveCallback) {
    _saveCallback = saveCallback;
    _content = document.getElementById('word-content');

    // Bind toolbar buttons
    document.querySelectorAll('#word-toolbar .tb-btn[data-cmd]').forEach(btn => {
      btn.addEventListener('mousedown', e => {
        e.preventDefault();
        const cmd = btn.dataset.cmd;
        if (cmd === 'createLink') {
          const url = prompt('Enter URL:', 'https://');
          if (url) document.execCommand('createLink', false, url);
        } else {
          document.execCommand(cmd, false, null);
        }
        updateToolbarState();
      });
    });

    // Font family
    document.getElementById('font-family').addEventListener('change', e => {
      document.execCommand('fontName', false, e.target.value);
    });

    // Font size
    document.getElementById('font-size').addEventListener('change', e => {
      document.execCommand('fontSize', false, e.target.value);
    });

    // Text color
    document.getElementById('text-color').addEventListener('input', e => {
      document.execCommand('foreColor', false, e.target.value);
    });

    // Background/highlight color
    document.getElementById('bg-color').addEventListener('input', e => {
      document.execCommand('hiliteColor', false, e.target.value);
    });

    // Image insert
    document.getElementById('tb-img').addEventListener('click', () => {
      document.getElementById('img-upload').click();
    });
    document.getElementById('img-upload').addEventListener('change', e => {
      const file = e.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = ev => {
        document.execCommand('insertHTML', false, `<img src="${ev.target.result}" style="max-width:100%;border-radius:8px;margin:8px 0">`);
      };
      reader.readAsDataURL(file);
      e.target.value = '';
    });

    // AI check button
    document.getElementById('tb-ai-check').addEventListener('click', runAICheck);

    // Toolbar state update on selection change
    _content.addEventListener('keyup', () => { updateToolbarState(); scheduleSave(); });
    _content.addEventListener('mouseup', updateToolbarState);
    _content.addEventListener('focus', updateToolbarState);

    // Hide AI popover on click outside
    document.addEventListener('click', e => {
      const popover = document.getElementById('ai-popover');
      if (!popover.contains(e.target)) popover.classList.add('hidden');
    });

    // AI bar close
    document.getElementById('ai-bar-close').addEventListener('click', () => {
      document.getElementById('ai-bar').classList.add('hidden');
      document.getElementById('ai-popover').classList.add('hidden');
    });
  }

  function load(pageData) {
    if (!_content) return;
    _content.innerHTML = pageData.content || '';
    _content.focus();
    updateToolbarState();
  }

  function save() {
    if (!_content) return {};
    return { content: _content.innerHTML };
  }

  function scheduleSave() {
    clearTimeout(_aiTimeout);
    _aiTimeout = setTimeout(() => {
      if (_saveCallback) _saveCallback();
    }, 600);
  }

  function updateToolbarState() {
    const cmds = ['bold', 'italic', 'underline', 'strikeThrough',
      'justifyLeft', 'justifyCenter', 'justifyRight', 'justifyFull'];
    cmds.forEach(cmd => {
      const btn = document.querySelector(`#word-toolbar [data-cmd="${cmd}"]`);
      if (btn) {
        const active = document.queryCommandState(cmd);
        btn.classList.toggle('active', active);
      }
    });
  }

  function runAICheck() {
    if (!_content) return;
    const text = _content.innerText || '';
    const issues = window.InkAI_Grammar.analyze(text);
    const summary = window.InkAI_Grammar.getSummary(issues);

    const bar = document.getElementById('ai-bar');
    const barMsg = document.getElementById('ai-bar-msg');
    bar.classList.remove('hidden');
    barMsg.textContent = summary;

    // Show popover with suggestions
    if (issues.length > 0) {
      showAISuggestions(issues);
    }
  }

  function showAISuggestions(issues) {
    const popover = document.getElementById('ai-popover');
    const list = document.getElementById('ai-suggestions-list');
    list.innerHTML = '';

    // Show first 8 issues
    const shown = issues.slice(0, 8);
    shown.forEach(issue => {
      const item = document.createElement('div');
      item.className = 'ai-suggestion-item';
      item.innerHTML = `
        <i class="fa-solid fa-${issue.type === 'spell' ? 'spell-check' : 'triangle-exclamation'}" style="color:${issue.type === 'spell' ? 'var(--accent-bright)' : '#ffaa00'}"></i>
        <div style="flex:1;min-width:0">
          <div style="font-weight:600;color:var(--txt);font-size:0.81rem">"${issue.word}"</div>
          <div style="font-size:0.72rem;color:var(--txt3);margin-top:2px">${issue.message}</div>
          ${issue.suggestions && issue.suggestions.length ? `<div style="font-size:0.75rem;color:var(--accent-bright);margin-top:3px">✓ ${issue.suggestions.slice(0,3).join(', ')}</div>` : ''}
        </div>
      `;
      // Click to replace
      if (issue.suggestions && issue.suggestions.length) {
        item.addEventListener('click', () => {
          replaceWordInEditor(issue.word, issue.suggestions[0]);
          popover.classList.add('hidden');
        });
        item.style.cursor = 'pointer';
      }
      list.appendChild(item);
    });

    if (issues.length > 8) {
      const more = document.createElement('div');
      more.style.cssText = 'text-align:center;font-size:0.72rem;color:var(--txt3);padding:6px';
      more.textContent = `+ ${issues.length - 8} more issues`;
      list.appendChild(more);
    }

    // Position near AI bar
    const bar = document.getElementById('ai-bar');
    const rect = bar.getBoundingClientRect();
    popover.style.top = (rect.bottom + 8) + 'px';
    popover.style.left = Math.min(rect.left, window.innerWidth - 300) + 'px';
    popover.classList.remove('hidden');
  }

  function replaceWordInEditor(original, replacement) {
    if (!_content) return;
    // Simple innerHTML replacement for the first occurrence
    const escaped = original.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp(`\\b${escaped}\\b`);

    // Walk text nodes and replace
    const walker = document.createTreeWalker(_content, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      const node = walker.currentNode;
      if (regex.test(node.textContent)) {
        node.textContent = node.textContent.replace(regex, replacement);
        scheduleSave();
        break;
      }
    }
  }

  return { init, load, save, runAICheck };

})();
