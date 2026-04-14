/* ═══════════════════════════════════════
   InkAI — AI Grammar & Spell Check Engine
   Powered by LanguageTool Public API
═══════════════════════════════════════ */

window.InkAI_Grammar = (() => {

  /**
   * Analyze plain text via LanguageTool API
   * @param {string} text
   * @returns {Promise<Array<{word, offset, length, type, message, suggestions}>>}
   */
  async function analyze(text) {
    if (!text.trim()) return [];
    
    try {
      const res = await fetch('https://api.languagetool.org/v2/check', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          text: text,
          language: 'auto'
        })
      });
      
      if (!res.ok) throw new Error('API down');
      const data = await res.json();
      
      return data.matches.map(m => ({
        word: text.substring(m.offset, m.offset + m.length),
        offset: m.offset,
        length: m.length,
        type: m.rule.issueType === 'misspelling' ? 'spell' : 'grammar',
        message: m.message,
        suggestions: m.replacements.slice(0, 4).map(r => r.value)
      }));
    } catch (err) {
      console.error('Grammar check failed:', err);
      // Fallback response safely
      return [];
    }
  }

  /**
   * Get a quick summary string for the AI bar
   */
  function getSummary(issues) {
    const spellCount = issues.filter(i => i.type === 'spell').length;
    const grammarCount = issues.filter(i => i.type !== 'spell').length;
    if (!issues.length) return '✓ No issues found — text looks great!';
    const parts = [];
    if (spellCount) parts.push(`${spellCount} spelling issue${spellCount > 1 ? 's' : ''}`);
    if (grammarCount) parts.push(`${grammarCount} grammar hint${grammarCount > 1 ? 's' : ''}`);
    return `AI found: ${parts.join(' • ')}`;
  }

  return { analyze, getSummary };

})();
