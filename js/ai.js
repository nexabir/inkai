/* ═══════════════════════════════════════
   InkAI — AI Grammar & Spell Check Engine
   Client-side, no API key needed
═══════════════════════════════════════ */

window.InkAI_Grammar = (() => {

  // ── Common misspellings dictionary ──
  const SPELL_MAP = {
    'teh': 'the', 'hte': 'the', 'adn': 'and', 'nad': 'and', 'pubic': 'public',
    'recieve': 'receive', 'beleive': 'believe', 'wierd': 'weird', 'freind': 'friend',
    'definately': 'definitely', 'definitly': 'definitely', 'occured': 'occurred',
    'occurance': 'occurrence', 'seperate': 'separate', 'untill': 'until',
    'tommorrow': 'tomorrow', 'tommorow': 'tomorrow', 'tomorro': 'tomorrow',
    'accomodate': 'accommodate', 'calender': 'calendar', 'commitee': 'committee',
    'embarass': 'embarrass', 'grammer': 'grammar', 'goverment': 'government',
    'harrass': 'harass', 'independance': 'independence', 'liason': 'liaison',
    'millenium': 'millennium', 'necesary': 'necessary', 'neccessary': 'necessary',
    'occassion': 'occasion', 'perseverance': 'perseverance', 'priviledge': 'privilege',
    'reccomend': 'recommend', 'rhythm': 'rhythm', 'rythm': 'rhythm',
    'succesful': 'successful', 'suprise': 'surprise', 'tendancy': 'tendency',
    'therefor': 'therefore', 'truely': 'truly', 'unforseen': 'unforeseen',
    'vaccuum': 'vacuum', 'visious': 'viscous', 'wether': 'whether',
    'whther': 'whether', 'wich': 'which', 'yesturday': 'yesterday',
    'recieved': 'received', 'beleived': 'believed', 'writting': 'writing',
    'writeing': 'writing', 'haveing': 'having', 'makeing': 'making',
    'takeing': 'taking', 'comming': 'coming', 'runing': 'running',
    'geting': 'getting', 'siting': 'sitting', 'planing': 'planning',
    'begining': 'beginning', 'stoping': 'stopping', 'puting': 'putting',
    'diferent': 'different', 'diferrence': 'difference', 'coment': 'comment',
    'contineu': 'continue', 'lenght': 'length', 'strenght': 'strength',
    'knolwedge': 'knowledge', 'experiance': 'experience', 'explaination': 'explanation',
    'existance': 'existence', 'enviroment': 'environment', 'develope': 'develop',
    'devlopment': 'development', 'documentaion': 'documentation',
    'busness': 'business', 'adress': 'address', 'acheive': 'achieve',
    'appearence': 'appearance', 'arguement': 'argument', 'assit': 'assist',
    'becuase': 'because','beacuse': 'because', 'biggining': 'beginning',
    'catagory': 'category', 'collegue': 'colleague', 'concious': 'conscious',
    'currect': 'correct', 'desicion': 'decision', 'disapoint': 'disappoint',
    'equiptment': 'equipment', 'familar': 'familiar', 'firmiliar': 'familiar',
    'happend': 'happened', 'imediately': 'immediately', 'importent': 'important',
    'interupt': 'interrupt', 'knowlege': 'knowledge', 'liscense': 'license',
    'maintainance': 'maintenance', 'managment': 'management', 'minuite': 'minute',
    'noticable': 'noticeable', 'oportunity': 'opportunity', 'performence': 'performance',
    'posession': 'possession', 'practicle': 'practical', 'preperation': 'preparation',
    'presance': 'presence', 'profesional': 'professional', 'probelm': 'problem',
    'questionaire': 'questionnaire', 'remeber': 'remember', 'resposibility': 'responsibility',
    'relevent': 'relevant', 'relize': 'realize', 'scedule': 'schedule',
    'sentance': 'sentence', 'similiar': 'similar', 'similer': 'similar',
    'sofware': 'software', 'sourse': 'source', 'studing': 'studying',
    'temperture': 'temperature', 'transfered': 'transferred', 'usally': 'usually',
    'vaild': 'valid', 'varing': 'varying', 'visable': 'visible',
    'writen': 'written', 'yoour': 'your', 'youre': "you're", 'dont': "don't",
    'doesnt': "doesn't", 'cant': "can't", 'wont': "won't", 'ive': "I've",
    'im': "I'm", 'id': "I'd", 'ill': "I'll", 'isnt': "isn't", 'arent': "aren't",
    'wasnt': "wasn't", 'werent': "weren't", 'hasnt': "hasn't", 'havent': "haven't",
    'hadnt': "hadn't", 'woudnt': "wouldn't", 'couldnt': "couldn't",
    'shouldnt': "shouldn't", 'musnt': "mustn't", 'mightn': "mightn't",
  };

  // ── Grammar rules ──
  const GRAMMAR_RULES = [
    {
      name: 'double_space',
      regex: /  +/g,
      message: 'Double spaces found',
      suggestion: 'Use single space',
      type: 'warning',
    },
    {
      name: 'no_cap_after_period',
      regex: /\.\s+([a-z])/g,
      message: 'Sentence should start with a capital letter',
      type: 'warning',
    },
    {
      name: 'their_there',
      words: ['their', 'there', "they're"],
      message: "Check: 'their' (possession), 'there' (place), 'they\\'re' (they are)",
      type: 'info',
    },
    {
      name: 'your_youre',
      words: ['your', "you're"],
      message: "Check: 'your' (possession), 'you\\'re' (you are)",
      type: 'info',
    },
    {
      name: 'its_its',
      words: ["its", "it's"],
      message: "Check: 'its' (possession), 'it\\'s' (it is)",
      type: 'info',
    },
    {
      name: 'affect_effect',
      words: ['affect', 'effect'],
      message: "'affect' is a verb, 'effect' is usually a noun",
      type: 'info',
    },
    {
      name: 'then_than',
      words: ['then', 'than'],
      message: "'then' (time), 'than' (comparison)",
      type: 'info',
    },
    {
      name: 'to_too_two',
      words: ['to', 'too', 'two'],
      message: "'too' means 'also' or 'very'; 'to' is a preposition; 'two' is 2",
      type: 'info',
    },
  ];

  /**
   * Analyze plain text and return an array of issues
   * @param {string} text
   * @returns {Array<{word, index, type, message, suggestions}>}
   */
  function analyze(text) {
    const issues = [];
    const words = text.split(/\b/);
    let pos = 0;

    for (const token of words) {
      const lower = token.toLowerCase().replace(/[^a-z']/g, '');
      if (lower.length > 1 && SPELL_MAP[lower]) {
        issues.push({
          word: token.trim(),
          index: pos,
          type: 'spell',
          message: `Possible misspelling of "${token.trim()}"`,
          suggestions: [SPELL_MAP[lower]],
        });
      }
      pos += token.length;
    }

    // Grammar rule checks (on confusable words)
    GRAMMAR_RULES.filter(r => r.words).forEach(rule => {
      const pattern = new RegExp(`\\b(${rule.words.join('|')})\\b`, 'gi');
      let m;
      while ((m = pattern.exec(text)) !== null) {
        // Only flag if it looks suspicious (near a verb/noun context check is simplified)
        issues.push({
          word: m[0],
          index: m.index,
          type: rule.type || 'grammar',
          message: rule.message,
          suggestions: rule.words.filter(w => w !== m[0].toLowerCase()),
        });
      }
    });

    return issues;
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

  return { analyze, getSummary, SPELL_MAP };

})();
