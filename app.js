/* ==========================================================================
   MOEX CBT Exam System Replica JavaScript Engine
   Includes: 8-Page Answer Sheet per Question, Punctuation Insertion,
   IndexedDB Auto-Save, Revision History & Word (.docx) Export
   ========================================================================== */

(function () {
  'use strict';

  // Questions Database
  const QUESTIONS = {
    1: {
      title: '第一題 (50.00分)',
      content: '範例題目僅供參考'
    },
    2: {
      title: '第二題 (50.00分)',
      content: '範例題目僅供參考'
    }
  };

  // Helper to create empty 8 pages
  function createEmptyQuestionPages() {
    const pages = {};
    for (let i = 1; i <= 8; i++) {
      pages[i] = '';
    }
    return pages;
  }

  // Application State
  let currentQ = 1;
  let currentPage = 1;
  let answers = {
    1: createEmptyQuestionPages(),
    2: createEmptyQuestionPages()
  };

  let activeTextarea = null;
  let db = null;
  let autoSaveTimer = null;
  let historyList = [];
  let selectedHistoryItem = null;
  let previewCurrentQ = 1;
  let fontSize = 18; // px

  // DOM Elements
  const selectQuestion = document.getElementById('selectQuestion');
  const selectPage = document.getElementById('selectPage');
  const btnPrevQuestion = document.getElementById('btnPrevQuestion');
  const btnNextQuestion = document.getElementById('btnNextQuestion');
  const questionTitle = document.getElementById('questionTitle');
  const questionBody = document.getElementById('questionBody');
  const questionContent = document.getElementById('questionContent');
  const btnToggleQuestion = document.getElementById('btnToggleQuestion');

  const paperPagesContainer = document.getElementById('paperPagesContainer');
  const autoSaveBadge = document.getElementById('autoSaveBadge');
  const autoSaveText = document.getElementById('autoSaveText');
  const manualSaveBadge = document.getElementById('manualSaveBadge');
  const manualSaveText = document.getElementById('manualSaveText');

  const statCurrentQ = document.getElementById('statCurrentQ');
  const statTotalChars = document.getElementById('statTotalChars');
  const statTotalLines = document.getElementById('statTotalLines');
  const statTotalPages = document.getElementById('statTotalPages');
  const answeredCountEl = document.getElementById('answeredCount');
  const unansweredCountEl = document.getElementById('unansweredCount');

  // History Drawer Elements
  const btnOpenHistory = document.getElementById('btnOpenHistory');
  const btnCloseHistory = document.getElementById('btnCloseHistory');
  const historyDrawer = document.getElementById('historyDrawer');
  const historyOverlay = document.getElementById('historyOverlay');
  const historyListEl = document.getElementById('historyList');
  const previewTimestamp = document.getElementById('previewTimestamp');
  const previewQTabs = document.getElementById('previewQTabs');
  const previewContent = document.getElementById('previewContent');
  const btnRestoreVersion = document.getElementById('btnRestoreVersion');
  const btnCreateSnapshot = document.getElementById('btnCreateSnapshot');
  const btnClearHistory = document.getElementById('btnClearHistory');
  const btnManualSave = document.getElementById('btnManualSave');
  const btnExportWord = document.getElementById('btnExportWord');

  // Timer
  const countdownTimerEl = document.getElementById('countdownTimer');
  let secondsRemaining = 4 * 3600 - 1; // 4 hours

  // Initialize DB
  function initIndexedDB() {
    return new Promise((resolve) => {
      try {
        const request = indexedDB.open('MOEX_CBT_EXAM_DB', 3);
        request.onupgradeneeded = function (e) {
          const dbInstance = e.target.result;
          if (!dbInstance.objectStoreNames.contains('answers')) {
            dbInstance.createObjectStore('answers', { keyPath: 'qId' });
          }
          if (!dbInstance.objectStoreNames.contains('history')) {
            const store = dbInstance.createObjectStore('history', { keyPath: 'id', autoIncrement: true });
            store.createIndex('timestamp', 'timestamp', { unique: false });
          }
        };
        request.onsuccess = function (e) {
          db = e.target.result;
          resolve(db);
        };
        request.onerror = function (e) {
          console.error('IndexedDB Error:', e);
          db = null;
          resolve(null);
        };
      } catch (e) {
        console.error('IndexedDB init Exception:', e);
        db = null;
        resolve(null);
      }
    });
  }

  // Helper for normalizing loaded answer state (handles legacy string vs 8-page object)
  function normalizeAnswerData(raw) {
    const pages = createEmptyQuestionPages();
    if (!raw) return pages;

    if (typeof raw === 'string') {
      pages[1] = raw;
    } else if (typeof raw === 'object') {
      for (let p = 1; p <= 8; p++) {
        pages[p] = raw[p] || '';
      }
    }
    return pages;
  }

  // Helper to get total combined text for a question
  function getQuestionCombinedText(qId) {
    const qAnswers = answers[qId] || {};
    let combined = [];
    for (let p = 1; p <= 8; p++) {
      const pText = (qAnswers[p] || '').trim();
      if (pText.length > 0) {
        combined.push(`【第 ${qId} 題 第 ${p} 頁】\n${qAnswers[p]}`);
      }
    }
    return combined.join('\n\n');
  }

  // Load Saved State (Always start fresh on page load, history remains in DB)
  async function loadState() {
    // Start fresh with clean empty exam paper
    answers = {
      1: createEmptyQuestionPages(),
      2: createEmptyQuestionPages()
    };
    renderEditor();
    updateStats();

    // Create an initial startup snapshot if history is empty
    setTimeout(() => {
      fetchHistoryList((items) => {
        if (items.length === 0) {
          addHistorySnapshot('自動儲存');
        }
      });
    }, 500);
  }

  // Save Answers to DB
  function saveCurrentAnswer(isManual = false) {
    if (!isManual) {
      updateAutoSaveStatus('saving');
    }

    // Save to LocalStorage fallback
    localStorage.setItem('MOEX_CBT_ANSWERS', JSON.stringify(answers));

    if (db) {
      const transaction = db.transaction(['answers'], 'readwrite');
      const store = transaction.objectStore('answers');
      store.put({ qId: 1, pages: answers[1] });
      store.put({ qId: 2, pages: answers[2] });
    }

    const now = new Date();
    const timeStr = now.toTimeString().split(' ')[0];

    if (isManual) {
      addHistorySnapshot('手動暫存');
      if (manualSaveText) manualSaveText.textContent = `手動暫存：${timeStr}`;
      if (manualSaveBadge) {
        manualSaveBadge.classList.remove('highlight');
        void manualSaveBadge.offsetWidth; // trigger reflow
        manualSaveBadge.classList.add('highlight');
      }
    } else {
      setTimeout(() => {
        updateAutoSaveStatus('saved', `已自動儲存 ${timeStr}`);
      }, 400);
    }
  }

  function updateAutoSaveStatus(state, text = '') {
    if (!autoSaveBadge) return;
    const dot = autoSaveBadge.querySelector('.status-dot');
    if (state === 'saving') {
      if (dot) dot.classList.add('saving');
      if (autoSaveText) autoSaveText.textContent = '儲存中...';
    } else {
      if (dot) dot.classList.remove('saving');
      if (autoSaveText) autoSaveText.textContent = text || '已自動儲存';
    }
  }

  // Add History Snapshot
  function addHistorySnapshot(type = '自動儲存') {
    const now = new Date();
    const timestampStr = now.toLocaleDateString() + ' ' + now.toTimeString().split(' ')[0];
    
    let totalChars = 0;
    [1, 2].forEach(q => {
      for (let p = 1; p <= 8; p++) {
        totalChars += (answers[q][p] || '').length;
      }
    });

    const snapshot = {
      timestamp: timestampStr,
      time: now.getTime(),
      type: type,
      qId: currentQ,
      answers: JSON.parse(JSON.stringify(answers)),
      charCount: totalChars
    };

    // Synchronous LocalStorage save fallback (Guaranteed in all environments)
    try {
      const saved = localStorage.getItem('MOEX_CBT_HISTORY');
      let list = saved ? JSON.parse(saved) : [];
      list.unshift(snapshot);
      if (list.length > 50) list.pop();
      localStorage.setItem('MOEX_CBT_HISTORY', JSON.stringify(list));
      historyList = list;
    } catch (e) {
      console.warn('LocalStorage save history warning:', e);
    }

    // Async IndexedDB write if available
    if (db && db.objectStoreNames.contains('history')) {
      try {
        const tx = db.transaction(['history'], 'readwrite');
        tx.objectStore('history').add(snapshot);
      } catch (e) {
        console.warn('IndexedDB history write warning:', e);
      }
    }
  }

  // Fetch History List
  function fetchHistoryList(callback) {
    if (db && db.objectStoreNames.contains('history')) {
      try {
        const tx = db.transaction(['history'], 'readonly');
        const store = tx.objectStore('history');
        const req = store.getAll();
        req.onsuccess = function () {
          const res = req.result ? req.result.sort((a, b) => b.time - a.time) : [];
          if (res && res.length > 0) {
            callback(res);
            return;
          }
          getLocalStorageHistory(callback);
        };
        req.onerror = function () {
          getLocalStorageHistory(callback);
        };
        return;
      } catch (e) {
        console.warn('IndexedDB fetchHistoryList exception:', e);
      }
    }

    getLocalStorageHistory(callback);
  }

  function getLocalStorageHistory(callback) {
    try {
      const saved = localStorage.getItem('MOEX_CBT_HISTORY');
      const list = saved ? JSON.parse(saved) : [];
      callback(list);
    } catch (e) {
      callback([]);
    }
  }

  // Render 8 Paper Pages per Question
  function renderEditor() {
    paperPagesContainer.innerHTML = '';
    activeTextarea = null;

    for (let p = 1; p <= 8; p++) {
      const page = document.createElement('div');
      page.className = 'paper-page';
      page.id = `paper-page-${p}`;

      const leftMargin = document.createElement('div');
      leftMargin.className = 'paper-left-margin';
      leftMargin.textContent = `第 ${currentQ} 題 第 ${p} 頁`;

      const centerArea = document.createElement('div');
      centerArea.className = 'paper-center-area';

      const textarea = document.createElement('textarea');
      textarea.className = 'paper-textarea';
      textarea.style.fontSize = `${fontSize}px`;
      textarea.placeholder = p === 1 ? '（請從本頁第 1 行依序開始登記作答內文...）' : '';
      textarea.value = answers[currentQ][p] || '';

      textarea.addEventListener('focus', function () {
        activeTextarea = textarea;
        currentPage = p;
        if (selectPage) selectPage.value = p;
        updateToolbarActiveStates();
      });

      textarea.addEventListener('keyup', updateToolbarActiveStates);
      textarea.addEventListener('click', updateToolbarActiveStates);

      // Keydown Listener for Enter Key Auto-Number Continuation & Soft Break
      textarea.addEventListener('keydown', function (e) {
        if (e.key === 'Backspace' && textarea.selectionStart === textarea.selectionEnd) {
          if (textarea.selectionStart === 0 && p > 1) {
            e.preventDefault();
            const prevPage = document.getElementById(`paper-page-${p - 1}`);
            if (prevPage) {
              const prevTextarea = prevPage.querySelector('.paper-textarea');
              prevTextarea.focus();
              prevTextarea.selectionStart = prevTextarea.selectionEnd = prevTextarea.value.length;
            }
            return;
          }

          const details = getActiveLineDetails();
          if (details) {
            const { currentLineText, cursorOffsetInLine } = details;
            const tag = parseLineNumberTag(currentLineText);
            
            if (tag && cursorOffsetInLine === tag.rawMatch.length) {
              e.preventDefault();
              const shift = -tag.rawMatch.length;
              const newLineText = currentLineText.substring(tag.rawMatch.length);
              applyLineChange(details, newLineText, shift);
              return;
            }
          }
        }

        if (e.key === 'Enter') {
          const details = getActiveLineDetails();
          if (!details) return;

          const { lines, lineIndex, currentLineText } = details;
          const tag = parseLineNumberTag(currentLineText);

          if (!e.shiftKey) {
            // Normal Enter: Auto-generate next tag
            if (tag) {
              e.preventDefault();
              const nextTag = generateNumberTag(tag.level, tag.num + 1);

              const startPos = textarea.selectionStart;
              const val = textarea.value;

              const beforeCursor = val.substring(0, startPos);
              const afterCursor = val.substring(startPos);

              const insertedText = '\n' + nextTag;
              textarea.value = beforeCursor + insertedText + afterCursor;

              const newCursorPos = startPos + insertedText.length;
              textarea.selectionStart = textarea.selectionEnd = newCursorPos;

              resequenceDocumentNumbers();
              textarea.dispatchEvent(new Event('input'));
              updateToolbarActiveStates();
            }
          } else {
            // Shift + Enter: Align to the current line's text start position (Soft break)
            e.preventDefault();
            const startPos = textarea.selectionStart;
            const val = textarea.value;

            const beforeCursor = val.substring(0, startPos);
            const afterCursor = val.substring(startPos);

            let indentStr = '';
            if (tag) {
              // Convert tag characters to spaces (full-width for Chinese/Full-width, half-width for numbers/English)
              for (let i = 0; i < tag.rawMatch.length; i++) {
                const char = tag.rawMatch[i];
                if (char === '　' || char === ' ') {
                  indentStr += char;
                } else if (/[一二三四五六七八九十、（）]/.test(char)) {
                  indentStr += '　';
                } else {
                  indentStr += ' '; // half-width space
                }
              }
            } else {
              // No tag, just match the existing leading spaces of the current line
              const match = currentLineText.match(/^([　\s]+)/);
              if (match) {
                indentStr = match[1];
              }
            }

            const insertedText = '\n' + indentStr;
            textarea.value = beforeCursor + insertedText + afterCursor;

            const newCursorPos = startPos + insertedText.length;
            textarea.selectionStart = textarea.selectionEnd = newCursorPos;
            textarea.dispatchEvent(new Event('input'));
          }
        }
      });

      textarea.addEventListener('input', function (e) {
        // Handle Auto-Jump to Next Page when full
        if (textarea.clientHeight > 0 && textarea.scrollHeight > textarea.clientHeight && p < 8) {
          const originalCursor = textarea.selectionStart;
          let overflowText = '';
          
          let didOverflow = false;
          while (textarea.scrollHeight > textarea.clientHeight && textarea.value.length > 0) {
            overflowText = textarea.value.slice(-1) + overflowText;
            textarea.value = textarea.value.slice(0, -1);
            didOverflow = true;
          }
          
          if (didOverflow) {
            const cursorOverflowed = originalCursor > textarea.value.length;
            
            let removedNewlines = 0;
            while (overflowText.startsWith('\n')) {
              overflowText = overflowText.substring(1);
              removedNewlines++;
            }
            let adjustedCursor = originalCursor;
            if (cursorOverflowed) {
              adjustedCursor -= removedNewlines;
            }

            const nextPage = document.getElementById(`paper-page-${p + 1}`);
            if (nextPage) {
              const nextPageTextarea = nextPage.querySelector('.paper-textarea');
              nextPageTextarea.value = overflowText + nextPageTextarea.value;
              answers[currentQ][p] = textarea.value;
              answers[currentQ][p + 1] = nextPageTextarea.value;
              
              if (cursorOverflowed) {
                // Cursor overflowed to the next page
                nextPageTextarea.focus();
                const newCursor = adjustedCursor - textarea.value.length;
                nextPageTextarea.selectionStart = nextPageTextarea.selectionEnd = Math.max(0, newCursor);
              } else {
                // Cursor is still on this page, restore it
                textarea.selectionStart = textarea.selectionEnd = adjustedCursor;
              }
              // Cascade input event to next page if it also overflowed
              nextPageTextarea.dispatchEvent(new Event('input'));
            }
          }
        }

        answers[currentQ][p] = textarea.value;
        resequenceDocumentNumbers();
        updateStats();

        // Debounce Auto-Save
        clearTimeout(autoSaveTimer);
        updateAutoSaveStatus('saving');
        autoSaveTimer = setTimeout(() => {
          saveCurrentAnswer(false);
          addHistorySnapshot('自動儲存');
        }, 2500);
      });

      centerArea.appendChild(textarea);

      const rightMargin = document.createElement('div');
      rightMargin.className = 'paper-right-margin';
      rightMargin.textContent = '（請從本頁第 1 行依序開始登記）';

      page.appendChild(leftMargin);
      page.appendChild(centerArea);
      page.appendChild(rightMargin);

      paperPagesContainer.appendChild(page);
    }

    // Set active textarea to page 1 by default
    const firstTextarea = paperPagesContainer.querySelector('.paper-textarea');
    if (firstTextarea) activeTextarea = firstTextarea;

    questionTitle.textContent = QUESTIONS[currentQ].title;
    questionContent.textContent = QUESTIONS[currentQ].content;
    selectQuestion.value = currentQ;
    if (selectPage) selectPage.value = 1;
    statCurrentQ.textContent = `第 ${currentQ} 題`;
  }

  // ==========================================================================
  // MOEX Legal Auto-Numbering Engine (一、, （一）, 1., (1))
  // ==========================================================================
  const CHINESE_NUMS = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十',
                        '十一', '十二', '十三', '十四', '十五', '十六', '十七', '十八', '十九', '二十'];

  function chineseToNum(str) {
    const idx = CHINESE_NUMS.indexOf(str);
    if (idx !== -1) return idx + 1;
    return parseInt(str, 10) || 1;
  }

  function numToChinese(num) {
    return CHINESE_NUMS[num - 1] || String(num);
  }

  function parseLineNumberTag(lineText) {
    // Level 1: 一、 二、 三、
    let match = lineText.match(/^(\s*)([一二三四五六七八九十]+)、/);
    if (match) {
      return { level: 1, num: chineseToNum(match[2]), rawMatch: match[0], indent: match[1] };
    }
    // Level 2: （一） （二）
    match = lineText.match(/^(\s*)（([一二三四五六七八九十]+)）/);
    if (match) {
      return { level: 2, num: chineseToNum(match[2]), rawMatch: match[0], indent: match[1] };
    }
    // Level 3: 1. 2. 3.
    match = lineText.match(/^(\s*)(\d+)\./);
    if (match) {
      return { level: 3, num: parseInt(match[2], 10), rawMatch: match[0], indent: match[1] };
    }
    // Level 4: (1) (2)
    match = lineText.match(/^(\s*)\((\d+)\)/);
    if (match) {
      return { level: 4, num: parseInt(match[2], 10), rawMatch: match[0], indent: match[1] };
    }
    return null;
  }

  function generateNumberTag(level, num) {
    switch (level) {
      case 1: return `${numToChinese(num)}、`;
      case 2: return `　　（${numToChinese(num)}）`;
      case 3: return `　　　　${num}.`;
      case 4: return `　　　　　　(${num})`;
      default: return `${numToChinese(num)}、`;
    }
  }

  function getActiveLineDetails() {
    if (!activeTextarea) {
      activeTextarea = paperPagesContainer.querySelector('.paper-textarea');
    }
    if (!activeTextarea) return null;

    const val = activeTextarea.value;
    const cursorPos = activeTextarea.selectionStart;

    const lines = val.split('\n');
    let currentOffset = 0;
    let lineIndex = 0;
    let lineStartOffset = 0;

    for (let i = 0; i < lines.length; i++) {
      const lineLen = lines[i].length;
      if (cursorPos >= currentOffset && cursorPos <= currentOffset + lineLen) {
        lineIndex = i;
        lineStartOffset = currentOffset;
        break;
      }
      currentOffset += lineLen + 1; // +1 for newline
    }

    return {
      textarea: activeTextarea,
      lines: lines,
      lineIndex: lineIndex,
      currentLineText: lines[lineIndex] || '',
      lineStartOffset: lineStartOffset,
      cursorOffsetInLine: cursorPos - lineStartOffset
    };
  }

  function resequenceDocumentNumbers() {
    let levelCounts = { 1: 0, 2: 0, 3: 0, 4: 0 };
    
    for (let p = 1; p <= 8; p++) {
      const page = document.getElementById(`paper-page-${p}`);
      if (!page) continue;
      const textarea = page.querySelector('.paper-textarea');
      if (!textarea) continue;
      
      const val = textarea.value;
      const cursorPos = textarea.selectionStart;
      const lines = val.split('\n');
      
      let changed = false;
      let newCursorPos = cursorPos;
      let currentOffset = 0;
      
      for (let i = 0; i < lines.length; i++) {
        const lineText = lines[i];
        const tag = parseLineNumberTag(lineText);
        
        if (tag) {
          levelCounts[tag.level]++;
          
          // Reset lower sub-levels
          for (let l = tag.level + 1; l <= 4; l++) {
            levelCounts[l] = 0;
          }
          
          const expectedNum = levelCounts[tag.level];
          const expectedTag = generateNumberTag(tag.level, expectedNum);
          
          if (tag.rawMatch !== expectedTag) {
            const textAfterTag = lineText.substring(tag.rawMatch.length);
            const newLineText = expectedTag + textAfterTag;
            const lenDiff = expectedTag.length - tag.rawMatch.length;
            
            lines[i] = newLineText;
            changed = true;
            
            if (cursorPos > currentOffset) {
              newCursorPos += lenDiff;
            }
          }
        }
        currentOffset += lines[i].length + 1;
      }
      
      if (changed) {
        textarea.value = lines.join('\n');
        if (textarea === activeTextarea) {
          textarea.selectionStart = textarea.selectionEnd = Math.max(0, newCursorPos);
        }
        answers[currentQ][p] = textarea.value;
      }
    }
  }

  // ==========================================================================
  // Single-Step Undo/Redo Engine (復原與取消復原僅限 1 個步驟)
  // ==========================================================================
  let singleUndoStep = null;
  let singleRedoStep = null;

  function saveUndoSnapshot() {
    if (!activeTextarea && paperPagesContainer) {
      activeTextarea = paperPagesContainer.querySelector('.paper-textarea');
    }
    if (!activeTextarea) return;

    singleUndoStep = {
      textarea: activeTextarea,
      val: activeTextarea.value,
      start: activeTextarea.selectionStart,
      end: activeTextarea.selectionEnd
    };
    singleRedoStep = null;
  }

  function performUndo() {
    if (!activeTextarea && paperPagesContainer) {
      activeTextarea = paperPagesContainer.querySelector('.paper-textarea');
    }
    if (!activeTextarea) return;

    if (!singleUndoStep) {
      document.execCommand('undo');
      return;
    }

    singleRedoStep = {
      textarea: activeTextarea,
      val: activeTextarea.value,
      start: activeTextarea.selectionStart,
      end: activeTextarea.selectionEnd
    };

    const target = singleUndoStep;
    singleUndoStep = null;

    activeTextarea = target.textarea;
    activeTextarea.focus();
    activeTextarea.value = target.val;
    activeTextarea.selectionStart = target.start;
    activeTextarea.selectionEnd = target.end;

    resequenceDocumentNumbers();
    activeTextarea.dispatchEvent(new Event('input'));
    updateToolbarActiveStates();
  }

  function performRedo() {
    if (!activeTextarea && paperPagesContainer) {
      activeTextarea = paperPagesContainer.querySelector('.paper-textarea');
    }
    if (!activeTextarea) return;

    if (!singleRedoStep) {
      document.execCommand('redo');
      return;
    }

    singleUndoStep = {
      textarea: activeTextarea,
      val: activeTextarea.value,
      start: activeTextarea.selectionStart,
      end: activeTextarea.selectionEnd
    };

    const target = singleRedoStep;
    singleRedoStep = null;

    activeTextarea = target.textarea;
    activeTextarea.focus();
    activeTextarea.value = target.val;
    activeTextarea.selectionStart = target.start;
    activeTextarea.selectionEnd = target.end;

    resequenceDocumentNumbers();
    activeTextarea.dispatchEvent(new Event('input'));
    updateToolbarActiveStates();
  }

  function applyLineChange(details, newLineText, relativeCursorShift = 0) {
    saveUndoSnapshot();
    const { textarea, lines, lineIndex, lineStartOffset, cursorOffsetInLine } = details;
    
    lines[lineIndex] = newLineText;
    textarea.value = lines.join('\n');

    // Restore focus and calculate exact new cursor position
    textarea.focus();
    const newCursorPos = Math.max(0, lineStartOffset + cursorOffsetInLine + relativeCursorShift);
    textarea.selectionStart = textarea.selectionEnd = newCursorPos;

    // Resequence all list numbers in this paper page
    resequenceDocumentNumbers();

    textarea.dispatchEvent(new Event('input'));
    updateToolbarActiveStates();
  }

  function toggleNumberedList() {
    const details = getActiveLineDetails();
    if (!details) return;

    const { lines, lineIndex, currentLineText } = details;
    const existingTag = parseLineNumberTag(currentLineText);

    if (existingTag) {
      // Toggle Off: Remove tag
      const newLineText = currentLineText.replace(existingTag.rawMatch, '').trimStart();
      const shift = -existingTag.rawMatch.length;
      applyLineChange(details, newLineText, shift);
    } else {
      // Toggle On: Find previous line's level or default to Level 1
      let targetLevel = 1;
      let targetNum = 1;

      for (let i = lineIndex - 1; i >= 0; i--) {
        const prevTag = parseLineNumberTag(lines[i]);
        if (prevTag) {
          targetLevel = prevTag.level;
          targetNum = prevTag.num + 1;
          break;
        }
      }

      const newTag = generateNumberTag(targetLevel, targetNum);
      const newLineText = newTag + currentLineText;
      const shift = newTag.length;
      applyLineChange(details, newLineText, shift);
    }
  }

  function changeHierarchyLevel(direction) {
    const details = getActiveLineDetails();
    if (!details) return;

    const { lines, lineIndex, currentLineText } = details;
    const existingTag = parseLineNumberTag(currentLineText);

    if (!existingTag) {
      toggleNumberedList();
      return;
    }

    let newLevel = existingTag.level + direction;
    if (newLevel < 1) newLevel = 1;
    if (newLevel > 4) newLevel = 4;

    const newTag = generateNumberTag(newLevel, 1);
    const newLineText = currentLineText.replace(existingTag.rawMatch, newTag);
    const shift = newTag.length - existingTag.rawMatch.length;

    applyLineChange(details, newLineText, shift);
  }

  function updateToolbarActiveStates() {
    const toolNumbering = document.getElementById('toolNumbering');
    if (!toolNumbering) return;

    const details = getActiveLineDetails();
    if (details && parseLineNumberTag(details.currentLineText)) {
      toolNumbering.classList.add('active');
    } else {
      toolNumbering.classList.remove('active');
    }
  }

  // Symbol & Tool Item Toolbar Actions
  function setupSymbolToolbar() {
    const symbolItems = document.querySelectorAll('.sym-item');
    symbolItems.forEach(item => {
      if (item.hasAttribute('data-symbol')) {
        item.addEventListener('click', function () {
          const symbol = item.getAttribute('data-symbol');
          insertAtCursor(symbol);
        });
      }
    });

    const toolCut = document.getElementById('toolCut');
    if (toolCut) {
      toolCut.addEventListener('click', () => {
        saveUndoSnapshot();
        if (activeTextarea) activeTextarea.focus();
        document.execCommand('cut');
      });
    }

    const toolCopy = document.getElementById('toolCopy');
    if (toolCopy) {
      toolCopy.addEventListener('click', () => {
        if (activeTextarea) activeTextarea.focus();
        document.execCommand('copy');
      });
    }

    const toolPaste = document.getElementById('toolPaste');
    if (toolPaste) {
      toolPaste.addEventListener('click', async () => {
        if (!activeTextarea) return;
        saveUndoSnapshot();
        activeTextarea.focus();
        try {
          const clipText = await navigator.clipboard.readText();
          if (clipText) insertAtCursor(clipText);
        } catch (e) {
          document.execCommand('paste');
        }
      });
    }

    const toolUndo = document.getElementById('toolUndo');
    if (toolUndo) {
      toolUndo.addEventListener('click', () => {
        performUndo();
      });
    }

    const toolRedo = document.getElementById('toolRedo');
    if (toolRedo) {
      toolRedo.addEventListener('click', () => {
        performRedo();
      });
    }

    const toolNumbering = document.getElementById('toolNumbering');
    if (toolNumbering) {
      toolNumbering.addEventListener('click', () => {
        toggleNumberedList();
      });
    }

    const toolPrevLevel = document.getElementById('toolPrevLevel');
    if (toolPrevLevel) {
      toolPrevLevel.addEventListener('click', () => {
        changeHierarchyLevel(-1); // 上一層 (Promote)
      });
    }

    const toolNextLevel = document.getElementById('toolNextLevel');
    if (toolNextLevel) {
      toolNextLevel.addEventListener('click', () => {
        changeHierarchyLevel(1); // 下一層 (Demote)
      });
    }
  }

  function insertAtCursor(text) {
    if (!activeTextarea) {
      activeTextarea = paperPagesContainer.querySelector('.paper-textarea');
    }
    if (!activeTextarea) return;

    saveUndoSnapshot();
    activeTextarea.focus();
    const start = activeTextarea.selectionStart;
    const end = activeTextarea.selectionEnd;
    const val = activeTextarea.value;

    activeTextarea.value = val.substring(0, start) + text + val.substring(end);
    
    activeTextarea.selectionStart = activeTextarea.selectionEnd = start + text.length;

    // Trigger input event logic
    activeTextarea.dispatchEvent(new Event('input'));
  }

  // Update Statistics
  function updateStats() {
    let totalChars = 0;
    let totalLines = 0;
    let maxFilledPage = 1;

    for (let p = 1; p <= 8; p++) {
      const pText = answers[currentQ][p] || '';
      if (pText.trim().length > 0) {
        maxFilledPage = p;
      }
      totalChars += pText.length;
      if (pText.length > 0) {
        totalLines += pText.split('\n').length;
      }
    }

    statTotalChars.textContent = totalChars;
    statTotalLines.textContent = totalLines;
    statTotalPages.textContent = maxFilledPage;

    let answered = 0;
    [1, 2].forEach(q => {
      let qHasText = false;
      for (let p = 1; p <= 8; p++) {
        if ((answers[q][p] || '').trim().length > 0) {
          qHasText = true;
          break;
        }
      }
      if (qHasText) answered++;
    });

    answeredCountEl.textContent = answered;
    unansweredCountEl.textContent = 2 - answered;
  }

  // Navigation
  function switchQuestion(newQ) {
    if (newQ < 1 || newQ > 2) return;
    saveCurrentAnswer(false);
    currentQ = newQ;
    currentPage = 1;
    renderEditor();
    updateStats();
  }

  function scrollToPage(pageNum) {
    const pageEl = document.getElementById(`paper-page-${pageNum}`);
    if (pageEl) {
      pageEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
      const ta = pageEl.querySelector('.paper-textarea');
      if (ta) ta.focus();
    }
  }

  btnPrevQuestion.addEventListener('click', () => switchQuestion(currentQ - 1));
  btnNextQuestion.addEventListener('click', () => switchQuestion(currentQ + 1));
  selectQuestion.addEventListener('change', (e) => switchQuestion(parseInt(e.target.value)));
  if (selectPage) {
    selectPage.addEventListener('change', (e) => scrollToPage(parseInt(e.target.value)));
  }

  btnManualSave.addEventListener('click', () => {
    saveCurrentAnswer(true);
    alert('【已手動暫存】成功保存目前試卷內文與版本快照！');
  });

  // Toggle Question Visibility
  btnToggleQuestion.addEventListener('click', () => {
    if (questionBody.style.display === 'none') {
      questionBody.style.display = 'block';
      btnToggleQuestion.textContent = '收合試題 ▲';
    } else {
      questionBody.style.display = 'none';
      btnToggleQuestion.textContent = '展開試題 ▼';
    }
  });

  // Countdown Timer
  function startTimer() {
    setInterval(() => {
      if (secondsRemaining <= 0) return;
      secondsRemaining--;
      const hrs = String(Math.floor(secondsRemaining / 3600)).padStart(2, '0');
      const mins = String(Math.floor((secondsRemaining % 3600) / 60)).padStart(2, '0');
      const secs = String(secondsRemaining % 60).padStart(2, '0');
      countdownTimerEl.textContent = `${hrs}:${mins}:${secs}`;
    }, 1000);
  }

  const btnBackToEditor = document.getElementById('btnBackToEditor');
  let currentHistoryFilter = 'all';

  // History Drawer Logic
  btnOpenHistory.addEventListener('click', () => {
    historyOverlay.style.display = 'block';
    historyDrawer.classList.add('open');
    currentHistoryFilter = 'all';
    setupHistoryFilterTabs();
    loadHistoryDrawer(currentHistoryFilter);
  });

  function closeHistoryDrawer() {
    historyOverlay.style.display = 'none';
    historyDrawer.classList.remove('open');
    if (activeTextarea) {
      activeTextarea.focus();
    }
  }

  btnCloseHistory.addEventListener('click', closeHistoryDrawer);
  historyOverlay.addEventListener('click', closeHistoryDrawer);
  if (btnBackToEditor) {
    btnBackToEditor.addEventListener('click', closeHistoryDrawer);
  }

  function setupHistoryFilterTabs() {
    const filterBtns = historyDrawer.querySelectorAll('.history-tab-btn');
    filterBtns.forEach(btn => {
      const f = btn.getAttribute('data-filter');
      if (f === currentHistoryFilter) {
        btn.classList.add('active');
      } else {
        btn.classList.remove('active');
      }
      btn.onclick = () => {
        filterBtns.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        currentHistoryFilter = f;
        loadHistoryDrawer(currentHistoryFilter);
      };
    });
  }

  function loadHistoryDrawer(filter = 'all') {
    fetchHistoryList((items) => {
      historyListEl.innerHTML = '';

      let filtered = items;
      if (filter === 'manual') {
        filtered = items.filter(i => i.type === '手動暫存');
      } else if (filter === 'auto') {
        filtered = items.filter(i => i.type === '自動儲存');
      }

      if (filtered.length === 0) {
        const filterName = filter === 'manual' ? '手動暫存' : filter === 'auto' ? '自動儲存' : '';
        const emptyNotice = document.createElement('div');
        emptyNotice.style.cssText = 'padding: 24px 14px; text-align: center; color: #718096; line-height: 1.6;';
        emptyNotice.innerHTML = `
          <div style="font-size: 13px; font-weight: bold; margin-bottom: 8px; color: #4a5568;">尚無「${filterName}」紀錄</div>
          <div style="font-size: 11.5px; margin-bottom: 14px; color: #718096;">${filter === 'manual' ? '請點擊選單列「【暫存】」進行手動備份' : '系統會在您打字時自動進行記錄'}</div>
          <button class="btn-small-blue" id="btnSwitchAllHistory">📋 檢視全部紀錄 (${items.length})</button>
        `;
        historyListEl.appendChild(emptyNotice);

        const btnSwitchAllHistory = document.getElementById('btnSwitchAllHistory');
        if (btnSwitchAllHistory) {
          btnSwitchAllHistory.addEventListener('click', () => {
            currentHistoryFilter = 'all';
            setupHistoryFilterTabs();
            loadHistoryDrawer('all');
          });
        }

        if (previewTimestamp) previewTimestamp.textContent = '未選擇版本';
        if (previewContent) previewContent.textContent = `當前篩選類別（${filterName}）暫無可用紀錄。請點擊左側「檢視全部紀錄」按鈕以查看歷程。`;
        if (btnRestoreVersion) btnRestoreVersion.disabled = true;
        return;
      }

      filtered.forEach((item, index) => {
        const div = document.createElement('div');
        div.className = `history-item ${index === 0 ? 'active' : ''}`;
        div.title = '雙擊直接載入此版本並回到作答區';
        
        const typeClass = item.type === '手動暫存' ? 'manual' : 'auto';
        div.innerHTML = `
          <span class="item-type ${typeClass}">${item.type}</span>
          <div class="item-time">${item.timestamp}</div>
          <div class="item-chars">全卷共 ${item.charCount || 0} 字</div>
        `;

        div.addEventListener('click', () => {
          document.querySelectorAll('.history-item').forEach(el => el.classList.remove('active'));
          div.classList.add('active');
          selectedHistoryItem = item;
          showPreview(item);
        });

        // Double click to restore & immediately return to answer sheet!
        div.addEventListener('dblclick', () => {
          selectedHistoryItem = item;
          restoreHistoryItem(item);
        });

        historyListEl.appendChild(div);
      });

      if (filtered.length > 0) {
        selectedHistoryItem = filtered[0];
        showPreview(filtered[0]);
      }
    });
  }

  function restoreHistoryItem(item) {
    if (!item) return;
    if (confirm(`確定要載入 ${item.timestamp} (${item.type}) 的版本並回到作答區嗎？現有未儲存變更將會被覆蓋。`)) {
      if (item.answers) {
        answers[1] = normalizeAnswerData(item.answers[1]);
        answers[2] = normalizeAnswerData(item.answers[2]);
      }
      saveCurrentAnswer(false);
      renderEditor();
      updateStats();
      closeHistoryDrawer();
      scrollToPage(currentPage || 1);
      if (activeTextarea) activeTextarea.focus();
    }
  }

  if (previewQTabs) {
    const pTabs = previewQTabs.querySelectorAll('.history-tab-btn');
    pTabs.forEach(btn => {
      btn.addEventListener('click', () => {
        pTabs.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        previewCurrentQ = parseInt(btn.getAttribute('data-preview-q'));
        if (selectedHistoryItem) {
          showPreview(selectedHistoryItem);
        }
      });
    });
  }

  function showPreview(item) {
    if (!item) {
      if (previewTimestamp) previewTimestamp.textContent = '未選擇版本';
      if (previewContent) previewContent.textContent = '請點選左側歷史紀錄項目以檢視詳細內容。';
      if (previewQTabs) previewQTabs.style.display = 'none';
      if (btnRestoreVersion) btnRestoreVersion.disabled = true;
      return;
    }

    if (previewQTabs) {
      previewQTabs.style.display = 'flex';
      const pTabs = previewQTabs.querySelectorAll('.history-tab-btn');
      pTabs.forEach(b => {
        if (parseInt(b.getAttribute('data-preview-q')) === previewCurrentQ) {
          b.classList.add('active');
        } else {
          b.classList.remove('active');
        }
      });
    }

    if (previewTimestamp) {
      previewTimestamp.textContent = `📅 ${item.timestamp || ''} (${item.type || '歷史紀錄'}) | 全卷字數：${item.charCount || 0} 字`;
    }

    let targetData = null;

    if (item.answers) {
      targetData = item.answers[previewCurrentQ] || item.answers[String(previewCurrentQ)];
    } else if (item.pages) {
      if (item.qId === previewCurrentQ) {
        targetData = item.pages;
      }
    }

    const itemText = getPreviewTextForQuestion(targetData);

    let contentStr = '';
    contentStr += `========================================\n  【第 ${previewCurrentQ} 題作答內容】\n========================================\n${itemText || '(本題尚無登記內容)'}`;

    if (previewContent) {
      previewContent.textContent = contentStr;
    }

    if (btnRestoreVersion) {
      btnRestoreVersion.disabled = false;
    }
  }

  function getPreviewTextForQuestion(qData) {
    if (!qData) return '';
    if (typeof qData === 'string') return qData.trim();

    const pages = qData.pages || qData;
    let parts = [];

    for (let p = 1; p <= 8; p++) {
      const pageText = (pages[p] || pages[String(p)] || '').trim();
      if (pageText.length > 0) {
        parts.push(`📄 【第 ${p} 頁】\n${pageText}`);
      }
    }
    return parts.join('\n\n');
  }

  btnRestoreVersion.addEventListener('click', () => {
    if (selectedHistoryItem) {
      restoreHistoryItem(selectedHistoryItem);
    }
  });

  btnCreateSnapshot.addEventListener('click', () => {
    saveCurrentAnswer(true);
    loadHistoryDrawer(currentHistoryFilter);
  });

  btnClearHistory.addEventListener('click', () => {
    if (confirm('確定要清除所有歷史版本紀錄嗎？（此動作不可復原）')) {
      if (db && db.objectStoreNames.contains('history')) {
        try {
          const tx = db.transaction(['history'], 'readwrite');
          tx.objectStore('history').clear();
        } catch (e) {
          console.warn('IndexedDB clear history exception:', e);
        }
      }
      try {
        localStorage.removeItem('MOEX_CBT_HISTORY');
      } catch (e) {}
      historyList = [];
      loadHistoryDrawer(currentHistoryFilter);
    }
  });

  // Export to Word (.docx)
  btnExportWord.addEventListener('click', async function () {
    if (typeof docx === 'undefined') {
      alert('docx 模組載入中，請稍後再試');
      return;
    }

    const { Document, Packer, Paragraph, TextRun, HeadingLevel } = docx;

    const children = [];

    [1, 2].forEach(q => {
      let qHasText = false;
      const qParagraphs = [];
      
      qParagraphs.push(
        new Paragraph({
          children: [
            new TextRun({
              text: `【第 ${q} 題】`,
              bold: true
            })
          ],
          outlineLevel: 0,
          spacing: { before: 200, after: 100 }
        })
      );

      for (let p = 1; p <= 8; p++) {
        const pText = (answers[q][p] || '').trimEnd();
        if (pText.length > 0) {
          qHasText = true;
          const lines = pText.split('\n');
          
          lines.forEach(line => {
            const trimmedLine = line.trimEnd();
            if (trimmedLine.length === 0) {
              qParagraphs.push(new Paragraph({ children: [] }));
              return;
            }
            
            const preservedLine = trimmedLine.replace(/ /g, '\u00A0');
            const tag = parseLineNumberTag(trimmedLine);
            
            const pOptions = {
              children: [
                new TextRun({ text: preservedLine })
              ]
            };
            
            if (tag) {
              pOptions.outlineLevel = tag.level; // level 1-4 becomes outlineLevel 1-4
            }
            
            qParagraphs.push(new Paragraph(pOptions));
          });
        }
      }

      if (qHasText) {
        children.push(...qParagraphs);
      } else {
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: `【第 ${q} 題】`,
                bold: true
              })
            ],
            outlineLevel: 0,
            spacing: { before: 200, after: 100 }
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: '(本題未作答)'
              })
            ]
          })
        );
      }
    });

    const doc = new Document({
      styles: {
        default: {
          document: {
            run: {
              size: 24, // 12pt = 24 half-points
              font: "新細明體"
            },
            paragraph: {
              spacing: { after: 0 }
            }
          }
        }
      },
      sections: [{
        properties: {},
        children: children
      }]
    });

    try {
      const blob = await Packer.toBlob(doc);
      if (window.showSaveFilePicker) {
        try {
          const handle = await window.showSaveFilePicker({
            suggestedName: '模擬作答_10100001.docx',
            types: [{
              description: 'Word Document',
              accept: { 'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'] }
            }]
          });
          const writable = await handle.createWritable();
          await writable.write(blob);
          await writable.close();
          alert('匯出成功！檔案已儲存。');
        } catch (pickerErr) {
          if (pickerErr.name !== 'AbortError') {
            console.error('SaveFilePicker error:', pickerErr);
            fallbackDownload(blob);
          }
        }
      } else {
        fallbackDownload(blob);
      }
    } catch (e) {
      console.error('Word export error:', e);
      alert('匯出時發生錯誤');
    }

    function fallbackDownload(blob) {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = '模擬作答_10100001.docx';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }
  });

  // Modals Controller
  function setupModals() {
    const btnGuide = document.getElementById('btnGuide');
    const btnNotice = document.getElementById('btnNotice');
    const btnFinishExam = document.getElementById('btnFinishExam');

    const modalGuide = document.getElementById('modalGuide');
    const modalNotice = document.getElementById('modalNotice');

    btnGuide.addEventListener('click', () => modalGuide.style.display = 'flex');
    btnNotice.addEventListener('click', () => modalNotice.style.display = 'flex');

    document.querySelectorAll('.modal-close, .modal-close-btn').forEach(btn => {
      btn.addEventListener('click', function () {
        this.closest('.modal-backdrop').style.display = 'none';
      });
    });

    btnFinishExam.addEventListener('click', () => {
      let count = 0;
      [1, 2].forEach(q => {
        for (let p = 1; p <= 8; p++) {
          if ((answers[q][p] || '').trim().length > 0) {
            count++;
            break;
          }
        }
      });
      if (confirm(`您目前已完成 ${count} / 2 題。確定要結束作答並交卷嗎？`)) {
        alert('【成功交卷】感謝使用考選部 CBT 線上模擬作答系統！您可以點擊「匯出 Word 檔」進行備份。');
      }
    });
  }

  // Single Question Preview Modal Controller (預覽頁面：只能看一題)
  function setupPreviewModal() {
    const btnPreviewPages = document.getElementById('btnPreviewPages');
    const modalPreviewPages = document.getElementById('modalPreviewPages');
    const pagesGrid = document.getElementById('pagesGrid');
    const previewModalTitle = document.getElementById('previewModalTitle');
    const previewSummaryQTitle = document.getElementById('previewSummaryQTitle');
    const previewSummaryAnswered = document.getElementById('previewSummaryAnswered');
    const previewSummaryChars = document.getElementById('previewSummaryChars');

    if (!btnPreviewPages || !modalPreviewPages) return;

    btnPreviewPages.addEventListener('click', () => {
      renderPagePreviews();
      modalPreviewPages.style.display = 'flex';
    });

    function renderPagePreviews() {
      pagesGrid.innerHTML = '';

      const qTitleText = currentQ === 1 ? '第一題' : '第二題';
      if (previewModalTitle) {
        previewModalTitle.textContent = `📖 ${QUESTIONS[currentQ].title} (8 頁) 試卷頁面預覽`;
      }
      if (previewSummaryQTitle) {
        previewSummaryQTitle.textContent = qTitleText;
      }

      let answeredPagesCount = 0;
      let totalQChars = 0;

      for (let p = 1; p <= 8; p++) {
        const text = (answers[currentQ][p] || '').trim();
        if (text.length > 0) {
          answeredPagesCount++;
          totalQChars += (answers[currentQ][p] || '').length;
        }
      }

      if (previewSummaryAnswered) previewSummaryAnswered.textContent = answeredPagesCount;
      if (previewSummaryChars) previewSummaryChars.textContent = totalQChars;

      // Render ONLY the 8 pages of currentQ!
      for (let p = 1; p <= 8; p++) {
        const rawText = answers[currentQ][p] || '';
        const trimmedText = rawText.trim();
        const charCount = rawText.length;
        const lineCount = rawText ? rawText.split('\n').length : 0;
        const hasContent = trimmedText.length > 0;

        const card = document.createElement('div');
        card.className = `page-card ${hasContent ? 'has-content' : ''}`;

        const header = document.createElement('div');
        header.className = 'page-card-header';
        header.innerHTML = `
          <span>第 ${currentQ} 題 - 第 ${p} 頁</span>
          <span class="page-badge">${hasContent ? `${charCount} 字` : '空白'}</span>
        `;

        const body = document.createElement('div');
        body.className = 'page-card-body';
        if (hasContent) {
          body.textContent = rawText;
        } else {
          body.innerHTML = '<div class="empty-page-notice">（點擊編輯此頁）</div>';
        }

        const footer = document.createElement('div');
        footer.className = 'page-card-footer';
        footer.innerHTML = `
          <span>字數: ${charCount}</span>
          <span>行數: ${lineCount}</span>
        `;

        card.appendChild(header);
        card.appendChild(body);
        card.appendChild(footer);

        card.addEventListener('click', () => {
          modalPreviewPages.style.display = 'none';
          scrollToPage(p);
        });

        pagesGrid.appendChild(card);
      }
    }
  }

  // App Startup
  window.addEventListener('DOMContentLoaded', async () => {
    await initIndexedDB();
    await loadState();
    setupSymbolToolbar();
    setupModals();
    setupPreviewModal();
    startTimer();
  });

})();
