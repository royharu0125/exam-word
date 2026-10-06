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
      content: '若本國僅與 A、B 兩國進行貿易，且本國與兩國的貿易量均相同。若 A 與 B 國的直接匯率為 31.64 與 0.36，請計算本國的名目有效匯率 (nominal effective exchange rate)。請詳述推導過程與法理依據。（完整試題請見紙本試題冊）'
    },
    2: {
      title: '第二題 (50.00分)',
      content: '請依相關法規及學說實務見解，詳述本題爭點及法律效果。（完整試題請見紙本試題冊）'
    },
    3: {
      title: '第三題 (50.00分)',
      content: '請依相關法規及學說實務見解，詳述本題爭點及法律效果。（完整試題請見紙本試題冊）'
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
    2: createEmptyQuestionPages(),
    3: createEmptyQuestionPages()
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
  let secondsRemaining = 3 * 3600 - 1; // 3 hours

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
      2: createEmptyQuestionPages(),
      3: createEmptyQuestionPages()
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
      Object.keys(QUESTIONS).forEach(qStr => {
        const q = Number(qStr);
        store.put({ qId: q, pages: answers[q] });
      });
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
    Object.keys(QUESTIONS).forEach(qStr => {
      const q = Number(qStr);
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

  // ==========================================================================
  // ContentEditable Auto Hanging Indent Helpers
  // ==========================================================================
  function getTagIndentWidth(tagStr) {
    let width = 0;
    for (let char of tagStr) {
      if (/[a-zA-Z0-9.\-() ]/.test(char) && char !== '（' && char !== '）' && char !== '　') {
        width += (fontSize / 2) + 2;
      } else {
        width += fontSize + 2;
      }
    }
    return width;
  }

  function escapeHTML(str) {
    return str.replace(/[&<>'"]/g, tag => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;'
    }[tag] || tag));
  }

  function getInheritedIndent(q, p) {
    if (p <= 1) return 0;
    let currentP = p - 1;
    while (currentP >= 1) {
      const pText = answers[q][currentP] || '';
      if (!pText) break;
      const pLines = pText.split('\n');
      const lastLine = pLines[pLines.length - 1];
      
      const tag = parseLineNumberTag(lastLine);
      if (tag) {
        return getTagIndentWidth(tag.rawMatch);
      }
      
      const spaceMatch = lastLine.match(/^([　\s]+)/);
      if (spaceMatch) {
        return getTagIndentWidth(spaceMatch[1]);
      }
      
      if (pLines.length > 1) {
        return 0;
      }
      currentP--;
    }
    return 0;
  }

  function renderTextToHTML(text, q, p) {
    if (!text) return '<div class="moex-line"><br></div>';
    
    let inheritedIndent = 0;
    if (q && p) {
      inheritedIndent = getInheritedIndent(q, p);
    }
    
    const lines = text.split('\n');
    return lines.map((line, index) => {
      const tag = parseLineNumberTag(line);
      let inlineStyle = '';
      if (tag) {
        const w = getTagIndentWidth(tag.rawMatch);
        inlineStyle = ` style="padding-left: ${w}px; text-indent: -${w}px;"`;
      } else {
        const spaceMatch = line.match(/^([　\s]+)/);
        if (spaceMatch) {
          const w = getTagIndentWidth(spaceMatch[1]);
          inlineStyle = ` style="padding-left: ${w}px; text-indent: -${w}px;"`;
        } else if (index === 0 && inheritedIndent > 0) {
          inlineStyle = ` style="padding-left: ${inheritedIndent}px;"`;
        }
      }
      return `<div class="moex-line"${inlineStyle}>${escapeHTML(line) || '<br>'}</div>`;
    }).join('');
  }

  function extractPlainText(editor) {
    const lines = [];
    for (let child of editor.childNodes) {
      if (child.nodeName === 'DIV') {
        lines.push(child.innerText === '\n' ? '' : child.textContent);
      } else if (child.nodeType === 3) {
        lines.push(child.textContent);
      }
    }
    return lines.join('\n');
  }

  function getCaretPosition(editor) {
    const selection = window.getSelection();
    if (!selection.rangeCount) return null;
    const range = selection.getRangeAt(0);
    
    let node = range.startContainer;
    let lineDiv = node.nodeType === 3 ? node.parentNode : node;
    
    while (lineDiv && lineDiv.parentNode !== editor && lineDiv !== editor) {
      lineDiv = lineDiv.parentNode;
    }
    
    if (!lineDiv) return null;
    if (lineDiv === editor) {
      if (editor.childNodes.length > 0) {
        lineDiv = editor.childNodes[0];
      } else {
        return { lineIndex: 0, offsetWithinLine: 0 };
      }
    }
    
    const lineIndex = Array.from(editor.childNodes).indexOf(lineDiv);
    if (lineIndex === -1) return null;
    
    const preCaretRange = range.cloneRange();
    preCaretRange.selectNodeContents(lineDiv);
    preCaretRange.setEnd(range.startContainer, range.startOffset);
    const offsetWithinLine = preCaretRange.toString().length;
    
    return { lineIndex, offsetWithinLine };
  }

  function setCaretPosition(editor, caretPos) {
    if (!caretPos) return;
    const { lineIndex, offsetWithinLine } = caretPos;
    const lineDiv = editor.children[lineIndex];
    if (!lineDiv) return;
    
    const selection = window.getSelection();
    const range = document.createRange();
    
    let currentOffset = 0;
    let found = false;
    
    function traverse(node) {
      if (found) return;
      if (node.nodeType === 3) {
        if (currentOffset + node.length >= offsetWithinLine) {
          range.setStart(node, offsetWithinLine - currentOffset);
          range.collapse(true);
          found = true;
        } else {
          currentOffset += node.length;
        }
      } else {
        for (let child of node.childNodes) {
          traverse(child);
        }
      }
    }
    
    traverse(lineDiv);
    
    if (!found) {
      range.selectNodeContents(lineDiv);
      range.collapse(false);
    }
    
    selection.removeAllRanges();
    selection.addRange(range);
  }

  function updateEditorLines(editor, q, p) {
    let changed = false;
    let inheritedIndent = 0;
    if (q && p) {
      inheritedIndent = getInheritedIndent(q, p);
    }

    const newNodes = [];
    Array.from(editor.childNodes).forEach((child, index) => {
      if (child.nodeName !== 'DIV') {
        const div = document.createElement('div');
        div.className = 'moex-line';
        div.textContent = child.textContent;
        editor.replaceChild(div, child);
        changed = true;
      } else {
        if (!child.classList.contains('moex-line')) {
          child.className = 'moex-line';
          changed = true;
        }
        const text = child.textContent;
        const tag = parseLineNumberTag(text);
        if (tag) {
          const w = getTagIndentWidth(tag.rawMatch);
          if (child.style.paddingLeft !== `${w}px`) {
            child.style.paddingLeft = `${w}px`;
            child.style.textIndent = `-${w}px`;
            changed = true;
          }
        } else {
          const spaceMatch = text.match(/^([　\s]+)/);
          if (spaceMatch) {
            const w = getTagIndentWidth(spaceMatch[1]);
            if (child.style.paddingLeft !== `${w}px` || child.style.textIndent !== `-${w}px`) {
              child.style.paddingLeft = `${w}px`;
              child.style.textIndent = `-${w}px`;
              changed = true;
            }
          } else if (index === 0 && inheritedIndent > 0) {
            if (child.style.paddingLeft !== `${inheritedIndent}px` || child.style.textIndent !== '') {
              child.style.paddingLeft = `${inheritedIndent}px`;
              child.style.textIndent = '';
              changed = true;
            }
          } else {
            if (child.style.paddingLeft || child.style.textIndent) {
              child.style.paddingLeft = '';
              child.style.textIndent = '';
              changed = true;
            }
          }
        }
      }
    });
    return changed;
  }

  function applyHangingIndents(editor) {
    let inheritedIndent = 0;
    Array.from(editor.childNodes).forEach(child => {
      if (child.nodeType === 3) {
        // Skip
      } else if (child.nodeName === 'DIV') {
        const text = child.textContent;
        const tag = parseLineNumberTag(text);
        
        if (tag) {
          const w = getTagIndentWidth(tag.rawMatch);
          child.style.paddingLeft = `${w}px`;
          child.style.textIndent = `-${w}px`;
          inheritedIndent = w;
        } else {
          const spaceMatch = text.match(/^([　\s]+)/);
          if (spaceMatch) {
            const w = getTagIndentWidth(spaceMatch[1]);
            child.style.paddingLeft = `${w}px`;
            child.style.textIndent = `-${w}px`;
            inheritedIndent = w;
          } else if (inheritedIndent > 0) {
            child.style.paddingLeft = `${inheritedIndent}px`;
            child.style.textIndent = '';
          } else {
            child.style.paddingLeft = '';
            child.style.textIndent = '';
          }
        }
      }
    });
  }

  // ==========================================================================
  // Clean Plain Text & Level-aware Copy / Cut / Paste Engine
  // ==========================================================================
  function cleanTextFromNumbering(rawText) {
    if (!rawText) return '';
    const lines = rawText.split(/\r?\n/);
    const cleanedLines = lines.map(line => {
      const tag = parseLineNumberTag(line);
      if (tag) {
        return line.substring(tag.rawMatch.length).replace(/^[　\s]+/, '');
      }
      return line.replace(/^[　\s]+/, '');
    });
    return cleanedLines.join('\n');
  }

  function getCleanSelectedText() {
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return '';
    return cleanTextFromNumbering(sel.toString());
  }

  function handlePasteText(rawClipboard) {
    if (!rawClipboard) return;
    if (!activeTextarea) {
      activeTextarea = paperPagesContainer.querySelector('.paper-textarea');
    }
    if (!activeTextarea) return;

    saveUndoSnapshot();
    activeTextarea.focus();

    // Clean source text: strip source numbering tags & leading indents
    const lines = rawClipboard.split(/\r?\n/).map(line => {
      const tag = parseLineNumberTag(line);
      if (tag) {
        return line.substring(tag.rawMatch.length).replace(/^[　\s]+/, '');
      }
      return line.replace(/^[　\s]+/, '');
    });

    const details = getActiveLineDetails();
    if (!details) {
      document.execCommand('insertText', false, lines.join('\n'));
      return;
    }

    const { currentLineText, lineIndex, cursorOffsetInLine } = details;
    const currentTag = parseLineNumberTag(currentLineText);

    const sel = window.getSelection();
    if (sel && !sel.isCollapsed) {
      document.execCommand('delete');
    }

    // If pasted into a numbered item and multiple lines exist:
    // "若將多段文字貼上到編號項目內，會依貼上位置的目前階層去建立後續項目；若來源為空白行仍維持為空白項目。"
    if (currentTag && lines.length > 1) {
      const allDocLines = extractPlainText(activeTextarea).split('\n');
      const refreshedDetails = getActiveLineDetails() || details;
      const refLineText = refreshedDetails.currentLineText;
      const refOffset = refreshedDetails.cursorOffsetInLine;
      const refIndex = refreshedDetails.lineIndex;

      const textBeforeCursor = refLineText.substring(0, refOffset);
      const textAfterCursor = refLineText.substring(refOffset);

      allDocLines[refIndex] = textBeforeCursor + lines[0];

      const insertedLines = [];
      for (let i = 1; i < lines.length; i++) {
        const lText = lines[i];
        if (lText.trim() === '') {
          insertedLines.push('');
        } else {
          const nextTag = generateNumberTag(currentTag.level, currentTag.num + i);
          insertedLines.push(nextTag + lText);
        }
      }
      if (insertedLines.length > 0) {
        insertedLines[insertedLines.length - 1] += textAfterCursor;
      } else {
        allDocLines[refIndex] += textAfterCursor;
      }

      allDocLines.splice(refIndex + 1, 0, ...insertedLines);
      const newDocText = allDocLines.join('\n');
      activeTextarea.innerHTML = newDocText.split('\n').map(l => `<div>${escapeHTML(l) || '<br>'}</div>`).join('');
      applyHangingIndents(activeTextarea);
      answers[currentQ][1] = newDocText;

      const targetLineIdx = refIndex + insertedLines.length;
      const targetOffset = allDocLines[targetLineIdx].length - textAfterCursor.length;
      setCaretPosition(activeTextarea, { lineIndex: targetLineIdx, offsetWithinLine: targetOffset });

      resequenceDocumentNumbers();
      activeTextarea.dispatchEvent(new Event('input'));
    } else {
      document.execCommand('insertText', false, lines.join('\n'));
      setTimeout(() => {
        resequenceDocumentNumbers();
      }, 0);
    }
  }

  // Render 8 Paper Pages per Question (Now as a Single Editor Layer over Background Pages)
  function renderEditor() {
    paperPagesContainer.innerHTML = '';
    activeTextarea = null;

    const editorWrapper = document.createElement('div');
    editorWrapper.className = 'exam-editor-wrapper';

    // The single editor
    const editor = document.createElement('div');
    editor.className = 'paper-textarea single-editor';
    editor.contentEditable = 'true';
    editor.style.fontSize = `${fontSize}px`;
    
    // Combine 8 pages into 1 for display
    let combinedText = '';
    for (let p = 1; p <= 8; p++) {
      if (answers[currentQ] && answers[currentQ][p]) {
        combinedText += answers[currentQ][p] + '\n';
      }
    }
    combinedText = combinedText.trimEnd();

    if (!combinedText) {
      editor.setAttribute('data-placeholder', '（請從本頁第 1 行依序開始登記作答內文...）');
      editor.innerHTML = '<div><br></div>';
    } else {
      editor.innerHTML = combinedText.split('\n').map(line => `<div>${escapeHTML(line) || '<br>'}</div>`).join('');
    }
    
    applyHangingIndents(editor);

    editor.addEventListener('focus', function () {
      activeTextarea = editor;
      updateToolbarActiveStates();
    });

    editor.addEventListener('keyup', updateToolbarActiveStates);
    editor.addEventListener('click', updateToolbarActiveStates);

    // Clean plain text Copy / Cut event handlers
    editor.addEventListener('copy', function (e) {
      const cleanText = getCleanSelectedText();
      if (cleanText) {
        e.preventDefault();
        e.clipboardData.setData('text/plain', cleanText);
      }
    });

    editor.addEventListener('cut', function (e) {
      const cleanText = getCleanSelectedText();
      if (cleanText) {
        e.preventDefault();
        e.clipboardData.setData('text/plain', cleanText);
        saveUndoSnapshot();
        document.execCommand('delete');
        setTimeout(() => {
          resequenceDocumentNumbers();
        }, 0);
      }
    });

    // Word-like Enter: auto-continue numbered list, or exit on empty title
    // Word-like Backspace: remove entire tag at once
    // Word-like cursor: snap cursor to after tag (tag is not editable)
    editor.addEventListener('keydown', function (e) {

      // === TAB / SHIFT+TAB: Level Indent / Outdent ===
      if (e.key === 'Tab') {
        const details = getActiveLineDetails();
        if (details) {
          const { currentLineText } = details;
          const tag = parseLineNumberTag(currentLineText);
          if (tag) {
            e.preventDefault();
            changeHierarchyLevel(e.shiftKey ? -1 : 1);
            return;
          }
        }
      }

      // === ENTER: Handle numbered list item creation or removal ===
      if (e.key === 'Enter' && !e.shiftKey && !e.isComposing) {
        const details = getActiveLineDetails();
        if (!details) return;

        const { lines, lineIndex, currentLineText, cursorOffsetInLine } = details;
        const tag = parseLineNumberTag(currentLineText);
        if (!tag) return; // not a numbered line, let browser handle normally

        // Check if there is typed text after the tag
        const textAfterTag = currentLineText.substring(tag.rawMatch.length);
        if (textAfterTag.trim() === '') {
          // 標題後面未打字：按 Enter 刪除該標題（不能未打字再開下一個標題）
          e.preventDefault();
          saveUndoSnapshot();
          const newLines = [...lines];
          newLines[lineIndex] = '';
          const newText = newLines.join('\n');
          editor.innerHTML = newText.split('\n').map(line => `<div>${escapeHTML(line) || '<br>'}</div>`).join('');
          applyHangingIndents(editor);
          answers[currentQ][1] = newText;
          editor.focus();
          setCaretPosition(editor, { lineIndex: lineIndex, offsetWithinLine: 0 });
          resequenceDocumentNumbers();
          editor.dispatchEvent(new Event('input'));
          return;
        }

        e.preventDefault();
        saveUndoSnapshot();

        // If cursor is at the very beginning (offset 0), insert an empty line above
        if (cursorOffsetInLine === 0) {
          const newLines = [...lines];
          newLines.splice(lineIndex, 0, '');
          const newText = newLines.join('\n');
          editor.innerHTML = newText.split('\n').map(line => `<div>${escapeHTML(line) || '<br>'}</div>`).join('');
          applyHangingIndents(editor);
          answers[currentQ][1] = newText;
          editor.focus();
          setCaretPosition(editor, { lineIndex: lineIndex + 1, offsetWithinLine: tag.rawMatch.length });
          resequenceDocumentNumbers();
          editor.dispatchEvent(new Event('input'));
          return;
        }

        // Effective cursor: if inside tag, treat as right after tag
        const effectiveOffset = Math.max(cursorOffsetInLine, tag.rawMatch.length);
        const textBeforeCursor = currentLineText.substring(0, effectiveOffset);
        const remainingTextAfterCursor = currentLineText.substring(effectiveOffset);

        const nextTag = generateNumberTag(tag.level, tag.num + 1);

        const newLines = [...lines];
        newLines[lineIndex] = textBeforeCursor;
        newLines.splice(lineIndex + 1, 0, nextTag + remainingTextAfterCursor);
        const newText = newLines.join('\n');
        editor.innerHTML = newText.split('\n').map(line => `<div>${escapeHTML(line) || '<br>'}</div>`).join('');
        applyHangingIndents(editor);
        answers[currentQ][1] = newText;

        editor.focus();
        setCaretPosition(editor, { lineIndex: lineIndex + 1, offsetWithinLine: nextTag.length });

        resequenceDocumentNumbers();
        editor.dispatchEvent(new Event('input'));
        return;
      }

      // === BACKSPACE: remove entire tag at once when cursor is at/within tag ===
      if (e.key === 'Backspace' && !e.isComposing) {
        const sel = window.getSelection();
        if (sel && !sel.isCollapsed) {
          // 圈選文字狀態：交由瀏覽器正常刪除選取的文字內容，隨後重整序號
          setTimeout(() => {
            resequenceDocumentNumbers();
          }, 0);
          return;
        }

        const details = getActiveLineDetails();
        if (!details) return;

        const { lines, lineIndex, currentLineText, cursorOffsetInLine } = details;
        const tag = parseLineNumberTag(currentLineText);
        if (!tag) return; // not a numbered line, let browser handle normally

        // If cursor is at or before the end of the tag, remove the whole tag
        if (cursorOffsetInLine <= tag.rawMatch.length) {
          e.preventDefault();
          saveUndoSnapshot();
          const textAfterTag = currentLineText.substring(tag.rawMatch.length);
          const newLines = [...lines];
          newLines[lineIndex] = textAfterTag;
          const newText = newLines.join('\n');
          editor.innerHTML = newText.split('\n').map(line => `<div>${escapeHTML(line) || '<br>'}</div>`).join('');
          applyHangingIndents(editor);
          answers[currentQ][1] = newText;
          editor.focus();
          setCaretPosition(editor, { lineIndex: lineIndex, offsetWithinLine: 0 });
          resequenceDocumentNumbers();
          editor.dispatchEvent(new Event('input'));
          return;
        }
        // If cursor is after the tag, let browser handle normally (delete one char)
        return;
      }

      // === DELETE: remove entire tag at once if cursor is before or inside tag ===
      if (e.key === 'Delete' && !e.isComposing) {
        const sel = window.getSelection();
        if (sel && !sel.isCollapsed) {
          // 圈選文字狀態：交由瀏覽器正常刪除選取的文字內容，隨後重整序號
          setTimeout(() => {
            resequenceDocumentNumbers();
          }, 0);
          return;
        }

        const details = getActiveLineDetails();
        if (!details) return;

        const { lines, lineIndex, currentLineText, cursorOffsetInLine } = details;
        const tag = parseLineNumberTag(currentLineText);
        if (tag && cursorOffsetInLine < tag.rawMatch.length) {
          e.preventDefault();
          saveUndoSnapshot();
          const textAfterTag = currentLineText.substring(tag.rawMatch.length);
          const newLines = [...lines];
          newLines[lineIndex] = textAfterTag;
          const newText = newLines.join('\n');
          editor.innerHTML = newText.split('\n').map(line => `<div>${escapeHTML(line) || '<br>'}</div>`).join('');
          applyHangingIndents(editor);
          answers[currentQ][1] = newText;
          editor.focus();
          setCaretPosition(editor, { lineIndex: lineIndex, offsetWithinLine: 0 });
          resequenceDocumentNumbers();
          editor.dispatchEvent(new Event('input'));
          return;
        }
      }

      // === Arrow Left / Home: snap cursor to after tag (don't let cursor enter tag) ===
      if (e.key === 'ArrowLeft' || e.key === 'Home') {
        setTimeout(() => {
          const details = getActiveLineDetails();
          if (!details) return;
          const { lineIndex, currentLineText, cursorOffsetInLine } = details;
          const tag = parseLineNumberTag(currentLineText);
          if (tag && cursorOffsetInLine > 0 && cursorOffsetInLine < tag.rawMatch.length) {
            setCaretPosition(editor, { lineIndex, offsetWithinLine: tag.rawMatch.length });
          }
        }, 0);
      }
    });

    editor.addEventListener('paste', function(e) {
      e.preventDefault();
      const text = (e.originalEvent || e).clipboardData.getData('text/plain');
      handlePasteText(text);
    });

    editor.addEventListener('input', function () {
      if (editor.getAttribute('data-placeholder')) {
        editor.removeAttribute('data-placeholder');
      }
      
      applyHangingIndents(editor);
      
      const plainText = extractPlainText(editor);
      // Because we still need to store it in answers[currentQ][1..8] for Word Export compatibility
      // let's just store the whole thing in page 1, and clear pages 2..8
      answers[currentQ][1] = plainText;
      for (let p = 2; p <= 8; p++) {
          answers[currentQ][p] = '';
      }

      updateStats();

      clearTimeout(autoSaveTimer);
      updateAutoSaveStatus('saving');
      autoSaveTimer = setTimeout(() => {
        saveCurrentAnswer(false);
        addHistorySnapshot('自動儲存');
      }, 2500);
    });

    // Create Background Pages
    const bgContainer = document.createElement('div');
    bgContainer.className = 'exam-bg-container';
    
    for (let p = 1; p <= 8; p++) {
      const page = document.createElement('div');
      page.className = 'paper-page-bg';
      
      const leftMargin = document.createElement('div');
      leftMargin.className = 'paper-left-margin';
      leftMargin.textContent = `第 ${currentQ} 題 第 ${p} 頁`;
      
      const centerArea = document.createElement('div');
      centerArea.className = 'paper-center-area';
      
      const rightMargin = document.createElement('div');
      rightMargin.className = 'paper-right-margin';
      rightMargin.textContent = '（請從本頁第 1 行依序開始登記）';
      
      page.appendChild(leftMargin);
      page.appendChild(centerArea);
      page.appendChild(rightMargin);
      bgContainer.appendChild(page);
    }
    
    editorWrapper.appendChild(bgContainer);
    editorWrapper.appendChild(editor);
    paperPagesContainer.appendChild(editorWrapper);

    const firstEditor = paperPagesContainer.querySelector('.paper-textarea');
    if (firstEditor) activeTextarea = firstEditor;

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

  function fullWidthToHalf(str) {
    return str.replace(/[０-９Ａ-Ｚａ-ｚ]/g, function(s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xfee0);
    });
  }

  function parseLineNumberTag(lineText) {
    let match;
    // Level 1: (一) or （一）
    match = lineText.match(/^([　\s]*)(?:\(|（)([一二三四五六七八九十]+)(?:\)|）)/);
    if (match) return { level: 1, num: chineseToNum(match[2]), rawMatch: match[0], indent: match[1] };
    
    // Level 2: 1. or １.
    match = lineText.match(/^([　\s]*)([0-9０-９]+)(?:\.|．)/);
    if (match) return { level: 2, num: parseInt(fullWidthToHalf(match[2]), 10), rawMatch: match[0], indent: match[1] };
    
    // Level 3: (1) or （1） or （１）
    match = lineText.match(/^([　\s]*)(?:\(|（)([0-9０-９]+)(?:\)|）)/);
    if (match) return { level: 3, num: parseInt(fullWidthToHalf(match[2]), 10), rawMatch: match[0], indent: match[1] };
    
    // Level 4: A. or Ａ.
    match = lineText.match(/^([　\s]*)([A-ZＡ-Ｚ])(?:\.|．)/);
    if (match) {
      const charCode = fullWidthToHalf(match[2]).charCodeAt(0);
      return { level: 4, num: charCode - 64, rawMatch: match[0], indent: match[1] };
    }
    
    // Level 5: a. or ａ.
    match = lineText.match(/^([　\s]*)([a-zａ-ｚ])(?:\.|．)/);
    if (match) {
      const charCode = fullWidthToHalf(match[2]).charCodeAt(0);
      return { level: 5, num: charCode - 96, rawMatch: match[0], indent: match[1] };
    }
    
    // Level 6: (a) or （a） or （ａ）
    match = lineText.match(/^([　\s]*)(?:\(|（)([a-zａ-ｚ])(?:\)|）)/);
    if (match) {
      const charCode = fullWidthToHalf(match[2]).charCodeAt(0);
      return { level: 6, num: charCode - 96, rawMatch: match[0], indent: match[1] };
    }

    // Fallback for previous Level 1: 一、
    match = lineText.match(/^([　\s]*)([一二三四五六七八九十]+)、/);
    if (match) return { level: 1, num: chineseToNum(match[2]), rawMatch: match[0], indent: match[1] };

    return null;
  }

  function generateNumberTag(level, num) {
    switch (level) {
      case 1: return `(${numToChinese(num)})`;
      case 2: return `　　${num}.`;
      case 3: return `　　　　(${num})`;
      case 4: return `　　　　　　${String.fromCharCode(num + 64)}.`;
      case 5: return `　　　　　　　　${String.fromCharCode(num + 96)}.`;
      case 6: return `　　　　　　　　　　(${String.fromCharCode(num + 96)})`;
      default: return `(${numToChinese(num)})`;
    }
  }

  function getActiveLineDetails() {
    if (!activeTextarea) {
      activeTextarea = paperPagesContainer.querySelector('.paper-textarea');
    }
    if (!activeTextarea) return null;

    const caretPos = getCaretPosition(activeTextarea);
    if (!caretPos) return null;

    const plainText = extractPlainText(activeTextarea);
    const lines = plainText.split('\n');
    const lineIndex = caretPos.lineIndex;

    return {
      textarea: activeTextarea,
      lines: lines,
      lineIndex: lineIndex,
      currentLineText: lines[lineIndex] || '',
      lineStartOffset: 0,
      cursorOffsetInLine: caretPos.offsetWithinLine
    };
  }

  function resequenceDocumentNumbers() {
    let levelCounts = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0 };
    
    const textarea = paperPagesContainer.querySelector('.single-editor');
    if (!textarea) return;
    
    const val = extractPlainText(textarea);
    const caretPos = getCaretPosition(textarea);
    const lines = val.split('\n');
    
    let changed = false;
    let newOffsetWithinLine = caretPos ? caretPos.offsetWithinLine : 0;
    
    for (let i = 0; i < lines.length; i++) {
      const lineText = lines[i];
      const tag = parseLineNumberTag(lineText);
      
      if (tag) {
        levelCounts[tag.level]++;
        
        for (let l = tag.level + 1; l <= 6; l++) {
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
          
          if (caretPos && caretPos.lineIndex === i) {
             newOffsetWithinLine += lenDiff;
          }
        }
      }
    }
    
    if (changed) {
      const newText = lines.join('\n');
      textarea.innerHTML = newText.split('\n').map(line => `<div>${escapeHTML(line) || '<br>'}</div>`).join('');
      applyHangingIndents(textarea);
      if (textarea === activeTextarea && caretPos) {
        setCaretPosition(textarea, { lineIndex: caretPos.lineIndex, offsetWithinLine: Math.max(0, newOffsetWithinLine) });
      }
      answers[currentQ][1] = newText;
    }
  }

  // ==========================================================================
  // Multi-Step Undo/Redo Engine (提供 4 次編輯還原機會)
  // ==========================================================================
  let undoStack = [];
  let redoStack = [];

  function saveUndoSnapshot() {
    if (!activeTextarea && paperPagesContainer) {
      activeTextarea = paperPagesContainer.querySelector('.paper-textarea');
    }
    if (!activeTextarea) return;

    undoStack.push({
      textarea: activeTextarea,
      val: extractPlainText(activeTextarea),
      caret: getCaretPosition(activeTextarea)
    });
    if (undoStack.length > 4) undoStack.shift();
    redoStack = [];
  }

  function performUndo() {
    if (!activeTextarea && paperPagesContainer) {
      activeTextarea = paperPagesContainer.querySelector('.paper-textarea');
    }
    if (!activeTextarea) return;

    if (undoStack.length === 0) {
      document.execCommand('undo');
      return;
    }

    redoStack.push({
      textarea: activeTextarea,
      val: extractPlainText(activeTextarea),
      caret: getCaretPosition(activeTextarea)
    });

    const target = undoStack.pop();

    activeTextarea = target.textarea;
    activeTextarea.focus();
    activeTextarea.innerHTML = target.val.split('\n').map(line => `<div>${escapeHTML(line) || '<br>'}</div>`).join('');
    applyHangingIndents(activeTextarea);
    setCaretPosition(activeTextarea, target.caret);

    answers[currentQ][1] = target.val;
    resequenceDocumentNumbers();
    activeTextarea.dispatchEvent(new Event('input'));
    updateToolbarActiveStates();
  }

  function performRedo() {
    if (!activeTextarea && paperPagesContainer) {
      activeTextarea = paperPagesContainer.querySelector('.paper-textarea');
    }
    if (!activeTextarea) return;

    if (redoStack.length === 0) {
      document.execCommand('redo');
      return;
    }

    undoStack.push({
      textarea: activeTextarea,
      val: extractPlainText(activeTextarea),
      caret: getCaretPosition(activeTextarea)
    });

    const target = redoStack.pop();

    activeTextarea = target.textarea;
    activeTextarea.focus();
    activeTextarea.innerHTML = target.val.split('\n').map(line => `<div>${escapeHTML(line) || '<br>'}</div>`).join('');
    applyHangingIndents(activeTextarea);
    setCaretPosition(activeTextarea, target.caret);

    answers[currentQ][1] = target.val;
    resequenceDocumentNumbers();
    activeTextarea.dispatchEvent(new Event('input'));
    updateToolbarActiveStates();
  }

  function applyLineChange(details, newLineText, relativeCursorShift = 0) {
    saveUndoSnapshot();
    const { textarea, lines, lineIndex, cursorOffsetInLine } = details;
    
    lines[lineIndex] = newLineText;
    const newText = lines.join('\n');
    textarea.innerHTML = newText.split('\n').map(line => `<div>${escapeHTML(line) || '<br>'}</div>`).join('');
    applyHangingIndents(textarea);
    answers[currentQ][1] = newText;

    textarea.focus();
    const newOffset = Math.max(0, cursorOffsetInLine + relativeCursorShift);
    setCaretPosition(textarea, { lineIndex: lineIndex, offsetWithinLine: newOffset });

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

    const { textarea, lines, lineIndex, currentLineText, cursorOffsetInLine } = details;
    const existingTag = parseLineNumberTag(currentLineText);

    if (!existingTag) {
      toggleNumberedList();
      return;
    }

    const oldParentLevel = existingTag.level;
    let newParentLevel = oldParentLevel + direction;
    if (newParentLevel < 1) newParentLevel = 1;
    if (newParentLevel > 6) newParentLevel = 6;

    const delta = newParentLevel - oldParentLevel;
    if (delta === 0) return; // 已達到最高或最低階層

    saveUndoSnapshot();

    // 1. 調整父層階層
    const parentNewTag = generateNumberTag(newParentLevel, existingTag.num);
    const parentTextAfterTag = currentLineText.substring(existingTag.rawMatch.length);
    lines[lineIndex] = parentNewTag + parentTextAfterTag;

    // 2. 子層之連動性（官方規範⑤）：向下尋找所有屬於該項目的子層（level > oldParentLevel），同步調整 delta
    for (let i = lineIndex + 1; i < lines.length; i++) {
      const childLine = lines[i];
      if (childLine.trim() === '') {
        break; // 空白行中斷
      }
      const childTag = parseLineNumberTag(childLine);
      if (!childTag) {
        break; // 非編號行中斷
      }
      if (childTag.level <= oldParentLevel) {
        break; // 遇到同階或更淺階層，代表子層範圍結束
      }
      // 此行是子層，同步位移階層
      const childNewLevel = Math.max(1, Math.min(6, childTag.level + delta));
      const childNewTag = generateNumberTag(childNewLevel, childTag.num);
      const childTextAfterTag = childLine.substring(childTag.rawMatch.length);
      lines[i] = childNewTag + childTextAfterTag;
    }

    const newText = lines.join('\n');
    textarea.innerHTML = newText.split('\n').map(line => `<div>${escapeHTML(line) || '<br>'}</div>`).join('');
    applyHangingIndents(textarea);
    answers[currentQ][1] = newText;

    textarea.focus();
    const shift = parentNewTag.length - existingTag.rawMatch.length;
    const newOffset = Math.max(0, cursorOffsetInLine + shift);
    setCaretPosition(textarea, { lineIndex: lineIndex, offsetWithinLine: newOffset });

    resequenceDocumentNumbers();
    textarea.dispatchEvent(new Event('input'));
    updateToolbarActiveStates();
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

  // Saved selection range (preserved when toolbar buttons are clicked)
  let savedSelection = null;
  function saveSelection() {
    const sel = window.getSelection();
    if (sel.rangeCount > 0) {
      savedSelection = sel.getRangeAt(0).cloneRange();
    }
  }
  function restoreSelection() {
    if (savedSelection && activeTextarea) {
      activeTextarea.focus();
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(savedSelection);
    }
  }

  // Symbol & Tool Item Toolbar Actions
  function setupSymbolToolbar() {
    // Prevent ALL toolbar items from stealing focus
    const allToolbarItems = document.querySelectorAll('.sym-item, .tool-icon');
    allToolbarItems.forEach(item => {
      item.addEventListener('mousedown', function (e) {
        e.preventDefault(); // prevent focus loss from editor
        saveSelection();
      });
    });

    const symbolItems = document.querySelectorAll('.sym-item');
    symbolItems.forEach(item => {
      if (item.hasAttribute('data-symbol')) {
        item.addEventListener('click', function () {
          restoreSelection();
          const symbol = item.getAttribute('data-symbol');
          insertAtCursor(symbol);
        });
      }
    });

    const toolCut = document.getElementById('toolCut');
    if (toolCut) {
      toolCut.addEventListener('click', () => {
        restoreSelection();
        if (activeTextarea) activeTextarea.focus();
        const cleanText = getCleanSelectedText();
        if (cleanText) {
          saveUndoSnapshot();
          if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(cleanText).catch(() => {});
          }
          document.execCommand('delete');
          setTimeout(() => {
            resequenceDocumentNumbers();
          }, 0);
        } else {
          document.execCommand('cut');
        }
      });
    }

    const toolCopy = document.getElementById('toolCopy');
    if (toolCopy) {
      toolCopy.addEventListener('click', () => {
        restoreSelection();
        if (activeTextarea) activeTextarea.focus();
        const cleanText = getCleanSelectedText();
        if (cleanText) {
          if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(cleanText).catch(() => {});
          }
        } else {
          document.execCommand('copy');
        }
      });
    }

    const toolPaste = document.getElementById('toolPaste');
    if (toolPaste) {
      toolPaste.addEventListener('click', async () => {
        restoreSelection();
        if (!activeTextarea) return;
        try {
          const clipText = await navigator.clipboard.readText();
          if (clipText) {
            handlePasteText(clipText);
            return;
          }
        } catch (e) {
          document.execCommand('paste');
        }
      });
    }

    const toolUndo = document.getElementById('toolUndo');
    if (toolUndo) {
      toolUndo.addEventListener('click', () => {
        restoreSelection();
        performUndo();
      });
    }

    const toolRedo = document.getElementById('toolRedo');
    if (toolRedo) {
      toolRedo.addEventListener('click', () => {
        restoreSelection();
        performRedo();
      });
    }

    const toolNumbering = document.getElementById('toolNumbering');
    if (toolNumbering) {
      toolNumbering.addEventListener('click', () => {
        restoreSelection();
        toggleNumberedList();
      });
    }

    const toolPrevLevel = document.getElementById('toolPrevLevel');
    if (toolPrevLevel) {
      toolPrevLevel.addEventListener('click', () => {
        restoreSelection();
        changeHierarchyLevel(-1); // 上一層 (Promote)
      });
    }

    const toolNextLevel = document.getElementById('toolNextLevel');
    if (toolNextLevel) {
      toolNextLevel.addEventListener('click', () => {
        restoreSelection();
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
    
    document.execCommand('insertText', false, text);
    
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
    const totalQCount = Object.keys(QUESTIONS).length;
    Object.keys(QUESTIONS).forEach(qStr => {
      const q = Number(qStr);
      let qHasText = false;
      for (let p = 1; p <= 8; p++) {
        if ((answers[q] && answers[q][p] || '').trim().length > 0) {
          qHasText = true;
          break;
        }
      }
      if (qHasText) answered++;
    });

    answeredCountEl.textContent = answered;
    unansweredCountEl.textContent = totalQCount - answered;
  }

  // Navigation
  function switchQuestion(newQ) {
    const totalQCount = Object.keys(QUESTIONS).length;
    if (newQ < 1 || newQ > totalQCount) return;
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
      if (ta) ta.focus({ preventScroll: true });
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
        Object.keys(QUESTIONS).forEach(qStr => {
          const q = Number(qStr);
          answers[q] = normalizeAnswerData(item.answers[q]);
        });
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

    Object.keys(QUESTIONS).map(Number).forEach(q => {
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
            
            let textToRender = preservedLine;
            const pOptions = {};
            
            if (tag) {
              pOptions.outlineLevel = tag.level;
              
              const indentSpaces = tag.indent;
              const tagContent = tag.rawMatch.substring(indentSpaces.length);
              
              let indentSpaceTwips = 0;
              for (let i = 0; i < indentSpaces.length; i++) {
                indentSpaceTwips += (indentSpaces.charCodeAt(i) > 255) ? 240 : 120;
              }
              
              let tagTwips = 0;
              for (let i = 0; i < tagContent.length; i++) {
                tagTwips += (tagContent.charCodeAt(i) > 255) ? 240 : 120;
              }
              
              pOptions.indent = { left: indentSpaceTwips + tagTwips, hanging: tagTwips };
              textToRender = preservedLine.substring(indentSpaces.length);
            } else {
              const spaceMatch = trimmedLine.match(/^([　\s]+)/);
              if (spaceMatch) {
                const indentSpaces = spaceMatch[1];
                let indentSpaceTwips = 0;
                for (let i = 0; i < indentSpaces.length; i++) {
                  indentSpaceTwips += (indentSpaces.charCodeAt(i) > 255) ? 240 : 120;
                }
                pOptions.indent = { left: indentSpaceTwips };
                textToRender = preservedLine.substring(indentSpaces.length);
              }
            }
            
            pOptions.children = [new TextRun({ text: textToRender })];
            
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
      const totalQCount = Object.keys(QUESTIONS).length;
      Object.keys(QUESTIONS).forEach(qStr => {
        const q = Number(qStr);
        for (let p = 1; p <= 8; p++) {
          if ((answers[q] && answers[q][p] || '').trim().length > 0) {
            count++;
            break;
          }
        }
      });
      if (confirm(`您目前已完成 ${count} / ${totalQCount} 題。確定要結束作答並交卷嗎？`)) {
        alert('【成功交卷】感謝使用考選部 CBT 線上模擬作答系統！您可以點擊「匯出 Word 檔」進行備份。接下來系統將為您清空並恢復初始狀態。');
        
        answers = {};
        Object.keys(QUESTIONS).forEach(qStr => {
          answers[Number(qStr)] = createEmptyQuestionPages();
        });
        currentQ = 1;
        currentPage = 1;
        
        if (selectQuestion) selectQuestion.value = 1;
        if (selectPage) selectPage.value = 1;
        
        renderEditor();
        updateStats();
        
        if (db) {
          try {
            const tx = db.transaction(['answers', 'history'], 'readwrite');
            tx.objectStore('answers').clear();
            tx.objectStore('history').clear();
          } catch(e) {
            console.error(e);
          }
        }
        historyList = [];
        
        secondsRemaining = 3 * 3600 - 1;
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

      const qTitleText = QUESTIONS[currentQ] ? (QUESTIONS[currentQ].title.split(' ')[0] || `第 ${currentQ} 題`) : `第 ${currentQ} 題`;
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
