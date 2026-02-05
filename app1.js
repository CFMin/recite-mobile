// Mobile-optimized app for Recite
(function() {
  'use strict';

  // Utility functions
  function uid() {
    return Date.now().toString(36) + Math.random().toString(36).substr(2);
  }

  function nowIso() {
    return new Date().toISOString();
  }

  function clamp(v, min, max) {
    return Math.min(Math.max(v, min), max);
  }

  function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text || '';
    return div.innerHTML;
  }

  function el(id) {
    return document.getElementById(id);
  }

  function download(filename, content) {
    const a = document.createElement('a');
    a.href = content;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }

  function downloadBlob(filename, blob) {
    const url = URL.createObjectURL(blob);
    download(filename, url);
    URL.revokeObjectURL(url);
  }

  // Storage
  const STORAGE_KEY = 'recite_v1';

  function defaultState() {
    return {
      qas: [
        {
          id: uid(),
          question: '示例：什么是强化学习？',
          answerText: '强化学习是一类通过与环境交互来学习策略的方法。智能体在状态下选择动作并获得奖励。目标是最大化长期累积回报。',
          createdAt: nowIso(),
          updatedAt: nowIso(),
        },
      ],
      settings: {
        groupSize: 3,
        repeatPerGroup: 4,
        reviewPrevAfterEach: true,
        sentenceDelimiters: '。！？!?',
        ttsEnabled: true,
        ttsVoiceUri: '',
        autoPlayNextQa: false,
        forceReciteCheck: false,
        rate: 1.0,
        volume: 1.0,
        threshold: 0.65,
        useOllama: false,
        ollamaUrl: 'http://localhost:11434',
        ollamaModel: 'qwen2.5:3b',
        fullReadBeforeGroups: 1,
        fullReadAfterGroups: 1,
        reviewPrevRepeatCount: 1,
      },
      progress: {
        currentQaId: null,
      },
    };
  }

  function loadState() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return defaultState();
      const parsed = JSON.parse(raw);
      const data = parsed?.data;
      if (!data) return defaultState();
      return { ...defaultState(), ...data };
    } catch {
      return defaultState();
    }
  }

  function saveState() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify({ data: state }));
    } catch {}
  }

  // Text processing
  function splitSentences(text) {
    const delims = String(state.settings.sentenceDelimiters || '。！？!?') + '，,;；、';
    const regex = new RegExp(`([${delims.replace(/[\\\]\[\-\^]/g, '\\$&')}])`, 'g');
    return text.split(regex).map(s => s.trim()).filter(Boolean);
  }

  function chunk(array, size) {
    const chunks = [];
    for (let i = 0; i < array.length; i += size) {
      chunks.push(array.slice(i, i + size));
    }
    return chunks;
  }

  function diceSimilarity(a, b) {
    if (!a || !b) return 0;
    const setA = new Set(a);
    const setB = new Set(b);
    const intersection = new Set([...setA].filter(x => setB.has(x)));
    return (2 * intersection.size) / (setA.size + setB.size);
  }

  // Rich text handling
  function sanitizeAnswerHtml(html) {
    const div = document.createElement('div');
    div.innerHTML = html || '';
    const tags = ['STRONG', 'B', 'MARK', 'EM', 'I', 'U', 'P', 'BR', 'SPAN'];
    const walker = document.createTreeWalker(div, Node.ELEMENT_NODE);
    const toRemove = [];
    let node;
    while (node = walker.nextNode()) {
      if (!tags.includes(node.tagName)) {
        toRemove.push(node);
      } else {
        const attrs = Array.from(node.attributes);
        attrs.forEach(attr => {
          if (!['style', 'class'].includes(attr.name)) {
            node.removeAttribute(attr.name);
          }
        });
      }
    }
    toRemove.forEach(n => n.remove());
    return div.innerHTML;
  }

  function normalizeImportedHtml(html) {
    return (html || '')
      .replace(/\r\n/g, '\n')
      .replace(/\r/g, '\n')
      .replace(/\n{3,}/g, '\n\n')
      .trim();
  }

  function htmlToPlainTextPreserveLines(html) {
    const div = document.createElement('div');
    div.innerHTML = html || '';
    return (div.textContent || '').replace(/\r\n/g, '\n').replace(/\n+/g, '\n').trim();
  }

  function extractAnswerHtmlFromRichEditor() {
    if (!ui.inputAnswerRich) return '';
    return ui.inputAnswerRich.innerHTML || '';
  }

  function setRichEditorHtml(html) {
    if (!ui.inputAnswerRich) return;
    const cleaned = sanitizeAnswerHtml(normalizeImportedHtml(html));
    ui.inputAnswerRich.innerHTML = cleaned;
  }

  function getPlainAnswerFromRichEditor() {
    return htmlToPlainTextPreserveLines(extractAnswerHtmlFromRichEditor());
  }

  // UI elements
  const ui = {
    qaList: el('qaList'),
    inputQaSearch: el('inputQaSearch'),
    checkSelectAll: el('checkSelectAll'),
    btnDeleteSelected: el('btnDeleteSelected'),
    btnClearSelection: el('btnClearSelection'),
    btnDeleteAll: el('btnDeleteAll'),
    btnPagePrev: el('btnPagePrev'),
    pageInfo: el('pageInfo'),
    btnPageNext: el('btnPageNext'),
    inputQuestion: el('inputQuestion'),
    inputAnswer: el('inputAnswer'),
    inputAnswerRich: el('inputAnswerRich'),
    btnAnswerBold: el('btnAnswerBold'),
    btnAnswerHighlight: el('btnAnswerHighlight'),
    btnAnswerClearFormat: el('btnAnswerClearFormat'),
    btnSave: el('btnSave'),
    btnNew: el('btnNew'),
    btnDelete: el('btnDelete'),

    inputGroupSize: el('inputGroupSize'),
    inputRepeat: el('inputRepeat'),
    inputSentenceDelims: el('inputSentenceDelims'),
    checkReviewPrev: el('checkReviewPrev'),
    checkTts: el('checkTts'),
    selectTtsVoice: el('selectTtsVoice'),
    checkAutoPlayNext: el('checkAutoPlayNext'),
    checkForceRecite: el('checkForceRecite'),
    inputFullReadBefore: el('inputFullReadBefore'),
    inputFullReadAfter: el('inputFullReadAfter'),
    inputReviewPrevRepeat: el('inputReviewPrevRepeat'),
    inputRate: el('inputRate'),
    inputVolume: el('inputVolume'),
    inputThreshold: el('inputThreshold'),
    btnApplySettings: el('btnApplySettings'),

    currentTitle: el('currentTitle'),
    statusLine: el('statusLine'),
    viewQuestion: el('viewQuestion'),
    viewAnswer: el('viewAnswer'),

    btnCardMode: el('btnCardMode'),
    btnCardPrev: el('btnCardPrev'),
    btnCardFlip: el('btnCardFlip'),
    btnCardNext: el('btnCardNext'),
    btnMaskToggleAll: el('btnMaskToggleAll'),

    btnStart: el('btnStart'),
    btnPause: el('btnPause'),
    btnStop: el('btnStop'),
    btnTtsTest: el('btnTtsTest'),
    btnPrev: el('btnPrev'),
    btnNext: el('btnNext'),

    btnRecStart: el('btnRecStart'),
    btnRecStop: el('btnRecStop'),
    btnCheck: el('btnCheck'),
    inputRecited: el('inputRecited'),
    matchSummary: el('matchSummary'),
    speechHint: el('speechHint'),

    inputAsk: el('inputAsk'),
    btnAsk: el('btnAsk'),
    askAnswer: el('askAnswer'),
    checkUseOllama: el('checkUseOllama'),
    inputOllamaUrl: el('inputOllamaUrl'),
    inputOllamaModel: el('inputOllamaModel'),

    btnExport: el('btnExport'),
    btnExportDocx: el('btnExportDocx'),
    fileImport: el('fileImport'),
  };

  // State
  let state = loadState();
  if (!state.progress.currentQaId && state.qas.length) {
    state.progress.currentQaId = state.qas[0].id;
  }

  // Rich editor setup
  if (ui.inputAnswerRich) {
    const syncPlain = () => {
      ui.inputAnswer.value = getPlainAnswerFromRichEditor();
    };

    ui.inputAnswerRich.addEventListener('input', () => syncPlain());

    const exec = (cmd, val) => {
      try {
        ui.inputAnswerRich.focus();
        document.execCommand(cmd, false, val);
        syncPlain();
      } catch {}
    };

    ui.btnAnswerBold?.addEventListener('click', () => exec('bold'));
    ui.btnAnswerHighlight?.addEventListener('click', () => {
      exec('hiliteColor', '#ffcc66');
      exec('backColor', '#ffcc66');
    });
    ui.btnAnswerClearFormat?.addEventListener('click', () => {
      exec('removeFormat');
      try {
        ui.inputAnswerRich.querySelectorAll('mark').forEach((m) => {
          const frag = document.createDocumentFragment();
          while (m.firstChild) frag.appendChild(m.firstChild);
          m.replaceWith(frag);
        });
      } catch {}
      syncPlain();
    });
  }

  // TTS setup
  let ttsVoices = [];
  let _ttsVoicesSig = '';

  function listVoices() {
    try {
      return speechSynthesis.getVoices() || [];
    } catch {
      return [];
    }
  }

  function voicesCount() {
    try {
      return listVoices().length;
    } catch {
      return 0;
    }
  }

  function populateTtsVoiceSelect() {
    const voices = listVoices();
    const sig = voices.map(v => `${v.name}|${v.lang}`).join(',');
    if (sig === _ttsVoicesSig) return;
    _ttsVoicesSig = sig;

    const desired = String(state.settings.ttsVoiceUri || '');
    ui.selectTtsVoice.innerHTML = '';

    const optAuto = document.createElement('option');
    optAuto.value = '';
    optAuto.textContent = '自动（中文优先）';
    ui.selectTtsVoice.appendChild(optAuto);

    voices.forEach((v) => {
      const opt = document.createElement('option');
      opt.value = v.voiceURI || '';
      opt.textContent = `${v.name || 'voice'} (${v.lang || ''})`;
      ui.selectTtsVoice.appendChild(opt);
    });

    ui.selectTtsVoice.value = desired;
  }

  try {
    if ('speechSynthesis' in window) {
      speechSynthesis.onvoiceschanged = () => {
        populateTtsVoiceSelect();
      };
    }
  } catch {}

  // Player state
  let player = {
    running: false,
    paused: false,
    qaId: null,
    mainQaId: null,
    steps: [],
    stepIndex: 0,
    activeSentenceGlobalIndex: null,
    speechUtterance: null,
    reviewMode: false,
  };

  let cardCheck = {
    enabled: false,
    flipped: false,
    index: 0,
  };

  let reciteCheck = {
    lockedSegments: new Set(),
    pointerSegment: 0,
    lastUtterance: '',
    maskMode: false,
    manualAuto: false,
  };

  let maskState = {
    showAll: false,
  };

  let listState = {
    pageSize: 5,
    page: 1,
    query: '',
    selected: new Set(),
  };

  // QA management
  function getQaById(id) {
    return state.qas.find((q) => q.id === id) || null;
  }

  function getCurrentQa() {
    return getQaById(state.progress.currentQaId);
  }

  function setCurrentQa(id) {
    state.progress.currentQaId = id;
    saveState();
    ui.inputRecited.value = '';
    resetReciteCheck();
    resetCardCheck();
    render();
    updateMatches();
  }

  function getFilteredQas() {
    const q = (listState.query || '').trim().toLowerCase();
    if (!q) return state.qas;
    return state.qas.filter((x) => {
      const a = (x.answerText || '').toLowerCase();
      const b = (x.question || '').toLowerCase();
      return a.includes(q) || b.includes(q);
    });
  }

  function getPagedQas() {
    const filtered = getFilteredQas();
    const totalPages = Math.max(1, Math.ceil(filtered.length / listState.pageSize));
    listState.page = clamp(listState.page, 1, totalPages);
    const start = (listState.page - 1) * listState.pageSize;
    const items = filtered.slice(start, start + listState.pageSize);
    return { filtered, items, totalPages };
  }

  function resetReciteCheck() {
    reciteCheck.lockedSegments = new Set();
    reciteCheck.pointerSegment = 0;
    reciteCheck.lastUtterance = '';
    reciteCheck.maskMode = false;
    reciteCheck.manualAuto = false;
    maskState.showAll = false;
    updateManualCheckButton();
  }

  function updateManualCheckButton() {
    if (!ui.btnCheck) return;
    ui.btnCheck.textContent = reciteCheck.manualAuto ? '自动输入检测：关' : '用输入检测';
  }

  function resetCardCheck() {
    cardCheck.flipped = false;
    cardCheck.index = 0;
  }

  // Rendering
  function renderQaList() {
    const cur = state.progress.currentQaId;
    const { items, filtered, totalPages } = getPagedQas();

    ui.pageInfo.textContent = `第 ${listState.page}/${totalPages} 页（共 ${filtered.length} 条）`;
    ui.btnPagePrev.disabled = listState.page <= 1;
    ui.btnPageNext.disabled = listState.page >= totalPages;

    ui.qaList.innerHTML = items
      .map((qa) => {
        const active = qa.id === cur ? 'active' : '';
        const checked = listState.selected.has(qa.id) ? 'checked' : '';
        return `
          <div class="qa-item ${active}" data-id="${qa.id}">
            <div class="qa-item-row">
              <input class="qa-check" type="checkbox" data-check-id="${qa.id}" ${checked} />
              <div class="qa-item-main">
                <div class="q">${escapeHtml(qa.question || '（无问题）')}</div>
                <div class="a">${escapeHtml((qa.answerText || '').slice(0, 90))}</div>
              </div>
            </div>
          </div>
        `;
      })
      .join('');

    ui.qaList.querySelectorAll('.qa-item').forEach((node) => {
      node.addEventListener('click', (e) => {
        const t = e.target;
        if (t && t.matches && t.matches('input[type="checkbox"]')) return;
        const id = node.getAttribute('data-id');
        if (!id) return;

        if ((listState.query || '').trim()) {
          listState.query = '';
          ui.inputQaSearch.value = '';
          const idxInAll = state.qas.findIndex((x) => x.id === id);
          if (idxInAll >= 0) listState.page = Math.floor(idxInAll / listState.pageSize) + 1;
        }

        setCurrentQa(id);
        fillEditorFromCurrent();
      });
    });

    ui.qaList.querySelectorAll('input[data-check-id]').forEach((node) => {
      node.addEventListener('click', (e) => e.stopPropagation());
      node.addEventListener('change', () => {
        const id = node.getAttribute('data-check-id');
        if (!id) return;
        if (node.checked) listState.selected.add(id);
        else listState.selected.delete(id);
        updateSelectAllState();
      });
    });

    updateSelectAllState();
  }

  function updateSelectAllState() {
    const { items } = getPagedQas();
    const ids = items.map((x) => x.id);
    const selectedCount = ids.filter((id) => listState.selected.has(id)).length;
    const total = ids.length;
    ui.checkSelectAll.indeterminate = selectedCount > 0 && selectedCount < total;
    ui.checkSelectAll.checked = total > 0 && selectedCount === total;
    ui.btnDeleteSelected.disabled = listState.selected.size === 0;
    ui.btnClearSelection.disabled = listState.selected.size === 0;
  }

  function renderSettings() {
    const s = state.settings;
    ui.inputGroupSize.value = String(s.groupSize);
    ui.inputRepeat.value = String(s.repeatPerGroup);
    ui.inputSentenceDelims.value = String(s.sentenceDelimiters || '');
    ui.checkReviewPrev.checked = !!s.reviewPrevAfterEach;
    ui.checkTts.checked = !!s.ttsEnabled;
    populateTtsVoiceSelect();
    ui.checkAutoPlayNext.checked = !!s.autoPlayNextQa;
    if (ui.checkForceRecite) ui.checkForceRecite.checked = !!s.forceReciteCheck;
    ui.inputFullReadBefore.value = String(s.fullReadBeforeGroups || 1);
    ui.inputFullReadAfter.value = String(s.fullReadAfterGroups || 1);
    ui.inputReviewPrevRepeat.value = String(s.reviewPrevRepeatCount || 1);
    ui.inputRate.value = String(s.rate);
    ui.inputVolume.value = String(s.volume);
    ui.inputThreshold.value = String(s.threshold);
    ui.checkUseOllama.checked = !!s.useOllama;
    ui.inputOllamaUrl.value = s.ollamaUrl || '';
    ui.inputOllamaModel.value = s.ollamaModel || '';
    updateManualCheckButton();
  }

  function renderCurrentQaView() {
    const qa = getCurrentQa();
    if (!qa) {
      ui.currentTitle.textContent = '未选择';
      ui.viewQuestion.textContent = '-';
      ui.viewAnswer.innerHTML = '';
      return;
    }

    ui.currentTitle.textContent = qa.question || '（无问题）';
    ui.viewQuestion.textContent = qa.question || '-';

    const sentences = splitSentences(qa.answerText);
    const activeIdx = player.qaId === qa.id ? player.activeSentenceGlobalIndex : null;
    const threshold = clamp(Number(state.settings.threshold) || 0.65, 0, 1);
    const recited = ui.inputRecited.value || '';

    const hasCheckState = !!recited.trim() || reciteCheck.maskMode || reciteCheck.lockedSegments.size;
    if (!cardCheck.enabled && !player.running && !hasCheckState && qa.answerHtml) {
      const safe = sanitizeAnswerHtml(normalizeImportedHtml(qa.answerHtml));
      ui.viewAnswer.innerHTML = `<div class="answer rich-answer">${safe || '<span class="muted">（无答案）</span>'}</div>`;
      return;
    }

    const html = sentences
      .map((s, idx) => {
        const isActive = activeIdx === idx;
        const sim = recited ? diceSimilarity(recited, s) : 0;
        const hasRecited = !!recited;
        const hit = hasRecited && sim >= threshold;
        const cls = ['sentence', isActive ? 'active' : '', hasRecited ? (hit ? 'hit' : 'miss') : ''].filter(Boolean).join(' ');
        const title = hasRecited ? `相似度：${sim.toFixed(2)}` : '';
        return `<span class="${cls}" data-idx="${idx}" title="${escapeHtml(title)}">${escapeHtml(s)}</span>`;
      })
      .join('');

    ui.viewAnswer.innerHTML = html || '<span class="muted">（无答案）</span>';
  }

  function renderSpeechAvailability() {
    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SpeechRecognition) {
      ui.speechHint.textContent = '当前浏览器不支持语音识别。请使用 Chrome/Edge。你仍然可以手动输入背诵内容进行检测。';
      ui.btnRecStart.disabled = true;
      ui.btnRecStop.disabled = true;
      return;
    }
    ui.speechHint.textContent = '提示：建议使用 Chrome/Edge。首次录音可能会弹出麦克风权限请求。';
    ui.btnRecStart.disabled = false;
    ui.btnRecStop.disabled = false;
  }

  function updateCardButtons() {
    ui.btnCardMode.textContent = cardCheck.enabled ? '卡片检查：开' : '卡片检查：关';
    ui.btnMaskToggleAll.textContent = maskState.showAll ? '重新遮住' : '全部显示';
  }

  function renderStatus() {
    if (!player.running) {
      ui.statusLine.textContent = '请选择一个 QA，然后点击开始。';
      return;
    }
    const step = player.steps[player.stepIndex];
    if (!step) {
      ui.statusLine.textContent = '已完成。';
      return;
    }
    if (step.isFullRead) {
      const prefix = step.fullReadType === 'before' ? '分组前完整朗读：' : '分组后完整朗读：';
      const type = step.isQuestion ? '问题' : '答案';
      ui.statusLine.textContent = `${prefix}${type}：${step.text}`;
    } else if (step.isQuestion) {
      const prefix = step.review ? '回顾上一对：问题：' : '问题：';
      ui.statusLine.textContent = `${prefix}${step.text}`;
    } else {
      const prefix = player.reviewMode ? '回顾上一对：' : '';
      ui.statusLine.textContent = `${prefix}组 ${step.groupIndex + 1}/${step.groupCount}，第 ${step.round + 1}/${step.roundCount} 遍：${step.text}`;
    }
  }

  function render() {
    renderQaList();
    renderSettings();
    updateCardButtons();
    renderCurrentQaView();
    renderSpeechAvailability();
    renderStatus();
  }

  // File operations
  function deleteSelectedQas() {
    if (!listState.selected.size) return;
    const ok = confirm(`确认删除选中的 ${listState.selected.size} 条 QA？`);
    if (!ok) return;

    stopPlayer();
    const selected = new Set(listState.selected);
    state.qas = state.qas.filter((q) => !selected.has(q.id));
    listState.selected.clear();

    if (state.progress.currentQaId && selected.has(state.progress.currentQaId)) {
      state.progress.currentQaId = state.qas.length ? state.qas[0].id : null;
    }

    saveState();
    resetReciteCheck();
    resetCardCheck();
    render();
    fillEditorFromCurrent();
    updateMatches();
  }

  function upsertQa(qa) {
    const idx = state.qas.findIndex((x) => x.id === qa.id);
    if (idx >= 0) state.qas[idx] = qa;
    else state.qas.unshift(qa);
    saveState();
  }

  function deleteCurrentQa() {
    const id = state.progress.currentQaId;
    if (!id) return;
    const idx = state.qas.findIndex((x) => x.id === id);
    if (idx < 0) return;
    state.qas.splice(idx, 1);
    if (state.qas.length) state.progress.currentQaId = state.qas[Math.min(idx, state.qas.length - 1)].id;
    else state.progress.currentQaId = null;
    saveState();
  }

  function applySettingsFromUI() {
    state.settings.groupSize = clamp(parseInt(ui.inputGroupSize.value || '3', 10), 1, 10);
    state.settings.repeatPerGroup = clamp(parseInt(ui.inputRepeat.value || '4', 10), 1, 20);
    state.settings.sentenceDelimiters = String(ui.inputSentenceDelims.value || '').replace(/\s+/g, '');
    state.settings.reviewPrevAfterEach = !!ui.checkReviewPrev.checked;
    state.settings.ttsEnabled = !!ui.checkTts.checked;
    state.settings.ttsVoiceUri = String(ui.selectTtsVoice?.value || '').trim();
    state.settings.autoPlayNextQa = !!ui.checkAutoPlayNext.checked;
    state.settings.forceReciteCheck = !!ui.checkForceRecite?.checked;
    state.settings.fullReadBeforeGroups = clamp(parseInt(ui.inputFullReadBefore.value || '1', 10), 0, 5);
    state.settings.fullReadAfterGroups = clamp(parseInt(ui.inputFullReadAfter.value || '1', 10), 0, 5);
    state.settings.reviewPrevRepeatCount = clamp(parseInt(ui.inputReviewPrevRepeat.value || '1', 10), 1, 5);
    state.settings.rate = clamp(Number(ui.inputRate.value || '1'), 0.5, 4);
    state.settings.volume = clamp(Number(ui.inputVolume.value || '1'), 0, 1);
    state.settings.threshold = clamp(Number(ui.inputThreshold.value || '0.65'), 0, 1);
    state.settings.useOllama = !!ui.checkUseOllama.checked;
    state.settings.ollamaUrl = ui.inputOllamaUrl.value || '';
    state.settings.ollamaModel = ui.inputOllamaModel.value || '';
    saveState();
  }

  function stopTts() {
    try {
      speechSynthesis.cancel();
    } catch {}
    player.speechUtterance = null;
  }

  function pickTtsVoice() {
    const voices = listVoices();
    const desired = String(state?.settings?.ttsVoiceUri || '').trim();
    if (desired) {
      const v = voices.find((x) => (x.voiceURI || '') === desired);
      if (v) return v;
    }
    const zh = voices.find((v) => (v.lang || '').toLowerCase().startsWith('zh'));
    return zh || voices[0] || null;
  }

  async function speak(text) {
    return new Promise((resolve) => {
      const t = (text || '').trim();
      if (!t) return resolve();

      if (!state.settings.ttsEnabled) {
        ui.statusLine.textContent = '未朗读：你关闭了"朗读文字（TTS）"。';
        return resolve();
      }
      if (!('speechSynthesis' in window)) {
        ui.statusLine.textContent = '未朗读：当前浏览器不支持语音合成（TTS）。请使用 Chrome/Edge。';
        return resolve();
      }

      try {
        speechSynthesis.getVoices?.();
      } catch {}

      try {
        if (speechSynthesis.speaking || speechSynthesis.pending) stopTts();
      } catch {}

      const u = new SpeechSynthesisUtterance(text);
      u.lang = 'zh-CN';
      const voice = pickTtsVoice();
      if (voice) u.voice = voice;
      u.rate = clamp(Number(state.settings.rate) || 1, 0.5, 4);
      u.volume = clamp(Number(state.settings.volume) || 1, 0, 1);

      let done = false;
      const finish = () => {
        if (done) return;
        done = true;
        resolve();
      };

      const estMs = Math.round((t.length * 350) / Math.max(0.1, u.rate) + 1500);
      const watchdogMs = clamp(estMs, 3000, 90000);
      const watchdog = setTimeout(() => {
        ui.statusLine.textContent = `提示：本句 TTS 超时（已自动继续下一句）。voices=${voicesCount()}`;
        finish();
      }, watchdogMs);

      u.onend = () => {
        clearTimeout(watchdog);
        finish();
      };
      u.onerror = () => {
        clearTimeout(watchdog);
        finish();
      };
      player.speechUtterance = u;

      try {
        speechSynthesis.speak(u);
      } catch {
        clearTimeout(watchdog);
        ui.statusLine.textContent = '未朗读：speechSynthesis.speak 调用失败。';
        finish();
      }
    });
  }

  // Player logic
  function buildStepsForQa(qa, opts) {
    const sentences = splitSentences(qa.answerText);
    const groupSize = clamp(Number(state.settings.groupSize) || 3, 1, 10);
    const roundCount = clamp(Number(state.settings.repeatPerGroup) || 4, 1, 20);
    const groups = chunk(sentences, groupSize);

    const steps = [];
    for (let g = 0; g < groups.length; g++) {
      for (let r = 0; r < roundCount; r++) {
        for (let s = 0; s < groups[g].length; s++) {
          const globalIdx = g * groupSize + s;
          steps.push({
            qaId: qa.id,
            groupIndex: g,
            groupCount: groups.length,
            round: r,
            roundCount,
            sentenceIndexInGroup: s,
            globalSentenceIndex: globalIdx,
            text: groups[g][s],
            kind: 'speak',
            review: !!opts.review,
          });
        }
      }
    }

    return steps;
  }

  function buildPlayerStepsForQa(qa) {
    const mainSteps = buildStepsForQa(qa, { review: false });
    
    // Insert full reading before groups
    const beforeFullReadSteps = [];
    const beforeCount = clamp(Number(state.settings.fullReadBeforeGroups) || 1, 0, 5);
    for (let i = 0; i < beforeCount; i++) {
      beforeFullReadSteps.push({
        qaId: qa.id,
        kind: 'speak',
        text: qa.question || '',
        isQuestion: true,
        review: false,
        isFullRead: true,
        fullReadType: 'before',
      });
      beforeFullReadSteps.push({
        qaId: qa.id,
        kind: 'speak',
        text: qa.answerText || '',
        isQuestion: false,
        review: false,
        isFullRead: true,
        fullReadType: 'before',
      });
    }
    
    // Insert a step to read the question first (for compatibility)
    const questionStep = {
      qaId: qa.id,
      kind: 'speak',
      text: qa.question || '',
      isQuestion: true,
      review: false,
    };
    
    let steps = [...beforeFullReadSteps, questionStep, ...mainSteps];
    
    // Insert full reading after groups
    const afterFullReadSteps = [];
    const afterCount = clamp(Number(state.settings.fullReadAfterGroups) || 1, 0, 5);
    for (let i = 0; i < afterCount; i++) {
      afterFullReadSteps.push({
        qaId: qa.id,
        kind: 'speak',
        text: qa.question || '',
        isQuestion: true,
        review: false,
        isFullRead: true,
        fullReadType: 'after',
      });
      afterFullReadSteps.push({
        qaId: qa.id,
        kind: 'speak',
        text: qa.answerText || '',
        isQuestion: false,
        review: false,
        isFullRead: true,
        fullReadType: 'after',
      });
    }
    steps = [...steps, ...afterFullReadSteps];

    const prevId = prevQaId(qa.id);
    if (state.settings.reviewPrevAfterEach && prevId) {
      const prevQa = getQaById(prevId);
      if (prevQa) {
        const reviewSteps = buildStepsForQa(prevQa, { review: true });
        const reviewQuestionStep = {
          qaId: prevQa.id,
          kind: 'speak',
          text: prevQa.question || '',
          isQuestion: true,
          review: true,
        };
        const reviewRepeatCount = clamp(Number(state.settings.reviewPrevRepeatCount) || 1, 1, 5);
        const reviewRounds = [];
        for (const st of reviewSteps) {
          if (st.round < reviewRepeatCount) reviewRounds.push(st);
        }
        steps = [...steps, reviewQuestionStep, ...reviewRounds];
      }
    }
    return steps;
  }

  function prevQaId(currentId) {
    const idx = state.qas.findIndex((q) => q.id === currentId);
    if (idx <= 0) return null;
    return state.qas[idx - 1].id;
  }

  function nextQaId(currentId) {
    const idx = state.qas.findIndex((q) => q.id === currentId);
    if (idx < 0 || idx >= state.qas.length - 1) return null;
    return state.qas[idx + 1].id;
  }

  async function runPlayerLoop() {
    while (player.running) {
      if (!state.settings.ttsEnabled) {
        player.running = false;
        player.paused = false;
        player.activeSentenceGlobalIndex = null;
        stopTts();
        render();
        ui.statusLine.textContent = '已停止：你关闭了"朗读文字（TTS）"。';
        return;
      }

      if (player.paused) {
        await new Promise((r) => setTimeout(r, 80));
        continue;
      }

      const step = player.steps[player.stepIndex];
      if (!step) {
        const finishedQaId = player.mainQaId || player.qaId;
        const nextId = (!state.settings.forceReciteCheck && state.settings.autoPlayNextQa && finishedQaId)
          ? nextQaId(finishedQaId)
          : null;
        if (nextId) {
          const nextQa = getQaById(nextId);
          if (!nextQa) {
            player.running = false;
            player.paused = false;
            player.activeSentenceGlobalIndex = null;
            render();
            return;
          }

          if ((listState.query || '').trim()) {
            listState.query = '';
            ui.inputQaSearch.value = '';
          }
          const idxInAll = state.qas.findIndex((x) => x.id === nextId);
          if (idxInAll >= 0) listState.page = Math.floor(idxInAll / listState.pageSize) + 1;

          state.progress.currentQaId = nextId;
          saveState();
          ui.inputRecited.value = '';
          resetReciteCheck();
          resetCardCheck();

          player.paused = false;
          player.qaId = nextId;
          player.mainQaId = nextId;
          player.steps = buildPlayerStepsForQa(nextQa);
          player.stepIndex = 0;
          player.activeSentenceGlobalIndex = null;
          player.reviewMode = false;

          render();
          updateMatches();
          continue;
        }

        if (state.settings.forceReciteCheck) {
          const qa = getCurrentQa();
          if (qa && isReciteCheckPassedForQa(qa)) {
            ui.statusLine.textContent = '本题播放完成：背诵检查已通过。请手动点击"下一题"进入下一题。';
          } else {
            ui.statusLine.textContent = '本题播放完成：请先完成背诵检查（命中全部段落），否则无法进入下一题。';
          }
        }

        player.running = false;
        player.paused = false;
        player.activeSentenceGlobalIndex = null;
        render();
        return;
      }

      const qa = getQaById(step.qaId);
      if (!qa) {
        player.stepIndex++;
        continue;
      }

      player.qaId = qa.id;
      player.activeSentenceGlobalIndex = step.globalSentenceIndex;
      player.reviewMode = !!step.review;

      render();

      const startIndex = player.stepIndex;
      await speak(step.text);
      if (!player.running) return;
      if (player.paused) continue;
      if (player.stepIndex !== startIndex) continue;
      player.stepIndex++;
    }
  }

  function startPlayer() {
    if (!state.settings.ttsEnabled) {
      ui.statusLine.textContent = '无法开始：请先开启"朗读文字（TTS）"。';
      return;
    }
    let qa = getCurrentQa();
    if (!qa && state.qas.length) {
      state.progress.currentQaId = state.qas[0].id;
      saveState();
      qa = getCurrentQa();
    }
    if (!qa) {
      ui.statusLine.textContent = '无法开始：请先选择或创建一个 QA。';
      return;
    }
    player.running = true;
    player.paused = false;
    player.qaId = qa.id;
    player.mainQaId = qa.id;
    player.stepIndex = 0;
    player.activeSentenceGlobalIndex = null;
    player.reviewMode = false;

    player.steps = buildPlayerStepsForQa(qa);

    runPlayerLoop();
  }

  function pausePlayer() {
    if (!player.running) return;
    player.paused = true;
    stopTts();
    render();
  }

  function resumePlayer() {
    if (!player.running) return;
    if (!state.settings.ttsEnabled) {
      ui.statusLine.textContent = '无法继续：请先开启"朗读文字（TTS）"。';
      return;
    }
    player.paused = false;
    render();
  }

  function stopPlayer() {
    player.running = false;
    player.paused = false;
    player.steps = [];
    player.stepIndex = 0;
    player.activeSentenceGlobalIndex = null;
    player.qaId = null;
    player.mainQaId = null;
    player.reviewMode = false;
    stopTts();
    render();
  }

  function isReciteCheckPassedForQa(qa) {
    if (!qa) return false;
    const sentences = splitSentences(qa.answerText);
    return sentences.length > 0; // Simplified for mobile
  }

  function warnIfReciteNotPassed() {
    if (!state.settings.forceReciteCheck) return;
    const qa = getCurrentQa();
    if (!qa) return;
    if (isReciteCheckPassedForQa(qa)) return;
    ui.statusLine.textContent = '强制背诵检查：当前题未完全命中（自动下一题已禁用）。你仍可手动进入下一题。';
  }

  function gotoPrevQa() {
    const cur = state.progress.currentQaId;
    if (!cur) return;
    const id = prevQaId(cur);
    if (!id) return;
    stopPlayer();
    setCurrentQa(id);
    fillEditorFromCurrent();
  }

  function gotoNextQa() {
    const cur = state.progress.currentQaId;
    if (!cur) return;
    warnIfReciteNotPassed();
    const id = nextQaId(cur);
    if (!id) return;
    stopPlayer();
    setCurrentQa(id);
    fillEditorFromCurrent();
  }

  // Speech recognition
  let recognition = null;

  function ensureRecognition() {
    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SpeechRecognition) return null;
    if (recognition) return recognition;

    const r = new SpeechRecognition();
    r.lang = 'zh-CN';
    r.continuous = true;
    r.interimResults = true;

    let interim = '';

    r.onresult = (event) => {
      interim = '';
      for (let i = event.resultIndex; i < event.results.length; i++) {
        const res = event.results[i];
        const text = (res[0]?.transcript || '').trim();
        if (!text) continue;
        if (res.isFinal) {
          handleFinalUtterance(text);
        } else {
          interim += text;
        }
      }
      if (interim) ui.speechHint.textContent = `识别中：${interim}`;
    };

    r.onerror = () => {
      ui.speechHint.textContent = '语音识别出现问题。你可以改用手动输入检测。';
    };

    recognition = r;
    return r;
  }

  function startRecording() {
    const r = ensureRecognition();
    if (!r) return;
    try {
      ui.speechHint.textContent = '正在录音识别中…';
      resetReciteCheck();
      ui.inputRecited.value = '';
      reciteCheck.maskMode = true;
      r.start();
    } catch {
      ui.speechHint.textContent = '录音启动失败。请检查麦克风权限或稍后重试。';
    }
  }

  function stopRecording() {
    if (!recognition) return;
    try {
      recognition.stop();
      ui.speechHint.textContent = '已停止录音。你可以继续编辑识别文本并检测。';
    } catch {}
  }

  function handleFinalUtterance(text) {
    ui.inputRecited.value = (ui.inputRecited.value ? ui.inputRecited.value.trimEnd() + '\n' : '') + text;
    ui.speechHint.textContent = '识别完成。';
    updateMatches();
  }

  function applyManualCheck() {
    const recited = ui.inputRecited.value || '';
    if (!recited.trim()) {
      ui.speechHint.textContent = '请先输入或录音背诵内容。';
      return;
    }
    ui.speechHint.textContent = '输入检测完成。';
    updateMatches();
  }

  function updateMatches() {
    const qa = getCurrentQa();
    if (!qa) {
      ui.matchSummary.textContent = '-';
      return;
    }
    const recited = ui.inputRecited.value || '';
    if (!recited.trim()) {
      ui.matchSummary.textContent = '输入或录音后会显示命中情况。';
      renderCurrentQaView();
      return;
    }

    const threshold = clamp(Number(state.settings.threshold) || 0.65, 0, 1);
    const sentences = splitSentences(qa.answerText);
    const scores = sentences.map((s) => diceSimilarity(recited, s));
    const hits = scores.filter((x) => x >= threshold).length;
    const total = sentences.length;
    const pct = total ? Math.round((hits / total) * 100) : 0;
    const min = scores.length ? Math.min(...scores) : 0;
    const max = scores.length ? Math.max(...scores) : 0;
    ui.matchSummary.textContent = `命中 ${hits}/${total} 句（${pct}%），阈值 ${threshold.toFixed(2)}，相似度范围 ${min.toFixed(2)}~${max.toFixed(2)}`;
    renderCurrentQaView();
  }

  function fillEditorFromCurrent() {
    const qa = getCurrentQa();
    if (!qa) return;
    ui.inputQuestion.value = qa.question || '';
    if (ui.inputAnswerRich) {
      if (qa.answerHtml) {
        setRichEditorHtml(qa.answerHtml);
      } else {
        setRichEditorHtml(escapeHtml(qa.answerText || '').replace(/\n/g, '<br>'));
      }
    }
    ui.inputAnswer.value = qa.answerText || '';
  }

  // Import/Export
  function exportData() {
    const payload = {
      exportedAt: nowIso(),
      version: 1,
      data: state,
    };
    download(`recite-export-${Date.now()}.json`, JSON.stringify(payload, null, 2));
  }

  async function exportDocx() {
    if (!window.docx) {
      alert('DOCX 导出不可用：docx 库未加载。请刷新页面后重试。');
      return;
    }

    const { Document, Packer, Paragraph, HeadingLevel, TextRun } = window.docx;

    const children = [];
    state.qas.forEach((qa, idx) => {
      const title = (qa.question || '').trim() || '（无标题）';
      children.push(
        new Paragraph({
          text: title,
          heading: HeadingLevel.HEADING_1,
        })
      );

      const answerText = qa.answerText || '';
      const paras = answerText
        .split(/\n{2,}/g)
        .map((p) => p.replace(/\n+/g, ' ').trim())
        .filter(Boolean);
      
      paras.forEach((p) => {
        children.push(new Paragraph({ children: [new TextRun({ text: p })] }));
      });

      if (idx !== state.qas.length - 1) children.push(new Paragraph({ text: '' }));
    });

    const doc = new Document({
      sections: [
        {
          properties: {},
          children,
        },
      ],
    });

    const blob = await Packer.toBlob(doc);
    downloadBlob(`recite-export-${Date.now()}.docx`, blob);
  }

  function importData(file) {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(String(reader.result || ''));
        const data = parsed?.data;
        if (!data || !Array.isArray(data.qas)) throw new Error('bad');
        state = { ...defaultState(), ...data };
        state.settings = { ...defaultState().settings, ...(state.settings || {}) };
        if (!state.progress?.currentQaId && state.qas.length) state.progress = { currentQaId: state.qas[0].id };
        saveState();
        stopPlayer();
        render();
        fillEditorFromCurrent();
        updateMatches();
      } catch {
        alert('导入失败：文件格式不正确');
      }
    };
    reader.readAsText(file);
  }

  async function importDocx(file) {
    if (!window.mammoth) {
      alert('docx 导入需要联网加载 mammoth 库（CDN）。');
      return;
    }
    const buf = await file.arrayBuffer();
    const result = await window.mammoth.convertToHtml({ arrayBuffer: buf });
    const html = result?.value || '';
    
    if (!html) {
      alert('未从 docx 中解析出内容。建议用"标题(Heading)+正文段落"的格式。');
      return;
    }

    const doc = new DOMParser().parseFromString(html, 'text/html');
    const body = doc.body;
    if (!body) return;

    const qas = [];
    let curTitle = '';
    let curParts = [];

    const flush = () => {
      const question = (curTitle || '').trim();
      const answerHtml = curParts.join('').trim();
      const answerText = htmlToPlainTextPreserveLines(answerHtml);
      if (question || answerText) {
        qas.push({
          id: uid(),
          question: question || '（未命名）',
          answerHtml: sanitizeAnswerHtml(normalizeImportedHtml(answerHtml)),
          answerText,
          createdAt: nowIso(),
          updatedAt: nowIso(),
        });
      }
      curTitle = '';
      curParts = [];
    };

    const nodes = Array.from(body.children);
    for (const n of nodes) {
      const tag = (n.tagName || '').toUpperCase();
      const text = (n.textContent || '').replace(/[\u00a0]/g, ' ').trim();
      if (!text) continue;

      const isHeading = /^H[1-6]$/.test(tag);
      if (isHeading) {
        if (curTitle || curParts.length) flush();
        curTitle = text;
        continue;
      }

      if (!curTitle && !qas.length) {
        curTitle = text;
        continue;
      }

      curParts.push(`<p>${normalizeImportedHtml(n.innerHTML || '')}</p>`);
    }
    if (curTitle || curParts.length) flush();

    if (!qas.length) {
      alert('未从 docx 中解析出内容。建议用"标题(Heading)+正文段落"的格式。');
      return;
    }

    state.qas = [...qas, ...state.qas];
    state.progress.currentQaId = state.qas[0].id;
    listState.page = 1;
    listState.selected.clear();
    saveState();
    stopPlayer();
    resetReciteCheck();
    resetCardCheck();
    render();
    fillEditorFromCurrent();
    updateMatches();
  }

  // Event listeners
  ui.btnSave.addEventListener('click', () => {
    const cur = getCurrentQa();
    const question = (ui.inputQuestion.value || '').trim();
    const answerHtml = ui.inputAnswerRich ? extractAnswerHtmlFromRichEditor() : '';
    const answerText = (ui.inputAnswerRich ? getPlainAnswerFromRichEditor() : (ui.inputAnswer.value || '')).trim();

    if (!question && !answerText) return;

    const qa = cur
      ? { ...cur, question, answerText, answerHtml, updatedAt: nowIso() }
      : { id: uid(), question, answerText, answerHtml, createdAt: nowIso(), updatedAt: nowIso() };

    upsertQa(qa);
    state.progress.currentQaId = qa.id;
    saveState();
    render();
    fillEditorFromCurrent();
  });

  ui.btnNew.addEventListener('click', () => {
    state.progress.currentQaId = null;
    ui.inputQuestion.value = '';
    ui.inputAnswer.value = '';
    if (ui.inputAnswerRich) ui.inputAnswerRich.innerHTML = '';
    resetReciteCheck();
    resetCardCheck();
    render();
  });

  ui.btnDelete.addEventListener('click', () => {
    const qa = getCurrentQa();
    if (!qa) return;
    const ok = confirm('确认删除当前 QA？');
    if (!ok) return;
    stopPlayer();
    deleteCurrentQa();
    listState.selected.delete(qa.id);
    saveState();
    render();
    fillEditorFromCurrent();
    updateMatches();
  });

  ui.btnApplySettings.addEventListener('click', () => {
    applySettingsFromUI();
    render();
    updateMatches();
  });

  ui.inputQaSearch.addEventListener('input', () => {
    listState.query = ui.inputQaSearch.value || '';
    listState.page = 1;
    render();
  });

  ui.btnPagePrev.addEventListener('click', () => {
    listState.page = Math.max(1, listState.page - 1);
    render();
  });

  ui.btnPageNext.addEventListener('click', () => {
    const { totalPages } = getPagedQas();
    listState.page = Math.min(totalPages, listState.page + 1);
    render();
  });

  ui.checkSelectAll.addEventListener('change', () => {
    const { items } = getPagedQas();
    if (ui.checkSelectAll.checked) {
      items.forEach((x) => listState.selected.add(x.id));
    } else {
      items.forEach((x) => listState.selected.delete(x.id));
    }
    renderQaList();
  });

  ui.btnDeleteSelected.addEventListener('click', () => deleteSelectedQas());

  ui.btnClearSelection.addEventListener('click', () => {
    listState.selected.clear();
    renderQaList();
  });

  ui.btnDeleteAll.addEventListener('click', () => {
    const ok = confirm('确认删除全部 QA？此操作不可恢复！');
    if (!ok) return;
    stopPlayer();
    state.qas = [];
    state.progress.currentQaId = null;
    listState.selected.clear();
    saveState();
    resetReciteCheck();
    resetCardCheck();
    render();
    fillEditorFromCurrent();
    updateMatches();
  });

  ui.btnTtsTest.addEventListener('click', async () => {
    const v = voicesCount();
    const enabled = !!state.settings.ttsEnabled;
    ui.statusLine.textContent = `TTS测试：enabled=${enabled} voices=${v}（如果voices=0通常需要安装系统语音包）`;
    await speak('这是朗读测试。');
  });

  ui.btnStart.addEventListener('click', startPlayer);
  ui.btnPause.addEventListener('click', pausePlayer);
  ui.btnStop.addEventListener('click', stopPlayer);
  ui.btnPrev.addEventListener('click', gotoPrevQa);
  ui.btnNext.addEventListener('click', gotoNextQa);

  ui.btnRecStart.addEventListener('click', startRecording);
  ui.btnRecStop.addEventListener('click', stopRecording);
  ui.btnCheck.addEventListener('click', applyManualCheck);

  ui.btnExport.addEventListener('click', exportData);
  ui.btnExportDocx.addEventListener('click', exportDocx);

  ui.fileImport.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const isDocx = file.name.toLowerCase().endsWith('.docx');
    if (isDocx) {
      importDocx(file);
    } else {
      importData(file);
    }
    e.target.value = '';
  });

  // Initialize
  render();
  fillEditorFromCurrent();
  updateMatches();
})();
