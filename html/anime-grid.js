const htmlEl = document.documentElement;
const EMOJI_FONT = '"Apple Color Emoji", "Segoe UI Emoji", "Noto Color Emoji", "Microsoft YaHei UI", "PingFang SC", sans-serif';
const TEXT_FONT = '"Microsoft YaHei UI", "PingFang SC", sans-serif';
const STAGE_W = 720;
const STAGE_H = 460;
const STAGE_PAD = 18;
const CELL_HEIGHT = 187;
const CELL_WIDTH = 120;
const PROJECT_URL = 'https://zmjkk1.github.io/anime-grid-emoji/html';
const clamp = (v, min, max) => Math.min(max, Math.max(min, v));
const uid = () => `emoji-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
const escapeHTML = text => String(text || '').replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]));
const clipText = (text, max = 18) => {
    const value = String(text || '').trim();
    return value.length > max ? `${value.slice(0, max)}…` : value;
};
const INSERT_OFFSETS = [
    { x: 0, y: 0 },
    { x: -0.12, y: -0.08 },
    { x: 0.12, y: -0.08 },
    { x: -0.14, y: 0.1 },
    { x: 0.14, y: 0.1 },
    { x: 0, y: -0.16 },
    { x: 0, y: 0.18 },
];

const getGraphemes = text => {
    if (!text) return [];
    if (window.Intl && Intl.Segmenter) {
        const segmenter = new Intl.Segmenter('zh-Hans', { granularity: 'grapheme' });
        return Array.from(segmenter.segment(text), item => item.segment);
    }
    return Array.from(text);
};

const createEmojiItem = (emoji, index = 0) => {
    const offset = INSERT_OFFSETS[index % INSERT_OFFSETS.length] || { x: 0, y: 0 };
    return { id: uid(), type: 'emoji', emoji, x: clamp(0.5 + offset.x, 0.1, 0.9), y: clamp(0.48 + offset.y, 0.08, 0.92), size: 0.19 };
};

const createTextItem = (text, index = 0) => {
    const offset = INSERT_OFFSETS[index % INSERT_OFFSETS.length] || { x: 0, y: 0 };
    return { id: uid(), type: 'text', text: String(text || '').trim(), x: clamp(0.5 + offset.x, 0.14, 0.86), y: clamp(0.48 + offset.y, 0.12, 0.88), size: 0.13 };
};

const createItemsFromText = clue => {
    const lines = String(clue || '')
        .split(/\n+/g)
        .map(line => line.trim())
        .filter(Boolean)
        .map(line => (/\s/.test(line) ? line.split(/\s+/g).filter(Boolean) : getGraphemes(line).filter(Boolean)));
    if (!lines.length) return [];
    const rows = lines.length;
    const cols = Math.max(...lines.map(line => line.length));
    const size = clamp(0.72 / Math.max(cols + 0.6, rows + 0.8), 0.11, 0.24);
    const items = [];
    lines.forEach((line, row) => {
        line.forEach((emoji, col) => {
            items.push({ id: uid(), emoji, x: (col + 1) / (line.length + 1), y: (row + 1) / (rows + 1), size });
        });
    });
    return items;
};

const createEmptyEntry = () => ({ clue: '', answer: '', items: [] });
const createCellTitle = index => `新格子${index + 1}`;
const normalizeItem = item => {
    if (!item || typeof item !== 'object') return null;
    const type = item.type === 'text' ? 'text' : 'emoji';
    const emoji = String(item.emoji || '').trim();
    const text = String(item.text || '').trim();
    if (type === 'emoji' && !emoji) return null;
    if (type === 'text' && !text) return null;
    return {
        id: String(item.id || uid()),
        type,
        emoji,
        text,
        x: clamp(Number(item.x), 0.05, 0.95),
        y: clamp(Number(item.y), 0.05, 0.95),
        size: clamp(Number(item.size), 0.08, 0.6),
    };
};
const normalizeEntry = value => {
    if (!value) return createEmptyEntry();
    if (typeof value === 'string') {
        const clue = value.trim();
        return !clue || /^\d+$/.test(clue) ? createEmptyEntry() : { clue, answer: '', items: createItemsFromText(clue) };
    }
    if (typeof value === 'object') {
        const clue = String(value.clue || '').trim();
        const items = Array.isArray(value.items) ? value.items.map(normalizeItem).filter(Boolean) : [];
        return { clue, answer: String(value.answer || '').trim(), items: items.length ? items : createItemsFromText(clue) };
    }
    return createEmptyEntry();
};
const cloneEntry = entry => ({ clue: String(entry.clue || ''), answer: String(entry.answer || ''), items: (entry.items || []).map(item => ({ ...item })) });
const getItemStoredText = item => (item.type === 'text' ? String(item.text || '').trim() : String(item.emoji || '').trim());
const serializeEntry = entry => ({
    clue: (entry.items || []).map(getItemStoredText).join(' '),
    answer: String(entry.answer || '').trim(),
    items: (entry.items || []).map(item => ({
        id: item.id,
        type: item.type === 'text' ? 'text' : 'emoji',
        emoji: item.emoji,
        text: item.text,
        x: Math.round(item.x * 1000) / 1000,
        y: Math.round(item.y * 1000) / 1000,
        size: Math.round(item.size * 1000) / 1000,
    })),
});

class AnimeGrid {
    constructor({ el, title, key, typeTexts, col, row, urlExt = '' }) {
        this.el = el;
        this.title = title;
        this.key = key;
        this.col = col;
        this.row = row;
        this.urlExt = urlExt;
        this.defaultTypes = typeTexts.trim().split(/\n+/g);
        this.types = [...this.defaultTypes];
        this.entries = [];
        this.currentIndex = null;
        this.currentEntryDraft = createEmptyEntry();
        this.showAnswers = false;
        this.selectedStageItemId = null;
        this.dragState = null;
        this.sortDraft = [];
        this.sortDragIndex = null;
        this.sortDropIndex = null;
        this.sortDropAfter = false;
        this.loadEntriesFromLocalStorage();
        el.innerHTML = this.generatorHTML();
        this.canvas = el.querySelector('.main-grid-canvas');
        this.ctx = this.canvas.getContext('2d');
        this.editorStageCanvas = el.querySelector('.editor-stage');
        this.editorStageCtx = this.editorStageCanvas.getContext('2d');
        this.downloadBtnEl = el.querySelector('.download-btn');
        this.answerToggleEl = el.querySelector('.toggle-answers-btn');
        this.answerToggleTextEl = el.querySelector('.toggle-answer-text');
        this.helperTextEl = el.querySelector('.helper-text');
        this.editorEl = el.querySelector('.editor-box');
        this.categoryNameEl = el.querySelector('.current-category');
        this.titleInputEl = el.querySelector('.title-input');
        this.answerInputEl = el.querySelector('.answer-input');
        this.selectionStatusEl = el.querySelector('.selection-status');
        this.sizeRangeEl = el.querySelector('.size-range');
        this.selectedTextFieldEl = el.querySelector('.selected-text-field');
        this.selectedTextInputEl = el.querySelector('.selected-text-input');
        this.emojiPickerEl = el.querySelector('.emoji-picker-box');
        this.emojiMartRootEl = el.querySelector('.emoji-mart-root');
        this.sortEl = el.querySelector('.sort-box');
        this.sortListEl = el.querySelector('.sort-list');
        this.outputEl = el.querySelector('.output-box');
        this.outputImageEl = this.outputEl.querySelector('img');
        this.emojiMartPicker = null;
        this.setupCanvas();
        this.bindEvents();
        this.updateUIState();
        this.renderEmojiPicker();
        this.draw();
    }

    generatorHTML() {
        return `<canvas class="main-grid-canvas"></canvas>
<div class="ctrl-box">
    <a class="ui-btn current download-btn" action="downloadImage">导出猜番图</a>
    <a class="ui-btn ghost" action="addCell">新增格子</a>
    <a class="ui-btn ghost" action="openSortBox">调整顺序</a>
    <a class="ui-btn ghost toggle-answers-btn" action="toggleAnswers"><span class="toggle-answer-text">显示答案</span></a>
    <a class="ui-btn ghost" action="clearAll">清空全部</a>
</div>
<p class="helper-text"></p>
<div class="editor-box ui-shadow" data-show="false">
    <div class="content-box editor-content-box">
        <div class="panel-head">
            <p class="panel-kicker">当前编辑</p>
            <h3 class="current-category"></h3>
        </div>
        <form class="editor-form">
            <label class="field">
                <span>格子标题</span>
                <input class="title-input" type="text" placeholder="例如：最喜欢">
            </label>
            <div class="editor-stage-box"><canvas class="editor-stage" width="${STAGE_W}" height="${STAGE_H}"></canvas></div>
            <p class="editor-stage-tip">拖动 emoji 或文本框自由组合；鼠标滚轮和下方滑杆都可以缩放。点中元素后可以删除、放大、缩小。</p>
            <div class="editor-toolbar">
                <a class="ui-btn ghost small" action="openEmojiPicker">插入 Emoji</a>
                <a class="ui-btn ghost small" action="insertTextBox">插入文本</a>
                <a class="ui-btn ghost small" action="scaleSelectionDown">缩小</a>
                <a class="ui-btn ghost small" action="scaleSelectionUp">放大</a>
                <a class="ui-btn ghost small" action="bringSelectionToFront">置于最上</a>
                <a class="ui-btn ghost small" action="deleteSelected">删除选中</a>
            </div>
            <p class="selection-status"></p>
            <label class="field size-field">
                <span>选中元素大小</span>
                <input class="size-range" type="range" min="8" max="60" step="1" value="19">
            </label>
            <label class="field selected-text-field" data-show="false">
                <span>文本框内容</span>
                <textarea class="selected-text-input" rows="3" placeholder="输入你想放进格子里的文字"></textarea>
            </label>
            <label class="field">
                <span>答案备注（可选）</span>
                <input class="answer-input" type="text" placeholder="例如：死亡笔记">
            </label>
        </form>
        <div class="foot foot-inline">
            <a class="ui-btn ghost" action="deleteCurrentCell">删除这个格子</a>
            <a class="ui-btn ghost" action="clearCurrent">清空这一格</a>
            <a class="ui-btn ghost" action="closeEditor">取消</a>
            <a class="ui-btn current" action="saveCurrent">保存</a>
        </div>
    </div>
</div>
<div class="emoji-picker-box ui-shadow" data-show="false">
    <div class="content-box emoji-content-box">
        <div class="panel-head emoji-panel-head">
            <div>
                <p class="panel-kicker">Emoji Mart</p>
                <h3 class="current-category">用官方分类和搜索来选 emoji</h3>
            </div>
            <a class="ui-btn ghost small emoji-close-btn" action="closeEmojiPicker">收起</a>
        </div>
        <div class="emoji-mart-root"></div>
        <div class="emoji-action-bar">
            <a class="ui-btn ghost small" action="scaleSelectionDown">缩小选中</a>
            <a class="ui-btn ghost small" action="scaleSelectionUp">放大选中</a>
            <a class="ui-btn current small" action="closeEmojiPicker">完成</a>
        </div>
    </div>
</div>
<div class="sort-box ui-shadow" data-show="false">
    <div class="content-box sort-content-box">
        <div class="panel-head emoji-panel-head">
            <div>
                <p class="panel-kicker">格子顺序</p>
                <h3 class="current-category">拖拽卡片，重新安排整张猜番表的顺序</h3>
            </div>
            <a class="ui-btn ghost small" action="closeSortBox">收起</a>
        </div>
        <div class="sort-body">
            <p class="sort-tip">按住卡片拖动即可换位；保存后，标题、emoji 排版和答案备注会一起跟着移动。</p>
            <div class="sort-list"></div>
        </div>
        <div class="foot foot-inline sort-foot">
            <a class="ui-btn ghost" action="closeSortBox">取消</a>
            <a class="ui-btn current" action="saveSortOrder">保存顺序</a>
        </div>
    </div>
</div>
<div class="output-box ui-shadow" data-show="false">
    <div class="content-box">
        <h3>图片已经生成</h3>
        <img alt="导出的猜番图">
        <p class="output-tip">如果你要发到社交平台，直接保存这张图片就可以了。</p>
        <div class="foot"><a class="ui-btn current" action="closeOutput">关闭</a></div>
    </div>
</div>`;
    }

    setupCanvas() {
        const bodyMargin = 20;
        const titleHeight = 54;
        const footerHeight = 20;
        const contentWidth = this.col * CELL_WIDTH;
        const contentHeight = this.getGridRowCount() * CELL_HEIGHT;
        const colWidth = Math.ceil(contentWidth / this.col);
        const rowHeight = CELL_HEIGHT;
        const labelHeight = 28;
        const width = contentWidth + bodyMargin * 2;
        const height = contentHeight + bodyMargin * 2 + titleHeight + footerHeight;
        const scale = Math.max(2, Math.ceil(window.devicePixelRatio || 1));
        this.bodyMargin = bodyMargin;
        this.titleHeight = titleHeight;
        this.colWidth = colWidth;
        this.rowHeight = rowHeight;
        this.labelHeight = labelHeight;
        this.imageWidth = colWidth - 2;
        this.imageHeight = rowHeight - labelHeight - 2;
        this.canvasScale = scale;
        this.width = width;
        this.height = height;
        this.canvas.width = width * scale;
        this.canvas.height = height * scale;
    }

    getGridRowCount() {
        return Math.max(1, Math.ceil(this.types.length / this.col));
    }

    bindEvents() {
        this.canvas.onclick = event => {
            const index = this.getIndexFromCanvasEvent(event);
            if (index !== null) this.openEditor(index);
        };
        this.el.onclick = event => {
            const target = event.target.closest('[action]');
            const action = target && target.getAttribute('action');
            if (action && typeof this[action] === 'function') this[action]();
        };
        this.editorEl.onclick = event => {
            if (event.target === this.editorEl) this.closeEditor();
        };
        this.emojiPickerEl.onclick = event => {
            if (event.target === this.emojiPickerEl) this.closeEmojiPicker();
        };
        this.sortEl.onclick = event => {
            if (event.target === this.sortEl) this.closeSortBox();
        };
        this.outputEl.onclick = event => {
            if (event.target === this.outputEl) this.closeOutput();
        };
        this.sortListEl.addEventListener('dragstart', this.onSortDragStart.bind(this));
        this.sortListEl.addEventListener('dragover', this.onSortDragOver.bind(this));
        this.sortListEl.addEventListener('drop', this.onSortDrop.bind(this));
        this.sortListEl.addEventListener('dragend', this.onSortDragEnd.bind(this));
        this.sortListEl.addEventListener('dragleave', this.onSortDragLeave.bind(this));
        this.editorStageCanvas.addEventListener('pointerdown', this.onStagePointerDown.bind(this));
        this.editorStageCanvas.addEventListener('pointermove', this.onStagePointerMove.bind(this));
        this.editorStageCanvas.addEventListener('pointerup', this.onStagePointerUp.bind(this));
        this.editorStageCanvas.addEventListener('pointercancel', this.onStagePointerUp.bind(this));
        this.editorStageCanvas.addEventListener('wheel', this.onStageWheel.bind(this), { passive: false });
        this.sizeRangeEl.addEventListener('input', event => this.setSelectedItemSize(Number(event.target.value) / 100));
        this.selectedTextInputEl.addEventListener('input', event => this.updateSelectedText(event.target.value));
        this.titleInputEl.addEventListener('input', event => {
            this.categoryNameEl.textContent = event.target.value.trim() || '未命名格子';
        });
        this.el.querySelector('.editor-form').onsubmit = event => {
            event.preventDefault();
            this.saveCurrent();
        };
    }
    getSceneArea() {
        return { left: STAGE_PAD, top: STAGE_PAD, width: this.editorStageCanvas.width - STAGE_PAD * 2, height: this.editorStageCanvas.height - STAGE_PAD * 2 };
    }

    getIndexFromCanvasEvent(event) {
        const rect = this.canvas.getBoundingClientRect();
        const clickX = (event.clientX - rect.left) / rect.width * this.width;
        const clickY = (event.clientY - rect.top) / rect.height * this.height;
        const x = Math.floor((clickX - this.bodyMargin) / this.colWidth);
        const gridY = Math.floor((clickY - this.bodyMargin - this.titleHeight) / this.rowHeight);
        if (x < 0 || gridY < 0 || x >= this.col || gridY >= this.getGridRowCount()) return null;
        return gridY * this.col + x;
    }

    getStagePoint(event) {
        const rect = this.editorStageCanvas.getBoundingClientRect();
        return { x: (event.clientX - rect.left) / rect.width * this.editorStageCanvas.width, y: (event.clientY - rect.top) / rect.height * this.editorStageCanvas.height };
    }

    loadEntriesFromLocalStorage() {
        const size = this.types.length;
        this.entries = new Array(size).fill(null).map(() => createEmptyEntry());
        if (!window.localStorage) return;
        const rawText = localStorage.getItem(this.key);
        if (!rawText) return;
        try {
            const rawState = JSON.parse(rawText);
            if (Array.isArray(rawState)) {
                this.entries = new Array(size).fill(null).map((_, index) => normalizeEntry(rawState[index]));
                return;
            }

            if (!rawState || typeof rawState !== 'object') return;

            if (Array.isArray(rawState.types) && rawState.types.length) {
                this.types = rawState.types.map((type, index) => {
                    const nextType = String(type || '').trim();
                    return nextType || createCellTitle(index);
                });
            }

            const nextSize = this.types.length;
            this.entries = new Array(nextSize).fill(null).map((_, index) => normalizeEntry(rawState.entries && rawState.entries[index]));
        } catch (error) {
            const legacyEntries = rawText.split(/,/g).map(normalizeEntry);
            this.entries = new Array(size).fill(null).map((_, index) => normalizeEntry(legacyEntries[index]));
        }
    }

    saveEntriesToLocalStorage() {
        if (!window.localStorage) return;
        localStorage.setItem(this.key, JSON.stringify({
            types: this.types,
            entries: this.entries.map(serializeEntry),
        }));
    }

    renderEmojiPicker() {
        if (this.emojiMartPicker || !window.EmojiMart || !this.emojiMartRootEl) return;

        const picker = new window.EmojiMart.Picker({
            data: async () => {
                const response = await fetch('https://cdn.jsdelivr.net/npm/@emoji-mart/data');
                return response.json();
            },
            onEmojiSelect: emoji => {
                const nativeEmoji = emoji && emoji.native;
                if (nativeEmoji) this.insertEmoji(nativeEmoji);
            },
            locale: 'zh',
            set: 'native',
            theme: 'light',
            dynamicWidth: true,
            searchPosition: 'sticky',
            skinTonePosition: 'search',
            previewPosition: 'bottom',
            navPosition: 'top',
            categories: ['frequent', 'people', 'nature', 'foods', 'activity', 'places', 'objects', 'symbols', 'flags'],
        });

        this.emojiMartPicker = picker;
        this.emojiMartRootEl.innerHTML = '';
        this.emojiMartRootEl.appendChild(picker);
    }

    updateUIState() {
        if (this.downloadBtnEl) this.downloadBtnEl.textContent = this.showAnswers ? '导出答案版' : '导出猜番图';
        if (this.answerToggleTextEl) this.answerToggleTextEl.textContent = this.showAnswers ? '隐藏答案' : '显示答案';
        if (this.answerToggleEl) {
            this.answerToggleEl.classList.toggle('current', this.showAnswers);
            this.answerToggleEl.classList.toggle('ghost', !this.showAnswers);
        }
        if (this.helperTextEl) {
            this.helperTextEl.textContent = this.showAnswers
                ? '当前是答案显示模式，导出时会带上答案备注。'
                : '点击任意格子进入编辑；你可以改标题、删除格子、新增格子，或在“调整顺序”里拖拽重排。现在既能摆 emoji，也能插入可拖动缩放的文本框。';
        }
    }

    syncPageLock() {
        const shouldLock = [this.editorEl, this.emojiPickerEl, this.sortEl, this.outputEl].some(el => el && el.getAttribute('data-show') === 'true');
        htmlEl.setAttribute('data-no-scroll', shouldLock ? 'true' : 'false');
    }

    getSelectedItem() {
        return this.currentEntryDraft.items.find(item => item.id === this.selectedStageItemId) || null;
    }

    updateSelectionUI() {
        const item = this.getSelectedItem();
        this.sizeRangeEl.disabled = !item;
        this.sizeRangeEl.value = item ? String(Math.round(item.size * 100)) : '19';
        const isTextItem = item && item.type === 'text';
        this.selectedTextFieldEl.setAttribute('data-show', isTextItem ? 'true' : 'false');
        const nextTextValue = isTextItem ? item.text : '';
        if (this.selectedTextInputEl.value !== nextTextValue) this.selectedTextInputEl.value = nextTextValue;
        this.selectionStatusEl.textContent = item
            ? (isTextItem
                ? `已选中文本框“${clipText(item.text, 14)}”。拖动它调整位置，滚轮或滑杆可以缩放，下方还能直接改文字。`
                : `已选中 ${item.emoji}。拖动它调整位置，滚轮或滑杆可以缩放。`)
            : '还没有选中任何元素。点舞台里的 emoji 或文本框选中，点空白区域取消选中。';
    }

    openEmojiPicker() {
        if (this.currentIndex === null) return;
        this.renderEmojiPicker();
        this.emojiPickerEl.setAttribute('data-show', 'true');
        this.syncPageLock();
    }

    closeEmojiPicker() {
        this.emojiPickerEl.setAttribute('data-show', 'false');
        this.syncPageLock();
    }

    createSortDraft() {
        return this.types.map((type, index) => ({
            id: uid(),
            type,
            entry: cloneEntry(normalizeEntry(this.entries[index])),
        }));
    }

    renderSortList() {
        if (!this.sortListEl) return;
        this.sortListEl.innerHTML = this.sortDraft.map((item, index) => {
            const elementCount = item.entry.items.length;
            const answerText = item.entry.answer ? `答案：${item.entry.answer}` : '未填写答案';
            return `<div class="sort-card" draggable="true" data-index="${index}">
    <div class="sort-order-badge">${index + 1}</div>
    <div class="sort-card-copy">
        <p class="sort-card-title">${escapeHTML(item.type)}</p>
        <p class="sort-card-meta">${elementCount ? `${elementCount} 个元素` : '还没有元素'} · ${escapeHTML(answerText)}</p>
    </div>
    <div class="sort-grip" aria-hidden="true">⋮⋮</div>
</div>`;
        }).join('');
        this.syncSortCardClasses();
    }

    syncSortCardClasses() {
        if (!this.sortListEl) return;
        Array.from(this.sortListEl.children).forEach(card => {
            const index = Number(card.dataset.index);
            card.classList.toggle('dragging', index === this.sortDragIndex);
            card.classList.toggle('drop-before', index === this.sortDropIndex && !this.sortDropAfter);
            card.classList.toggle('drop-after', index === this.sortDropIndex && this.sortDropAfter);
        });
    }

    openSortBox() {
        if (this.types.length < 2) {
            window.alert('至少需要两个格子才能调整顺序。');
            return;
        }
        this.sortDraft = this.createSortDraft();
        this.sortDragIndex = null;
        this.sortDropIndex = null;
        this.sortDropAfter = false;
        this.renderSortList();
        this.sortEl.setAttribute('data-show', 'true');
        this.syncPageLock();
    }

    closeSortBox() {
        this.sortEl.setAttribute('data-show', 'false');
        this.sortDragIndex = null;
        this.sortDropIndex = null;
        this.sortDropAfter = false;
        this.sortDraft = [];
        this.syncPageLock();
    }

    saveSortOrder() {
        if (!this.sortDraft.length) return;
        this.types = this.sortDraft.map((item, index) => String(item.type || '').trim() || createCellTitle(index));
        this.entries = this.sortDraft.map(item => normalizeEntry(serializeEntry(item.entry)));
        this.setupCanvas();
        this.saveEntriesToLocalStorage();
        this.draw();
        this.closeSortBox();
    }

    openEditor(index) {
        this.currentIndex = index;
        this.currentEntryDraft = cloneEntry(normalizeEntry(this.entries[index]));
        this.selectedStageItemId = this.currentEntryDraft.items.at(-1)?.id || null;
        this.dragState = null;
        this.categoryNameEl.textContent = this.types[index];
        this.titleInputEl.value = this.types[index];
        this.answerInputEl.value = this.currentEntryDraft.answer;
        this.editorEl.setAttribute('data-show', 'true');
        this.syncPageLock();
        this.renderEditorStage();
        this.titleInputEl.focus();
    }

    closeEditor() {
        this.editorEl.setAttribute('data-show', 'false');
        this.closeEmojiPicker();
        this.currentIndex = null;
        this.currentEntryDraft = createEmptyEntry();
        this.selectedStageItemId = null;
        this.dragState = null;
        this.syncPageLock();
    }

    saveCurrent() {
        if (this.currentIndex === null) return;
        const nextTitle = this.titleInputEl.value.trim() || createCellTitle(this.currentIndex);
        this.types[this.currentIndex] = nextTitle;
        this.currentEntryDraft.answer = this.answerInputEl.value.trim();
        this.entries[this.currentIndex] = normalizeEntry(serializeEntry(this.currentEntryDraft));
        this.setupCanvas();
        this.saveEntriesToLocalStorage();
        this.draw();
        this.closeEditor();
    }

    clearCurrent() {
        if (this.currentIndex === null) return;
        this.entries[this.currentIndex] = createEmptyEntry();
        this.saveEntriesToLocalStorage();
        this.draw();
        this.closeEditor();
    }

    addCell() {
        const index = this.types.length;
        this.types.push(createCellTitle(index));
        this.entries.push(createEmptyEntry());
        this.setupCanvas();
        this.saveEntriesToLocalStorage();
        this.draw();
        this.openEditor(index);
    }

    moveSortDraftItem(fromIndex, toIndex) {
        if (fromIndex === toIndex || fromIndex < 0 || toIndex < 0 || fromIndex >= this.sortDraft.length || toIndex > this.sortDraft.length) return;
        const nextDraft = [...this.sortDraft];
        const [moved] = nextDraft.splice(fromIndex, 1);
        nextDraft.splice(toIndex, 0, moved);
        this.sortDraft = nextDraft;
    }

    getSortCardFromEvent(event) {
        return event.target.closest('.sort-card');
    }

    onSortDragStart(event) {
        const card = this.getSortCardFromEvent(event);
        if (!card) return;
        this.sortDragIndex = Number(card.dataset.index);
        this.sortDropIndex = null;
        this.sortDropAfter = false;
        if (event.dataTransfer) {
            event.dataTransfer.effectAllowed = 'move';
            event.dataTransfer.setData('text/plain', String(this.sortDragIndex));
        }
        this.syncSortCardClasses();
    }

    onSortDragOver(event) {
        if (this.sortDragIndex === null) return;
        const card = this.getSortCardFromEvent(event);
        if (!card) return;
        event.preventDefault();
        const targetIndex = Number(card.dataset.index);
        const rect = card.getBoundingClientRect();
        this.sortDropIndex = targetIndex;
        this.sortDropAfter = event.clientY > rect.top + rect.height / 2;
        this.syncSortCardClasses();
    }

    onSortDrop(event) {
        if (this.sortDragIndex === null) return;
        const card = this.getSortCardFromEvent(event);
        event.preventDefault();
        if (!card) return;
        const targetIndex = Number(card.dataset.index);
        let nextIndex = targetIndex + (this.sortDropAfter ? 1 : 0);
        if (this.sortDragIndex < nextIndex) nextIndex -= 1;
        this.moveSortDraftItem(this.sortDragIndex, nextIndex);
        this.sortDragIndex = null;
        this.sortDropIndex = null;
        this.sortDropAfter = false;
        this.renderSortList();
    }

    onSortDragEnd() {
        this.sortDragIndex = null;
        this.sortDropIndex = null;
        this.sortDropAfter = false;
        this.syncSortCardClasses();
    }

    onSortDragLeave(event) {
        if (!this.sortListEl.contains(event.relatedTarget)) {
            this.sortDropIndex = null;
            this.sortDropAfter = false;
            this.syncSortCardClasses();
        }
    }

    deleteCurrentCell() {
        if (this.currentIndex === null) return;
        if (this.types.length <= 1) {
            window.alert('至少需要保留一个格子。');
            return;
        }

        this.types.splice(this.currentIndex, 1);
        this.entries.splice(this.currentIndex, 1);
        this.types = this.types.map((type, index) => String(type || '').trim() || createCellTitle(index));
        this.setupCanvas();
        this.saveEntriesToLocalStorage();
        this.draw();
        this.closeEditor();
    }

    clearAll() {
        const hasContent = this.entries.some(entry => entry.items.length || entry.answer || entry.clue);
        if (!hasContent) return;
        if (!window.confirm('要清空所有格子的 emoji / 文本框排版和答案备注吗？')) return;
        this.entries = new Array(this.types.length).fill(null).map(() => createEmptyEntry());
        this.saveEntriesToLocalStorage();
        this.draw();
    }

    toggleAnswers() {
        this.showAnswers = !this.showAnswers;
        this.updateUIState();
        this.draw();
    }

    insertEmoji(emoji) {
        if (this.currentIndex === null) return;
        const item = createEmojiItem(emoji, this.currentEntryDraft.items.length);
        this.currentEntryDraft.items.push(item);
        this.selectedStageItemId = item.id;
        this.renderEditorStage();
    }

    insertTextBox() {
        if (this.currentIndex === null) return;
        const text = window.prompt('请输入文本框内容', 'TEXT');
        if (text === null) return;
        const nextText = String(text).trim();
        if (!nextText) {
            window.alert('文本框内容不能为空。');
            return;
        }
        const item = createTextItem(nextText, this.currentEntryDraft.items.length);
        this.currentEntryDraft.items.push(item);
        this.selectedStageItemId = item.id;
        this.renderEditorStage();
    }

    updateSelectedText(text) {
        const item = this.getSelectedItem();
        if (!item || item.type !== 'text') return;
        item.text = String(text || '').slice(0, 120);
        if (!item.text.trim()) item.text = '文本';
        this.renderEditorStage();
    }

    deleteSelected() {
        if (!this.selectedStageItemId) return;
        this.currentEntryDraft.items = this.currentEntryDraft.items.filter(item => item.id !== this.selectedStageItemId);
        this.selectedStageItemId = this.currentEntryDraft.items.at(-1)?.id || null;
        this.renderEditorStage();
    }

    bringSelectionToFront() {
        const item = this.getSelectedItem();
        if (!item) return;
        this.currentEntryDraft.items = this.currentEntryDraft.items.filter(other => other.id !== item.id).concat(item);
        this.renderEditorStage();
    }

    scaleSelectionBy(delta) {
        const item = this.getSelectedItem();
        if (!item) return;
        item.size = clamp(item.size + delta, 0.08, 0.6);
        this.renderEditorStage();
    }

    scaleSelectionUp() { this.scaleSelectionBy(0.02); }
    scaleSelectionDown() { this.scaleSelectionBy(-0.02); }
    setSelectedItemSize(size) {
        const item = this.getSelectedItem();
        if (!item) return;
        item.size = clamp(size, 0.08, 0.6);
        this.renderEditorStage();
    }

    drawRoundedRectPath(ctx, x, y, width, height, radius) {
        const nextRadius = Math.min(radius, width / 2, height / 2);
        ctx.beginPath();
        if (typeof ctx.roundRect === 'function') {
            ctx.roundRect(x, y, width, height, nextRadius);
            return;
        }
        ctx.moveTo(x + nextRadius, y);
        ctx.arcTo(x + width, y, x + width, y + height, nextRadius);
        ctx.arcTo(x + width, y + height, x, y + height, nextRadius);
        ctx.arcTo(x, y + height, x, y, nextRadius);
        ctx.arcTo(x, y, x + width, y, nextRadius);
        ctx.closePath();
    }

    wrapTextByContext(ctx, text, maxWidth) {
        const rawLines = String(text || '').split(/\n/g).map(line => line.trim()).filter(Boolean);
        if (!rawLines.length) return [''];
        const result = [];
        rawLines.forEach(rawLine => {
            const useWords = /\s/.test(rawLine);
            const pieces = useWords ? rawLine.split(/\s+/g).filter(Boolean) : getGraphemes(rawLine);
            let current = '';
            pieces.forEach(piece => {
                const next = useWords ? (current ? `${current} ${piece}` : piece) : `${current}${piece}`;
                if (!current || ctx.measureText(next).width <= maxWidth) {
                    current = next;
                    return;
                }
                result.push(current);
                current = piece;
            });
            if (current) result.push(current);
        });
        return result.slice(0, 6);
    }

    getTextItemMetrics(ctx, item, area) {
        const fontSize = clamp(item.size * Math.min(area.width, area.height), 16, area.height * 0.5);
        const maxTextWidth = clamp(fontSize * 4.8, 92, area.width * 0.82);
        const lineHeight = Math.round(fontSize * 1.18);
        ctx.save();
        ctx.font = `700 ${fontSize}px ${TEXT_FONT}`;
        const lines = this.wrapTextByContext(ctx, item.text, maxTextWidth);
        const textWidth = Math.max(...lines.map(line => ctx.measureText(line).width), fontSize);
        ctx.restore();
        const paddingX = Math.max(16, Math.round(fontSize * 0.4));
        const paddingY = Math.max(12, Math.round(fontSize * 0.28));
        const width = Math.min(maxTextWidth + paddingX * 2, textWidth + paddingX * 2);
        const height = lines.length * lineHeight + paddingY * 2;
        const centerX = area.left + item.x * area.width;
        const centerY = area.top + item.y * area.height;
        return {
            centerX,
            centerY,
            fontSize,
            lineHeight,
            lines,
            width,
            height,
            paddingX,
            paddingY,
            left: centerX - width / 2,
            top: centerY - height / 2,
            right: centerX + width / 2,
            bottom: centerY + height / 2,
        };
    }

    getItemBounds(ctx, item, area) {
        if (item.type === 'text') return this.getTextItemMetrics(ctx, item, area);
        const fontSize = item.size * Math.min(area.width, area.height);
        ctx.save();
        ctx.font = `600 ${fontSize}px ${EMOJI_FONT}`;
        const width = Math.max(fontSize * 0.72, ctx.measureText(item.emoji).width);
        ctx.restore();
        const centerX = area.left + item.x * area.width;
        const centerY = area.top + item.y * area.height;
        const height = fontSize * 0.92;
        return { centerX, centerY, width, height, left: centerX - width / 2, top: centerY - height / 2, right: centerX + width / 2, bottom: centerY + height / 2 };
    }

    drawSceneItems(ctx, items, area, { selectionId = null, clip = true } = {}) {
        if (!items.length) return;
        if (clip) {
            ctx.save();
            ctx.beginPath();
            ctx.rect(area.left, area.top, area.width, area.height);
            ctx.clip();
        }
        items.forEach(item => {
            const bounds = this.getItemBounds(ctx, item, area);
            if (item.type === 'text') {
                ctx.save();
                this.drawRoundedRectPath(ctx, bounds.left, bounds.top, bounds.width, bounds.height, 18);
                ctx.fillStyle = 'rgba(255, 250, 243, 0.96)';
                ctx.fill();
                ctx.strokeStyle = 'rgba(126, 95, 63, 0.24)';
                ctx.lineWidth = 2;
                ctx.stroke();
                ctx.fillStyle = '#2d261f';
                ctx.font = `700 ${bounds.fontSize}px ${TEXT_FONT}`;
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';
                const startY = bounds.centerY - (bounds.lines.length - 1) * bounds.lineHeight / 2;
                bounds.lines.forEach((line, index) => ctx.fillText(line, bounds.centerX, startY + index * bounds.lineHeight, bounds.width - bounds.paddingX * 2));
                ctx.restore();
                return;
            }
            const fontSize = item.size * Math.min(area.width, area.height);
            ctx.save();
            ctx.font = `600 ${fontSize}px ${EMOJI_FONT}`;
            ctx.textAlign = 'center';
            ctx.textBaseline = 'middle';
            ctx.fillText(item.emoji, bounds.centerX, bounds.centerY);
            ctx.restore();
        });
        if (clip) ctx.restore();
        if (!selectionId) return;
        const selected = items.find(item => item.id === selectionId);
        if (!selected) return;
        const bounds = this.getItemBounds(ctx, selected, area);
        ctx.save();
        ctx.strokeStyle = '#7e5f3f';
        ctx.lineWidth = 3;
        ctx.setLineDash([8, 6]);
        ctx.strokeRect(bounds.left - 8, bounds.top - 8, bounds.width + 16, bounds.height + 16);
        ctx.setLineDash([]);
        ctx.fillStyle = '#7e5f3f';
        ctx.beginPath();
        ctx.arc(bounds.right + 4, bounds.top - 4, 6, 0, Math.PI * 2);
        ctx.fill();
        ctx.restore();
    }
    drawEmptyHint(ctx, area) {
        ctx.save();
        ctx.strokeStyle = '#d5d5d5';
        ctx.lineWidth = 1.5;
        ctx.strokeRect(area.left + 12, area.top + 12, area.width - 24, area.height - 24);
        ctx.fillStyle = '#b8b8b8';
        ctx.font = `400 14px ${TEXT_FONT}`;
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText('点击打开编辑器，插入并摆放 emoji / 文本框', area.left + area.width / 2, area.top + area.height / 2, area.width - 32);
        ctx.restore();
    }

    renderEditorStage() {
        const ctx = this.editorStageCtx;
        const canvas = this.editorStageCanvas;
        const area = this.getSceneArea();
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        ctx.fillStyle = '#f8efe3';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        ctx.save();
        ctx.fillStyle = '#fffaf3';
        ctx.fillRect(area.left, area.top, area.width, area.height);
        ctx.strokeStyle = 'rgba(42, 39, 35, 0.18)';
        ctx.lineWidth = 2;
        ctx.strokeRect(area.left, area.top, area.width, area.height);
        for (let i = 1; i < 4; i += 1) {
            ctx.beginPath();
            ctx.moveTo(area.left + area.width / 4 * i, area.top);
            ctx.lineTo(area.left + area.width / 4 * i, area.top + area.height);
            ctx.strokeStyle = 'rgba(42, 39, 35, 0.08)';
            ctx.stroke();
        }
        for (let i = 1; i < 4; i += 1) {
            ctx.beginPath();
            ctx.moveTo(area.left, area.top + area.height / 4 * i);
            ctx.lineTo(area.left + area.width, area.top + area.height / 4 * i);
            ctx.strokeStyle = 'rgba(42, 39, 35, 0.08)';
            ctx.stroke();
        }
        if (!this.currentEntryDraft.items.length) {
            ctx.fillStyle = '#a69480';
            ctx.font = `600 26px ${TEXT_FONT}`;
            ctx.textAlign = 'center';
            ctx.textBaseline = 'middle';
            ctx.fillText('从 Emoji 面板选图标，或插入一个文本框', area.left + area.width / 2, area.top + area.height / 2 - 16);
            ctx.font = `400 18px ${TEXT_FONT}`;
            ctx.fillText('两种元素都能拖动和缩放，做出更自由的排版', area.left + area.width / 2, area.top + area.height / 2 + 22);
        }
        this.drawSceneItems(ctx, this.currentEntryDraft.items, area, { selectionId: this.selectedStageItemId, clip: true });
        ctx.restore();
        this.updateSelectionUI();
        this.editorStageCanvas.classList.toggle('dragging', Boolean(this.dragState));
    }

    getItemAtStagePoint(point) {
        const area = this.getSceneArea();
        for (let index = this.currentEntryDraft.items.length - 1; index >= 0; index -= 1) {
            const item = this.currentEntryDraft.items[index];
            const bounds = this.getItemBounds(this.editorStageCtx, item, area);
            if (point.x >= bounds.left && point.x <= bounds.right && point.y >= bounds.top && point.y <= bounds.bottom) return item;
        }
        return null;
    }

    onStagePointerDown(event) {
        if (this.currentIndex === null) return;
        const point = this.getStagePoint(event);
        const item = this.getItemAtStagePoint(point);
        if (!item) {
            this.selectedStageItemId = null;
            this.dragState = null;
            this.renderEditorStage();
            return;
        }
        this.selectedStageItemId = item.id;
        this.bringSelectionToFront();
        const area = this.getSceneArea();
        const selected = this.getSelectedItem();
        this.dragState = {
            pointerId: event.pointerId,
            offsetX: point.x - (area.left + selected.x * area.width),
            offsetY: point.y - (area.top + selected.y * area.height),
        };
        this.editorStageCanvas.setPointerCapture(event.pointerId);
        this.renderEditorStage();
    }

    onStagePointerMove(event) {
        if (!this.dragState || event.pointerId !== this.dragState.pointerId) return;
        const item = this.getSelectedItem();
        if (!item) return;
        const point = this.getStagePoint(event);
        const area = this.getSceneArea();
        item.x = clamp((point.x - this.dragState.offsetX - area.left) / area.width, 0.04, 0.96);
        item.y = clamp((point.y - this.dragState.offsetY - area.top) / area.height, 0.05, 0.95);
        this.renderEditorStage();
    }

    onStagePointerUp(event) {
        if (!this.dragState || event.pointerId !== this.dragState.pointerId) return;
        this.dragState = null;
        if (this.editorStageCanvas.hasPointerCapture(event.pointerId)) this.editorStageCanvas.releasePointerCapture(event.pointerId);
        this.renderEditorStage();
    }

    onStageWheel(event) {
        if (!this.getSelectedItem()) return;
        event.preventDefault();
        this.scaleSelectionBy(event.deltaY < 0 ? 0.02 : -0.02);
    }

    draw() {
        this.setupCanvas();
        const row = this.getGridRowCount();
        const { ctx, canvasScale, width, height, bodyMargin, titleHeight, contentWidth, contentHeight, col, rowHeight, colWidth, labelHeight } = this;
        ctx.setTransform(1, 0, 0, 1, 0, 0);
        ctx.clearRect(0, 0, this.canvas.width, this.canvas.height);
        ctx.fillStyle = '#fffdf8';
        ctx.fillRect(0, 0, this.canvas.width, this.canvas.height);
        ctx.save();
        ctx.setTransform(canvasScale, 0, 0, canvasScale, 0, 0);
        ctx.fillStyle = '#1d1d1d';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.font = `700 28px ${TEXT_FONT}`;
        ctx.fillText(this.title, width / 2, bodyMargin + 14);
        ctx.textAlign = 'left';
        ctx.fillStyle = '#767676';
        ctx.font = `12px ${TEXT_FONT}`;
        ctx.fillText(this.showAnswers ? '当前导出模式：答案显示版' : '当前导出模式：猜番隐藏答案版', bodyMargin, height - 10);
        ctx.textAlign = 'right';
        ctx.fillText(PROJECT_URL, width - bodyMargin, height - 10);
        ctx.translate(bodyMargin, bodyMargin + titleHeight);
        ctx.strokeStyle = '#222';
        ctx.lineWidth = 2;
        ctx.fillStyle = '#fff';
        ctx.fillRect(0, 0, contentWidth, contentHeight);
        for (let y = 0; y <= row; y += 1) {
            ctx.beginPath();
            ctx.moveTo(0, y * rowHeight);
            ctx.lineTo(contentWidth, y * rowHeight);
            ctx.stroke();
            if (y === row) continue;
            ctx.save();
            ctx.globalAlpha = 0.25;
            ctx.beginPath();
            ctx.moveTo(0, y * rowHeight + rowHeight - labelHeight);
            ctx.lineTo(contentWidth, y * rowHeight + rowHeight - labelHeight);
            ctx.stroke();
            ctx.restore();
        }
        for (let x = 0; x <= col; x += 1) {
            ctx.beginPath();
            ctx.moveTo(x * colWidth, 0);
            ctx.lineTo(x * colWidth, contentHeight);
            ctx.stroke();
        }
        ctx.textAlign = 'center';
        ctx.fillStyle = '#222';
        ctx.font = `500 16px ${TEXT_FONT}`;
        this.types.forEach((type, index) => {
            const x = index % col;
            const y = Math.floor(index / col);
            ctx.fillText(type, x * colWidth + colWidth / 2, y * rowHeight + rowHeight - labelHeight / 2 + 1);
        });
        this.entries.forEach((entry, index) => this.drawEntry(index, normalizeEntry(entry)));
        ctx.restore();
    }

    drawEntry(index, entry) {
        const { ctx, col, colWidth, rowHeight, imageWidth, imageHeight } = this;
        const x = index % col;
        const y = Math.floor(index / col);
        const left = x * colWidth + 1;
        const top = y * rowHeight + 1;
        const area = { left, top, width: imageWidth, height: imageHeight };
        ctx.save();
        ctx.fillStyle = '#fff';
        ctx.fillRect(left, top, imageWidth, imageHeight);
        if (!entry.items.length) {
            if (this.showAnswers && entry.answer) {
                this.drawAnswerText({ text: entry.answer, centerX: left + imageWidth / 2, centerY: top + imageHeight / 2, maxWidth: imageWidth - 20, maxHeight: imageHeight - 20, color: '#2d261f' });
                ctx.restore();
                return;
            }
            this.drawEmptyHint(ctx, area);
            ctx.restore();
            return;
        }
        this.drawSceneItems(ctx, entry.items, area, { clip: true });
        if (this.showAnswers && entry.answer) this.drawAnswerOverlay({ text: entry.answer, left, top, width: imageWidth });
        ctx.restore();
    }

    drawAnswerText({ text, centerX, centerY, maxWidth, maxHeight, color = '#fff' }) {
        this.drawFittedText({ text, centerX, centerY, maxWidth, maxHeight, minFontSize: 14, maxFontSize: 24, maxLines: 3, fontWeight: 700, fontFamily: TEXT_FONT, color });
    }

    drawFittedText({ text, centerX, centerY, maxWidth, maxHeight, minFontSize, maxFontSize, maxLines, fontWeight, fontFamily, color }) {
        const { ctx } = this;
        const preparedText = String(text || '').trim();
        if (!preparedText) return;
        let best = null;
        for (let fontSize = maxFontSize; fontSize >= minFontSize; fontSize -= 2) {
            const lineHeight = Math.round(fontSize * 1.16);
            ctx.font = `${fontWeight} ${fontSize}px ${fontFamily}`;
            const lines = this.wrapText(preparedText, maxWidth).slice(0, maxLines);
            if (lines.length * lineHeight <= maxHeight) {
                best = { fontSize, lineHeight, lines };
                break;
            }
        }
        if (!best) {
            best = { fontSize: minFontSize, lineHeight: Math.round(minFontSize * 1.16), lines: this.wrapText(preparedText, maxWidth).slice(0, maxLines) };
        }
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillStyle = color;
        ctx.font = `${fontWeight} ${best.fontSize}px ${fontFamily}`;
        const startY = centerY - (best.lines.length - 1) * best.lineHeight / 2;
        best.lines.forEach((line, index) => ctx.fillText(line, centerX, startY + index * best.lineHeight, maxWidth + 8));
    }

    drawAnswerOverlay({ text, left, top, width }) {
        const { ctx } = this;
        const overlayHeight = 42;
        ctx.save();
        ctx.fillStyle = 'rgba(31, 28, 24, 0.84)';
        ctx.fillRect(left, top + this.imageHeight - overlayHeight, width, overlayHeight);
        this.drawAnswerText({ text, centerX: left + width / 2, centerY: top + this.imageHeight - overlayHeight / 2, maxWidth: width - 14, maxHeight: overlayHeight - 8, color: '#fffaf3' });
        ctx.restore();
    }

    wrapText(text, maxWidth) {
        return this.wrapTextByContext(this.ctx, text, maxWidth);
    }

    showOutput(imgURL) {
        this.outputImageEl.src = imgURL;
        this.outputEl.setAttribute('data-show', 'true');
        this.syncPageLock();
    }

    closeOutput() {
        this.outputEl.setAttribute('data-show', 'false');
        this.syncPageLock();
    }

    downloadImage() {
        const suffix = this.showAnswers ? '答案版' : '猜番版';
        const fileName = `[猜番表]${this.title}-${suffix}.jpg`;
        const imgURL = this.canvas.toDataURL('image/jpeg', 0.92);
        const linkEl = document.createElement('a');
        linkEl.download = fileName;
        linkEl.href = imgURL;
        linkEl.dataset.downloadurl = ['image/jpeg', fileName, imgURL].join(':');
        document.body.appendChild(linkEl);
        linkEl.click();
        document.body.removeChild(linkEl);
        this.showOutput(imgURL);
    }
}
