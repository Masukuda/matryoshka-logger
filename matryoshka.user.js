// ==UserScript==
// @name         Matryoshka Backtest Auto-Logger
// @namespace    http://tampermonkey.net/
// @version      2.1
// @description  Экспорт сделок из TradingView в Excel и Google Таблицы (Стратегия «Матрёшка»)
// @match        https://*.tradingview.com/*
// @grant        unsafeWindow
// @updateURL    https://raw.githubusercontent.com/Masukuda/matryoshka-logger/main/matryoshka.user.js
// @downloadURL  https://raw.githubusercontent.com/Masukuda/matryoshka-logger/main/matryoshka.user.js
// @run-at       document-idle
// ==/UserScript==

(function() {
    'use strict';

    if (window.self !== window.top) return;

    const pageWindow = typeof unsafeWindow !== 'undefined' ? unsafeWindow : window;
    const STORAGE_KEY = 'matryoshka_logger_state_v2';
    const isMac = /Mac|iPod|iPhone|iPad/i.test(navigator.platform || navigator.userAgent);

    // --- Конфигурация и состояние ---

    const defaultState = {
        decimalSeparator: '.',
        formulaMode: 'en_comma',
        excelRow: 6,
        trend: 'Тренд вверх',
        zone: 'Ордер-блок',
        liq: 'да',
        conf: 'Слом структуры',
        customHotkey: null,
        lastSavedTradeSig: '',
        lastSavedRow: 6,
        batchQueue: [],
        posTop: 20,
        posRight: 20
    };

    function loadState() {
        try {
            const saved = localStorage.getItem(STORAGE_KEY);
            return saved ? { ...defaultState, ...JSON.parse(saved) } : { ...defaultState };
        } catch (e) {
            return { ...defaultState };
        }
    }

    function saveState(patch) {
        try {
            state = { ...state, ...patch };
            localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
        } catch (e) {}
    }

    let state = loadState();
    if (!Array.isArray(state.batchQueue)) state.batchQueue = [];
    let currentTrade = null;
    let isReEditingSameTrade = false;
    let isRecordingHotkey = false;

    // --- Расчет сессии, времени и результата ---

    function getHourInTimezone(dateObj, timeZone) {
        try {
            const parts = new Intl.DateTimeFormat('en-US', {
                timeZone,
                hour: 'numeric',
                minute: 'numeric',
                hour12: false
            }).formatToParts(dateObj);
            const h = parseInt(parts.find(p => p.type === 'hour')?.value || '0', 10) % 24;
            const m = parseInt(parts.find(p => p.type === 'minute')?.value || '0', 10);
            return h + (m / 60);
        } catch (e) {
            return dateObj.getUTCHours();
        }
    }

    function detectSession(timestamp) {
        if (!timestamp) return 'Лондон';
        const date = new Date(timestamp * 1000);
        const nyHour = getHourInTimezone(date, 'America/New_York');
        const lonHour = getHourInTimezone(date, 'Europe/London');

        if (nyHour >= 8 && nyHour < 17) return 'Нью-Йорк';
        if (lonHour >= 8 && nyHour < 8) return 'Лондон';
        return 'Азия';
    }

    function formatTimeframe(rawInterval) {
        const str = String(rawInterval || '1').toUpperCase().trim();
        if (/^\d+$/.test(str)) {
            const mins = parseInt(str, 10);
            if (mins >= 60 && mins % 60 === 0) return 'H' + (mins / 60);
            return 'M' + mins;
        }
        if (str === 'D' || str === '1D') return 'D1';
        if (str === 'W' || str === '1W') return 'W1';
        if (str === 'M' || str === '1M') return 'MN';
        return str;
    }

    function formatDate(timestamp) {
        const d = new Date(timestamp * 1000);
        const day = String(d.getDate()).padStart(2, '0');
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const year = d.getFullYear();
        return `${day}.${month}.${year}`;
    }

    function calculateOutcome(bars, startIndex, endIndex, isLong, stopPrice, targetPrice) {
        if (!bars || typeof bars.valueAt !== 'function') return 'Тейк';
        if (startIndex == null || endIndex == null) return 'Тейк';

        const from = Math.min(startIndex, endIndex);
        const to = Math.max(startIndex, endIndex);

        for (let i = from; i <= to; i++) {
            const bar = bars.valueAt(i);
            if (!bar || !Array.isArray(bar)) continue;

            const high = Number(bar[2]);
            const low = Number(bar[3]);
            const close = Number(bar[4]);
            if (isNaN(high) || isNaN(low)) continue;

            const hitTarget = isLong ? (high >= targetPrice) : (low <= targetPrice);
            const hitStop = isLong ? (low <= stopPrice) : (high >= stopPrice);

            if (hitTarget && hitStop) {
                if (i === from && ((isLong && close >= targetPrice) || (!isLong && close <= targetPrice))) {
                    return 'Тейк';
                }
                return 'Стоп';
            }
            if (hitStop) return 'Стоп';
            if (hitTarget) return 'Тейк';
        }
        return 'Безубыток';
    }

    function readWatched(val) {
        if (val != null && typeof val === 'object' && typeof val.value === 'function') {
            return val.value();
        }
        return val;
    }

    // --- Извлечение сделки из TradingView ---

    function extractTrade() {
        try {
            const c = pageWindow._exposed_chartWidgetCollection;
            if (!c) {
                showToast('График TradingView еще не загружен', 'error');
                return null;
            }

            const widget = typeof c.activeChartWidget === 'function'
                ? c.activeChartWidget()
                : (c.activeChartWidget?.value ? c.activeChartWidget.value() : c.activeChartWidget);
            if (!widget) return null;

            const rawModel = typeof widget.model === 'function'
                ? widget.model()
                : (widget.model || widget._model);
            const model = rawModel?.m_model || rawModel;
            if (!model) return null;

            const ms = typeof model.mainSeries === 'function' ? model.mainSeries() : model.mainSeries;
            const bars = typeof ms.bars === 'function' ? ms.bars() : ms.bars;
            const tools = typeof model.allLineTools === 'function' ? model.allLineTools() : model.allLineTools;
            const list = Array.isArray(tools) ? tools : Object.values(tools || {});

            const getToolName = (t) => {
                if (!t) return '';
                const name = (typeof t.toolname === 'function' ? t.toolname() : t.toolname) || t._toolName || '';
                return typeof name === 'string' ? name : '';
            };

            let tool = null;
            try {
                if (typeof model.selection === 'function') {
                    const sel = model.selection();
                    const selSources = typeof sel?.lineTools === 'function' ? sel.lineTools() :
                                       (typeof sel?.sources === 'function' ? sel.sources() : null);
                    if (Array.isArray(selSources)) {
                        tool = selSources.find(t => getToolName(t).includes('LineToolRiskReward'));
                    }
                }
            } catch (e) {}

            if (!tool) {
                tool = list.slice().reverse().find(t => getToolName(t).includes('LineToolRiskReward'));
            }

            if (!tool) {
                showToast('Нарисуйте инструмент «Длинная/Короткая позиция» на графике!', 'error');
                return null;
            }

            const rawProps = typeof tool.properties === 'function' ? tool.properties() : tool.properties;
            const props = (rawProps && typeof rawProps.state === 'function') ? rawProps.state() : rawProps;
            const points = typeof tool.points === 'function' ? tool.points() : (tool._points || tool.points || []);

            const toolNameStr = getToolName(tool);
            const isLong = toolNameStr.includes('Long');

            const entry = Number(points[0]?.price || 0);
            const time = points[0]?.time || Math.floor(Date.now() / 1000);

            let symInfo = {};
            try {
                symInfo = (typeof ms.symbolInfo === 'function' ? ms.symbolInfo() : ms._symbolInfo) || {};
            } catch (e) {}

            const priceStep = (typeof ms.priceStep === 'function' ? ms.priceStep() : ms.priceStep) ||
                              (symInfo.minmov && symInfo.pricescale ? symInfo.minmov / symInfo.pricescale : 0.001);

            const stopLevel = Number(readWatched(props?.stopLevel) ?? readWatched(rawProps?.stopLevel) ?? 0);
            const profitLevel = Number(readWatched(props?.profitLevel) ?? readWatched(rawProps?.profitLevel) ?? 0);

            const stop = isLong ? entry - (stopLevel * priceStep) : entry + (stopLevel * priceStep);
            const target = isLong ? entry + (profitLevel * priceStep) : entry - (profitLevel * priceStep);
            const rrNum = (stopLevel > 0) ? (profitLevel / stopLevel) : 0;

            const startIdx = points[0]?.index ?? 0;
            const endIdx = points[1]?.index ?? startIdx;
            const calculatedOutcome = calculateOutcome(bars, startIdx, endIdx, isLong, stop, target);

            const decimals = priceStep < 1
                ? (priceStep.toFixed(8).replace(/0+$/, '').split('.')[1]?.length || 2)
                : 2;

            const rawSymbol = readWatched(props?.symbol) || symInfo.name || 'XAUUSD';
            const cleanSymbol = String(rawSymbol).split(':').pop();
            const rawInterval = readWatched(props?.interval) || '1';

            const entryRounded = Number(entry.toFixed(decimals));
            const stopRounded = Number(stop.toFixed(decimals));
            const targetRounded = Number(target.toFixed(decimals));

            const signature = `${cleanSymbol}_${time}_${entryRounded}_${stopRounded}_${targetRounded}`;

            return {
                signature,
                symbol: cleanSymbol,
                timeframe: formatTimeframe(rawInterval),
                direction: isLong ? 'LONG' : 'SHORT',
                entryPrice: entryRounded,
                stopPrice: stopRounded,
                targetPrice: targetRounded,
                riskPts: Number(Math.abs(entryRounded - stopRounded).toFixed(decimals)),
                potPts: Number(Math.abs(targetRounded - entryRounded).toFixed(decimals)),
                riskReward: rrNum.toFixed(2),
                riskRewardNum: rrNum,
                candleDate: formatDate(time),
                session: detectSession(time),
                autoResult: calculatedOutcome
            };
        } catch (e) {
            console.error('Ошибка парсинга TV:', e);
            showToast('Ошибка парсинга TV: ' + e.message, 'error');
            return null;
        }
    }

    // --- Генерация TSV и формул ---

    function formatNumForSheet(num) {
        if (num == null || num === '') return '';
        const str = String(num);
        return state.decimalSeparator === ',' ? str.replace('.', ',') : str.replace(',', '.');
    }

    function sanitizeCell(str) {
        return String(str || '').replace(/[\t\r\n]+/g, ' ').trim();
    }

    function buildCalculatedColumns(trade, result, rowNum) {
        const r = Math.max(6, parseInt(rowNum, 10) || 6);
        const mode = state.formulaMode;

        if (mode === 'empty') {
            return { colL: '', colM: '', colN: '', colP: '', colQ: '', colR: '' };
        }

        if (mode === 'val') {
            const colL = formatNumForSheet(trade.riskPts.toFixed(2));
            const colM = formatNumForSheet(trade.potPts.toFixed(2));
            const colN = formatNumForSheet(trade.riskReward);
            let colP = '';
            if (result === 'Тейк') colP = formatNumForSheet(trade.riskReward);
            else if (result === 'Стоп') colP = formatNumForSheet('-1.00');
            else if (result === 'Безубыток') colP = formatNumForSheet('0.00');
            return { colL, colM, colN, colP, colQ: '', colR: '' };
        }

        if (mode === 'ru_semi') {
            return {
                colL: `=ЕСЛИОШИБКА(ABS(I${r}-J${r});"")`,
                colM: `=ЕСЛИОШИБКА(ABS(K${r}-I${r});"")`,
                colN: `=ЕСЛИОШИБКА(M${r}/L${r};"")`,
                colP: `=ЕСЛИ(O${r}="Тейк";N${r};ЕСЛИ(O${r}="Стоп";-1;ЕСЛИ(O${r}="Безубыток";0;"")))`,
                colQ: `=ЕСЛИ(СЧЁТ($P$6:P${r})=0;"";СУММ($P$6:P${r}))`,
                colR: `=ЕСЛИ(Q${r}="";"";МАКС($Q$6:Q${r})-Q${r})`
            };
        }

        const s = (mode === 'en_semi') ? ';' : ',';
        return {
            colL: `=IFERROR(ABS(I${r}-J${r})${s}"")`,
            colM: `=IFERROR(ABS(K${r}-I${r})${s}"")`,
            colN: `=IFERROR(M${r}/L${r}${s}"")`,
            colP: `=IF(O${r}="Тейк"${s}N${r}${s}IF(O${r}="Стоп"${s}-1${s}IF(O${r}="Безубыток"${s}0${s}"")))`,
            colQ: `=IF(COUNT($P$6:P${r})=0${s}""${s}SUM($P$6:P${r}))`,
            colR: `=IF(Q${r}=""${s}""${s}MAX($Q$6:Q${r})-Q${r})`
        };
    }

    function buildRowTsv(record) {
        const calc = buildCalculatedColumns(record.trade, record.result, record.row);
        const tsvArray = [
            record.trade.candleDate,
            record.session,
            record.trend,
            record.trade.timeframe,
            record.entryZone,
            record.liquidityTaken,
            record.confirmation,
            formatNumForSheet(record.trade.entryPrice),
            formatNumForSheet(record.trade.stopPrice),
            formatNumForSheet(record.trade.targetPrice),
            calc.colL,
            calc.colM,
            calc.colN,
            record.result,
            calc.colP,
            calc.colQ,
            calc.colR,
            record.execError,
            record.commentStr,
            record.imgUrl
        ];
        return tsvArray.join('\t');
    }

    function getBatchStartRow() {
        if (state.batchQueue.length > 0 && state.batchQueue[0]?.row) {
            return state.batchQueue[0].row;
        }
        return state.excelRow;
    }

    function getHotkeyLabel() {
        if (state.customHotkey?.label) return state.customHotkey.label;
        return isMac ? '⌥Q / Fn+F2' : 'F2 / Alt+Q';
    }

    // --- Интерфейс ---

    const modalHtml = `
    <div id="bt-toast" style="display:none; position:fixed; bottom:24px; right:24px; z-index:10000000; background:#141414; color:#e0e0e0; padding:10px 16px; border-radius:6px; border-left:4px solid #d6b658; box-shadow:0 8px 24px rgba(0,0,0,0.85); font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif; font-size:12px; font-weight:500;"></div>

    <button type="button" id="bt-fab-trigger" title="Открыть экспорт сделки (Матрёшка)" style="position:fixed; bottom:18px; left:58px; z-index:9999990; background:#111; color:#d6b658; border:1px solid #332d1f; border-radius:6px; padding:5px 10px; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif; font-size:11px; font-weight:bold; cursor:pointer; box-shadow:0 4px 12px rgba(0,0,0,0.6); transition:0.2s;">
        🪆 Матрёшка <span id="bt-fab-hk" style="color:#888; font-weight:normal; margin-left:3px;">(${getHotkeyLabel()})</span>
    </button>

    <div id="bt-modal" style="display:none; position:fixed; top:${state.posTop}px; right:${state.posRight}px; z-index:9999999; background:#0a0a0a; color:#e0e0e0; padding:15px; border-radius:8px; box-shadow:0 12px 32px rgba(0,0,0,0.85); border:1px solid #332d1f; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif; width:390px; user-select:none;">

        <div id="bt-drag-handle" style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px; cursor:move; padding-bottom:6px; border-bottom:1px solid #1f1f1f;">
            <div style="text-align:left;">
                <div style="font-size:9px; letter-spacing:1.5px; color:#d6b658; text-transform:uppercase;">Стратегия «Матрёшка»</div>
                <div style="font-family:'Georgia', 'Playfair Display', serif; font-size:17px; font-weight:bold; color:#d6b658; letter-spacing:0.5px;">ЭКСПОРТ В ТАБЛИЦУ</div>
            </div>
            <button type="button" id="bt-hk-record-btn" class="bt-mini-btn" title="Кликните, чтобы назначить любую удобную клавишу открытия окна" style="padding:4px 8px; font-size:10px;">
                ⌨ Хоткей: <b id="bt-hk-label" style="color:#d6b658;">${getHotkeyLabel()}</b>
            </button>
        </div>

        <div style="background:#121212; padding:7px 8px; border-radius:6px; border:1px solid #222; margin-bottom:9px; font-size:10px;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
                <div style="display:flex; align-items:center; gap:4px;">
                    <span style="color:#9e9e9e; font-weight:bold;">ДРОБЬ:</span>
                    <div id="bt-sep-group" style="display:flex; gap:2px;">
                        <button type="button" class="bt-mini-btn" data-sep="." title="Числа через точку (2335.60)">Точка ( . )</button>
                        <button type="button" class="bt-mini-btn" data-sep="," title="Числа через запятую (2335,60)">Запятая ( , )</button>
                    </div>
                </div>
                <div style="display:flex; align-items:center; gap:3px;">
                    <span style="color:#9e9e9e; font-weight:bold;">СТРОКА B:</span>
                    <button type="button" id="bt-row-dec" class="bt-mini-btn" style="padding:2px 6px;">−</button>
                    <input type="number" id="bt-row-num" min="6" max="9999" value="${state.excelRow}" style="width:40px; background:#1a1a1a; color:#d6b658; border:1px solid #332d1f; border-radius:3px; text-align:center; font-weight:bold; font-size:11px; padding:1px 0; outline:none;">
                    <button type="button" id="bt-row-inc" class="bt-mini-btn" style="padding:2px 6px;">+</button>
                </div>
            </div>
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <span style="color:#9e9e9e; font-weight:bold;">ФОРМУЛЫ:</span>
                <div id="bt-fmode-group" style="display:flex; gap:3px;">
                    <button type="button" class="bt-mini-btn" data-fmode="en_comma" title="=IFERROR(ABS(I6-J6),&quot;&quot;)">EN ( , )</button>
                    <button type="button" class="bt-mini-btn" data-fmode="ru_semi" title="=ЕСЛИОШИБКА(ABS(I6-J6);&quot;&quot;)">RU ( ; )</button>
                    <button type="button" class="bt-mini-btn" data-fmode="en_semi" title="=IFERROR(ABS(I6-J6);&quot;&quot;)">EN ( ; )</button>
                    <button type="button" class="bt-mini-btn" data-fmode="val" title="Готовые числа без формул">Числа</button>
                    <button type="button" class="bt-mini-btn" data-fmode="empty" title="Оставлять пустыми">Пусто</button>
                </div>
            </div>
        </div>

        <div style="background:#141414; border:1px solid #2a2a2a; border-radius:6px; padding:8px 10px; margin-bottom:9px; font-size:12px; line-height:1.5;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:3px;">
                <span id="bt-sym-tf" style="color:#d6b658; font-weight:bold;"></span>
                <span id="bt-badge" style="background:#089981; color:#fff; font-size:10px; padding:2px 6px; border-radius:4px; font-weight:bold;"></span>
            </div>
            <div style="color:#a0a0a0; display:flex; justify-content:space-between;">
                <span>Вход: <b id="bt-en" style="color:#e0e0e0;"></b></span>
                <span>Стоп: <b id="bt-st" style="color:#e0e0e0;"></b></span>
                <span>Цель: <b id="bt-tg" style="color:#d6b658;"></b></span>
            </div>
            <div style="color:#787b86; margin-top:2px; font-size:11px; display:flex; justify-content:space-between;">
                <span id="bt-time"></span>
                <span id="bt-pts-preview" style="color:#9e9e9e;"></span>
            </div>
        </div>

        <div style="display:flex; justify-content:space-between; align-items:flex-end; margin-bottom:3px;">
            <div class="bt-section-title" style="margin:0;">ТОРГОВАЯ СЕССИЯ</div>
            <span style="font-size:9px; color:#089981;">АВТО (DST) ✓</span>
        </div>
        <div class="bt-group" id="bt-session" style="display:grid; grid-template-columns:repeat(3, 1fr); gap:4px; margin-bottom:8px;">
            <button type="button" class="bt-chip" data-val="Азия">Азия</button>
            <button type="button" class="bt-chip" data-val="Лондон">Лондон</button>
            <button type="button" class="bt-chip" data-val="Нью-Йорк">Нью-Йорк</button>
        </div>

        <div style="display:flex; justify-content:space-between; align-items:flex-end; margin-bottom:3px;">
            <div class="bt-section-title" style="margin:0;">ФАЗА РЫНКА <span style="color:#666; font-weight:normal;">(A / S / D / F)</span></div>
            <span style="font-size:9px; color:#089981;">АВТО ПО LONG/SHORT ✓</span>
        </div>
        <div class="bt-group" id="bt-trend" style="display:grid; grid-template-columns:1fr 1fr; gap:4px; margin-bottom:8px;">
            <button type="button" class="bt-chip" data-val="Тренд вверх">A. Тренд вверх</button>
            <button type="button" class="bt-chip" data-val="Тренд вниз">S. Тренд вниз</button>
            <button type="button" class="bt-chip" data-val="Боковик">D. Боковик</button>
            <button type="button" class="bt-chip" data-val="Импульс/новости">F. Импульс/новости</button>
        </div>

        <div class="bt-section-title">ЗОНА ВХОДА <span style="color:#666; font-weight:normal;">(1–5)</span></div>
        <div class="bt-group" id="bt-zone" style="display:grid; grid-template-columns:1fr 1fr; gap:4px; margin-bottom:8px;">
            <button type="button" class="bt-chip" data-val="Ордер-блок">1. Ордер-блок</button>
            <button type="button" class="bt-chip" data-val="Имбаланс">2. Имбаланс</button>
            <button type="button" class="bt-chip" data-val="Материнская свеча">3. Мать-свеча</button>
            <button type="button" class="bt-chip" data-val="Зона ликвидности">4. Ликвидность</button>
            <button type="button" class="bt-chip" data-val="Правильный откат" style="grid-column: span 2;">5. Правильный откат</button>
        </div>

        <div style="display:grid; grid-template-columns:1fr 2fr; gap:10px; margin-bottom:8px;">
            <div>
                <div class="bt-section-title">ЛИКВИДНОСТЬ <span style="color:#666; font-weight:normal;">(Space)</span></div>
                <div class="bt-group" id="bt-liq" style="display:flex; gap:4px;">
                    <button type="button" class="bt-chip" style="flex:1;" data-val="да">Да</button>
                    <button type="button" class="bt-chip" style="flex:1;" data-val="нет">Нет</button>
                </div>
            </div>
            <div>
                <div class="bt-section-title">ПОДТВЕРЖДЕНИЕ <span style="color:#666; font-weight:normal;">(Z / X / C / V / B)</span></div>
                <select id="bt-conf" style="width:100%; background:#1a1a1a; color:#d6b658; border:1px solid #332d1f; padding:5px; border-radius:4px; font-size:11px; outline:none; cursor:pointer;">
                    <option value="Слом структуры">Z. Слом структуры</option>
                    <option value="Реакция от зоны">X. Реакция от зоны</option>
                    <option value="Поглощение">C. Поглощение</option>
                    <option value="Снятие ликвидности">V. Снятие ликвидности</option>
                    <option value="Без подтверждения">B. Без подтверждения</option>
                </select>
            </div>
        </div>

        <div style="display:flex; justify-content:space-between; align-items:flex-end; margin-bottom:3px;">
            <div class="bt-section-title" style="margin:0;">РЕЗУЛЬТАТ <span style="color:#666; font-weight:normal;">(Q / W / E / R)</span></div>
            <span style="font-size:9px; color:#089981;">АВТОРАСЧЕТ ✓</span>
        </div>
        <div class="bt-group" id="bt-result" style="display:grid; grid-template-columns:repeat(4, 1fr); gap:4px; margin-bottom:8px;">
            <button type="button" class="bt-chip" data-val="Тейк" style="--ac:#089981;">Тейк (Q)</button>
            <button type="button" class="bt-chip" data-val="Стоп" style="--ac:#f23645;">Стоп (W)</button>
            <button type="button" class="bt-chip" data-val="Безубыток">БУ (E)</button>
            <button type="button" class="bt-chip" data-val="Пропуск">Пропуск (R)</button>
        </div>

        <div style="display:flex; justify-content:space-between; align-items:flex-end; margin-bottom:3px;">
            <div class="bt-section-title" style="margin:0;">ОШИБКА ИСПОЛНЕНИЯ</div>
            <span id="bt-err-hint" style="font-size:9px; color:#d6b658;"></span>
        </div>
        <select id="bt-err" style="width:100%; background:#1a1a1a; color:#e0e0e0; border:1px solid #332d1f; padding:5px; border-radius:4px; font-size:11px; outline:none; cursor:pointer; margin-bottom:8px;">
            <option value="—">—</option>
            <option value="Ранний вход">Ранний вход</option>
            <option value="Нет подтверждения">Нет подтверждения</option>
            <option value="Вход без снятия ликвидности">Вход без снятия ликвидности</option>
            <option value="Стоп слишком узкий">Стоп слишком узкий</option>
            <option value="Стоп слишком широкий">Стоп слишком широкий</option>
            <option value="Ранний выход">Ранний выход</option>
            <option value="Пересидел цель">Пересидел цель</option>
            <option value="Нарушен минимальный R:R">Нарушен минимальный R:R</option>
            <option value="Вход вне сессии">Вход вне сессии</option>
            <option value="Другое">Другое</option>
        </select>

        <input type="text" id="bt-comment" placeholder="Комментарий (колонка T, нажмите Tab для ввода)..." style="width:100%; box-sizing:border-box; background:#141414; border:1px solid #332d1f; color:#e0e0e0; padding:6px 8px; border-radius:4px; font-size:11px; margin-bottom:5px; outline:none;">
        <input type="text" id="bt-img" placeholder="Скриншот (колонка U, вставьте через Ctrl+V / ⌘V)" style="width:100%; box-sizing:border-box; background:#141414; border:1px solid #332d1f; color:#e0e0e0; padding:6px 8px; border-radius:4px; font-size:11px; margin-bottom:9px; outline:none;">

        <div style="display:flex; gap:6px; margin-bottom:8px;">
            <button type="button" id="bt-save-btn" style="flex:1; background:#d6b658; color:#0a0a0a; border:none; padding:9px; border-radius:4px; cursor:pointer; font-weight:bold; font-size:12px; transition:0.2s;">Скопировать в B<span id="bt-btn-row">${state.excelRow}</span> (Enter)</button>
            <button type="button" id="bt-close-btn" style="background:#1a1a1a; color:#a0a0a0; border:1px solid #332d1f; padding:9px 13px; border-radius:4px; cursor:pointer; transition:0.2s;">✕</button>
        </div>

        <div style="background:#121212; padding:6px 8px; border-radius:4px; border:1px solid #222; font-size:10px;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:4px;">
                <span style="color:#9e9e9e;">Серия: <b id="bt-batch-count" style="color:#d6b658;">0</b> сд. &#8226; <span id="bt-live-stats" style="color:#a0a0a0;">WR: 0% &#8226; 0.00 R</span></span>
            </div>
            <div style="display:flex; gap:4px;">
                <button type="button" id="bt-copy-batch" class="bt-mini-btn" style="flex:1; padding:4px 6px;">📋 Скопировать всю пачку (для B<span id="bt-batch-start-row">${getBatchStartRow()}</span>)</button>
                <button type="button" id="bt-clear-batch" class="bt-mini-btn" style="padding:4px 8px; color:#f23645;" title="Очистить пачку (повторное нажатие сбросит строку на B6)">🗑 Сброс</button>
            </div>
        </div>
    </div>

    <style>
        .bt-section-title { font-size: 10px; color: #787b86; margin-bottom: 3px; font-weight: bold; letter-spacing: 0.5px; }
        .bt-chip { background: #1a1a1a; color: #a0a0a0; border: 1px solid #2a2a2a; padding: 5px 6px; font-size: 11px; border-radius: 4px; cursor: pointer; text-align: center; transition: 0.15s ease; }
        .bt-chip:hover { border-color: #d6b658; color: #e0e0e0; }
        .bt-chip.active { background: var(--ac, #d6b658); color: #0a0a0a; border-color: var(--ac, #d6b658); font-weight: bold; }
        .bt-mini-btn { background: #1a1a1a; color: #999; border: 1px solid #2a2a2a; padding: 2px 6px; font-size: 10px; border-radius: 3px; cursor: pointer; transition: 0.15s; }
        .bt-mini-btn:hover { border-color: #d6b658; color: #e0e0e0; }
        .bt-mini-btn.active { background: #d6b658; color: #0a0a0a; border-color: #d6b658; font-weight: bold; }
        #bt-save-btn:hover { background: #e8c767; box-shadow: 0 0 10px rgba(214, 182, 88, 0.3); }
        #bt-fab-trigger:hover { border-color: #d6b658; background: #1a1a1a; }
        #bt-img:focus, #bt-comment:focus, #bt-err:focus, #bt-conf:focus, #bt-row-num:focus { border-color: #d6b658; }
    </style>
    `;

    document.body.insertAdjacentHTML('beforeend', modalHtml);

    const modalEl = document.getElementById('bt-modal');
    const toastEl = document.getElementById('bt-toast');
    const rowInputEl = document.getElementById('bt-row-num');
    const saveBtnEl = document.getElementById('bt-save-btn');
    const copyBatchBtnEl = document.getElementById('bt-copy-batch');
    const hkRecordBtn = document.getElementById('bt-hk-record-btn');
    const hkLabelEl = document.getElementById('bt-hk-label');
    const fabTriggerEl = document.getElementById('bt-fab-trigger');
    const fabHkEl = document.getElementById('bt-fab-hk');
    let toastTimer = null;

    function showToast(msg, type = 'success') {
        if (!toastEl) return;
        toastEl.innerText = msg;
        toastEl.style.borderLeftColor = type === 'error' ? '#f23645' : '#089981';
        toastEl.style.display = 'block';
        clearTimeout(toastTimer);
        toastTimer = setTimeout(() => {
            toastEl.style.display = 'none';
        }, 2800);
    }

    function setGroupValue(groupId, val) {
        const group = document.getElementById(groupId);
        if (!group) return;
        group.querySelectorAll('.bt-chip').forEach(b => {
            b.classList.toggle('active', b.dataset.val === val);
        });
    }

    function getGroupValue(groupId) {
        return document.querySelector(`#${groupId} .active`)?.dataset.val || '';
    }

    function autoSuggestError() {
        if (!currentTrade) return;
        const result = getGroupValue('bt-result');
        const liq = getGroupValue('bt-liq');
        const conf = document.getElementById('bt-conf').value;
        const errSelect = document.getElementById('bt-err');
        const hintEl = document.getElementById('bt-err-hint');

        if (result === 'Тейк' || result === 'Пропуск') {
            errSelect.value = '—';
            hintEl.innerText = '';
            return;
        }

        if (result === 'Стоп') {
            if (liq === 'нет') {
                errSelect.value = 'Вход без снятия ликвидности';
                hintEl.innerText = 'АВТО: нет ликвидности';
                return;
            }
            if (conf === 'Без подтверждения') {
                errSelect.value = 'Нет подтверждения';
                hintEl.innerText = 'АВТО: без подтверждения';
                return;
            }
            if (currentTrade.riskRewardNum > 0 && currentTrade.riskRewardNum < 2.5) {
                errSelect.value = 'Нарушен минимальный R:R';
                hintEl.innerText = 'АВТО: низкий R:R';
                return;
            }
        }
        hintEl.innerText = '';
    }

    function computeBatchStats() {
        let wins = 0, losses = 0, totalR = 0;
        for (const item of state.batchQueue) {
            if (!item || typeof item !== 'object') continue;
            if (item.result === 'Тейк') {
                wins++;
                totalR += Number(item.rr || item.trade?.riskRewardNum || 0);
            } else if (item.result === 'Стоп') {
                losses++;
                totalR -= 1;
            }
        }
        const closed = wins + losses;
        const wr = closed > 0 ? Math.round((wins / closed) * 100) : 0;
        const sign = totalR > 0 ? '+' : '';
        return `WR: ${wr}% (${wins}Т/${losses}С) \u2022 ${sign}${totalR.toFixed(2)} R`;
    }

    function updateUIControls() {
        document.querySelectorAll('#bt-sep-group .bt-mini-btn').forEach(btn => {
            btn.classList.toggle('active', btn.dataset.sep === state.decimalSeparator);
        });
        document.querySelectorAll('#bt-fmode-group .bt-mini-btn').forEach(btn => {
            btn.classList.toggle('active', btn.dataset.fmode === state.formulaMode);
        });

        const targetRow = isReEditingSameTrade ? state.lastSavedRow : state.excelRow;
        const batchStartRow = getBatchStartRow();
        const hkText = getHotkeyLabel();

        hkLabelEl.innerText = hkText;
        fabHkEl.innerText = `(${hkText})`;
        rowInputEl.value = targetRow;
        document.getElementById('bt-batch-count').innerText = state.batchQueue.length;
        document.getElementById('bt-live-stats').innerText = computeBatchStats();
        copyBatchBtnEl.innerHTML = `📋 Скопировать всю пачку (для B<span id="bt-batch-start-row">${batchStartRow}</span>)`;

        if (isReEditingSameTrade) {
            saveBtnEl.innerHTML = `Обновить B<span id="bt-btn-row">${targetRow}</span> (повтор сделки \u2022 Enter)`;
        } else {
            saveBtnEl.innerHTML = `Скопировать в B<span id="bt-btn-row">${targetRow}</span> (Enter)`;
        }

        if (!currentTrade) return;
        document.getElementById('bt-en').innerText = formatNumForSheet(currentTrade.entryPrice);
        document.getElementById('bt-st').innerText = formatNumForSheet(currentTrade.stopPrice);
        document.getElementById('bt-tg').innerText = formatNumForSheet(currentTrade.targetPrice);
        document.getElementById('bt-pts-preview').innerText =
            `Риск: ${formatNumForSheet(currentTrade.riskPts)} п. \u2022 Потенциал: ${formatNumForSheet(currentTrade.potPts)} п.`;

        const badge = document.getElementById('bt-badge');
        badge.innerText = `${currentTrade.direction} (R:R ${formatNumForSheet(currentTrade.riskReward)})`;
    }

    setGroupValue('bt-trend', state.trend);
    setGroupValue('bt-zone', state.zone);
    setGroupValue('bt-liq', state.liq);
    document.getElementById('bt-conf').value = state.conf;
    updateUIControls();

    // --- Обработчики событий и хоткеев ---

    fabTriggerEl.onclick = () => openModal();

    hkRecordBtn.onclick = (e) => {
        e.stopPropagation();
        if (isRecordingHotkey) {
            isRecordingHotkey = false;
            saveState({ customHotkey: null });
            hkRecordBtn.classList.remove('active');
            updateUIControls();
            showToast('Восстановлен стандартный хоткей: ' + getHotkeyLabel());
            return;
        }
        isRecordingHotkey = true;
        hkRecordBtn.classList.add('active');
        hkLabelEl.innerText = 'Нажмите клавишу...';
    };

    document.getElementById('bt-sep-group').addEventListener('click', e => {
        const btn = e.target.closest('.bt-mini-btn');
        if (!btn) return;
        const newSep = btn.dataset.sep;
        const patch = { decimalSeparator: newSep };

        if (newSep === ',' && state.formulaMode === 'en_comma') {
            patch.formulaMode = 'ru_semi';
        } else if (newSep === '.' && (state.formulaMode === 'ru_semi' || state.formulaMode === 'en_semi')) {
            patch.formulaMode = 'en_comma';
        }

        saveState(patch);
        updateUIControls();
    });

    document.getElementById('bt-fmode-group').addEventListener('click', e => {
        const btn = e.target.closest('.bt-mini-btn');
        if (!btn) return;
        const newMode = btn.dataset.fmode;
        const patch = { formulaMode: newMode };

        if (newMode === 'en_comma') {
            patch.decimalSeparator = '.';
        } else if (newMode === 'ru_semi' || newMode === 'en_semi') {
            patch.decimalSeparator = ',';
        }

        saveState(patch);
        updateUIControls();
    });

    function setExcelRow(val) {
        const num = Math.max(6, parseInt(val, 10) || 6);
        isReEditingSameTrade = false;
        saveState({ excelRow: num });
        updateUIControls();
    }
    document.getElementById('bt-row-dec').onclick = () => setExcelRow(state.excelRow - 1);
    document.getElementById('bt-row-inc').onclick = () => setExcelRow(state.excelRow + 1);
    rowInputEl.addEventListener('change', e => setExcelRow(e.target.value));

    document.querySelectorAll('.bt-group').forEach(group => {
        group.addEventListener('click', e => {
            const chip = e.target.closest('.bt-chip');
            if (!chip) return;
            group.querySelectorAll('.bt-chip').forEach(b => b.classList.remove('active'));
            chip.classList.add('active');

            const val = chip.dataset.val;
            if (group.id === 'bt-trend') saveState({ trend: val });
            if (group.id === 'bt-zone') saveState({ zone: val });
            if (group.id === 'bt-liq') {
                saveState({ liq: val });
                if (val === 'нет' && getGroupValue('bt-result') === 'Пропуск') {
                    document.getElementById('bt-conf').value = 'Без подтверждения';
                }
                autoSuggestError();
            }
            if (group.id === 'bt-result') {
                if (val === 'Пропуск' && getGroupValue('bt-liq') === 'нет') {
                    document.getElementById('bt-conf').value = 'Без подтверждения';
                }
                autoSuggestError();
            }
        });
    });

    document.getElementById('bt-conf').addEventListener('change', e => {
        saveState({ conf: e.target.value });
        autoSuggestError();
    });

    copyBatchBtnEl.onclick = () => {
        if (!state.batchQueue.length) {
            showToast('Пачка сделок пуста', 'error');
            return;
        }
        const startRow = getBatchStartRow();
        const allRows = state.batchQueue.map(item => {
            if (typeof item === 'string') return item;
            if (item.trade) return buildRowTsv(item);
            return item.tsv || '';
        }).join('\n');

        navigator.clipboard.writeText(allRows).then(() => {
            showToast(`✓ Скопировано сделок: ${state.batchQueue.length}! Поставьте курсор в B${startRow} и нажмите ${isMac ? '⌘V' : 'Ctrl+V'}`);
        });
    };

    document.getElementById('bt-clear-batch').onclick = () => {
        isReEditingSameTrade = false;
        if (state.batchQueue.length > 0) {
            saveState({ batchQueue: [], lastSavedTradeSig: '' });
            updateUIControls();
            showToast(`Пачка очищена (следующая строка: B${state.excelRow})`);
        } else {
            saveState({ batchQueue: [], excelRow: 6, lastSavedTradeSig: '' });
            updateUIControls();
            showToast('Счетчик строк сброшен на B6');
        }
    };

    (function initDrag() {
        const handle = document.getElementById('bt-drag-handle');
        let isDragging = false, startX = 0, startY = 0, initialTop = 0, initialRight = 0;

        handle.addEventListener('mousedown', e => {
            if (e.target.closest('#bt-hk-record-btn')) return;
            isDragging = true;
            startX = e.clientX;
            startY = e.clientY;
            const rect = modalEl.getBoundingClientRect();
            initialTop = rect.top;
            initialRight = window.innerWidth - rect.right;
            e.preventDefault();
        });

        window.addEventListener('mousemove', e => {
            if (!isDragging) return;
            const newTop = Math.max(10, Math.min(window.innerHeight - 100, initialTop + (e.clientY - startY)));
            const newRight = Math.max(10, Math.min(window.innerWidth - 400, initialRight - (e.clientX - startX)));
            modalEl.style.top = `${newTop}px`;
            modalEl.style.right = `${newRight}px`;
        });

        window.addEventListener('mouseup', () => {
            if (!isDragging) return;
            isDragging = false;
            saveState({
                posTop: parseInt(modalEl.style.top, 10) || 20,
                posRight: parseInt(modalEl.style.right, 10) || 20
            });
        });
    })();

    function openModal() {
        currentTrade = extractTrade();
        if (!currentTrade) return;

        isReEditingSameTrade = (
            state.lastSavedTradeSig === currentTrade.signature &&
            state.batchQueue.length > 0
        );

        const badge = document.getElementById('bt-badge');
        badge.style.background = currentTrade.direction === 'LONG' ? '#089981' : '#f23645';

        document.getElementById('bt-sym-tf').innerText = `${currentTrade.symbol} \u2022 ${currentTrade.timeframe}`;
        document.getElementById('bt-time').innerText = `Дата: ${currentTrade.candleDate}`;

        const currentTrend = getGroupValue('bt-trend');
        if (currentTrend === 'Тренд вверх' || currentTrend === 'Тренд вниз') {
            const autoTrend = currentTrade.direction === 'LONG' ? 'Тренд вверх' : 'Тренд вниз';
            setGroupValue('bt-trend', autoTrend);
            saveState({ trend: autoTrend });
        }

        setGroupValue('bt-session', currentTrade.session);
        setGroupValue('bt-result', currentTrade.autoResult);
        document.getElementById('bt-err').value = '—';

        updateUIControls();
        autoSuggestError();

        if (document.activeElement && typeof document.activeElement.blur === 'function') {
            document.activeElement.blur();
        }

        modalEl.style.display = 'block';
    }

    function saveTrade() {
        if (!currentTrade) return;

        const rowNum = isReEditingSameTrade ? state.lastSavedRow : state.excelRow;

        const record = {
            row: rowNum,
            trade: currentTrade,
            session: getGroupValue('bt-session') || currentTrade.session,
            trend: getGroupValue('bt-trend'),
            entryZone: getGroupValue('bt-zone'),
            liquidityTaken: getGroupValue('bt-liq'),
            confirmation: document.getElementById('bt-conf').value,
            result: getGroupValue('bt-result'),
            execError: document.getElementById('bt-err').value,
            commentStr: sanitizeCell(document.getElementById('bt-comment').value),
            imgUrl: sanitizeCell(document.getElementById('bt-img').value),
            rr: currentTrade.riskRewardNum
        };

        const tsvString = buildRowTsv(record);
        record.tsv = tsvString;

        navigator.clipboard.writeText(tsvString).then(() => {
            const newBatch = [...state.batchQueue];

            if (isReEditingSameTrade && newBatch.length > 0) {
                newBatch[newBatch.length - 1] = record;
            } else {
                newBatch.push(record);
            }

            const nextRow = isReEditingSameTrade ? state.excelRow : (rowNum + 1);

            saveState({
                batchQueue: newBatch,
                excelRow: nextRow,
                lastSavedRow: rowNum,
                lastSavedTradeSig: currentTrade.signature
            });

            isReEditingSameTrade = false;
            modalEl.style.display = 'none';
            document.getElementById('bt-img').value = '';
            document.getElementById('bt-comment').value = '';
            updateUIControls();
            showToast(`✓ Строка B${rowNum} скопирована в буфер! (Следующая: B${nextRow})`);
        }).catch(err => {
            console.error('Ошибка буфера:', err);
            showToast('Не удалось скопировать в буфер', 'error');
        });
    }

    function matchesOpenHotkey(e) {
        if (state.customHotkey) {
            const hk = state.customHotkey;
            return (
                e.code === hk.code &&
                !!e.altKey === !!hk.alt &&
                !!e.ctrlKey === !!hk.ctrl &&
                !!e.shiftKey === !!hk.shift &&
                !!e.metaKey === !!hk.meta
            );
        }
        if (e.key === 'F2' || e.code === 'F2') return true;
        if (e.altKey && !e.ctrlKey && !e.metaKey && e.code === 'KeyQ') return true;
        return false;
    }

    window.addEventListener('paste', e => {
        if (modalEl.style.display !== 'block') return;
        const activeTag = document.activeElement?.tagName;
        if (activeTag === 'INPUT' || activeTag === 'TEXTAREA') return;

        const pasted = (e.clipboardData || window.clipboardData)?.getData('text') || '';
        if (/^https?:\/\//i.test(pasted.trim())) {
            e.preventDefault();
            document.getElementById('bt-img').value = pasted.trim();
            showToast('Ссылка на скриншот вставлена ✓');
        }
    });

    window.addEventListener('keydown', e => {
        if (isRecordingHotkey) {
            e.preventDefault();
            e.stopPropagation();
            if (['Shift', 'Control', 'Alt', 'Meta'].includes(e.key)) return;

            if (e.key === 'Escape') {
                isRecordingHotkey = false;
                hkRecordBtn.classList.remove('active');
                updateUIControls();
                return;
            }

            const parts = [];
            if (e.metaKey) parts.push(isMac ? '⌘' : 'Win+');
            if (e.ctrlKey) parts.push(isMac ? '⌃' : 'Ctrl+');
            if (e.altKey) parts.push(isMac ? '⌥' : 'Alt+');
            if (e.shiftKey) parts.push(isMac ? '⇧' : 'Shift+');

            const cleanKeyName = e.code.replace(/^Key/, '').replace(/^Digit/, '');
            parts.push(cleanKeyName);

            const customHotkey = {
                code: e.code,
                alt: !!e.altKey,
                ctrl: !!e.ctrlKey,
                shift: !!e.shiftKey,
                meta: !!e.metaKey,
                label: parts.join('')
            };

            isRecordingHotkey = false;
            hkRecordBtn.classList.remove('active');
            saveState({ customHotkey });
            updateUIControls();
            showToast(`Новый хоткей сохранен: ${customHotkey.label}`);
            return;
        }

        const isModalOpen = modalEl.style.display === 'block';
        const inInput = e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.isContentEditable;

        if (matchesOpenHotkey(e) && !inInput) {
            e.preventDefault();
            e.stopPropagation();
            openModal();
            return;
        }

        if (!isModalOpen) return;

        if (e.key === 'Enter') {
            e.preventDefault();
            e.stopPropagation();
            saveTrade();
            return;
        }

        if (e.key === 'Escape') {
            e.preventDefault();
            e.stopPropagation();
            if (inInput) {
                e.target.blur();
            } else {
                modalEl.style.display = 'none';
            }
            return;
        }

        if (e.key === 'Tab' && !inInput) {
            e.preventDefault();
            document.getElementById('bt-comment').focus();
            return;
        }

        if (!inInput && !e.ctrlKey && !e.altKey && !e.metaKey) {
            const zoneMap = {
                '1': 'Ордер-блок',
                '2': 'Имбаланс',
                '3': 'Материнская свеча',
                '4': 'Зона ликвидности',
                '5': 'Правильный откат'
            };
            const resMap = {
                'q': 'Тейк', 'й': 'Тейк',
                'w': 'Стоп', 'ц': 'Стоп',
                'e': 'Безубыток', 'у': 'Безубыток',
                'r': 'Пропуск', 'к': 'Пропуск'
            };
            const trendMap = {
                'a': 'Тренд вверх', 'ф': 'Тренд вверх',
                's': 'Тренд вниз', 'ы': 'Тренд вниз',
                'd': 'Боковик', 'в': 'Боковик',
                'f': 'Импульс/новости', 'а': 'Импульс/новости'
            };
            const confMap = {
                'z': 'Слом структуры', 'я': 'Слом структуры',
                'x': 'Реакция от зоны', 'ч': 'Реакция от зоны',
                'c': 'Поглощение', 'с': 'Поглощение',
                'v': 'Снятие ликвидности', 'м': 'Снятие ликвидности',
                'b': 'Без подтверждения', 'и': 'Без подтверждения'
            };

            const keyLower = e.key.toLowerCase();

            if (zoneMap[e.key]) {
                e.preventDefault();
                setGroupValue('bt-zone', zoneMap[e.key]);
                saveState({ zone: zoneMap[e.key] });
            } else if (resMap[keyLower]) {
                e.preventDefault();
                setGroupValue('bt-result', resMap[keyLower]);
                if (resMap[keyLower] === 'Пропуск' && getGroupValue('bt-liq') === 'нет') {
                    document.getElementById('bt-conf').value = 'Без подтверждения';
                }
                autoSuggestError();
            } else if (trendMap[keyLower]) {
                e.preventDefault();
                setGroupValue('bt-trend', trendMap[keyLower]);
                saveState({ trend: trendMap[keyLower] });
            } else if (confMap[keyLower]) {
                e.preventDefault();
                document.getElementById('bt-conf').value = confMap[keyLower];
                saveState({ conf: confMap[keyLower] });
                autoSuggestError();
            } else if (e.code === 'Space') {
                e.preventDefault();
                const nextLiq = getGroupValue('bt-liq') === 'да' ? 'нет' : 'да';
                setGroupValue('bt-liq', nextLiq);
                saveState({ liq: nextLiq });
                if (nextLiq === 'нет' && getGroupValue('bt-result') === 'Пропуск') {
                    document.getElementById('bt-conf').value = 'Без подтверждения';
                }
                autoSuggestError();
            }
        }
    }, true);

    document.getElementById('bt-save-btn').onclick = saveTrade;
    document.getElementById('bt-close-btn').onclick = () => { modalEl.style.display = 'none'; };

})();
