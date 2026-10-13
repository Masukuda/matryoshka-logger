// ==UserScript==
// @name         Matryoshka Backtest Auto-Logger
// @namespace    http://tampermonkey.net/
// @version      4.1
// @description  Экспорт сделок из TradingView в Excel / Google Таблицы (Стратегия «Матрёшка»)
// @match        https://*.tradingview.com/*
// @grant        unsafeWindow
// @updateURL    https://raw.githubusercontent.com/Masukuda/matryoshka-logger/main/matryoshka.user.js
// @downloadURL  https://raw.githubusercontent.com/Masukuda/matryoshka-logger/main/matryoshka.user.js
// @run-at       document-idle
// ==/UserScript==

(function() {
    'use strict';

    // Не запускаем скрипт внутри вложенных технических iframe TradingView
    if (window.self !== window.top) return;

    const pageWindow = typeof unsafeWindow !== 'undefined' ? unsafeWindow : window;
    const STORAGE_KEY = 'matryoshka_logger_state_v4';

    // Настройки по умолчанию
    const defaultState = {
        decimalSeparator: ',', // ',' для русскоязычного Excel/Sheets, '.' для US
        trend: 'Тренд вверх',
        zone: 'Ордер-блок',
        liq: 'да',
        conf: 'Слом структуры',
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
        } catch (e) {
            console.warn('Не удалось сохранить состояние в localStorage:', e);
        }
    }

    let state = loadState();
    let currentTrade = null;

    // 1. Точное определение сессии с учетом летнего/зимнего времени (DST)
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

        // Нью-Йорк: с 08:00 до 17:00 по времени Нью-Йорка (учитывает EDT/EST автоматически)
        if (nyHour >= 8 && nyHour < 17) {
            return 'Нью-Йорк';
        }
        // Лондон: с 08:00 по Лондону до открытия Нью-Йорка (учитывает BST/GMT автоматически)
        if (lonHour >= 8 && nyHour < 8) {
            return 'Лондон';
        }
        // Азия (Токио + Сидней): после закрытия Нью-Йорка (17:00 NY) и до открытия Лондона (08:00 LON)
        return 'Азия';
    }

    // 2. Форматирование таймфрейма (1 -> M1, 60 -> H1, 240 -> H4, 1D -> D1)
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

    // 3. Форматирование даты в DD.MM.YYYY
    function formatDate(timestamp) {
        const d = new Date(timestamp * 1000);
        const day = String(d.getDate()).padStart(2, '0');
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const year = d.getFullYear();
        return `${day}.${month}.${year}`;
    }

    // 4. Авто-расчет результата (Тейк / Стоп / Безубыток)
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

    // Безопасное чтение примитива или WatchedValue (.value()), НЕ теряющее контекст this
    function readWatched(val) {
        if (val != null && typeof val === 'object' && typeof val.value === 'function') {
            return val.value();
        }
        return val;
    }

    // 5. Извлечение сделки из ядра TradingView (с сохранением контекста this)
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

            // 1) Сначала проверяем, выделен ли конкретный инструмент на графике
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

            // 2) Если ничего не выделено — берем последний нарисованный
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

            // Шаг цены (Price Step) из ядра TradingView
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
            const rr = (stopLevel > 0) ? (profitLevel / stopLevel).toFixed(2) : '0.00';

            const startIdx = points[0]?.index ?? 0;
            const endIdx = points[1]?.index ?? startIdx;
            const calculatedOutcome = calculateOutcome(bars, startIdx, endIdx, isLong, stop, target);

            // Точное вычисление знаков после запятой (без бага с 1e-6 и float-погрешностями)
            const decimals = priceStep < 1
                ? (priceStep.toFixed(8).replace(/0+$/, '').split('.')[1]?.length || 3)
                : 2;

            // Очистка тикера от префикса биржи (OANDA:XAUUSD -> XAUUSD)
            const rawSymbol = readWatched(props?.symbol) || symInfo.name || 'XAUUSD';
            const cleanSymbol = String(rawSymbol).split(':').pop();
            const rawInterval = readWatched(props?.interval) || '1';

            return {
                symbol: cleanSymbol,
                timeframe: formatTimeframe(rawInterval),
                direction: isLong ? 'LONG' : 'SHORT',
                entryPrice: Number(entry.toFixed(decimals)),
                stopPrice: Number(stop.toFixed(decimals)),
                targetPrice: Number(target.toFixed(decimals)),
                riskReward: rr,
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

    // Форматирование числа с учетом выбранного разделителя (точка или запятая)
    function formatNumForSheet(num) {
        if (num == null || num === '') return '';
        const str = String(num);
        return state.decimalSeparator === ',' ? str.replace('.', ',') : str.replace(',', '.');
    }

    // Очистка текста от табуляции и переносов строк, чтобы не сбивать колонки Excel
    function sanitizeCell(str) {
        return String(str || '').replace(/[\t\r\n]+/g, ' ').trim();
    }

    // 6. UI Модального окна
    const modalHtml = `
    <div id="bt-toast" style="display:none; position:fixed; bottom:24px; right:24px; z-index:10000000; background:#141414; color:#e0e0e0; padding:10px 16px; border-radius:6px; border-left:4px solid #d6b658; box-shadow:0 8px 24px rgba(0,0,0,0.85); font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif; font-size:12px; font-weight:500;"></div>

    <div id="bt-modal" style="display:none; position:fixed; top:${state.posTop}px; right:${state.posRight}px; z-index:9999999; background:#0a0a0a; color:#e0e0e0; padding:18px; border-radius:8px; box-shadow:0 12px 32px rgba(0,0,0,0.85); border:1px solid #332d1f; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif; width:365px; user-select:none;">

        <div id="bt-drag-handle" style="text-align:center; margin-bottom:12px; cursor:move; padding-bottom:6px; border-bottom:1px solid #1f1f1f;">
            <div style="font-size:10px; letter-spacing:2px; color:#d6b658; text-transform:uppercase; margin-bottom:2px;">Стратегия «Матрёшка»</div>
            <div style="font-family:'Georgia', 'Playfair Display', serif; font-size:20px; font-weight:bold; color:#d6b658; letter-spacing:1px;">ЭКСПОРТ В ТАБЛИЦУ</div>
        </div>

        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px; background:#121212; padding:6px 10px; border-radius:5px; border:1px solid #222;">
            <span style="font-size:10px; color:#9e9e9e; font-weight:bold;">РАЗДЕЛИТЕЛЬ ДРОБИ (EXCEL / SHEETS):</span>
            <div id="bt-sep-group" style="display:flex; gap:4px;">
                <button type="button" class="bt-sep-btn" data-sep=",">Запятая ( , )</button>
                <button type="button" class="bt-sep-btn" data-sep=".">Точка ( . )</button>
            </div>
        </div>

        <div style="background:#141414; border:1px solid #2a2a2a; border-radius:6px; padding:10px; margin-bottom:12px; font-size:12px; line-height:1.6;">
            <div style="display:flex; justify-content:space-between; margin-bottom:4px;">
                <span id="bt-sym-tf" style="color:#d6b658; font-weight:bold;"></span>
                <span id="bt-badge" style="background:#089981; color:#fff; font-size:10px; padding:2px 6px; border-radius:4px; font-weight:bold;"></span>
            </div>
            <div style="color:#a0a0a0; display:flex; justify-content:space-between;">
                <span>Вход: <b id="bt-en" style="color:#e0e0e0;"></b></span>
                <span>Стоп: <b id="bt-st" style="color:#e0e0e0;"></b></span>
                <span>Цель: <b id="bt-tg" style="color:#d6b658;"></b></span>
            </div>
            <div style="color:#787b86; margin-top:4px; font-size:11px; display:flex; justify-content:space-between;">
                <span id="bt-time"></span>
                <span style="color:#5c6068;">Вставка ссылки: Ctrl+V</span>
            </div>
        </div>

        <div style="display:flex; justify-content:space-between; align-items:flex-end; margin-bottom:4px;">
            <div class="bt-section-title" style="margin:0;">ТОРГОВАЯ СЕССИЯ</div>
            <span style="font-size:9px; color:#089981;">АВТО (DST) ✓</span>
        </div>
        <div class="bt-group" id="bt-session" style="display:grid; grid-template-columns:repeat(3, 1fr); gap:4px; margin-bottom:10px;">
            <button type="button" class="bt-chip" data-val="Азия">Азия</button>
            <button type="button" class="bt-chip" data-val="Лондон">Лондон</button>
            <button type="button" class="bt-chip" data-val="Нью-Йорк">Нью-Йорк</button>
        </div>

        <div class="bt-section-title">ФАЗА РЫНКА</div>
        <div class="bt-group" id="bt-trend" style="display:grid; grid-template-columns:1fr 1fr; gap:4px; margin-bottom:10px;">
            <button type="button" class="bt-chip" data-val="Тренд вверх">Тренд вверх</button>
            <button type="button" class="bt-chip" data-val="Тренд вниз">Тренд вниз</button>
            <button type="button" class="bt-chip" data-val="Боковик">Боковик</button>
            <button type="button" class="bt-chip" data-val="Импульс/новости">Импульс/новости</button>
        </div>

        <div class="bt-section-title">ЗОНА ВХОДА</div>
        <div class="bt-group" id="bt-zone" style="display:grid; grid-template-columns:1fr 1fr; gap:4px; margin-bottom:10px;">
            <button type="button" class="bt-chip" data-val="Ордер-блок">Ордер-блок</button>
            <button type="button" class="bt-chip" data-val="Имбаланс">Имбаланс</button>
            <button type="button" class="bt-chip" data-val="Материнская свеча">Мать-свеча</button>
            <button type="button" class="bt-chip" data-val="Зона ликвидности">Ликвидность</button>
            <button type="button" class="bt-chip" data-val="Правильный откат" style="grid-column: span 2;">Правильный откат</button>
        </div>

        <div style="display:grid; grid-template-columns:1fr 2fr; gap:10px; margin-bottom:10px;">
            <div>
                <div class="bt-section-title">ЛИКВИДНОСТЬ</div>
                <div class="bt-group" id="bt-liq" style="display:flex; gap:4px;">
                    <button type="button" class="bt-chip" style="flex:1;" data-val="да">Да</button>
                    <button type="button" class="bt-chip" style="flex:1;" data-val="нет">Нет</button>
                </div>
            </div>
            <div>
                <div class="bt-section-title">ПОДТВЕРЖДЕНИЕ</div>
                <select id="bt-conf" style="width:100%; background:#1a1a1a; color:#d6b658; border:1px solid #332d1f; padding:5px; border-radius:4px; font-size:11px; outline:none; cursor:pointer;">
                    <option value="Слом структуры">Слом структуры</option>
                    <option value="Реакция от зоны">Реакция от зоны</option>
                    <option value="Поглощение">Поглощение</option>
                    <option value="Снятие ликвидности">Снятие ликвидности</option>
                    <option value="Без подтверждения">Без подтверждения</option>
                </select>
            </div>
        </div>

        <div style="display:flex; justify-content:space-between; align-items:flex-end; margin-bottom:4px;">
            <div class="bt-section-title" style="margin:0;">РЕЗУЛЬТАТ СДЕЛКИ</div>
            <span style="font-size:9px; color:#089981;">АВТОРАСЧЕТ ✓</span>
        </div>
        <div class="bt-group" id="bt-result" style="display:grid; grid-template-columns:repeat(4, 1fr); gap:4px; margin-bottom:10px;">
            <button type="button" class="bt-chip" data-val="Тейк" style="--ac:#089981;">Тейк</button>
            <button type="button" class="bt-chip" data-val="Стоп" style="--ac:#f23645;">Стоп</button>
            <button type="button" class="bt-chip" data-val="Безубыток">БУ</button>
            <button type="button" class="bt-chip" data-val="Пропуск">Пропуск</button>
        </div>

        <div class="bt-section-title">ОШИБКА ИСПОЛНЕНИЯ</div>
        <select id="bt-err" style="width:100%; background:#1a1a1a; color:#e0e0e0; border:1px solid #332d1f; padding:6px; border-radius:4px; font-size:11px; outline:none; cursor:pointer; margin-bottom:10px;">
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

        <input type="text" id="bt-comment" placeholder="Комментарий к сделке (колонка 19)..." style="width:100%; box-sizing:border-box; background:#141414; border:1px solid #332d1f; color:#e0e0e0; padding:8px; border-radius:4px; font-size:11px; margin-bottom:6px; outline:none;">
        <input type="text" id="bt-img" placeholder="Ссылка на скриншот (колонка 20, Alt+S -> Ctrl+V)" style="width:100%; box-sizing:border-box; background:#141414; border:1px solid #332d1f; color:#e0e0e0; padding:8px; border-radius:4px; font-size:11px; margin-bottom:14px; outline:none;">

        <div style="display:flex; gap:8px;">
            <button type="button" id="bt-save-btn" style="flex:1; background:#d6b658; color:#0a0a0a; border:none; padding:10px; border-radius:4px; cursor:pointer; font-weight:bold; font-size:13px; transition:0.2s;">Скопировать строку (Enter)</button>
            <button type="button" id="bt-close-btn" style="background:#1a1a1a; color:#a0a0a0; border:1px solid #332d1f; padding:10px 14px; border-radius:4px; cursor:pointer; transition:0.2s;">✕</button>
        </div>
    </div>

    <style>
        .bt-section-title { font-size: 10px; color: #787b86; margin-bottom: 5px; font-weight: bold; letter-spacing: 0.5px; }
        .bt-chip { background: #1a1a1a; color: #a0a0a0; border: 1px solid #2a2a2a; padding: 6px; font-size: 11px; border-radius: 4px; cursor: pointer; text-align: center; transition: 0.15s ease; }
        .bt-chip:hover { border-color: #d6b658; color: #e0e0e0; }
        .bt-chip.active { background: var(--ac, #d6b658); color: #0a0a0a; border-color: var(--ac, #d6b658); font-weight: bold; }
        .bt-sep-btn { background: #1a1a1a; color: #888; border: 1px solid #2a2a2a; padding: 3px 8px; font-size: 10px; border-radius: 3px; cursor: pointer; transition: 0.15s; }
        .bt-sep-btn.active { background: #d6b658; color: #0a0a0a; border-color: #d6b658; font-weight: bold; }
        #bt-save-btn:hover { background: #e8c767; box-shadow: 0 0 10px rgba(214, 182, 88, 0.3); }
        #bt-img:focus, #bt-comment:focus, #bt-err:focus, #bt-conf:focus { border-color: #d6b658; }
    </style>
    `;

    document.body.insertAdjacentHTML('beforeend', modalHtml);

    const modalEl = document.getElementById('bt-modal');
    const toastEl = document.getElementById('bt-toast');
    let toastTimer = null;

    function showToast(msg, type = 'success') {
        if (!toastEl) return;
        toastEl.innerText = msg;
        toastEl.style.borderLeftColor = type === 'error' ? '#f23645' : '#089981';
        toastEl.style.display = 'block';
        clearTimeout(toastTimer);
        toastTimer = setTimeout(() => {
            toastEl.style.display = 'none';
        }, 2600);
    }

    function setGroupValue(groupId, val) {
        const group = document.getElementById(groupId);
        if (!group) return;
        group.querySelectorAll('.bt-chip').forEach(b => {
            b.classList.toggle('active', b.dataset.val === val);
        });
    }

    function updatePricePreview() {
        document.querySelectorAll('.bt-sep-btn').forEach(btn => {
            btn.classList.toggle('active', btn.dataset.sep === state.decimalSeparator);
        });
        if (!currentTrade) return;
        document.getElementById('bt-en').innerText = formatNumForSheet(currentTrade.entryPrice);
        document.getElementById('bt-st').innerText = formatNumForSheet(currentTrade.stopPrice);
        document.getElementById('bt-tg').innerText = formatNumForSheet(currentTrade.targetPrice);
        const badge = document.getElementById('bt-badge');
        badge.innerText = `${currentTrade.direction} (R:R ${formatNumForSheet(currentTrade.riskReward)})`;
    }

    setGroupValue('bt-trend', state.trend);
    setGroupValue('bt-zone', state.zone);
    setGroupValue('bt-liq', state.liq);
    document.getElementById('bt-conf').value = state.conf;
    updatePricePreview();

    document.getElementById('bt-sep-group').addEventListener('click', e => {
        const btn = e.target.closest('.bt-sep-btn');
        if (!btn) return;
        saveState({ decimalSeparator: btn.dataset.sep });
        updatePricePreview();
    });

    document.querySelectorAll('.bt-group').forEach(group => {
        group.addEventListener('click', e => {
            const chip = e.target.closest('.bt-chip');
            if (!chip) return;
            group.querySelectorAll('.bt-chip').forEach(b => b.classList.remove('active'));
            chip.classList.add('active');

            if (group.id === 'bt-trend') saveState({ trend: chip.dataset.val });
            if (group.id === 'bt-zone') saveState({ zone: chip.dataset.val });
            if (group.id === 'bt-liq') saveState({ liq: chip.dataset.val });
        });
    });

    document.getElementById('bt-conf').addEventListener('change', e => {
        saveState({ conf: e.target.value });
    });

    // Перетаскивание окна мышкой за шапку
    (function initDrag() {
        const handle = document.getElementById('bt-drag-handle');
        let isDragging = false, startX = 0, startY = 0, initialTop = 0, initialRight = 0;

        handle.addEventListener('mousedown', e => {
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
            const newRight = Math.max(10, Math.min(window.innerWidth - 380, initialRight - (e.clientX - startX)));
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

        const badge = document.getElementById('bt-badge');
        badge.style.background = currentTrade.direction === 'LONG' ? '#089981' : '#f23645';

        document.getElementById('bt-sym-tf').innerText = `${currentTrade.symbol} | ${currentTrade.timeframe}`;
        document.getElementById('bt-time').innerText = `Дата свечи: ${currentTrade.candleDate}`;

        updatePricePreview();
        setGroupValue('bt-session', currentTrade.session);
        setGroupValue('bt-result', currentTrade.autoResult);

        document.getElementById('bt-err').value = '—';
        modalEl.style.display = 'block';

        setTimeout(() => document.getElementById('bt-img').focus(), 40);
    }

    function saveTrade() {
        if (!currentTrade) return;

        const session = document.querySelector('#bt-session .active')?.dataset.val || currentTrade.session;
        const trend = document.querySelector('#bt-trend .active')?.dataset.val || '';
        const entryZone = document.querySelector('#bt-zone .active')?.dataset.val || '';
        const liquidityTaken = document.querySelector('#bt-liq .active')?.dataset.val || '';
        const confirmation = document.getElementById('bt-conf').value;
        const result = document.querySelector('#bt-result .active')?.dataset.val || '';
        const execError = document.getElementById('bt-err').value;

        const commentStr = sanitizeCell(document.getElementById('bt-comment').value);
        const imgUrl = sanitizeCell(document.getElementById('bt-img').value);

        // Строка для Excel / Google Таблиц (20 колонок: комментарий — 19-я, скриншот — 20-я)
        const tsvArray = [
            currentTrade.candleDate,                     // 1. Дата
            session,                                     // 2. Сессия
            trend,                                       // 3. Фаза рынка
            currentTrade.timeframe,                      // 4. Таймфрейм
            entryZone,                                   // 5. Зона входа
            liquidityTaken,                              // 6. Ликвидность
            confirmation,                                // 7. Подтверждение
            formatNumForSheet(currentTrade.entryPrice),  // 8. Цена входа
            formatNumForSheet(currentTrade.stopPrice),   // 9. Стоп-лосс
            formatNumForSheet(currentTrade.targetPrice), // 10. Тейк-профит
            '',                                          // 11. Пропуск: Риск (формула)
            '',                                          // 12. Пропуск: Потенциал (формула)
            '',                                          // 13. Пропуск: R:R (формула)
            result,                                      // 14. Результат сделки
            '',                                          // 15. Пропуск: Итог (формула)
            '',                                          // 16. Пропуск: Накопительно (формула)
            '',                                          // 17. Пропуск: Просадка (формула)
            execError,                                   // 18. Ошибка исполнения
            commentStr,                                  // 19. Комментарий
            imgUrl                                       // 20. Ссылка на скриншот (отдельная ячейка -> кликабельная!)
        ];

        const tsvString = tsvArray.join('\t');

        navigator.clipboard.writeText(tsvString).then(() => {
            modalEl.style.display = 'none';
            document.getElementById('bt-img').value = '';
            document.getElementById('bt-comment').value = '';
            showToast('✓ Сделка скопирована в буфер обмена!');
        }).catch(err => {
            console.error('Ошибка буфера:', err);
            showToast('Не удалось скопировать в буфер', 'error');
        });
    }

    // Умный перехват Ctrl+V при открытом окне
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

    // Горячие клавиши
    window.addEventListener('keydown', e => {
        const isModalOpen = modalEl.style.display === 'block';
        const inInput = e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.isContentEditable;

        if (e.key === 'F2' && !inInput) {
            e.preventDefault();
            e.stopPropagation();
            openModal();
        } else if (e.key === 'Enter' && isModalOpen) {
            e.preventDefault();
            e.stopPropagation();
            saveTrade();
        } else if (e.key === 'Escape' && isModalOpen) {
            e.preventDefault();
            e.stopPropagation();
            modalEl.style.display = 'none';
        }
    }, true);

    document.getElementById('bt-save-btn').onclick = saveTrade;
    document.getElementById('bt-close-btn').onclick = () => { modalEl.style.display = 'none'; };

})();
