// ==UserScript==
// @name         Matryoshka Backtest Auto-Logger
// @namespace    http://tampermonkey.net/
// @version      3.1
// @description  Сделать экспорт сделок в Excel проще
// @match        https://*.tradingview.com/*
// @grant        unsafeWindow
// @updateURL    https://raw.githubusercontent.com/Masukuda/matryoshka-logger/main/matryoshka.user.js
// @downloadURL  https://raw.githubusercontent.com/Masukuda/matryoshka-logger/main/matryoshka.user.js
// @run-at       document-idle
// ==/UserScript==

(function() {
    'use strict';

    const pageWindow = typeof unsafeWindow !== 'undefined' ? unsafeWindow : window;

    // 1. Определение сессии
    function detectSession(timestamp) {
        if (!timestamp) return 'Не определено';
        const hour = new Date(timestamp * 1000).getUTCHours();
        if (hour >= 0 && hour < 7) return 'Азия';
        if (hour >= 7 && hour < 13) return 'Лондон';
        return 'Нью-Йорк';
    }

    // 2. Авто-расчет результата (Тейк/Стоп/БУ)
    function calculateOutcome(bars, startIndex, endIndex, isLong, stopPrice, targetPrice) {
        if (!bars || typeof bars.valueAt !== 'function') return 'Тейк';
        const from = Math.min(startIndex, endIndex);
        const to = Math.max(startIndex, endIndex);
        let orderFilled = false;

        for (let i = from; i <= to; i++) {
            const bar = bars.valueAt(i);
            if (!bar || !Array.isArray(bar)) continue;
            const high = bar[2], low = bar[3];

            if (!orderFilled) {
                if (isLong) {
                    if (low <= targetPrice && high >= targetPrice && low > stopPrice) orderFilled = true;
                    orderFilled = true;
                } else {
                    orderFilled = true;
                }
            }

            if (isLong) {
                const hitTarget = high >= targetPrice;
                const hitStop = low <= stopPrice;
                if (hitTarget && hitStop) return 'Стоп';
                if (hitStop) return 'Стоп';
                if (hitTarget) return 'Тейк';
            } else {
                const hitTarget = low <= targetPrice;
                const hitStop = high >= stopPrice;
                if (hitTarget && hitStop) return 'Стоп';
                if (hitStop) return 'Стоп';
                if (hitTarget) return 'Тейк';
            }
        }
        return 'Безубыток';
    }

    // Форматирование даты в DD.MM.YYYY
    function formatDate(timestamp) {
        const d = new Date(timestamp * 1000);
        const day = String(d.getDate()).padStart(2, '0');
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const year = d.getFullYear();
        return `${day}.${month}.${year}`;
    }

    // 3. Извлечение сделки
    function extractTrade() {
        try {
            const c = pageWindow._exposed_chartWidgetCollection;
            if (!c) return null;

            const widget = typeof c.activeChartWidget === 'function' ? c.activeChartWidget() : (c.activeChartWidget?.value ? c.activeChartWidget.value() : c.activeChartWidget);
            if (!widget) return null;

            const rawModel = typeof widget.model === 'function' ? widget.model() : (widget.model || widget._model);
            const model = rawModel?.m_model || rawModel;
            const ms = model.mainSeries();
            const bars = typeof ms.bars === 'function' ? ms.bars() : ms.bars;
            const tools = model.allLineTools();
            const list = Array.isArray(tools) ? tools : Object.values(tools || {});

            const tool = list.slice().reverse().find(t => {
                const name = (typeof t.toolname === 'function' ? t.toolname() : t.toolname) || t._toolName || '';
                return typeof name === 'string' && name.includes('LineToolRiskReward');
            });

            if (!tool) {
                alert('Нарисуйте инструмент Long/Short на графике!');
                return null;
            }

            let props = typeof tool.properties === 'function' ? (typeof tool.properties().state === 'function' ? tool.properties().state() : tool.properties()) : tool.properties;
            let points = typeof tool.points === 'function' ? tool.points() : (tool._points || tool.points);

            const toolNameStr = (typeof tool.toolname === 'function' ? tool.toolname() : tool.toolname) || tool._toolName || '';
            const isLong = toolNameStr.includes('Long');

            const entry = points[0]?.price || 0;
            const time = points[0]?.time || Math.floor(Date.now() / 1000);

            // Идеальное вычисление шага цены (Price Step) из ядра TradingView
            let symInfo = {};
            try { symInfo = (typeof ms.symbolInfo === 'function' ? ms.symbolInfo() : ms._symbolInfo) || {}; } catch(e){}
            const priceStep = (typeof ms.priceStep === 'function' ? ms.priceStep() : ms.priceStep) ||
                              (symInfo.minmov && symInfo.pricescale ? symInfo.minmov / symInfo.pricescale : 0.001);

            // Идеальный расчет уровней
            const stop = isLong ? entry - (props.stopLevel * priceStep) : entry + (props.stopLevel * priceStep);
            const target = isLong ? entry + (props.profitLevel * priceStep) : entry - (props.profitLevel * priceStep);
            const rr = (props.stopLevel && props.stopLevel > 0) ? (props.profitLevel / props.stopLevel).toFixed(2) : "0.00";

            const calculatedOutcome = calculateOutcome(bars, points[0]?.index ?? 0, points[1]?.index ?? (points[0]?.index ?? 0), isLong, stop, target);

            // Динамическое определение знаков после запятой (для золота 3, для форекса 5 и т.д.)
            const decimals = priceStep < 1 ? priceStep.toString().split('.')[1]?.length || 3 : 2;

            return {
                symbol: props.symbol || 'XAUUSD',
                timeframe: 'M' + (props.interval || '1'),
                direction: isLong ? 'LONG' : 'SHORT',
                entryPrice: Number(entry.toFixed(decimals)),
                stopPrice: Number(stop.toFixed(decimals)),
                targetPrice: Number(target.toFixed(decimals)),
                riskReward: Number(rr),
                candleDate: formatDate(time),
                session: detectSession(time),
                autoResult: calculatedOutcome
            };
        } catch (e) {
            console.error('Ошибка парсинга TV:', e);
            return null;
        }
    }

    // 4. UI Модалки
    const modalHtml = `
    <div id="bt-modal" style="display:none; position:fixed; top:20px; right:20px; z-index:9999999; background:#0a0a0a; color:#e0e0e0; padding:20px; border-radius:8px; box-shadow:0 12px 32px rgba(0,0,0,0.8); border:1px solid #332d1f; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif; width:360px;">

        <div style="text-align:center; margin-bottom:16px;">
            <div style="font-size:10px; letter-spacing:2px; color:#d6b658; text-transform:uppercase; margin-bottom:4px;">Стратегия «Матрёшка»</div>
            <div style="font-family:'Georgia', 'Playfair Display', serif; font-size:22px; font-weight:bold; color:#d6b658; letter-spacing:1px;">ЭКСПОРТ В EXCEL</div>
        </div>

        <div style="background:#141414; border:1px solid #2a2a2a; border-radius:6px; padding:10px; margin-bottom:16px; font-size:12px; line-height:1.6;">
            <div style="display:flex; justify-content:space-between; margin-bottom:4px;">
                <span id="bt-sym-tf" style="color:#d6b658; font-weight:bold;"></span>
                <span id="bt-badge" style="background:#089981; color:#fff; font-size:10px; padding:2px 6px; border-radius:4px; font-weight:bold;"></span>
            </div>
            <div style="color:#a0a0a0; display:flex; justify-content:space-between;">
                <span>Вход: <b id="bt-en" style="color:#e0e0e0;"></b></span>
                <span>Стоп: <b id="bt-st" style="color:#e0e0e0;"></b></span>
                <span>Цель: <b id="bt-tg" style="color:#d6b658;"></b></span>
            </div>
            <div style="color:#787b86; margin-top:4px; font-size:11px;" id="bt-time"></div>
        </div>

        <div class="bt-section-title">ФАЗА РЫНКА</div>
        <div class="bt-group" id="bt-trend" style="display:grid; grid-template-columns:1fr 1fr; gap:4px; margin-bottom:12px;">
            <button class="bt-chip active" data-val="Тренд вверх">Тренд вверх</button>
            <button class="bt-chip" data-val="Тренд вниз">Тренд вниз</button>
            <button class="bt-chip" data-val="Боковик">Боковик</button>
            <button class="bt-chip" data-val="Импульс/новости">Импульс/новости</button>
        </div>

        <div class="bt-section-title">ЗОНА ВХОДА</div>
        <div class="bt-group" id="bt-zone" style="display:grid; grid-template-columns:1fr 1fr; gap:4px; margin-bottom:12px;">
            <button class="bt-chip active" data-val="Ордер-блок">Ордер-блок</button>
            <button class="bt-chip" data-val="Имбаланс">Имбаланс</button>
            <button class="bt-chip" data-val="Материнская свеча">Мать-свеча</button>
            <button class="bt-chip" data-val="Зона ликвидности">Ликвидность</button>
            <button class="bt-chip" data-val="Правильный откат" style="grid-column: span 2;">Правильный откат</button>
        </div>

        <div style="display:grid; grid-template-columns:1fr 2fr; gap:10px; margin-bottom:12px;">
            <div>
                <div class="bt-section-title">ЛИКВИДНОСТЬ</div>
                <div class="bt-group" id="bt-liq" style="display:flex; gap:4px;">
                    <button class="bt-chip active" data-val="да">Да</button>
                    <button class="bt-chip" data-val="нет">Нет</button>
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
        <div class="bt-group" id="bt-result" style="display:grid; grid-template-columns:repeat(4, 1fr); gap:4px; margin-bottom:12px;">
            <button class="bt-chip" data-val="Тейк" style="--ac:#089981;">Тейк</button>
            <button class="bt-chip" data-val="Стоп" style="--ac:#f23645;">Стоп</button>
            <button class="bt-chip" data-val="Безубыток">БУ</button>
            <button class="bt-chip" data-val="Пропуск">Пропуск</button>
        </div>

        <div class="bt-section-title">ОШИБКА ИСПОЛНЕНИЯ</div>
        <select id="bt-err" style="width:100%; background:#1a1a1a; color:#e0e0e0; border:1px solid #332d1f; padding:6px; border-radius:4px; font-size:11px; outline:none; cursor:pointer; margin-bottom:12px;">
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

        <input type="text" id="bt-img" placeholder="Ссылка на скриншот (Alt+S -> Ctrl+V)" style="width:100%; box-sizing:border-box; background:#141414; border:1px solid #332d1f; color:#e0e0e0; padding:8px; border-radius:4px; font-size:11px; margin-bottom:6px; outline:none;">
        <input type="text" id="bt-comment" placeholder="Комментарий к сделке..." style="width:100%; box-sizing:border-box; background:#141414; border:1px solid #332d1f; color:#e0e0e0; padding:8px; border-radius:4px; font-size:11px; margin-bottom:16px; outline:none;">

        <div style="display:flex; gap:8px;">
            <button id="bt-save-btn" style="flex:1; background:#d6b658; color:#0a0a0a; border:none; padding:10px; border-radius:4px; cursor:pointer; font-weight:bold; font-size:13px; transition:0.2s;">Копировать для Excel (Enter)</button>
            <button id="bt-close-btn" style="background:#1a1a1a; color:#a0a0a0; border:1px solid #332d1f; padding:10px 14px; border-radius:4px; cursor:pointer; transition:0.2s;">✕</button>
        </div>
    </div>

    <style>
        .bt-section-title { font-size: 10px; color: #787b86; margin-bottom: 6px; font-weight: bold; letter-spacing: 0.5px; }
        .bt-chip { background: #1a1a1a; color: #a0a0a0; border: 1px solid #2a2a2a; padding: 6px; font-size: 11px; border-radius: 4px; cursor: pointer; text-align: center; transition: 0.15s ease; }
        .bt-chip:hover { border-color: #d6b658; color: #e0e0e0; }
        .bt-chip.active { background: var(--ac, #d6b658); color: #0a0a0a; border-color: var(--ac, #d6b658); font-weight: bold; }
        #bt-save-btn:hover { background: #e8c767; box-shadow: 0 0 10px rgba(214, 182, 88, 0.3); }
        #bt-img:focus, #bt-comment:focus, #bt-err:focus { border-color: #d6b658; }
    </style>
    `;

    document.body.insertAdjacentHTML('beforeend', modalHtml);

    // Логика переключения
    document.querySelectorAll('.bt-group').forEach(group => {
        group.addEventListener('click', e => {
            if (e.target.classList.contains('bt-chip')) {
                group.querySelectorAll('.bt-chip').forEach(b => b.classList.remove('active'));
                e.target.classList.add('active');
            }
        });
    });

    let currentTrade = null;

    function openModal() {
        currentTrade = extractTrade();
        if (!currentTrade) return;

        const badge = document.getElementById('bt-badge');
        badge.innerText = `${currentTrade.direction} (R:R ${currentTrade.riskReward})`;
        badge.style.background = currentTrade.direction === 'LONG' ? '#089981' : '#f23645';

        document.getElementById('bt-sym-tf').innerText = `${currentTrade.symbol} | ${currentTrade.timeframe} | ${currentTrade.session}`;
        document.getElementById('bt-en').innerText = currentTrade.entryPrice;
        document.getElementById('bt-st').innerText = currentTrade.stopPrice;
        document.getElementById('bt-tg').innerText = currentTrade.targetPrice;
        document.getElementById('bt-time').innerText = currentTrade.candleDate;

        const resultGroup = document.getElementById('bt-result');
        resultGroup.querySelectorAll('.bt-chip').forEach(b => {
            b.classList.remove('active');
            if (b.dataset.val === currentTrade.autoResult) b.classList.add('active');
        });

        document.getElementById('bt-err').value = "—";

        document.getElementById('bt-modal').style.display = 'block';
        setTimeout(() => document.getElementById('bt-img').focus(), 50);
    }

    function saveTrade() {
        const trend = document.querySelector('#bt-trend .active')?.dataset.val || "";
        const entryZone = document.querySelector('#bt-zone .active')?.dataset.val || "";
        const liquidityTaken = document.querySelector('#bt-liq .active')?.dataset.val || "";
        const confirmation = document.getElementById('bt-conf').value;
        const result = document.querySelector('#bt-result .active')?.dataset.val || "";
        const execError = document.getElementById('bt-err').value;

        let commentStr = document.getElementById('bt-comment').value;
        const imgUrl = document.getElementById('bt-img').value;
        if (imgUrl) commentStr += (commentStr ? " " : "") + imgUrl;

        // Строка для Excel (Табуляция \t пропускает ячейки с формулами)
        const tsvArray = [
            currentTrade.candleDate,
            currentTrade.session,
            trend,
            currentTrade.timeframe,
            entryZone,
            liquidityTaken,
            confirmation,
            currentTrade.entryPrice,
            currentTrade.stopPrice,
            currentTrade.targetPrice,
            "", // Пропуск: Риск
            "", // Пропуск: Потенциал
            "", // Пропуск: R:R
            result,
            "", // Пропуск: Итог
            "", // Пропуск: Накопительно
            "", // Пропуск: Просадка
            execError,
            commentStr
        ];

        const tsvString = tsvArray.join('\t');

        navigator.clipboard.writeText(tsvString).then(() => {
            document.getElementById('bt-modal').style.display = 'none';
            document.getElementById('bt-img').value = '';
            document.getElementById('bt-comment').value = '';
        }).catch(err => {
            console.error('Ошибка буфера:', err);
            alert('Не удалось скопировать. Проверьте консоль.');
        });
    }

    // Хоткеи
    window.addEventListener('keydown', (e) => {
        if (e.key === 'F2' && e.target.tagName !== 'INPUT' && e.target.tagName !== 'TEXTAREA') {
            openModal();
        } else if (e.key === 'Enter' && document.getElementById('bt-modal').style.display === 'block') {
            saveTrade();
        } else if (e.key === 'Escape') {
            document.getElementById('bt-modal').style.display = 'none';
        }
    });

    document.getElementById('bt-save-btn').onclick = saveTrade;
    document.getElementById('bt-close-btn').onclick = () => document.getElementById('bt-modal').style.display = 'none';

})();
