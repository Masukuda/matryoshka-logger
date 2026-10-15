// ==UserScript==
// @name         Matryoshka Backtest Auto-Logger
// @namespace    http://tampermonkey.net/
// @version      2.2.1
// @description  Экспорт сделок из TradingView в Excel и Google Таблицы
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
    const STORAGE_KEY = 'matryoshka_logger_v2_rowfree';
    const isMac = /Mac|iPod|iPhone|iPad/i.test(navigator.platform || navigator.userAgent);

    const LOGO_SVG = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" style="flex-shrink:0;"><circle cx="12" cy="7" r="4.5" stroke="#e5c158" stroke-width="2"/><path d="M5.5 19.5C5.5 15.634 8.41015 12.5 12 12.5C15.5899 12.5 18.5 15.634 18.5 19.5C18.5 20.8807 17.3807 22 16 22H8C6.61929 22 5.5 20.8807 5.5 19.5Z" stroke="#e5c158" stroke-width="2"/><circle cx="12" cy="17.5" r="1.5" fill="#e5c158"/></svg>`;
    const BOLT = `<span class="bt-bolt" title="Определено автоматически по данным графика">⚡</span>`;

    const PRESET_SCHEMAS = {
        bt_main: {
            name: 'Бэктест «Матрёшка»',
            pasteCol: 'B',
            startRow: 6,
            sessionStyle: 'std',
            tfStyle: 'lat',
            columns: [
                'date', 'session', 'trend', 'tf', 'zone', 'liq', 'conf',
                'entry', 'stop', 'target', 'risk_pts', 'pot_pts', 'rr',
                'result', 'r_out', 'cum_r', 'dd_r', 'error', 'comment', 'screenshot'
            ]
        },
        bt_30: {
            name: 'Бэктест 30 сделок',
            pasteCol: 'B',
            startRow: 4,
            sessionStyle: 'eu_us',
            tfStyle: 'cyr',
            columns: [
                'date', 'time_msk', 'session', 'tf', 'dir_lower',
                'flow_d', 'flow_4h', 'flow_1h', 'flow_15m',
                'struct_30', 'zone_30', 'fuel_30', 'touch_30', 'against_flow',
                'swept_30', 'swept_level_30', 'candle_col',
                'entry', 'stop', 'target', 'risk_pts', 'rr', 'lots_30',
                'check_30', 'res_30', 'manual_r_30', 'r_out', 'profit_usd',
                'hit_1r', 'hit_3r', 'screenshot', 'comment_30'
            ]
        },
        dz_poi: {
            name: 'ДЗ: Журнал POI',
            pasteCol: 'B',
            startRow: 6,
            sessionStyle: 'std',
            tfStyle: 'lat',
            columns: [
                'aoi_num', 'date', 'poi_type', 'liq', 'react', 'rr',
                'result', 'r_out', 'cum_r', 'error_poi', 'comment', 'screenshot'
            ]
        },
        dz_aoi: {
            name: 'ДЗ: Таблица AOI',
            pasteCol: 'B',
            startRow: 6,
            sessionStyle: 'std',
            tfStyle: 'lat',
            columns: [
                'aoi_num', 'date', 'trend_h4', 'f_ret', 'f_ob', 'f_fvg', 'f_liq',
                'factors_cnt', 'valid_aoi', 'upper_bound', 'lower_bound',
                'reached', 'aoi_res', 'comment', 'screenshot'
            ]
        },
        office: {
            name: 'Торговый офис (Недели 1–4)',
            pasteCol: 'B',
            startRow: 16,
            sessionStyle: 'std',
            tfStyle: 'lat',
            columns: [
                'time_slot', 'dir_title', 'tf_ok', 'tf_ok', 'tf_ok', 'tf_ok', 'form_m1',
                'lots_off', 'entry', 'stop', 'target', 'safe_pct', 'safe_rr', 'risk_pct_col', 'rr',
                'res_off', 'r_out_off', 'prof_pct_off', 'move_pts', 'tf_ok', 'emotion', 'screenshot', 'comment'
            ]
        },
        prices_only: {
            name: 'Только 3 цены (Вход/Стоп/Цель)',
            pasteCol: 'B',
            startRow: 2,
            sessionStyle: 'std',
            tfStyle: 'lat',
            columns: ['entry', 'stop', 'target']
        }
    };

    const TOKEN_LABELS = {
        symbol: 'Инструмент (Тикер)',
        date: 'Дата входа',
        time_msk: 'Время входа (МСК)',
        time_slot: 'Временной слот (30м)',
        session: 'Торговая сессия',
        tf: 'Таймфрейм',
        dir_title: 'Направление (Лонг/Шорт)',
        dir_lower: 'Сторона (лонг/шорт)',
        trend: 'Фаза рынка',
        trend_h4: 'Тренд H4',
        flow_d: 'Поток D1',
        flow_4h: 'Поток H4',
        flow_1h: 'Поток H1',
        flow_15m: 'Поток M15',
        struct_30: 'Структура перед входом',
        zone: 'Зона интереса (основная)',
        zone_30: 'Вид зоны (детальный)',
        poi_type: 'Тип POI (OB / Имб / МС)',
        form_m1: 'Формация входа M1',
        candle_col: 'Импульс свечи (индикатор)',
        liq: 'Снятие ликвидности (да/нет)',
        swept_30: 'Что снято перед входом',
        swept_level_30: 'Уровень снятия (цена)',
        fuel_30: 'Топливо (да/нет)',
        touch_30: 'Касание зоны (1/2/3+)',
        against_flow: 'Против потока (да/нет)',
        conf: 'Подтверждение входа',
        react: 'Реакция от зоны (да/нет)',
        entry: 'Цена входа',
        stop: 'Цена стопа',
        target: 'Цена цели (тейк)',
        upper_bound: 'Верхняя граница зоны',
        lower_bound: 'Нижняя граница зоны',
        risk_pts: 'Риск в пунктах / $',
        pot_pts: 'Потенциал в пунктах',
        rr: 'Соотношение R:R',
        lots_30: 'Объём (лот, фьючерс/форекс)',
        lots_off: 'Объём (лот от % риска)',
        check_30: 'Соблюдение чек-листа (да/нет)',
        result: 'Результат (Тейк/Стоп/БУ)',
        res_30: 'Исход (с ручным выходом)',
        res_off: 'Результат (с сейфом)',
        manual_r_30: 'Фактический R (вручную)',
        r_out: 'Итог сделки, R',
        r_out_off: 'Итог R (с сейфом)',
        profit_usd: 'Финансовый итог, $',
        prof_pct_off: 'Прибыль / Убыток, %',
        move_pts: 'Взятое движение, п.',
        cum_r: 'Накопительный итог, R',
        dd_r: 'Просадка, R',
        hit_1r: 'Достижение 1:1 (да/нет)',
        hit_3r: 'Достижение 1:3 (да/нет)',
        error: 'Причина ошибки',
        error_poi: 'Ошибка выбора POI',
        emotion: 'Эмоциональное состояние',
        comment: 'Комментарий',
        comment_30: 'Комментарий (30 сд.)',
        screenshot: 'Скриншот (ссылка)',
        aoi_num: '№ зоны / области',
        tf_ok: 'ТФ по системе (да)',
        safe_pct: 'Доля сейфа (0.5)',
        safe_rr: 'Уровень сейфа RR (3)',
        risk_pct_col: 'Риск на сделку % (0.003)',
        f_ret: 'Фактор: Откат 50–62%',
        f_ob: 'Фактор: OB H4',
        f_fvg: 'Фактор: Имбаланс H4',
        f_liq: 'Фактор: Ликвидность рядом',
        factors_cnt: 'Количество факторов (формула)',
        valid_aoi: 'Валидность зоны (формула)',
        reached: 'Цена дошла до зоны',
        aoi_res: 'Реакция области',
        empty: 'Пустая ячейка (пропуск)'
    };

    const defaultState = {
        presetKey: 'bt_main',
        customColumns: null,
        customPasteCol: 'B',
        customStartRow: 2,
        outputProfile: 'ru_form',
        decimalSep: ',',
        trend: 'Тренд вверх',
        zone: 'Ордер-блок',
        liq: 'да',
        conf: 'Слом структуры',
        aoiNum: 1,
        flowD: 'вверх',
        flow4H: 'вверх',
        flow1H: 'вверх',
        flow15M: 'вверх',
        perTpl: {},
        minimized: false,
        posTop: 20,
        posRight: 20
    };

    function loadState() {
        try {
            const raw = localStorage.getItem(STORAGE_KEY);
            const parsed = raw ? { ...defaultState, ...JSON.parse(raw) } : { ...defaultState };
            if (parsed.outputProfile !== 'ru_form' && parsed.outputProfile !== 'en_form') {
                parsed.outputProfile = 'ru_form';
            }
            return parsed;
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
    if (!state.perTpl || typeof state.perTpl !== 'object') state.perTpl = {};

    function getTplBucket(key = state.presetKey) {
        if (!state.perTpl[key]) {
            state.perTpl[key] = { batchQueue: [], cumR: 0, maxCumR: 0, lastSavedSig: '' };
        }
        return state.perTpl[key];
    }

    function saveTplBucket(patch, key = state.presetKey) {
        const b = getTplBucket(key);
        state.perTpl[key] = { ...b, ...patch };
        saveState({ perTpl: state.perTpl });
    }

    let allChartTrades = [];
    let activeTradeIdx = 0;
    let currentTrade = null;
    let manualSessionOverride = null;
    let screenshotUrl = '';
    let lastSelectedToolId = null;
    let lastChartSelectedId = null;
    let lastHoveredToolId = null;
    let isSettingsOpen = false;
    let showColumnEditor = false;

    function getPlotRow(plotsObj, idx) {
        if (!plotsObj || idx == null) return null;
        try {
            if (typeof plotsObj.valueAt === 'function') {
                const v = plotsObj.valueAt(idx);
                if (Array.isArray(v)) return v;
                if (v && Array.isArray(v.value)) return v.value;
            }
            if (typeof plotsObj.search === 'function') {
                const s = plotsObj.search(idx);
                if (s && Array.isArray(s.value)) return s.value;
            }
        } catch (e) {}
        return null;
    }

    function findPlotRowByIndexOrTime(plotsObj, targetTime, hintIdx) {
        if (!plotsObj) return null;
        try {
            if (hintIdx != null) {
                const r0 = getPlotRow(plotsObj, hintIdx);
                if (r0 && (!targetTime || r0[0] === targetTime)) return { index: hintIdx, val: r0 };
                for (let d = 1; d <= 6; d++) {
                    const rA = getPlotRow(plotsObj, hintIdx - d);
                    if (rA && rA[0] === targetTime) return { index: hintIdx - d, val: rA };
                    const rB = getPlotRow(plotsObj, hintIdx + d);
                    if (rB && rB[0] === targetTime) return { index: hintIdx + d, val: rB };
                }
                if (r0) return { index: hintIdx, val: r0 };
            }

            if (!targetTime) return null;
            const first = typeof plotsObj.firstIndex === 'function' ? plotsObj.firstIndex() : 0;
            const last = typeof plotsObj.lastIndex === 'function' ? plotsObj.lastIndex() : 1000;
            let lo = first ?? 0, hi = last ?? 1000;
            while (lo <= hi) {
                const mid = (lo + hi) >> 1;
                const row = getPlotRow(plotsObj, mid);
                if (!row) break;
                if (row[0] === targetTime) return { index: mid, val: row };
                if (row[0] < targetTime) lo = mid + 1;
                else hi = mid - 1;
            }
        } catch (e) {}
        return null;
    }

    function readMatryoshkaIndicators(model, entryTime, triggerIdx, isLong) {
        const result = { flows: null, candleColor: null, detectedStructure: null };
        if (!model || typeof model.dataSources !== 'function') return result;

        try {
            const sources = model.dataSources();
            for (const s of sources) {
                const meta = typeof s.metaInfo === 'function' ? s.metaInfo() : s._metaInfo;
                const title = meta?.description || meta?.name || '';

                if (title.includes('Матрёшка Зоны интереса')) {
                    const g = typeof s.graphics === 'function' ? s.graphics() : s._graphics;
                    const tcContainer = g?.dwgtablecells?.()?.get('tableCells');
                    const cellMap = tcContainer?._primitivesDataById || tcContainer;
                    if (cellMap instanceof Map) {
                        const rowsByTid = {};
                        for (const cell of cellMap.values()) {
                            if (!cell || cell.tid == null) continue;
                            const key = `${cell.tid}_${cell.row}`;
                            if (!rowsByTid[key]) rowsByTid[key] = {};
                            rowsByTid[key][cell.col] = String(cell.t || '').trim();
                        }
                        const flows = {};
                        for (const rObj of Object.values(rowsByTid)) {
                            const label = rObj[0];
                            const sym = rObj[1];
                            if (!label || !sym) continue;
                            const dirWord = sym === '↑' ? 'вверх' : (sym === '↓' ? 'вниз' : '—');
                            if (label === 'Д') flows.d = dirWord;
                            else if (label === '4Ч') flows.h4 = dirWord;
                            else if (label === '1Ч') flows.h1 = dirWord;
                            else if (label === '15м') flows.m15 = dirWord;
                        }
                        if (Object.keys(flows).length > 0) result.flows = flows;
                    }

                    const plots = typeof s.plots === 'function' ? s.plots() : s._data?.bars?.();
                    const foundBar = findPlotRowByIndexOrTime(plots, entryTime, triggerIdx);
                    if (foundBar && plots) {
                        let countCont = 0, hasChange = false, hasFakeout = false, hasInnerChange = false;
                        for (let idx = Math.max(0, foundBar.index - 40); idx <= foundBar.index; idx++) {
                            const v = getPlotRow(plots, idx);
                            if (!v) continue;
                            if (v[5] === 1) countCont++;
                            if (v[6] === 1) hasChange = true;
                            if (v[7] === 1) hasFakeout = true;
                            if (v[9] === 1) hasInnerChange = true;
                        }
                        if (hasFakeout) {
                            result.detectedStructure = isLong ? 'ложный пробой ↓ · слом ↑' : 'ложный пробой ↑ · слом ↓';
                        } else if (hasChange && countCont >= 2) {
                            result.detectedStructure = isLong ? 'пробой ↓ · пробой ↓ · слом ↑' : 'пробой ↑ · пробой ↑ · слом ↓';
                        } else if (hasChange) {
                            result.detectedStructure = isLong ? 'слом ↑' : 'слом ↓';
                        } else if (hasInnerChange) {
                            result.detectedStructure = isLong ? 'внутр. слом ↑' : 'внутр. слом ↓';
                        } else if (countCont >= 2) {
                            result.detectedStructure = isLong ? 'пробой ↑ · пробой ↑' : 'пробой ↓ · пробой ↓';
                        } else if (countCont === 1) {
                            result.detectedStructure = isLong ? 'пробой ↑' : 'пробой ↓';
                        }
                    }
                }

                if (title.includes('Матрёшка Модели входа')) {
                    const plots = typeof s.plots === 'function' ? s.plots() : s._data?.bars?.();
                    const foundBar = findPlotRowByIndexOrTime(plots, entryTime, triggerIdx);
                    if (foundBar && plots) {
                        const candidates = [foundBar.index, foundBar.index - 1, foundBar.index + 1];
                        let matchedCode = 0;
                        for (const cIdx of candidates) {
                            const r = getPlotRow(plots, cIdx);
                            const code = Number(r?.[3] || 0);
                            if ((isLong && code > 0) || (!isLong && code < 0)) {
                                matchedCode = code;
                                break;
                            }
                        }
                        const colorMap = {
                            '3': 'Лонг 3', '2': 'Лонг 2', '1': 'Лонг 1',
                            '-1': 'Шорт 1', '-2': 'Шорт 2', '-3': 'Шорт 3'
                        };
                        if (colorMap[String(matchedCode)]) {
                            result.candleColor = colorMap[String(matchedCode)];
                        }
                    }
                }
            }
        } catch (e) {}

        return result;
    }

    function getToolId(t) {
        if (!t) return null;
        try {
            if (typeof t.id === 'function') return t.id();
            if (t._id && typeof t._id.value === 'function') return t._id.value();
            if (typeof t._id === 'string') return t._id;
        } catch (e) {}
        return null;
    }

    function getToolName(t) {
        if (!t) return '';
        try {
            const name = (typeof t.toolname === 'function' ? t.toolname() : t.toolname) || t._toolName || '';
            return typeof name === 'string' ? name : '';
        } catch (e) {
            return '';
        }
    }

    function isSupportedTool(t) {
        const n = getToolName(t);
        return n.includes('LineToolRiskReward') || (state.presetKey === 'dz_aoi' && n.includes('LineToolRectangle'));
    }

    function getTvModel() {
        try {
            const c = pageWindow._exposed_chartWidgetCollection;
            if (!c) return null;
            const widget = typeof c.activeChartWidget === 'function'
                ? c.activeChartWidget()
                : (c.activeChartWidget?.value ? c.activeChartWidget.value() : c.activeChartWidget);
            if (!widget) return null;
            const rawModel = typeof widget.model === 'function' ? widget.model() : (widget.model || widget._model);
            return rawModel?.m_model || rawModel || null;
        } catch (e) {
            return null;
        }
    }

    function readWatched(val) {
        if (val != null && typeof val === 'object' && typeof val.value === 'function') return val.value();
        return val;
    }

    function detectCurrentlySelectedToolId(model) {
        if (!model) return null;
        try {
            const sel = typeof model.selection === 'function' ? model.selection() : model._selection;
            if (sel) {
                if (typeof sel.lineDataSources === 'function') {
                    const lds = sel.lineDataSources();
                    if (Array.isArray(lds)) {
                        const found = lds.find(isSupportedTool);
                        if (found) return getToolId(found);
                    }
                }
                if (typeof sel.isSelected === 'function') {
                    const tools = typeof model.allLineTools === 'function' ? model.allLineTools() : [];
                    const list = Array.isArray(tools) ? tools : Object.values(tools || {});
                    const selTool = list.find(t => isSupportedTool(t) && sel.isSelected(t));
                    if (selTool) return getToolId(selTool);
                }
            }
        } catch (e) {}
        return null;
    }

    function getMskHourDecimal(dateObj) {
        const utcH = (dateObj.getUTCHours() + 3) % 24;
        return utcH + (dateObj.getUTCMinutes() / 60);
    }

    function detectSession(timestamp) {
        if (!timestamp) return { std: 'Лондон', eu_us: 'Европа' };
        const d = new Date(timestamp * 1000);
        const mskH = getMskHourDecimal(d);

        let std = 'Азия';
        if (mskH >= 15.0 && mskH < 23.0) {
            std = 'Нью-Йорк';
        } else if (mskH >= 8.0 && mskH < 15.0) {
            std = 'Лондон';
        }

        let eu_us = 'Азия';
        if (mskH >= 15.5 && mskH < 23.0) {
            eu_us = 'Америка';
        } else if (mskH >= 8.0 && mskH < 15.5) {
            eu_us = 'Европа';
        }

        return { std, eu_us };
    }

    function formatMskTime(timestamp) {
        const d = new Date(timestamp * 1000);
        const h = String((d.getUTCHours() + 3) % 24).padStart(2, '0');
        const m = String(d.getUTCMinutes()).padStart(2, '0');
        return `${h}:${m}`;
    }

    function formatHalfHourSlot(timestamp) {
        const [h, m] = formatMskTime(timestamp).split(':').map(Number);
        const slotM = m >= 30 ? '30' : '00';
        const slotH = String(Math.max(7, Math.min(23, h || 10))).padStart(2, '0');
        return `${slotH}:${slotM}`;
    }

    function formatDate(timestamp) {
        const d = new Date(timestamp * 1000);
        try {
            return new Intl.DateTimeFormat('ru-RU', {
                timeZone: 'Europe/Moscow', day: '2-digit', month: '2-digit', year: 'numeric'
            }).format(d);
        } catch (e) {
            return `${String(d.getDate()).padStart(2, '0')}.${String(d.getMonth() + 1).padStart(2, '0')}.${d.getFullYear()}`;
        }
    }

    function formatShortDate(timestamp) {
        return formatDate(timestamp).slice(0, 5);
    }

    function formatTf(rawInterval, style) {
        const str = String(rawInterval || '1').toUpperCase().trim();
        if (style === 'cyr') {
            if (str === '1') return 'М1';
            if (str === '5') return '5м';
            if (str === '15') return '15м';
            return 'М1';
        }
        if (/^\d+$/.test(str)) {
            const mins = parseInt(str, 10);
            return (mins >= 60 && mins % 60 === 0) ? 'H' + (mins / 60) : 'M' + mins;
        }
        if (str === 'D' || str === '1D') return 'D1';
        if (str === 'W' || str === '1W') return 'W1';
        return str;
    }

    function normalizeUnixSeconds(val) {
        const n = Number(val || 0);
        if (!n || isNaN(n)) return 0;
        const sec = n > 1e11 ? Math.floor(n / 1000) : Math.floor(n);
        return sec > 946684800 ? sec : 0;
    }

    function resolveBarTimestamp(model, bars, barIdx, tool, points) {
        const pt0 = tool?._points?.[0] || points?.[0];
        const rawPtTime = normalizeUnixSeconds(pt0?.time_t || pt0?.time?.getTime?.() || pt0?.time);

        const row = getPlotRow(bars, barIdx);
        if (row) {
            const rt = normalizeUnixSeconds(row[0]);
            if (rt) return rt;
        }

        try {
            const ts = typeof model.timeScale === 'function' ? model.timeScale() : model._timeScale;
            if (ts && barIdx != null && typeof ts.indexToTimePoint === 'function') {
                const tp = normalizeUnixSeconds(ts.indexToTimePoint(barIdx));
                if (tp) return tp;
            }
        } catch (e) {}

        if (rawPtTime) return rawPtTime;
        return Math.floor(Date.now() / 1000);
    }

    function analyzeBars(bars, startIndex, endIndex, isLong, entryPrice, stopPrice, targetPrice) {
        const risk = Math.abs(entryPrice - stopPrice);
        const price1R = isLong ? entryPrice + risk : entryPrice - risk;
        const price3R = isLong ? entryPrice + (risk * 3) : entryPrice - (risk * 3);

        if (!bars || startIndex == null) {
            return { outcome: 'Тейк', reached1R: true, reached3R: true, entryCandleExtreme: null, triggerIdx: startIndex ?? 0 };
        }

        const anchorIdx = startIndex;
        const to = Math.max(anchorIdx, endIndex ?? anchorIdx);

        let triggerIdx = anchorIdx;
        for (let i = anchorIdx; i <= to; i++) {
            const b = getPlotRow(bars, i);
            if (!b) continue;
            const hi = Number(b[2]), lo = Number(b[3]);
            if (!isNaN(hi) && !isNaN(lo) && lo <= entryPrice && hi >= entryPrice) {
                triggerIdx = i;
                break;
            }
        }

        let sweepExtreme = null;
        for (let i = Math.max(0, triggerIdx - 2); i <= triggerIdx; i++) {
            const b = getPlotRow(bars, i);
            if (!b) continue;
            const hi = Number(b[2]), lo = Number(b[3]);
            if (isNaN(hi) || isNaN(lo)) continue;
            sweepExtreme = sweepExtreme === null
                ? (isLong ? lo : hi)
                : (isLong ? Math.min(sweepExtreme, lo) : Math.max(sweepExtreme, hi));
        }

        let reached1R = false;
        let reached3R = false;
        let outcome = 'Безубыток';
        let outcomeDecided = false;

        for (let i = triggerIdx; i <= to; i++) {
            const bar = getPlotRow(bars, i);
            if (!bar) continue;
            const high = Number(bar[2]);
            const low = Number(bar[3]);
            const close = Number(bar[4]);
            if (isNaN(high) || isNaN(low)) continue;

            const checkHigh = (i === triggerIdx) ? Math.max(entryPrice, close) : high;
            const checkLow = (i === triggerIdx) ? Math.min(entryPrice, close) : low;

            if (isLong ? (checkHigh >= price1R) : (checkLow <= price1R)) reached1R = true;
            if (isLong ? (checkHigh >= price3R) : (checkLow <= price3R)) reached3R = true;

            if (!outcomeDecided) {
                const hitTarget = isLong ? (checkHigh >= targetPrice) : (checkLow <= targetPrice);
                const hitStop = isLong ? (checkLow <= stopPrice) : (checkHigh >= stopPrice);

                if (hitTarget && hitStop) {
                    outcome = ((isLong && close >= targetPrice) || (!isLong && close <= targetPrice)) ? 'Тейк' : 'Стоп';
                    outcomeDecided = true;
                } else if (hitStop) {
                    outcome = 'Стоп';
                    outcomeDecided = true;
                } else if (hitTarget) {
                    outcome = 'Тейк';
                    outcomeDecided = true;
                }
            }
        }

        return { outcome, reached1R, reached3R, entryCandleExtreme: sweepExtreme, triggerIdx };
    }

    function parseSingleTool(tool, model, ms, bars, symInfo, priceStep) {
        const rawProps = typeof tool.properties === 'function' ? tool.properties() : tool.properties;
        const props = (rawProps && typeof rawProps.state === 'function') ? rawProps.state() : rawProps;
        const points = typeof tool.points === 'function' ? tool.points() : (tool._points || tool.points || []);

        const toolNameStr = getToolName(tool);
        const isRect = toolNameStr.includes('Rectangle');
        const isLong = !toolNameStr.includes('Short');

        const p0 = Number(points[0]?.price ?? tool._points?.[0]?.price ?? 0);
        const p1 = Number(points[1]?.price ?? tool._points?.[1]?.price ?? p0);
        const startIdx = points[0]?.index ?? tool._points?.[0]?.index ?? null;
        const endIdx = points[1]?.index ?? tool._points?.[1]?.index ?? startIdx;

        let entry = p0, stop = p0, target = p1, rrNum = 3.0;
        let barAnalysis = { outcome: 'Тейк', reached1R: true, reached3R: true, entryCandleExtreme: null, triggerIdx: startIdx };

        if (isRect) {
            entry = Math.max(p0, p1);
            stop = Math.min(p0, p1);
            target = entry;
        } else {
            const stopLevel = Number(readWatched(props?.stopLevel) ?? readWatched(rawProps?.stopLevel) ?? 0);
            const profitLevel = Number(readWatched(props?.profitLevel) ?? readWatched(rawProps?.profitLevel) ?? 0);
            stop = isLong ? entry - (stopLevel * priceStep) : entry + (stopLevel * priceStep);
            target = isLong ? entry + (profitLevel * priceStep) : entry - (profitLevel * priceStep);
            rrNum = (stopLevel > 0) ? (profitLevel / stopLevel) : 0;
            barAnalysis = analyzeBars(bars, startIdx, endIdx, isLong, entry, stop, target);
        }

        const effectiveIdx = barAnalysis.triggerIdx ?? startIdx;
        const time = resolveBarTimestamp(model, bars, effectiveIdx ?? startIdx, tool, points);

        const rawSymbol = readWatched(props?.symbol) || symInfo.name || 'XAUUSD';
        const cleanSymbol = String(rawSymbol).split(':').pop();
        const rawInterval = readWatched(props?.interval) || '15';

        const decimals = symInfo.pricescale
            ? Math.max(0, Math.round(Math.log10(symInfo.pricescale)))
            : (priceStep < 1 ? Math.min(8, priceStep.toFixed(8).replace(/0+$/, '').split('.')[1]?.length || 2) : 2);

        const entryR = Number(entry.toFixed(decimals));
        const stopR = Number(stop.toFixed(decimals));
        const targetR = Number(target.toFixed(decimals));
        const extremeR = barAnalysis.entryCandleExtreme != null ? Number(barAnalysis.entryCandleExtreme.toFixed(decimals)) : '';
        const sess = detectSession(time);
        const id = getToolId(tool);

        const ind = readMatryoshkaIndicators(model, time, effectiveIdx, isLong);
        const signature = `${id}_${startIdx}_${endIdx}_${entryR}_${stopR}_${targetR}`;

        return {
            id,
            signature,
            serverTime: Number(tool._serverUpdateTime || 0),
            time,
            symbol: cleanSymbol,
            timeframeLat: formatTf(rawInterval, 'lat'),
            timeframeCyr: formatTf(rawInterval, 'cyr'),
            direction: isLong ? 'LONG' : 'SHORT',
            entryPrice: entryR,
            stopPrice: stopR,
            targetPrice: targetR,
            entryExtreme: extremeR,
            upperBound: Math.max(entryR, stopR),
            lowerBound: Math.min(entryR, stopR),
            riskPts: Number(Math.abs(entryR - stopR).toFixed(decimals)),
            potPts: Number(Math.abs(targetR - entryR).toFixed(decimals)),
            riskReward: Number(rrNum.toFixed(2)),
            candleDate: formatDate(time),
            shortDate: formatShortDate(time),
            mskTime: formatMskTime(time),
            officeTimeSlot: formatHalfHourSlot(time),
            sessionStd: sess.std,
            sessionEuUs: sess.eu_us,
            autoResult: barAnalysis.outcome,
            reached1R: barAnalysis.reached1R,
            reached3R: barAnalysis.reached3R,
            ind
        };
    }

    function scanChartTrades(silent = false) {
        try {
            const model = getTvModel();
            if (!model) {
                if (!silent) showToast('График TradingView еще не загружен', 'error');
                return false;
            }

            const ms = typeof model.mainSeries === 'function' ? model.mainSeries() : model.mainSeries;
            const bars = typeof ms.bars === 'function' ? ms.bars() : ms.bars;
            const tools = typeof model.allLineTools === 'function' ? model.allLineTools() : model.allLineTools;
            const list = Array.isArray(tools) ? tools : Object.values(tools || {});

            const matchingTools = list.filter(isSupportedTool);
            if (!matchingTools.length) {
                if (!silent) showToast('Нарисуйте на графике инструмент «Длинная или Короткая позиция»!', 'error');
                return false;
            }

            let symInfo = {};
            try {
                symInfo = (typeof ms.symbolInfo === 'function' ? ms.symbolInfo() : ms._symbolInfo) || {};
            } catch (e) {}

            const priceStep = (typeof ms.priceStep === 'function' ? ms.priceStep() : ms.priceStep) ||
                              (symInfo.minmov && symInfo.pricescale ? symInfo.minmov / symInfo.pricescale : 0.01);

            const parsed = matchingTools.map(t => parseSingleTool(t, model, ms, bars, symInfo, priceStep));
            parsed.sort((a, b) => a.time - b.time);
            allChartTrades = parsed;

            const chartSelId = detectCurrentlySelectedToolId(model);
            if (chartSelId !== lastChartSelectedId) {
                lastChartSelectedId = chartSelId;
                if (chartSelId) lastSelectedToolId = chartSelId;
            }

            let targetId = (lastSelectedToolId && parsed.some(p => p.id === lastSelectedToolId))
                ? lastSelectedToolId
                : chartSelId;

            if (!targetId && lastHoveredToolId && parsed.some(p => p.id === lastHoveredToolId)) {
                targetId = lastHoveredToolId;
            }
            if (!targetId) {
                let newest = parsed[parsed.length - 1];
                for (const p of parsed) {
                    if (p.serverTime > newest.serverTime) newest = p;
                }
                targetId = newest.id;
            }

            const idx = parsed.findIndex(p => p.id === targetId);
            activeTradeIdx = idx >= 0 ? idx : (parsed.length - 1);
            currentTrade = allChartTrades[activeTradeIdx];

            if (currentTrade?.ind?.flows) {
                const fl = currentTrade.ind.flows;
                saveState({
                    flowD: fl.d || state.flowD,
                    flow4H: fl.h4 || state.flow4H,
                    flow1H: fl.h1 || state.flow1H,
                    flow15M: fl.m15 || state.flow15M
                });
            }

            return true;
        } catch (e) {
            if (!silent) showToast('Ошибка чтения инструмента на графике', 'error');
            return false;
        }
    }

    function getActiveSep() {
        return state.outputProfile === 'en_form' ? '.' : ',';
    }

    function numStr(val) {
        if (val == null || val === '') return '';
        const s = String(val);
        return getActiveSep() === ',' ? s.replace(/\./g, ',') : s.replace(/,/g, '.');
    }

    function cleanText(s) {
        return String(s || '').replace(/[\t\r\n]+/g, ' ').trim();
    }

    function getActiveColumns() {
        if (state.presetKey === 'custom' && Array.isArray(state.customColumns) && state.customColumns.length) {
            return [...state.customColumns];
        }
        return [...(PRESET_SCHEMAS[state.presetKey] || PRESET_SCHEMAS.bt_main).columns];
    }

    function getActivePasteCol() {
        if (state.presetKey === 'custom') return state.customPasteCol || 'B';
        return PRESET_SCHEMAS[state.presetKey]?.pasteCol || 'B';
    }

    function getActiveStartRow() {
        if (state.presetKey === 'custom') return Number(state.customStartRow) || 2;
        return PRESET_SCHEMAS[state.presetKey]?.startRow || 6;
    }

    function getActiveSession() {
        if (manualSessionOverride) return manualSessionOverride;
        if (!currentTrade) return 'Лондон';
        const isEuUs = (PRESET_SCHEMAS[state.presetKey]?.sessionStyle === 'eu_us');
        return isEuUs ? currentTrade.sessionEuUs : currentTrade.sessionStd;
    }

    function colLetterToIndex(letter) {
        const clean = String(letter || 'B').toUpperCase().replace(/[^A-Z]/g, '') || 'B';
        let idx = 0;
        for (let i = 0; i < clean.length; i++) {
            idx = idx * 26 + (clean.charCodeAt(i) - 64);
        }
        return idx - 1;
    }

    function colIndexToLetter(idx) {
        let col = '', n = idx;
        while (n >= 0) {
            col = String.fromCharCode((n % 26) + 65) + col;
            n = Math.floor(n / 26) - 1;
        }
        return col;
    }

    function buildDynamicFormula(tok, cols, pasteCol, startRow, fallbackVal) {
        const isRu = (state.outputProfile === 'ru_form');
        const s = isRu ? ';' : ',';
        const IF = isRu ? 'ЕСЛИ' : 'IF';
        const OR = isRu ? 'ИЛИ' : 'OR';
        const IFERROR = isRu ? 'ЕСЛИОШИБКА' : 'IFERROR';
        const COUNT = isRu ? 'СЧЁТ' : 'COUNT';
        const COUNTIF = isRu ? 'СЧЁТЕСЛИ' : 'COUNTIF';
        const SUM = isRu ? 'СУММ' : 'SUM';
        const MAX = isRu ? 'МАКС' : 'MAX';
        const INT = isRu ? 'ЦЕЛОЕ' : 'INT';
        const ROUND = isRu ? 'ОКРУГЛ' : 'ROUND';

        const baseOffset = colLetterToIndex(pasteCol);
        const findCol = (...names) => {
            for (const name of names) {
                const i = cols.indexOf(name);
                if (i !== -1) return colIndexToLetter(baseOffset + i);
            }
            return null;
        };
        const ref = (colLet) => isRu ? `ИНДЕКС(${colLet}:${colLet};СТРОКА())` : `INDEX(${colLet}:${colLet},ROW())`;
        const vRef = (...names) => {
            const colLet = findCol(...names);
            return colLet ? ref(colLet) : null;
        };

        if (tok === 'risk_pts') {
            const en = vRef('entry'), st = vRef('stop');
            if (!en || !st) return fallbackVal;
            if (state.presetKey === 'bt_30') {
                return `=${IF}(${OR}(${en}=""${s}${st}="")${s}""${s}ABS(${en}-${st}))`;
            }
            return `=${IFERROR}(ABS(${en}-${st})${s}"")`;
        }

        if (tok === 'pot_pts') {
            const en = vRef('entry'), tg = vRef('target');
            if (!en || !tg) return fallbackVal;
            return `=${IFERROR}(ABS(${tg}-${en})${s}"")`;
        }

        if (tok === 'rr') {
            const pot = vRef('pot_pts'), rsk = vRef('risk_pts');
            const en = vRef('entry'), st = vRef('stop'), tg = vRef('target');
            if (state.presetKey === 'bt_30' && en && tg && rsk) {
                return `=${IF}(${OR}(${rsk}=""${s}${rsk}=0${s}${tg}="")${s}""${s}ABS(${tg}-${en})/${rsk})`;
            }
            if (pot && rsk) return `=${IFERROR}(${pot}/${rsk}${s}"")`;
            if (en && st && tg) return `=${IFERROR}(ABS(${tg}-${en})/ABS(${en}-${st})${s}"")`;
            return fallbackVal;
        }

        if (tok === 'lots_30') {
            const rsk = vRef('risk_pts');
            if (!rsk || state.presetKey !== 'bt_30') return fallbackVal;
            return `=${IF}(${OR}(${rsk}=""${s}${rsk}=0)${s}""${s}${INT}(('Параметры'!$B$5/${rsk}/'Параметры'!$B$6)/'Параметры'!$B$7)*'Параметры'!$B$7)`;
        }

        if (tok === 'r_out') {
            const rrRef = vRef('rr');
            const res30Ref = vRef('res_30');
            const manRRef = vRef('manual_r_30');
            if (res30Ref && rrRef) {
                const tail = manRRef || '""';
                return `=${IF}(${res30Ref}=""${s}""${s}${IF}(${res30Ref}="цель"${s}${rrRef}${s}${IF}(${res30Ref}="стоп"${s}-1${s}${IF}(${res30Ref}="безубыток"${s}0${s}${tail}))))`;
            }
            const resRef = vRef('result');
            if (resRef && rrRef) {
                return `=${IF}(${resRef}="Тейк"${s}${rrRef}${s}${IF}(${resRef}="Стоп"${s}-1${s}${IF}(${resRef}="Безубыток"${s}0${s}"")))`;
            }
            return fallbackVal;
        }

        if (tok === 'profit_usd') {
            const rOutRef = vRef('r_out'), lotsRef = vRef('lots_30'), rskRef = vRef('risk_pts');
            if (!rOutRef || !lotsRef || !rskRef || state.presetKey !== 'bt_30') return fallbackVal;
            return `=${IF}(${OR}(${rOutRef}=""${s}${lotsRef}="")${s}""${s}${rOutRef}*${lotsRef}*'Параметры'!$B$6*${rskRef})`;
        }

        if (tok === 'cum_r') {
            const rCol = findCol('r_out', 'r_out_off');
            if (!rCol) return fallbackVal;
            const rCur = ref(rCol);
            return `=${IF}(${COUNT}($${rCol}$${startRow}:${rCur})=0${s}""${s}${SUM}($${rCol}$${startRow}:${rCur}))`;
        }

        if (tok === 'dd_r') {
            const cumCol = findCol('cum_r');
            if (!cumCol) return fallbackVal;
            const cumCur = ref(cumCol);
            return `=${IF}(${cumCur}=""${s}""${s}${MAX}($${cumCol}$${startRow}:${cumCur})-${cumCur})`;
        }

        if (tok === 'factors_cnt') {
            const c1 = findCol('f_ret'), c4 = findCol('f_liq');
            if (!c1 || !c4) return fallbackVal;
            return `=${COUNTIF}(${ref(c1)}:${ref(c4)}${s}"да")`;
        }

        if (tok === 'valid_aoi') {
            const c1 = findCol('f_ret'), c4 = findCol('f_liq');
            if (!c1 || !c4) return fallbackVal;
            return `=${IF}(${COUNTIF}(${ref(c1)}:${ref(c4)}${s}"да")>=2${s}"ДА"${s}"нет")`;
        }

        if (tok === 'lots_off') {
            const en = vRef('entry'), st = vRef('stop'), rskPct = vRef('risk_pct_col');
            if (!en || !st || !rskPct || state.presetKey !== 'office') return fallbackVal;
            return `=${IF}(${OR}(${en}=""${s}${st}=""${s}${en}=${st})${s}""${s}${ROUND}('Итоги месяца'!$B$2*${rskPct}/(ABS(${en}-${st})/0.01*'Итоги месяца'!$B$3)${s}2))`;
        }

        if (tok === 'r_out_off') {
            const res = vRef('res_off'), sfPct = vRef('safe_pct'), sfRr = vRef('safe_rr'), rr = vRef('rr');
            if (!res || !sfPct || !sfRr || !rr) return fallbackVal;
            return `=${IF}(${res}=""${s}""${s}${IF}(${res}="Стоп"${s}-1${s}${IF}(${res}="Тейк"${s}${IF}(${sfPct}=""${s}0${s}${sfPct})*${IF}(${sfRr}=""${s}0${s}${sfRr})+(1-${IF}(${sfPct}=""${s}0${s}${sfPct}))*${rr}${s}${IF}(${res}="Сейф/БУ"${s}${IF}(${sfPct}=""${s}0${s}${sfPct})*${IF}(${sfRr}=""${s}0${s}${sfRr})${s}${IF}(${res}="Бу"${s}0${s}"")))))`;
        }

        if (tok === 'prof_pct_off') {
            const rOutOff = vRef('r_out_off'), rskPct = vRef('risk_pct_col');
            if (!rOutOff || !rskPct) return fallbackVal;
            return `=${IF}(${rOutOff}=""${s}""${s}${rOutOff}*${rskPct})`;
        }

        if (tok === 'move_pts') {
            const res = vRef('res_off', 'result'), dir = vRef('dir_title'), en = vRef('entry'), st = vRef('stop'), tg = vRef('target');
            if (!res || !dir || !en || !st || !tg) return fallbackVal;
            return `=${IF}(${res}=""${s}""${s}${IF}(${dir}="Боковик"${s}0${s}${IF}(${res}="Тейк"${s}${IF}(${dir}="Лонг"${s}${tg}-${en}${s}${en}-${tg})${s}${IF}(${res}="Стоп"${s}${IF}(${dir}="Лонг"${s}${st}-${en}${s}${en}-${st})${s}0))))`;
        }

        return fallbackVal;
    }

    const addTokenOptionsHtml = Object.entries(TOKEN_LABELS)
        .map(([k, v]) => `<option value="${k}">+ Добавить переменную: ${v}</option>`)
        .join('');

    const uiHtml = `
    <div id="bt-toast" style="display:none; position:fixed; bottom:24px; right:24px; z-index:10000000; background:#141417; color:#f4f4f5; padding:11px 16px; border-radius:8px; border-left:4px solid #089981; box-shadow:0 12px 32px rgba(0,0,0,0.9); font-family:-apple-system,BlinkMacSystemFont,'Inter','Segoe UI',Roboto,sans-serif; font-size:12px; font-weight:500; max-width:370px;"></div>

    <button type="button" id="bt-fab-btn" style="position:fixed; bottom:18px; left:58px; z-index:9999990; display:flex; align-items:center; gap:6px; background:#141417; color:#e5c158; border:1px solid rgba(229,193,88,0.45); border-radius:7px; padding:6px 12px; font-family:-apple-system,BlinkMacSystemFont,'Inter',sans-serif; font-size:12px; font-weight:600; cursor:pointer; box-shadow:0 4px 16px rgba(0,0,0,0.75); transition:0.15s;">
        ${LOGO_SVG}
        <span>Матрёшка</span>
        <span style="color:#8a8a93; font-weight:400; font-size:11px;">(${isMac ? '⌃Q' : 'F2'})</span>
    </button>

    <div id="bt-modal" class="bt-no-scroll" tabindex="-1" style="display:none; outline:none; position:fixed; top:${state.posTop}px; right:${state.posRight}px; z-index:9999999; background:#0d0d10; color:#f4f4f5; padding:13px; border-radius:12px; box-shadow:0 22px 55px rgba(0,0,0,0.92), 0 0 0 1px rgba(229,193,88,0.25); font-family:-apple-system,BlinkMacSystemFont,'Inter','Segoe UI',Roboto,sans-serif; width:388px; max-height:95vh; overflow-y:auto; box-sizing:border-box; user-select:none;">

        <div id="bt-header-drag" style="display:flex; justify-content:space-between; align-items:center; padding-bottom:9px; margin-bottom:9px; border-bottom:1px solid #222228; cursor:move;">
            <div style="display:flex; align-items:center; gap:7px;">
                ${LOGO_SVG}
                <span id="bt-active-tpl-title" style="font-size:13px; font-weight:700; color:#f4f4f5;"></span>
            </div>

            <div style="display:flex; align-items:center; gap:4px;">
                <div style="display:flex; align-items:center; background:#16161a; border:1px solid #2a2a32; border-radius:5px; padding:1px 3px;" title="Переключение между сделками на графике (Стрелки ← / →)">
                    <button type="button" id="bt-prev-trade" class="bt-icon-btn">◀</button>
                    <span id="bt-trade-counter" style="font-size:11px; font-weight:700; color:#e5c158; padding:0 5px;">1/1</span>
                    <button type="button" id="bt-next-trade" class="bt-icon-btn">▶</button>
                </div>
                <button type="button" id="bt-open-settings" class="bt-top-btn" title="Выбор журнала и настройка переменных">⚙️</button>
                <button type="button" id="bt-min-btn" class="bt-top-btn" title="Свернуть или развернуть окно">—</button>
                <button type="button" id="bt-close-x" class="bt-top-btn" title="Закрыть окно (Esc)">✕</button>
            </div>
        </div>

        <div id="bt-trade-card" style="background:#141418; border:1px solid #26262e; border-radius:8px; padding:8px 10px; margin-bottom:9px; transition:border-color 0.2s ease;">
            <div style="display:flex; justify-content:space-between; align-items:center; gap:4px; margin-bottom:6px; white-space:nowrap;">
                <div style="display:flex; align-items:center; gap:6px; font-size:11.5px; min-width:0; overflow:hidden;">
                    <b id="bt-info-sym" style="color:#ffffff; white-space:nowrap;">XAUUSD · M5</b>
                    <button type="button" id="bt-session-pill" title="Сессия определена автоматически по времени входа. Кликните для ручного переключения" style="background:rgba(229,193,88,0.14); color:#f3d26b; border:1px solid rgba(229,193,88,0.4); border-radius:4px; padding:1px 6px; font-size:10.5px; font-weight:600; cursor:pointer; white-space:nowrap;">⚡ Европа</button>
                    <span id="bt-info-date" style="color:#9ca3af; font-size:10.5px; white-space:nowrap;"></span>
                </div>
                <span id="bt-info-badge" style="background:#089981; color:#fff; font-size:10.5px; padding:2px 7px; border-radius:4px; font-weight:700; white-space:nowrap; flex-shrink:0;"></span>
            </div>

            <div id="bt-quick-copy-prices" title="Кликните, чтобы скопировать только 3 цены: Вход + Стоп + Цель" style="display:grid; grid-template-columns:1fr 1fr 1fr; gap:4px; background:#0e0e11; padding:6px 8px; border-radius:5px; cursor:pointer; border:1px solid #222228; font-size:11.5px; white-space:nowrap; transition:0.12s;">
                <div><span style="color:#8a8a93; font-size:10px;">Вход </span><b id="bt-info-en" style="color:#ffffff;"></b></div>
                <div style="text-align:center;"><span style="color:#8a8a93; font-size:10px;">Стоп </span><b id="bt-info-st" style="color:#ffffff;"></b></div>
                <div style="text-align:right;"><span style="color:#8a8a93; font-size:10px;">Цель </span><b id="bt-info-tg" style="color:#f3d26b;"></b></div>
            </div>
        </div>

        <div id="bt-collapsible-body">

            <div id="bt-settings-drawer" style="display:none; background:#141418; border:1px solid #3d3522; border-radius:8px; padding:11px;">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px; padding-bottom:7px; border-bottom:1px solid #26262e;">
                    <span style="font-size:11.5px; font-weight:700; color:#f3d26b;">⚙️ СХЕМА ЖУРНАЛА И ПЕРЕМЕННЫХ</span>
                    <button type="button" id="bt-close-settings-btn" style="background:#e5c158; color:#09090b; border:none; border-radius:5px; padding:4px 10px; font-size:11px; font-weight:700; cursor:pointer;">
                        ← Готово
                    </button>
                </div>

                <div class="bt-lbl">ШАБЛОН ЖУРНАЛА</div>
                <select id="bt-preset-select" class="bt-sel" style="border-color:#e5c158; color:#f3d26b; font-weight:600; margin-bottom:10px;">
                    <option value="bt_main">📘 1. Бэктест «Матрёшка» (20 переменных)</option>
                    <option value="bt_30">📊 2. Бэктест 30 сделок (30_trades.xlsx)</option>
                    <option value="dz_poi">🎯 3. ДЗ: Журнал POI</option>
                    <option value="dz_aoi">🗺️ 4. ДЗ: Таблица AOI</option>
                    <option value="office">🏢 5. Торговый офис (Недели 1–4)</option>
                    <option value="prices_only">⚡ 6. Только 3 цены (Вход / Стоп / Цель)</option>
                    <option value="custom">✨ 7. Универсальный конструктор (Свой журнал)</option>
                </select>

                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <button type="button" id="bt-toggle-cols-btn" class="bt-mini-pill" style="width:100%; text-align:center; padding:5px;">
                        🔧 Конструктор переменных и столбцов...
                    </button>
                </div>

                <div id="bt-cols-editor-wrap" style="display:none; margin-top:8px; padding-top:8px; border-top:1px solid #26262e;">
                    <div style="display:grid; grid-template-columns:1fr 1fr; gap:6px; margin-bottom:7px;">
                        <div>
                            <div class="bt-lbl">СТАРТОВЫЙ СТОЛБЕЦ ВСТАВКИ</div>
                            <select id="bt-custom-paste-col" class="bt-sel" style="margin-bottom:0; height:26px;">
                                <option value="A">Столбец A (с 1-й кол.)</option>
                                <option value="B">Столбец B (после №)</option>
                                <option value="C">Столбец C</option>
                                <option value="D">Столбец D</option>
                            </select>
                        </div>
                        <div>
                            <div class="bt-lbl">1-Я СТРОКА ДАННЫХ (ДЛЯ СУММ)</div>
                            <input type="number" id="bt-custom-start-row" min="1" max="500" value="${getActiveStartRow()}" style="width:100%; height:26px; box-sizing:border-box; background:#16161b; color:#f4f4f5; border:1px solid #282830; border-radius:5px; padding:0 6px; font-size:11px; outline:none;">
                        </div>
                    </div>

                    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:5px; white-space:nowrap;">
                        <span style="font-size:10px; font-weight:700; color:#9ca3af;">ПОРЯДОК ПЕРЕМЕННЫХ (ПЕРЕТАЩИТЕ МЫШКОЙ):</span>
                        <button type="button" id="bt-reset-schema" class="bt-mini-pill" style="padding:2px 7px; font-size:10px; white-space:nowrap;">↺ Сбросить</button>
                    </div>
                    <div id="bt-schema-pills" class="bt-no-scroll" style="display:flex; flex-wrap:wrap; gap:4px; max-height:135px; overflow-y:auto; padding:6px; background:#0b0b0e; border-radius:6px; border:1px solid #27272e; margin-bottom:6px;"></div>

                    <select id="bt-add-token-select" class="bt-sel" style="margin-bottom:0; font-size:11px; color:#f3d26b;">
                        <option value="">+ Добавить переменную графика / сделки...</option>
                        ${addTokenOptionsHtml}
                    </select>
                </div>
            </div>

            <div id="bt-main-view">
                <div id="bt-dynamic-fields"></div>

                <div id="bt-extra-inputs" style="position:relative; display:flex; gap:6px; margin-top:4px; margin-bottom:7px;">
                    <div style="position:relative; flex:1; display:flex; align-items:center;">
                        <input type="text" id="bt-inp-comment" autocomplete="off" placeholder="Комментарий к сделке (Tab)..." style="width:100%; box-sizing:border-box; background:#141418; border:1px solid #2a2a32; color:#ffffff; padding:7px 44px 7px 9px; border-radius:6px; font-size:11.5px; outline:none;">
                        <button type="button" id="bt-clear-comment" title="Очистить комментарий" style="position:absolute; right:22px; background:transparent; border:none; color:#71717a; cursor:pointer; font-size:11px; padding:2px 4px;">✕</button>
                        <button type="button" id="bt-toggle-presets" title="Выбрать готовый шаблон комментария" style="position:absolute; right:4px; background:transparent; border:none; color:#e5c158; cursor:pointer; font-size:11px; padding:2px 5px;">▾</button>
                    </div>
                    <div id="bt-comment-menu" class="bt-no-scroll" style="display:none; position:absolute; bottom:36px; left:0; width:285px; max-height:185px; overflow-y:auto; background:#18181e; border:1px solid #3d3522; border-radius:6px; box-shadow:0 10px 25px rgba(0,0,0,0.85); z-index:100; padding:4px;">
                        <div class="bt-cm-item" style="padding:5px 8px; font-size:11px; color:#e4e4e7; cursor:pointer; border-radius:4px;">7 из 7 — по закрытой свече</div>
                        <div class="bt-cm-item" style="padding:5px 8px; font-size:11px; color:#e4e4e7; cursor:pointer; border-radius:4px;">рано — приманка не снята</div>
                        <div class="bt-cm-item" style="padding:5px 8px; font-size:11px; color:#e4e4e7; cursor:pointer; border-radius:4px;">зона против потока</div>
                        <div class="bt-cm-item" style="padding:5px 8px; font-size:11px; color:#e4e4e7; cursor:pointer; border-radius:4px;">стоп за тень — не за зону</div>
                        <div class="bt-cm-item" style="padding:5px 8px; font-size:11px; color:#e4e4e7; cursor:pointer; border-radius:4px;">цель дальше ближайшей 0.5</div>
                        <div class="bt-cm-item" style="padding:5px 8px; font-size:11px; color:#e4e4e7; cursor:pointer; border-radius:4px;">догонял движение</div>
                        <div class="bt-cm-item" style="padding:5px 8px; font-size:11px; color:#e4e4e7; cursor:pointer; border-radius:4px;">внутри проторговки</div>
                        <div class="bt-cm-item" style="padding:5px 8px; font-size:11px; color:#e4e4e7; cursor:pointer; border-radius:4px;">по бледной свече</div>
                        <div class="bt-cm-item" style="padding:5px 8px; font-size:11px; color:#e4e4e7; cursor:pointer; border-radius:4px;">зона с касанием 2</div>
                        <div class="bt-cm-item" style="padding:5px 8px; font-size:11px; color:#e4e4e7; cursor:pointer; border-radius:4px;">по правилам — рынок не пошёл</div>
                    </div>
                    <button type="button" id="bt-screen-pill" title="Нажмите Alt+S в TradingView, затем Ctrl+V в окне" style="background:#16161a; color:#d4d4d8; border:1px solid #2a2a32; border-radius:6px; padding:0 10px; font-size:11.5px; font-weight:500; cursor:pointer; white-space:nowrap; transition:0.15s;">
                        📸 Скрин
                    </button>
                </div>

                <div style="display:grid; grid-template-columns:1fr 1fr; gap:4px; margin-bottom:7px; background:#121216; padding:3px; border-radius:6px; border:1px solid #22222a;">
                    <button type="button" class="bt-prof-btn" data-prof="ru_form" title="Для русских таблиц: дробь через запятую и формулы =ЕСЛИ(...; СТРОКА())">
                        🇷🇺 RU ( , + ЕСЛИ )
                    </button>
                    <button type="button" class="bt-prof-btn" data-prof="en_form" title="Для международных таблиц: дробь через точку и формулы =IF(..., ROW())">
                        🌐 EN ( . + IF )
                    </button>
                </div>

                <div style="display:flex; gap:6px; align-items:stretch;">
                    <button type="button" id="bt-main-copy-btn" style="flex:1; background:linear-gradient(180deg, #ebd074 0%, #cfa538 100%); color:#09090b; border:none; padding:9px 12px; border-radius:6px; cursor:pointer; font-weight:700; font-size:12.5px; display:flex; justify-content:center; align-items:center; gap:6px; box-shadow:0 3px 12px rgba(229,193,88,0.28);">
                        <span id="bt-main-btn-text">Скопировать строку (в стлб. B)</span>
                        <kbd style="background:rgba(0,0,0,0.18); color:#09090b; padding:1px 5px; border-radius:3px; font-size:10.5px; font-family:inherit;">↵</kbd>
                    </button>

                    <div style="display:flex; align-items:center; background:#141418; border:1px solid #2a2a32; border-radius:6px; overflow:hidden;">
                        <button type="button" id="bt-copy-batch-btn" title="Скопировать всю накопленную серию сделок сразу" style="background:transparent; color:#e4e4e7; border:none; padding:0 10px; height:100%; font-size:11.5px; font-weight:600; cursor:pointer; display:flex; align-items:center; gap:4px;">
                            📦 <span id="bt-batch-count" style="color:#f3d26b;">0</span>
                        </button>
                        <button type="button" id="bt-clear-batch-btn" title="Очистить накопленную серию сделок" style="background:rgba(242,54,69,0.14); color:#f23645; border:none; border-left:1px solid #2a2a32; padding:0 8px; height:100%; font-size:11px; cursor:pointer;">✕</button>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <style>
        .bt-no-scroll::-webkit-scrollbar { width:4px; height:4px; }
        .bt-no-scroll::-webkit-scrollbar-track { background:transparent; }
        .bt-no-scroll::-webkit-scrollbar-thumb { background:#2e2e38; border-radius:4px; }
        .bt-section-box { background:#121216; border:1px solid #22222a; border-radius:7px; padding:7px 8px; margin-bottom:7px; }
        .bt-lbl { font-size:9.5px; color:#9ca3af; height:13px; line-height:13px; margin-bottom:4px; font-weight:700; text-transform:uppercase; letter-spacing:0.35px; display:flex; align-items:center; gap:3px; white-space:nowrap; overflow:hidden; }
        .bt-bolt { color:#e5c158; font-size:10px; line-height:1; flex-shrink:0; }
        .bt-grid-1 { display:grid; grid-template-columns:1fr; gap:4px; margin-bottom:6px; }
        .bt-grid-2 { display:grid; grid-template-columns:1fr 1fr; gap:5px; margin-bottom:6px; }
        .bt-grid-3 { display:grid; grid-template-columns:1fr 1fr 1fr; gap:5px; margin-bottom:6px; }
        .bt-grid-4 { display:grid; grid-template-columns:repeat(4, 1fr); gap:5px; margin-bottom:6px; }
        .bt-grid-5 { display:grid; grid-template-columns:repeat(5, 1fr); gap:4px; margin-bottom:6px; }

        .bt-chip { position:relative; box-sizing:border-box; width:100%; height:28px; background:#16161b; color:#a1a1aa; border:1px solid #282830; padding:0 7px; font-size:11px; font-weight:500; border-radius:5px; cursor:pointer; display:flex; justify-content:space-between; align-items:center; transition:0.12s ease; white-space:nowrap; }
        .bt-chip.bt-chip-centered { justify-content:center; text-align:center; padding:0 4px; }
        .bt-chip:hover { border-color:#52525b; color:#f4f4f5; background:#1c1c22; }
        .bt-chip.active {
            background: var(--ac-bg, rgba(229, 193, 88, 0.16));
            color: var(--ac-fg, #f7d97e);
            border-color: var(--ac-bd, rgba(229, 193, 88, 0.62));
            font-weight: 600;
        }
        .bt-chip kbd { font-size:9.5px; font-family:inherit; color:#71717a; background:#1e1e24; border:1px solid #2e2e38; padding:0px 4px; border-radius:3px; margin-left:3px; font-weight:600; }
        .bt-chip.active kbd { background:rgba(255,255,255,0.08); border-color:transparent; color:inherit; }

        .bt-flow-pill { width:100%; height:28px; border-radius:5px; font-size:11px; font-weight:600; cursor:pointer; display:flex; justify-content:space-between; align-items:center; padding:0 7px; box-sizing:border-box; transition:0.12s ease; border:1px solid #282830; background:#16161b; color:#d4d4d8; }
        .bt-flow-pill[data-dir="вверх"] { background:rgba(16, 185, 129, 0.14); color:#34d399; border-color:rgba(16, 185, 129, 0.45); }
        .bt-flow-pill[data-dir="вниз"] { background:rgba(242, 54, 69, 0.14); color:#f87171; border-color:rgba(242, 54, 69, 0.45); }
        .bt-flow-pill[data-dir="—"] { background:#16161b; color:#71717a; border-color:#282830; }

        .bt-prof-btn { background:transparent; color:#8a8a93; border:1px solid transparent; border-radius:4px; padding:4px 4px; font-size:10.5px; font-weight:600; cursor:pointer; transition:0.12s; white-space:nowrap; }
        .bt-prof-btn:hover { color:#d4d4d8; }
        .bt-prof-btn.active { background:rgba(229, 193, 88, 0.18); color:#f7d97e; border-color:rgba(229, 193, 88, 0.55); }

        .bt-struct-preset { background:#18181e; color:#a1a1aa; border:1px solid #282832; border-radius:4px; padding:2px 5px; font-size:9.5px; cursor:pointer; transition:0.12s; white-space:nowrap; }
        .bt-struct-preset:hover { color:#f3d26b; border-color:#e5c158; }

        .bt-sel { width:100%; height:28px; box-sizing:border-box; background:#16161b; color:#f4f4f5; border:1px solid #282830; padding:0 6px; border-radius:5px; font-size:11px; font-weight:500; outline:none; margin-bottom:6px; cursor:pointer; }
        .bt-sel:focus, #bt-inp-comment:focus, #f-swept-lvl:focus { border-color:#e5c158; }
        .bt-mini-pill { background:#1c1c22; color:#d4d4d8; border:1px solid #2e2e38; padding:3px 8px; font-size:10.5px; border-radius:4px; cursor:pointer; transition:0.12s; }
        .bt-mini-pill.active { background:rgba(229, 193, 88, 0.2); color:#f7d97e; border-color:#e5c158; font-weight:700; }
        .bt-top-btn { background:#16161a; color:#d4d4d8; border:1px solid #2a2a32; border-radius:5px; padding:2px 7px; font-size:11.5px; cursor:pointer; transition:0.12s; }
        .bt-top-btn:hover, .bt-top-btn.active { color:#f3d26b; border-color:#e5c158; background:rgba(229,193,88,0.14); }
        .bt-icon-btn { background:transparent; color:#d4d4d8; border:none; cursor:pointer; font-size:10.5px; padding:2px 5px; }
        .bt-icon-btn:hover { color:#f3d26b; }
        .bt-drag-pill { background:#1a1a20; color:#f4f4f5; border:1px solid #383844; border-radius:5px; padding:3px 7px; font-size:10.5px; cursor:grab; display:inline-flex; align-items:center; gap:5px; transition:transform 0.1s ease, border-color 0.1s ease; }
        .bt-drag-pill.dragging-active { background:#2a2515; border-color:#f3d26b; color:#f3d26b; transform:scale(1.05); z-index:10; cursor:grabbing; box-shadow:0 4px 12px rgba(0,0,0,0.6); }
        .bt-del-col { color:#f23645; font-weight:bold; cursor:pointer; padding:0 2px; font-size:11px; }
		.bt-cm-item:hover, .bt-cc-item:hover { background:rgba(229,193,88,0.18); color:#f3d26b !important; }
        #bt-cc-btn:hover { border-color:#52525b; }
        #bt-quick-copy-prices:hover { border-color:rgba(229,193,88,0.55); }
    </style>
    `;

    document.body.insertAdjacentHTML('beforeend', uiHtml);

    const modalEl = document.getElementById('bt-modal');
    const toastEl = document.getElementById('bt-toast');
    const tradeCardEl = document.getElementById('bt-trade-card');
    const collapsibleBodyEl = document.getElementById('bt-collapsible-body');
    const mainViewEl = document.getElementById('bt-main-view');
    const dynFieldsEl = document.getElementById('bt-dynamic-fields');
    const settingsDrawerEl = document.getElementById('bt-settings-drawer');
    const openSettingsBtn = document.getElementById('bt-open-settings');
    const mainCopyBtn = document.getElementById('bt-main-copy-btn');
    const mainBtnText = document.getElementById('bt-main-btn-text');
    const screenPillBtn = document.getElementById('bt-screen-pill');
    const sessionPillBtn = document.getElementById('bt-session-pill');
    const commentInpEl = document.getElementById('bt-inp-comment');
    const commentMenuEl = document.getElementById('bt-comment-menu');
    let toastTimer = null;

    function showToast(msg, type = 'success') {
        toastEl.innerText = msg;
        toastEl.style.borderLeftColor = type === 'error' ? '#f23645' : '#089981';
        toastEl.style.display = 'block';
        clearTimeout(toastTimer);
        toastTimer = setTimeout(() => { toastEl.style.display = 'none'; }, 3000);
    }

    function pulseTradeCard(color = '#e5c158') {
        if (!tradeCardEl) return;
        tradeCardEl.style.borderColor = color;
        setTimeout(() => { tradeCardEl.style.borderColor = '#26262e'; }, 380);
    }

    function toggleSettingsView(forceState) {
        isSettingsOpen = typeof forceState === 'boolean' ? forceState : !isSettingsOpen;
        if (state.minimized && isSettingsOpen) saveState({ minimized: false });
        settingsDrawerEl.style.display = isSettingsOpen ? 'block' : 'none';
        mainViewEl.style.display = isSettingsOpen ? 'none' : 'block';
        openSettingsBtn.classList.toggle('active', isSettingsOpen);
        if (isSettingsOpen) renderSettingsDrawer();
        updateHeaderUI();
    }

    function makeChips(id, options, activeVal, cols = 2) {
        const htmlOpts = options.map(o => {
            const val = typeof o === 'string' ? o : o.val;
            const label = typeof o === 'string' ? o : o.label;
            const kbd = o.kbd ? `<kbd>${o.kbd}</kbd>` : '';
            const centeredClass = o.kbd ? '' : 'bt-chip-centered';
            let styleVars = '';
            if (o.tone === 'green') {
                styleVars = `style="--ac-bg:rgba(16,185,129,0.18); --ac-fg:#34d399; --ac-bd:#10b981;"`;
            } else if (o.tone === 'red') {
                styleVars = `style="--ac-bg:rgba(242,54,69,0.2); --ac-fg:#f87171; --ac-bd:#f23645;"`;
            }
            const isAct = val === activeVal ? 'active' : '';
            return `<button type="button" class="bt-chip ${centeredClass} ${isAct}" data-val="${val}" ${styleVars}><span>${label}</span>${kbd}</button>`;
        }).join('');
        return `<div class="bt-group bt-grid-${cols}" id="${id}">${htmlOpts}</div>`;
    }

	const CANDLE_COLORS = {
        'Лонг 3': { hex: '#72f238', hint: 'салатовая' },
        'Лонг 2': { hex: '#4bb8d4', hint: 'голубая' },
        'Лонг 1': { hex: '#3b7dff', hint: 'синяя' },
        'Шорт 3': { hex: '#f7e84b', hint: 'жёлтая' },
        'Шорт 2': { hex: '#eb3b4b', hint: 'красная' },
        'Шорт 1': { hex: '#8a2bbd', hint: 'фиолетовая' }
    };

    function makeCandleColorSelect(activeVal) {
        const info = CANDLE_COLORS[activeVal] || CANDLE_COLORS['Лонг 3'];
        const itemsHtml = Object.entries(CANDLE_COLORS).map(([k, v]) => `
            <div class="bt-cc-item" data-val="${k}" style="display:flex; align-items:center; justify-content:space-between; padding:5px 8px; border-radius:4px; cursor:pointer; font-size:11px; color:#f4f4f5;">
                <div style="display:flex; align-items:center; gap:7px;">
                    <span style="width:13px; height:13px; border-radius:3px; background:${v.hex}; display:inline-block; box-shadow:0 0 0 1px rgba(255,255,255,0.2);"></span>
                    <span style="font-weight:600;">${k}</span>
                </div>
                <span style="color:#8a8a93; font-size:10px;">${v.hint}</span>
            </div>
        `).join('');

        return `
            <div id="bt-cc-wrap" style="position:relative; margin-bottom:6px;">
                <input type="hidden" id="f-candle-col" value="${activeVal}">
                <button type="button" id="bt-cc-btn" style="width:100%; height:28px; box-sizing:border-box; background:#16161b; color:#f4f4f5; border:1px solid #282830; padding:0 7px; border-radius:5px; font-size:11px; font-weight:500; cursor:pointer; display:flex; align-items:center; justify-content:space-between; outline:none; transition:0.12s;">
                    <span style="display:flex; align-items:center; gap:6px; min-width:0; overflow:hidden;">
                        <span id="bt-cc-dot" style="width:12px; height:12px; border-radius:3px; background:${info.hex}; flex-shrink:0; box-shadow:0 0 0 1px rgba(255,255,255,0.22);"></span>
                        <span id="bt-cc-label" style="font-weight:600;">${activeVal}</span>
                        <span id="bt-cc-hint" style="color:#8a8a93; font-size:10px;">${info.hint}</span>
                    </span>
                    <span style="color:#8a8a93; font-size:10px;">▾</span>
                </button>
                <div id="bt-cc-menu" style="display:none; position:absolute; top:31px; left:0; width:100%; background:#18181e; border:1px solid #3d3522; border-radius:6px; box-shadow:0 10px 25px rgba(0,0,0,0.92); z-index:120; padding:3px; box-sizing:border-box;">
                    ${itemsHtml}
                </div>
            </div>
        `;
    }

    function setCandleColVal(val) {
        const info = CANDLE_COLORS[val];
        if (!info) return;
        const inp = document.getElementById('f-candle-col');
        const dot = document.getElementById('bt-cc-dot');
        const lbl = document.getElementById('bt-cc-label');
        const hnt = document.getElementById('bt-cc-hint');
        if (inp) inp.value = val;
        if (dot) dot.style.background = info.hex;
        if (lbl) lbl.innerText = val;
        if (hnt) hnt.innerText = info.hint;
    }

    function makeSelect(id, options, activeVal, mb = '6px') {
        const htmlOpts = options.map(o => {
            const val = typeof o === 'string' ? o : o.val;
            const label = typeof o === 'string' ? o : o.label;
            return `<option value="${val}" ${val === activeVal ? 'selected' : ''}>${label}</option>`;
        }).join('');
        return `<select class="bt-sel" id="${id}" style="margin-bottom:${mb};">${htmlOpts}</select>`;
    }

    function getChipVal(id, fallback = '') {
        return document.querySelector(`#${id} .active`)?.dataset.val || fallback;
    }

    function setChipVal(id, val) {
        const grp = document.getElementById(id);
        if (!grp) return;
        grp.querySelectorAll('.bt-chip').forEach(b => b.classList.toggle('active', b.dataset.val === val));
    }

    function makeFlowPill(id, tfLabel, dirVal) {
        const arrowText = dirVal === 'вверх' ? '↑ вверх' : (dirVal === 'вниз' ? '↓ вниз' : '—');
        return `<button type="button" class="bt-flow-pill" id="${id}" data-tf="${tfLabel}" data-dir="${dirVal}" title="Кликните, чтобы переключить направление потока ${tfLabel}">
            <span style="color:#9ca3af; font-weight:500;">${tfLabel}</span>
            <span class="bt-flow-val">${arrowText}</span>
        </button>`;
    }

    function getFlowVal(id, fallback = 'вверх') {
        return document.getElementById(id)?.dataset.dir || fallback;
    }

    function updateScreenshotPillUI() {
        if (screenshotUrl) {
            screenPillBtn.innerText = '📸 Скрин ✓';
            screenPillBtn.style.background = 'rgba(16, 185, 129, 0.18)';
            screenPillBtn.style.color = '#34d399';
            screenPillBtn.style.borderColor = '#10b981';
            screenPillBtn.title = `Прикреплен: ${screenshotUrl} (Кликните, чтобы очистить)`;
        } else {
            screenPillBtn.innerText = '📸 Скрин';
            screenPillBtn.style.background = '#16161a';
            screenPillBtn.style.color = '#d4d4d8';
            screenPillBtn.style.borderColor = '#2a2a32';
            screenPillBtn.title = 'Нажмите Alt+S в TradingView, затем Ctrl+V в окне';
        }
    }

    function sync30TradesChecklist() {
        const chkGrp = document.getElementById('f-check30');
        if (!chkGrp) return;

        const fuel = getChipVal('f-fuel30', 'да');
        const touch = getChipVal('f-touch30', '1');
        const against = getChipVal('f-against30', 'нет');
        const swept = document.getElementById('f-swept')?.value || 'приманка';

        if (swept === 'ничего не снято') {
            setChipVal('f-check30', 'нет');
            const lvlInp = document.getElementById('f-swept-lvl');
            if (lvlInp) lvlInp.value = '';
        } else if (against === 'да' || touch !== '1' || fuel === 'нет') {
            setChipVal('f-check30', 'нет');
        } else {
            setChipVal('f-check30', 'да');
        }
    }

    function autoSuggestError() {
        if (!currentTrade) return;
        const errSel = document.getElementById('f-err');
        const errHint = document.getElementById('bt-err-reason');
        if (!errSel) return;

        const res = getChipVal('f-res', currentTrade.autoResult);
        const liq = getChipVal('f-liq', 'да');
        const conf = document.getElementById('f-conf')?.value || '';
        const isPoiErr = getActiveColumns().includes('error_poi');

        if (res === 'Тейк' || res === 'Пропуск') {
            errSel.value = '—';
            if (errHint) errHint.innerText = '';
            return;
        }
        if (res === 'Стоп') {
            if (liq === 'нет') {
                errSel.value = isPoiErr ? 'Без снятия' : 'Вход без снятия ликвидности';
                if (errHint) errHint.innerText = '⚡ нет снятия';
                return;
            }
            if (conf === 'Без подтверждения') {
                errSel.value = isPoiErr ? 'Без реакции' : 'Нет подтверждения';
                if (errHint) errHint.innerText = '⚡ без подтв.';
                return;
            }
            if (currentTrade.riskReward > 0 && currentTrade.riskReward < 2.5) {
                errSel.value = isPoiErr ? 'R:R ниже нормы' : 'Нарушен минимальный R:R';
                if (errHint) errHint.innerText = `⚡ R:R ${numStr(currentTrade.riskReward)} < 2.5`;
                return;
            }
        }
        errSel.value = '—';
        if (errHint) errHint.innerText = '';
    }

    function renderMainFields() {
        const cols = getActiveColumns();
        const has = tok => cols.includes(tok);
        const t = currentTrade || {
            direction: 'LONG', sessionStd: 'Лондон', sessionEuUs: 'Европа', autoResult: 'Тейк', reached1R: true, reached3R: true, entryExtreme: '', ind: {}
        };
        const isLong = t.direction === 'LONG';
        const ind = t.ind || {};

        let html = '';

        if (has('aoi_num')) {
            html += `
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:7px;">
                    <span class="bt-lbl" style="margin:0;">№ ОБЛАСТИ / ЗОНЫ</span>
                    <input type="number" id="f-aoi-num" value="${state.aoiNum}" min="1" max="999" style="width:54px; background:#151519; color:#f3d26b; border:1px solid #282830; padding:3px; border-radius:4px; font-weight:700; text-align:center;">
                </div>
            `;
        }

        if (has('trend')) {
            const autoTrend = (state.trend === 'Тренд вверх' || state.trend === 'Тренд вниз')
                ? (isLong ? 'Тренд вверх' : 'Тренд вниз')
                : state.trend;
            html += `
                <div class="bt-lbl">ФАЗА РЫНКА ${BOLT}</div>
                ${makeChips('f-trend', [
                    { val: 'Тренд вверх', label: 'Тренд вверх', kbd: 'A' },
                    { val: 'Тренд вниз', label: 'Тренд вниз', kbd: 'S' },
                    { val: 'Боковик', label: 'Боковик', kbd: 'D' },
                    { val: 'Импульс/новости', label: 'Импульс / новости', kbd: 'F' }
                ], autoTrend, 2)}
            `;
        }

        if (has('trend_h4') || has('f_ret') || has('f_ob') || has('f_fvg') || has('f_liq')) {
            const h4Def = ind.flows?.h4 ? (ind.flows.h4 === 'вверх' ? 'Вверх' : 'Вниз') : (isLong ? 'Вверх' : 'Вниз');
            html += `<div class="bt-lbl">ТРЕНД H4 ${BOLT} И ФАКТОРЫ ОБЛАСТИ</div>` +
                makeChips('f-h4-trend', ['Вверх', 'Вниз'], h4Def, 2) +
                `<div style="display:grid; grid-template-columns:repeat(4, 1fr); gap:5px; margin-bottom:6px;">
                    <div><div style="font-size:10px; color:#9ca3af; text-align:center; margin-bottom:2px;">Откат</div>${makeChips('f-f1', ['да', 'нет'], 'да', 1)}</div>
                    <div><div style="font-size:10px; color:#9ca3af; text-align:center; margin-bottom:2px;">OB H4</div>${makeChips('f-f2', ['да', 'нет'], 'да', 1)}</div>
                    <div><div style="font-size:10px; color:#9ca3af; text-align:center; margin-bottom:2px;">Имб.</div>${makeChips('f-f3', ['да', 'нет'], 'нет', 1)}</div>
                    <div><div style="font-size:10px; color:#9ca3af; text-align:center; margin-bottom:2px;">Ликв.</div>${makeChips('f-f4', ['да', 'нет'], 'да', 1)}</div>
                </div>` +
                `<div style="display:grid; grid-template-columns:1fr 2fr; gap:6px;">
                    <div><div class="bt-lbl">ЦЕНА ДОШЛА</div>${makeChips('f-reached', [{val:'да', label:'Да'}, {val:'нет', label:'Нет'}], 'да', 2)}</div>
                    <div><div class="bt-lbl">РЕЗУЛЬТАТ ОБЛАСТИ</div>${makeChips('f-aoi-res', ['Отбой', 'Пробита', 'Ждёт'], 'Отбой', 3)}</div>
                </div>`;
        }

        if (has('zone')) {
            html += `<div class="bt-lbl">ЗОНА ИНТЕРЕСА</div>` +
                makeChips('f-zone', [
                    { val: 'Ордер-блок', label: 'Ордер-блок', kbd: '1' },
                    { val: 'Имбаланс', label: 'Имбаланс', kbd: '2' },
                    { val: 'Материнская свеча', label: 'Материнская свеча', kbd: '3' },
                    { val: 'Зона ликвидности', label: 'Зона ликвидности', kbd: '4' },
                    { val: 'Правильный откат', label: 'Правильный откат', kbd: '5' }
                ], state.zone, 2);
        }

        if (has('poi_type')) {
            html += `<div class="bt-lbl">ТИП ТОЧКИ ВХОДА (POI)</div>` +
                makeChips('f-poi-type', [
                    { val: 'OB', label: 'OB', kbd: '1' },
                    { val: 'Имбаланс', label: 'Имбаланс', kbd: '2' },
                    { val: '50% МС', label: '50% МС', kbd: '3' }
                ], 'OB', 3);
        }

        if (has('form_m1')) {
            html += `<div class="bt-lbl">ФОРМАЦИЯ ВХОДА НА M1</div>` +
                makeChips('f-form-m1', [
                    { val: 'SCOB', label: 'SCOB', kbd: '1' },
                    { val: 'Свип', label: 'Свип', kbd: '2' },
                    { val: 'Пин-бар (IFS)', label: 'Пин-бар', kbd: '3' },
                    { val: 'Поглощение', label: 'Поглощение', kbd: '4' },
                    { val: 'Правильный откат', label: 'Правильный откат', kbd: '5' }
                ], 'SCOB', 2);
        }

        const hasAnyFlowOr30 = has('flow_d') || has('flow_4h') || has('flow_1h') || has('flow_15m') ||
                               has('struct_30') || has('zone_30') || has('candle_col') ||
                               has('swept_30') || has('swept_level_30') || has('fuel_30') ||
                               has('touch_30') || has('against_flow') || has('check_30') ||
                               has('hit_1r') || has('hit_3r');

        if (hasAnyFlowOr30) {
            const defCandleCol = ind.candleColor || (isLong ? 'Лонг 3' : 'Шорт 3');
            const defStruct = ind.detectedStructure || (isLong ? 'пробой ↓ · пробой ↓ · слом ↑' : 'пробой ↑ · пробой ↑ · слом ↓');

            const stSlom = isLong ? 'слом ↑' : 'слом ↓';
            const st2Prob = isLong ? 'пробой ↓ · пробой ↓ · слом ↑' : 'пробой ↑ · пробой ↑ · слом ↓';
            const stFake = isLong ? 'ложный пробой ↓ · слом ↑' : 'ложный пробой ↑ · слом ↓';
            const stCont = isLong ? 'пробой ↑ · пробой ↑' : 'пробой ↓ · пробой ↓';

            html += `
                <div class="bt-section-box">
                    <div class="bt-lbl">1. ПОТОКИ В МОМЕНТ ВХОДА (КЛИК ДЛЯ СМЕНЫ) ${ind.flows ? BOLT : ''}</div>
                    <div style="display:grid; grid-template-columns:repeat(4, 1fr); gap:5px; margin-bottom:7px;">
                        ${makeFlowPill('f-fl-d', 'Д', state.flowD)}
                        ${makeFlowPill('f-fl-4h', '4Ч', state.flow4H)}
                        ${makeFlowPill('f-fl-1h', '1Ч', state.flow1H)}
                        ${makeFlowPill('f-fl-15m', '15м', state.flow15M)}
                    </div>

                    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:3px;">
                        <div class="bt-lbl" style="margin:0;">СТРУКТУРА ПЕРЕД ВХОДОМ ${ind.detectedStructure ? BOLT : ''}</div>
                        <div style="display:flex; gap:3px;">
                            <button type="button" class="bt-struct-preset" data-st="${stSlom}">Слом</button>
                            <button type="button" class="bt-struct-preset" data-st="${st2Prob}">2 проб.+слом</button>
                            <button type="button" class="bt-struct-preset" data-st="${stFake}">Ложн. пробой</button>
                            <button type="button" class="bt-struct-preset" data-st="${stCont}">2 пробоя</button>
                        </div>
                    </div>
                    ${makeSelect('f-struct', ['пробой ↑', 'пробой ↑ · пробой ↑', 'пробой ↑ · пробой ↑ · слом ↓', 'пробой ↓', 'пробой ↓ · пробой ↓', 'пробой ↓ · пробой ↓ · слом ↑', 'слом ↑', 'слом ↓', 'ложный пробой ↑ · слом ↓', 'ложный пробой ↓ · слом ↑', 'внутр. слом ↑', 'внутр. слом ↓', 'смена ↑ на 4Ч', 'смена ↓ на 4Ч'], defStruct, '0px')}
                </div>

                <div class="bt-section-box">
                    <div style="display:grid; grid-template-columns:1.1fr 0.9fr; gap:6px;">
                        <div>
                            <div class="bt-lbl">2. ВИД ЗОНЫ (ДЕТАЛЬНЫЙ)</div>
                            ${makeSelect('f-zone30', ['Разворотная зона', 'Сильный покупатель', 'Средний покупатель', 'Красивый покупатель', 'Сильный продавец', 'Средний продавец', 'Красивый продавец', 'Зона входа', 'Проторговка', 'Зона 15м 0.5', 'Зона 1Ч 0.5', 'Зона 4Ч 0.5', 'Зона Д 0.5'], 'Разворотная зона')}
                        </div>
                        <div>
							<div class="bt-lbl">СВЕЧА ВХОДА ${ind.candleColor ? BOLT : ''}</div>
                            ${makeCandleColorSelect(defCandleCol)}
                        </div>
                    </div>

                    <div style="display:grid; grid-template-columns:1fr 1.25fr 1.1fr; gap:6px;">
                        <div>
                            <div class="bt-lbl">ТОПЛИВО</div>
                            ${makeChips('f-fuel30', [{val:'да', label:'Да'}, {val:'нет', label:'Нет'}], 'да', 2)}
                        </div>
                        <div>
                            <div class="bt-lbl">КАСАНИЕ ЗОНЫ</div>
                            ${makeChips('f-touch30', [{val:'1', label:'1'}, {val:'2', label:'2'}, {val:'3 и больше', label:'3+'}], '1', 3)}
                        </div>
                        <div>
                            <div class="bt-lbl">ПРОТИВ ПОТОКА</div>
                            ${makeChips('f-against30', [{val:'нет', label:'Нет'}, {val:'да', label:'Да', tone:'red'}], 'нет', 2)}
                        </div>
                    </div>

                    <div style="display:grid; grid-template-columns:1.15fr 1fr; gap:6px; align-items:start;">
                        <div>
                            <div class="bt-lbl">ЧТО СНЯТО ПЕРЕД ВХОДОМ</div>
                            ${makeSelect('f-swept', ['приманка', 'полка', 'уровень 15м', 'уровень 1Ч', 'уровень 4Ч', 'уровень Д', 'склейка 15м/1Ч', 'склейка Д · 4Ч', 'ничего не снято'], 'приманка', '0px')}
                        </div>
                        <div>
                            <div class="bt-lbl">ЦЕНА УРОВНЯ СНЯТИЯ ${BOLT}</div>
                            <div style="display:flex; gap:3px; height:28px;">
                                <input type="text" id="f-swept-lvl" value="${t.entryExtreme ? numStr(t.entryExtreme) : ''}" placeholder="Цена..." style="width:100%; height:28px; box-sizing:border-box; background:#16161b; color:#f4f4f5; border:1px solid #282830; padding:0 6px; border-radius:5px; font-size:11px; outline:none;">
                                <button type="button" id="bt-fill-swept-lvl" title="Подставить экстремум свечи входа (${t.entryExtreme})" style="height:28px; background:#18181d; color:#f3d26b; border:1px solid #282830; border-radius:5px; padding:0 6px; cursor:pointer; font-size:11px;">↺</button>
                            </div>
                        </div>
                    </div>
                </div>

                <div style="display:grid; grid-template-columns:1.1fr 1fr 1fr; gap:6px;">
                    <div>
                        <div class="bt-lbl">ЧЕК-ЛИСТ 7/7 ${BOLT}</div>
                        ${makeChips('f-check30', [{val:'да', label:'Да', tone:'green'}, {val:'нет', label:'Нет', tone:'red'}], 'да', 2)}
                    </div>
                    <div>
                        <div class="bt-lbl">ДОШЛА ДО 1:1 ${BOLT}</div>
                        ${makeChips('f-hit1r', [{val:'да', label:'Да'}, {val:'нет', label:'Нет'}], t.reached1R ? 'да' : 'нет', 2)}
                    </div>
                    <div>
                        <div class="bt-lbl">ДОШЛА ДО 1:3 ${BOLT}</div>
                        ${makeChips('f-hit3r', [{val:'да', label:'Да'}, {val:'нет', label:'Нет'}], t.reached3R ? 'да' : 'нет', 2)}
                    </div>
                </div>
            `;
        }

        if (has('liq') || has('conf') || has('react')) {
            html += `<div style="display:grid; grid-template-columns:1.25fr 1.75fr; gap:7px;">`;
            if (has('liq')) {
                html += `<div><div class="bt-lbl">СНЯТИЕ ЛИКВИДНОСТИ</div>${makeChips('f-liq', [
                    { val: 'да', label: 'Да' },
                    { val: 'нет', label: 'Нет' }
                ], state.liq, 2)}</div>`;
            }
            if (has('conf')) {
                html += `<div><div class="bt-lbl">ПОДТВЕРЖДЕНИЕ <span style="color:#71717a;">(Z–B)</span></div>${makeSelect('f-conf', [
                    { val: 'Слом структуры', label: 'Слом структуры [Z]' },
                    { val: 'Реакция от зоны', label: 'Реакция от зоны [X]' },
                    { val: 'Поглощение', label: 'Поглощение [C]' },
                    { val: 'Снятие ликвидности', label: 'Снятие ликвидности [V]' },
                    { val: 'Без подтверждения', label: 'Без подтверждения [B]' }
                ], state.conf)}</div>`;
            } else if (has('react')) {
                html += `<div><div class="bt-lbl">РЕАКЦИЯ ОТ ЗОНЫ</div>${makeChips('f-react', [
                    { val: 'да', label: 'Да' },
                    { val: 'нет', label: 'Нет' }
                ], 'да', 2)}</div>`;
            }
            html += `</div>`;
        }

        if (has('res_30')) {
            const defRes30 = t.autoResult === 'Тейк' ? 'цель' : (t.autoResult === 'Стоп' ? 'стоп' : 'безубыток');
            html += `
                <div class="bt-lbl">ИСХОД СДЕЛКИ ${BOLT} <span style="color:#71717a; font-weight:500;">(Q / W / E / R)</span></div>
                ${makeChips('f-res30', [
                    { val: 'цель', label: 'Цель', kbd: 'Q', tone: 'green' },
                    { val: 'стоп', label: 'Стоп', kbd: 'W', tone: 'red' },
                    { val: 'безубыток', label: 'БУ', kbd: 'E' },
                    { val: 'частичная фиксация', label: 'Частич.' },
                    { val: 'закрыл руками', label: 'Руками', kbd: 'R' }
                ], defRes30, 5)}
                <div id="bt-manual-r-wrap" style="display:none; justify-content:space-between; align-items:center; background:#141418; padding:5px 8px; border-radius:5px; border:1px solid #3d3522; margin-bottom:6px;">
                    <span style="font-size:10.5px; color:#f3d26b;">Фактический итог в R (например 1,4):</span>
                    <input type="text" id="f-manual-r30" placeholder="0,0" style="width:60px; background:#0d0d10; color:#ffffff; border:1px solid #4a3f23; border-radius:4px; padding:3px 6px; font-size:11px; text-align:center; outline:none;">
                </div>
            `;
        } else if (has('result') || has('res_off') || has('r_out')) {
            html += `<div class="bt-lbl">РЕЗУЛЬТАТ СДЕЛКИ ${BOLT}</div>` +
                makeChips('f-res', [
                    { val: 'Тейк', label: 'Тейк', kbd: 'Q', tone: 'green' },
                    { val: 'Стоп', label: 'Стоп', kbd: 'W', tone: 'red' },
                    { val: 'Безубыток', label: 'БУ', kbd: 'E' },
                    { val: 'Пропуск', label: 'Пропуск', kbd: 'R' }
                ], t.autoResult, 4);
        }

        if (has('error') || has('error_poi')) {
            const errList = has('error_poi')
                ? ['—', 'Без снятия', 'Без реакции', 'Один фактор', 'Область на M15', 'Вход по касанию', 'Зона отработана', 'R:R ниже нормы', 'Догонял цену', 'Другое']
                : ['—', 'Ранний вход', 'Нет подтверждения', 'Вход без снятия ликвидности', 'Стоп слишком узкий', 'Стоп слишком широкий', 'Ранний выход', 'Пересидел цель', 'Нарушен минимальный R:R', 'Вход вне сессии', 'Другое'];
            html += `
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <div class="bt-lbl">ПРИЧИНА ОШИБКИ ${BOLT}</div>
                    <span id="bt-err-reason" style="font-size:10px; color:#f3d26b; font-weight:600; margin-bottom:4px;"></span>
                </div>
                ${makeSelect('f-err', errList, '—')}
            `;
        }

        dynFieldsEl.innerHTML = html;

        const zone5Btn = dynFieldsEl.querySelector('#f-zone .bt-chip:nth-child(5)');
        if (zone5Btn) zone5Btn.style.gridColumn = 'span 2';

        const form5Btn = dynFieldsEl.querySelector('#f-form-m1 .bt-chip:nth-child(5)');
        if (form5Btn) form5Btn.style.gridColumn = 'span 2';

        document.getElementById('bt-extra-inputs').style.display =
            (has('comment') || has('comment_30') || has('screenshot')) ? 'flex' : 'none';

        dynFieldsEl.querySelectorAll('.bt-flow-pill').forEach(pill => {
            pill.onclick = () => {
                const order = ['вверх', 'вниз', '—'];
                const curr = pill.dataset.dir || 'вверх';
                const next = order[(order.indexOf(curr) + 1) % order.length];
                pill.dataset.dir = next;
                pill.querySelector('.bt-flow-val').innerText = next === 'вверх' ? '↑ вверх' : (next === 'вниз' ? '↓ вниз' : '—');
                if (pill.id === 'f-fl-d') saveState({ flowD: next });
                if (pill.id === 'f-fl-4h') saveState({ flow4H: next });
                if (pill.id === 'f-fl-1h') saveState({ flow1H: next });
                if (pill.id === 'f-fl-15m') saveState({ flow15M: next });
            };
        });

		const ccBtn = document.getElementById('bt-cc-btn');
        const ccMenu = document.getElementById('bt-cc-menu');
        if (ccBtn && ccMenu) {
            ccBtn.onclick = e => {
                e.stopPropagation();
                ccMenu.style.display = ccMenu.style.display === 'block' ? 'none' : 'block';
            };
            ccMenu.querySelectorAll('.bt-cc-item').forEach(item => {
                item.onclick = () => {
                    setCandleColVal(item.dataset.val);
                    ccMenu.style.display = 'none';
                };
            });
        }

        dynFieldsEl.querySelectorAll('.bt-struct-preset').forEach(btn => {
            btn.onclick = () => {
                const sel = document.getElementById('f-struct');
                if (sel) sel.value = btn.dataset.st;
            };
        });

        const fillSweptBtn = document.getElementById('bt-fill-swept-lvl');
        if (fillSweptBtn) {
            fillSweptBtn.onclick = () => {
                const inp = document.getElementById('f-swept-lvl');
                if (inp && currentTrade?.entryExtreme) inp.value = numStr(currentTrade.entryExtreme);
            };
        }

        const sweptSel = document.getElementById('f-swept');
        if (sweptSel) {
            sweptSel.addEventListener('change', () => sync30TradesChecklist());
        }

        dynFieldsEl.querySelectorAll('.bt-group').forEach(grp => {
            grp.addEventListener('click', e => {
                const chip = e.target.closest('.bt-chip');
                if (!chip) return;
                grp.querySelectorAll('.bt-chip').forEach(b => b.classList.remove('active'));
                chip.classList.add('active');

                const val = chip.dataset.val;
                if (grp.id === 'f-trend') saveState({ trend: val });
                if (grp.id === 'f-zone') saveState({ zone: val });
                if (grp.id === 'f-fuel30' || grp.id === 'f-touch30' || grp.id === 'f-against30') {
                    sync30TradesChecklist();
                }
                if (grp.id === 'f-res30') {
                    const manWrap = document.getElementById('bt-manual-r-wrap');
                    if (manWrap) {
                        manWrap.style.display = (val === 'частичная фиксация' || val === 'закрыл руками') ? 'flex' : 'none';
                    }
                }
                if (grp.id === 'f-liq') {
                    saveState({ liq: val });
                    if (val === 'нет' && getChipVal('f-res') === 'Пропуск') {
                        const confEl = document.getElementById('f-conf');
                        if (confEl) confEl.value = 'Без подтверждения';
                    }
                    autoSuggestError();
                }
                if (grp.id === 'f-res') {
                    if (val === 'Пропуск' && getChipVal('f-liq') === 'нет') {
                        const confEl = document.getElementById('f-conf');
                        if (confEl) confEl.value = 'Без подтверждения';
                    }
                    autoSuggestError();
                }
            });
        });

        const confSel = document.getElementById('f-conf');
        if (confSel) {
            confSel.addEventListener('change', e => {
                saveState({ conf: e.target.value });
                autoSuggestError();
            });
        }

        autoSuggestError();
    }

    function syncLiveChartChanges() {
        if (modalEl.style.display !== 'block') return;

        const prevId = currentTrade?.id;
        const prevSig = currentTrade?.signature;

        if (!scanChartTrades(true) || !currentTrade) return;

        if (currentTrade.id !== prevId) {
            manualSessionOverride = null;
            if (commentInpEl) commentInpEl.value = '';
            renderMainFields();
            updateHeaderUI();
            pulseTradeCard('#089981');
            return;
        }

        if (currentTrade.signature !== prevSig) {
            setChipVal('f-res', currentTrade.autoResult);
            const defRes30 = currentTrade.autoResult === 'Тейк' ? 'цель' : (currentTrade.autoResult === 'Стоп' ? 'стоп' : 'безубыток');
            setChipVal('f-res30', defRes30);
            setChipVal('f-hit1r', currentTrade.reached1R ? 'да' : 'нет');
            setChipVal('f-hit3r', currentTrade.reached3R ? 'да' : 'нет');
			if (document.getElementById('f-candle-col')) {
                const isLong = currentTrade.direction === 'LONG';
                setCandleColVal(currentTrade.ind?.candleColor || (isLong ? 'Лонг 3' : 'Шорт 3'));
            }
            if (currentTrade.entryExtreme && document.getElementById('f-swept-lvl')) {
                document.getElementById('f-swept-lvl').value = numStr(currentTrade.entryExtreme);
            }
            updateHeaderUI();
            autoSuggestError();
            pulseTradeCard('#e5c158');
        }
    }

    setInterval(() => {
        const model = getTvModel();
        if (!model) return;

        const selId = detectCurrentlySelectedToolId(model);
        if (selId !== lastChartSelectedId) {
            lastChartSelectedId = selId;
            if (selId) lastSelectedToolId = selId;
        }

        try {
            const hov = typeof model.hoveredSource === 'function' ? model.hoveredSource() : model._hoveredSource;
            if (hov && isSupportedTool(hov)) lastHoveredToolId = getToolId(hov);
        } catch (e) {}

        if (modalEl.style.display === 'block') {
            syncLiveChartChanges();
        }
    }, 250);

    window.addEventListener('pointerup', e => {
        if (modalEl.style.display !== 'block') return;
        if (modalEl.contains(e.target)) return;
        setTimeout(syncLiveChartChanges, 30);
    }, true);

    function renderSettingsDrawer() {
        document.getElementById('bt-preset-select').value = state.presetKey;
        document.getElementById('bt-custom-paste-col').value = getActivePasteCol();
        document.getElementById('bt-custom-start-row').value = getActiveStartRow();

        const colsWrap = document.getElementById('bt-cols-editor-wrap');
        const toggleColsBtn = document.getElementById('bt-toggle-cols-btn');
        const shouldShowCols = showColumnEditor || state.presetKey === 'custom';
        colsWrap.style.display = shouldShowCols ? 'block' : 'none';
        toggleColsBtn.innerText = shouldShowCols
            ? '▲ Скрыть конструктор переменных'
            : `🔧 Конструктор переменных и столбцов (${getActiveColumns().length} шт.)`;

        const cols = getActiveColumns();
        const baseColIdx = colLetterToIndex(getActivePasteCol());
        const pillsContainer = document.getElementById('bt-schema-pills');

        pillsContainer.innerHTML = cols.map((tok, idx) => {
            const colLet = colIndexToLetter(baseColIdx + idx);
            return `
                <div class="bt-drag-pill" data-idx="${idx}" data-tok="${tok}" title="Столбец ${colLet} в таблице. Перетащите мышкой для смены порядка">
                    <span class="bt-pill-num" style="color:#e5c158; font-weight:700;">${colLet}</span>
                    <span>${TOKEN_LABELS[tok] || tok}</span>
                    <span class="bt-del-col" data-del="${idx}" title="Удалить переменную">✕</span>
                </div>
            `;
        }).join('');

        pillsContainer.querySelectorAll('.bt-del-col').forEach(delBtn => {
            delBtn.addEventListener('mousedown', e => e.stopPropagation());
            delBtn.addEventListener('click', e => {
                e.stopPropagation();
                const i = parseInt(delBtn.dataset.del, 10);
                const updated = getActiveColumns();
                updated.splice(i, 1);
                saveState({
                    presetKey: 'custom',
                    customColumns: updated,
                    customPasteCol: getActivePasteCol(),
                    customStartRow: getActiveStartRow()
                });
                renderMainFields();
                renderSettingsDrawer();
                updateHeaderUI();
            });
        });

        pillsContainer.querySelectorAll('.bt-drag-pill').forEach(pill => {
            pill.addEventListener('mousedown', e => {
                if (e.button !== 0 || e.target.classList.contains('bt-del-col')) return;
                e.preventDefault();
                e.stopPropagation();

                const activePill = pill;
                activePill.classList.add('dragging-active');

                const updateColLetters = () => {
                    const bIdx = colLetterToIndex(getActivePasteCol());
                    Array.from(pillsContainer.children).forEach((el, i) => {
                        const numSpan = el.querySelector('.bt-pill-num');
                        if (numSpan) numSpan.innerText = colIndexToLetter(bIdx + i);
                    });
                };

                const onMouseMove = moveEvt => {
                    const targetEl = document.elementFromPoint(moveEvt.clientX, moveEvt.clientY);
                    const overPill = targetEl?.closest('.bt-drag-pill');
                    if (overPill && overPill !== activePill && pillsContainer.contains(overPill)) {
                        const allPills = Array.from(pillsContainer.children);
                        const activeIdx = allPills.indexOf(activePill);
                        const overIdx = allPills.indexOf(overPill);
                        if (activeIdx < overIdx) {
                            pillsContainer.insertBefore(activePill, overPill.nextSibling);
                        } else {
                            pillsContainer.insertBefore(activePill, overPill);
                        }
                        updateColLetters();
                    }
                };

                const onMouseUp = () => {
                    window.removeEventListener('mousemove', onMouseMove, true);
                    window.removeEventListener('mouseup', onMouseUp, true);
                    activePill.classList.remove('dragging-active');

                    const newOrder = Array.from(pillsContainer.querySelectorAll('.bt-drag-pill'))
                        .map(el => el.dataset.tok);

                    saveState({
                        presetKey: 'custom',
                        customColumns: newOrder,
                        customPasteCol: getActivePasteCol(),
                        customStartRow: getActiveStartRow()
                    });
                    renderMainFields();
                    renderSettingsDrawer();
                    updateHeaderUI();
                };

                window.addEventListener('mousemove', onMouseMove, true);
                window.addEventListener('mouseup', onMouseUp, true);
            });
        });
    }

    function computeBatchStats() {
        const bucket = getTplBucket();
        let wins = 0, losses = 0, totalR = 0;
        for (const item of bucket.batchQueue) {
            if (!item || typeof item !== 'object') continue;
            if (item.result === 'Тейк' || item.result === 'цель') {
                wins++;
                totalR += Number(item.rr || 0);
            } else if (item.result === 'Стоп' || item.result === 'стоп') {
                losses++;
                totalR -= 1;
            }
        }
        const closed = wins + losses;
        const wr = closed > 0 ? Math.round((wins / closed) * 100) : 0;
        const sign = totalR > 0 ? '+' : '';
        return `Серия (${PRESET_SCHEMAS[state.presetKey]?.name || 'Свой журнал'}): ${bucket.batchQueue.length} сд. | WR: ${wr}% (${wins}Т/${losses}С) | ${sign}${totalR.toFixed(2)} R`;
    }

    function updateHeaderUI() {
        const schemaObj = PRESET_SCHEMAS[state.presetKey];
        const bucket = getTplBucket();

        document.getElementById('bt-active-tpl-title').innerText =
            schemaObj ? schemaObj.name : `Свой журнал (${getActiveColumns().length} перем.)`;

        collapsibleBodyEl.style.display = state.minimized ? 'none' : 'block';

        document.querySelectorAll('.bt-prof-btn').forEach(btn => {
            btn.classList.toggle('active', btn.dataset.prof === state.outputProfile);
        });

        const isReEdit = currentTrade && bucket.lastSavedSig === currentTrade.signature && bucket.batchQueue.length > 0;
        const pasteCol = getActivePasteCol();

        mainBtnText.innerText = isReEdit
            ? `Обновить сделку (в стлб. ${pasteCol})`
            : `Скопировать строку (в стлб. ${pasteCol})`;

        document.getElementById('bt-batch-count').innerText = bucket.batchQueue.length;
        document.getElementById('bt-copy-batch-btn').title =
            `${computeBatchStats()}\nКликните, чтобы скопировать всю пачку (вставьте в любую пустую строку столбца ${pasteCol})`;

        updateScreenshotPillUI();

        if (!currentTrade) return;

        const total = allChartTrades.length || 1;
        document.getElementById('bt-trade-counter').innerText = `${activeTradeIdx + 1}/${total}`;

        document.getElementById('bt-info-sym').innerText = `${currentTrade.symbol} · ${currentTrade.timeframeLat}`;
        sessionPillBtn.innerText = (manualSessionOverride ? '' : '⚡ ') + getActiveSession();
        document.getElementById('bt-info-date').innerText = `${currentTrade.shortDate} · ${currentTrade.mskTime} МСК`;
        document.getElementById('bt-info-date').title = `Дата входа: ${currentTrade.candleDate} (${currentTrade.mskTime} МСК)`;

        const badge = document.getElementById('bt-info-badge');
        badge.innerText = `${currentTrade.direction} · ${numStr(currentTrade.riskReward)}R`;
        badge.style.background = currentTrade.direction === 'LONG' ? '#089981' : '#f23645';

        document.getElementById('bt-info-en').innerText = numStr(currentTrade.entryPrice);
        document.getElementById('bt-info-st').innerText = numStr(currentTrade.stopPrice);
        document.getElementById('bt-info-tg').innerText = numStr(currentTrade.targetPrice);
    }

    function buildAndCopyRow() {
        if (!currentTrade) return;
        const cols = getActiveColumns();
        const t = currentTrade;
        const isLong = t.direction === 'LONG';
        const hasRes30 = cols.includes('res_30');
        const pasteCol = getActivePasteCol();
        const startRow = getActiveStartRow();

        const resStd = getChipVal('f-res', t.autoResult);
        const res30Val = getChipVal('f-res30', t.autoResult === 'Тейк' ? 'цель' : (t.autoResult === 'Стоп' ? 'стоп' : 'безубыток'));
        const activeResult = hasRes30 ? res30Val : resStd;

        const bucket = getTplBucket();
        const isReEdit = bucket.lastSavedSig === t.signature && bucket.batchQueue.length > 0;

        const manualRStr = cleanText(document.getElementById('f-manual-r30')?.value || '');
        const manualRNum = manualRStr ? parseFloat(manualRStr.replace(',', '.')) : NaN;

        let rOut = '';
        let newCum = bucket.cumR;
        let newMax = bucket.maxCumR;

        if (activeResult === 'Тейк' || activeResult === 'цель') {
            rOut = t.riskReward;
            newCum = Number((bucket.cumR + rOut).toFixed(2));
        } else if (activeResult === 'Стоп' || activeResult === 'стоп') {
            rOut = -1;
            newCum = Number((bucket.cumR - 1).toFixed(2));
        } else if (activeResult === 'Безубыток' || activeResult === 'безубыток') {
            rOut = 0;
        } else if (!isNaN(manualRNum)) {
            rOut = manualRNum;
            newCum = Number((bucket.cumR + rOut).toFixed(2));
        }

        if (newCum > newMax) newMax = newCum;
        const dd = Number(Math.max(0, newMax - newCum).toFixed(2));

        const aoiInp = document.getElementById('f-aoi-num');
        const aoiNum = aoiInp ? (parseInt(aoiInp.value, 10) || 1) : state.aoiNum;

        const flD = getFlowVal('f-fl-d', state.flowD);
        const fl4H = getFlowVal('f-fl-4h', state.flow4H);
        const fl1H = getFlowVal('f-fl-1h', state.flow1H);
        const fl15M = getFlowVal('f-fl-15m', state.flow15M);

        const f1 = getChipVal('f-f1', 'да'), f2 = getChipVal('f-f2', 'да'), f3 = getChipVal('f-f3', 'нет'), f4 = getChipVal('f-f4', 'да');
        const factorsCnt = [f1, f2, f3, f4].filter(x => x === 'да').length;

        const stopUsd = Math.max(0.01, t.riskPts);
        const lots30 = Math.floor((30 / stopUsd / 100) * 100) / 100;
        const profitUsd = rOut === '' ? '' : Number((Number(rOut) * lots30 * 100 * stopUsd).toFixed(2));
        const lotsOff = Number(((10000 * 0.003) / (stopUsd / 0.01 * 1)).toFixed(2));
        const rOutOff = resStd === 'Стоп' ? -1 : (resStd === 'Тейк' ? Number((0.5 * 3 + 0.5 * t.riskReward).toFixed(2)) : 0);

        const comment = cleanText(commentInpEl?.value || '');
        const chk30 = getChipVal('f-check30', 'да');
        const img = cleanText(screenshotUrl);
        const sweptVal = document.getElementById('f-swept')?.value || 'приманка';
        const sweptLvlRaw = cleanText(document.getElementById('f-swept-lvl')?.value || '');
        const sweptLvlFormatted = (sweptVal === 'ничего не снято' || !sweptLvlRaw) ? '' : numStr(sweptLvlRaw);

        const schemaStyle = PRESET_SCHEMAS[state.presetKey] || PRESET_SCHEMAS.bt_main;
        const f = (tok, fallback) => buildDynamicFormula(tok, cols, pasteCol, startRow, fallback);

        const tokenValues = {
            symbol: t.symbol,
            date: t.candleDate,
            time_msk: t.mskTime,
            time_slot: t.officeTimeSlot,
            session: getActiveSession(),
            trend: getChipVal('f-trend', isLong ? 'Тренд вверх' : 'Тренд вниз'),
            trend_h4: getChipVal('f-h4-trend', isLong ? 'Вверх' : 'Вниз'),
            tf: schemaStyle.tfStyle === 'cyr' ? t.timeframeCyr : t.timeframeLat,
            dir_lower: isLong ? 'лонг' : 'шорт',
            dir_title: isLong ? 'Лонг' : 'Шорт',
            zone: getChipVal('f-zone', state.zone),
            zone_30: document.getElementById('f-zone30')?.value || 'Разворотная зона',
            poi_type: getChipVal('f-poi-type', 'OB'),
            form_m1: getChipVal('f-form-m1', 'SCOB'),
            liq: getChipVal('f-liq', state.liq),
            react: getChipVal('f-react', 'да'),
            conf: document.getElementById('f-conf')?.value || state.conf,
            entry: numStr(t.entryPrice),
            stop: numStr(t.stopPrice),
            target: numStr(t.targetPrice),
            upper_bound: numStr(t.upperBound),
            lower_bound: numStr(t.lowerBound),
            risk_pts: f('risk_pts', numStr(t.riskPts)),
            pot_pts: f('pot_pts', numStr(t.potPts)),
            rr: f('rr', numStr(t.riskReward)),
            lots_30: f('lots_30', numStr(lots30)),
            lots_off: f('lots_off', numStr(lotsOff)),
            result: resStd,
            res_30: res30Val,
            manual_r_30: (res30Val === 'частичная фиксация' || res30Val === 'закрыл руками') && !isNaN(manualRNum) ? numStr(manualRNum) : '',
            res_off: resStd === 'Безубыток' ? 'Сейф/БУ' : resStd,
            r_out: f('r_out', rOut === '' ? '' : numStr(rOut)),
            r_out_off: f('r_out_off', numStr(rOutOff)),
            profit_usd: f('profit_usd', profitUsd === '' ? '' : numStr(profitUsd)),
            prof_pct_off: f('prof_pct_off', numStr(Number((rOutOff * 0.003).toFixed(4)))),
            move_pts: f('move_pts', numStr(resStd === 'Тейк' ? t.potPts : (resStd === 'Стоп' ? -t.riskPts : 0))),
            cum_r: f('cum_r', resStd === 'Пропуск' ? numStr(bucket.cumR) : numStr(newCum)),
            dd_r: f('dd_r', resStd === 'Пропуск' ? numStr(0) : numStr(dd)),
            hit_1r: getChipVal('f-hit1r', t.reached1R ? 'да' : 'нет'),
            hit_3r: getChipVal('f-hit3r', t.reached3R ? 'да' : 'нет'),
            error: document.getElementById('f-err')?.value || '—',
            error_poi: document.getElementById('f-err')?.value || '—',
            comment: comment,
            comment_30: comment,
            screenshot: img,
            aoi_num: aoiNum,
            flow_d: flD,
            flow_4h: fl4H,
            flow_1h: fl1H,
            flow_15m: fl15M,
            struct_30: document.getElementById('f-struct')?.value || (isLong ? 'слом ↑' : 'слом ↓'),
            swept_30: sweptVal,
            swept_level_30: sweptLvlFormatted,
            candle_col: document.getElementById('f-candle-col')?.value || (isLong ? 'Лонг 3' : 'Шорт 3'),
            against_flow: getChipVal('f-against30', 'нет'),
            fuel_30: getChipVal('f-fuel30', 'да'),
            touch_30: getChipVal('f-touch30', '1'),
            check_30: chk30,
            tf_ok: 'да',
            emotion: 'Спокоен',
            safe_pct: numStr(0.5),
            safe_rr: numStr(3),
            risk_pct_col: numStr(0.003),
            f_ret: f1,
            f_ob: f2,
            f_fvg: f3,
            f_liq: f4,
            factors_cnt: f('factors_cnt', factorsCnt),
            valid_aoi: f('valid_aoi', factorsCnt >= 2 ? 'ДА' : 'нет'),
            reached: getChipVal('f-reached', 'да'),
            aoi_res: getChipVal('f-aoi-res', 'Отбой'),
            empty: ''
        };

        const rowArr = cols.map(tok => tokenValues[tok] ?? '');
        const tsvString = rowArr.join('\t');

        navigator.clipboard.writeText(tsvString).then(() => {
            const newBatch = [...bucket.batchQueue];
            const batchItem = { tsv: tsvString, result: activeResult, rr: t.riskReward };
            if (isReEdit && newBatch.length > 0) {
                newBatch[newBatch.length - 1] = batchItem;
            } else {
                newBatch.push(batchItem);
            }

            saveTplBucket({
                batchQueue: newBatch,
                lastSavedSig: t.signature,
                cumR: newCum,
                maxCumR: newMax
            });

            saveState({
                aoiNum: cols.includes('valid_aoi') ? aoiNum + 1 : aoiNum,
                flowD: flD, flow4H: fl4H, flow1H: fl1H, flow15M: fl15M
            });

            if (commentInpEl) commentInpEl.value = '';
            if (commentMenuEl) commentMenuEl.style.display = 'none';
            screenshotUrl = '';
            manualSessionOverride = null;
            updateHeaderUI();
            modalEl.style.display = 'none';
            showToast(`✓ Скопировано! Вставьте (Ctrl+V) в любую строку столбца ${pasteCol}`);
        }).catch(() => {
            showToast('Ошибка доступа к буферу обмена', 'error');
        });
    }

    document.querySelectorAll('.bt-prof-btn').forEach(btn => {
        btn.onclick = () => {
            const prof = btn.dataset.prof;
            const sep = prof === 'en_form' ? '.' : ',';
            saveState({ outputProfile: prof, decimalSep: sep });
            updateHeaderUI();
            const labels = {
                ru_form: '🇷🇺 RU (Запятая + =ЕСЛИ)',
                en_form: '🌐 EN (Точка + =IF)'
            };
            showToast(`Включен профиль: ${labels[prof]}`);
        };
    });

    document.getElementById('bt-custom-paste-col').onchange = e => {
        saveState({
            presetKey: 'custom',
            customColumns: getActiveColumns(),
            customPasteCol: e.target.value
        });
        renderSettingsDrawer();
        updateHeaderUI();
    };

    document.getElementById('bt-custom-start-row').onchange = e => {
        const r = Math.max(1, parseInt(e.target.value, 10) || 2);
        saveState({
            presetKey: 'custom',
            customColumns: getActiveColumns(),
            customStartRow: r
        });
        renderSettingsDrawer();
    };

    sessionPillBtn.onclick = () => {
        const isEuUs = (PRESET_SCHEMAS[state.presetKey]?.sessionStyle === 'eu_us');
        const list = isEuUs ? ['Азия', 'Европа', 'Америка'] : ['Азия', 'Лондон', 'Нью-Йорк'];
        const curr = getActiveSession();
        const next = list[(list.indexOf(curr) + 1) % list.length];
        manualSessionOverride = next;
        updateHeaderUI();
    };

    screenPillBtn.onclick = () => {
        if (screenshotUrl) {
            screenshotUrl = '';
            updateScreenshotPillUI();
            showToast('Ссылка на скриншот удалена');
        } else {
            showToast('Нажмите Alt+S на графике, а затем Ctrl+V прямо в окне!');
        }
    };

    document.getElementById('bt-min-btn').onclick = () => {
        saveState({ minimized: !state.minimized });
        updateHeaderUI();
    };

    openSettingsBtn.onclick = () => toggleSettingsView();
    document.getElementById('bt-close-settings-btn').onclick = () => toggleSettingsView(false);

    document.getElementById('bt-toggle-cols-btn').onclick = () => {
        showColumnEditor = !showColumnEditor;
        renderSettingsDrawer();
    };

    document.getElementById('bt-preset-select').onchange = e => {
        const val = e.target.value;
        showColumnEditor = (val === 'custom');
        saveState({
            presetKey: val,
            customColumns: val === 'custom' ? getActiveColumns() : null
        });
        manualSessionOverride = null;
        renderMainFields();
        renderSettingsDrawer();
        updateHeaderUI();
        if (val !== 'custom') {
            toggleSettingsView(false);
        }
    };

    document.getElementById('bt-reset-schema').onclick = () => {
        const baseKey = state.presetKey === 'custom' ? 'bt_main' : state.presetKey;
        showColumnEditor = false;
        saveState({ presetKey: baseKey, customColumns: null });
        renderMainFields();
        renderSettingsDrawer();
        updateHeaderUI();
        showToast('Схема переменных возвращена к шаблону');
    };

    document.getElementById('bt-add-token-select').onchange = e => {
        const tok = e.target.value;
        if (!tok) return;
        const updated = [...getActiveColumns(), tok];
        saveState({
            presetKey: 'custom',
            customColumns: updated,
            customPasteCol: getActivePasteCol(),
            customStartRow: getActiveStartRow()
        });
        e.target.value = '';
        renderMainFields();
        renderSettingsDrawer();
        updateHeaderUI();
    };

    document.getElementById('bt-copy-batch-btn').onclick = () => {
        const bucket = getTplBucket();
        if (!bucket.batchQueue.length) {
            showToast('Пачка сделок для этого журнала пока пуста', 'error');
            return;
        }
        const allTsv = bucket.batchQueue.map(x => x.tsv).join('\n');
        const pasteCol = getActivePasteCol();
        navigator.clipboard.writeText(allTsv).then(() => {
            showToast(`✓ Скопировано сделок: ${bucket.batchQueue.length}! Вставьте в любую пустую строку столбца ${pasteCol} (${computeBatchStats()})`);
        });
    };

    document.getElementById('bt-clear-batch-btn').onclick = () => {
        saveTplBucket({ batchQueue: [], cumR: 0, maxCumR: 0, lastSavedSig: '' });
        updateHeaderUI();
        showToast('Пачка сделок и счетчик серии очищены');
    };

    function switchChartTrade(delta) {
        if (!allChartTrades.length) return;
        activeTradeIdx = (activeTradeIdx + delta + allChartTrades.length) % allChartTrades.length;
        currentTrade = allChartTrades[activeTradeIdx];
        lastSelectedToolId = currentTrade.id;
        lastChartSelectedId = detectCurrentlySelectedToolId(getTvModel());
        manualSessionOverride = null;
        if (commentInpEl) commentInpEl.value = '';
        if (commentMenuEl) commentMenuEl.style.display = 'none';
        renderMainFields();
        updateHeaderUI();
        pulseTradeCard('#089981');
    }

    document.getElementById('bt-prev-trade').onclick = () => switchChartTrade(-1);
    document.getElementById('bt-next-trade').onclick = () => switchChartTrade(1);

    document.getElementById('bt-quick-copy-prices').onclick = () => {
        if (!currentTrade) return;
        const tsv = [numStr(currentTrade.entryPrice), numStr(currentTrade.stopPrice), numStr(currentTrade.targetPrice)].join('\t');
        navigator.clipboard.writeText(tsv).then(() => {
            showToast('✓ Скопированы 3 цены: Вход + Стоп + Цель!');
        });
    };

    document.getElementById('bt-clear-comment').onclick = () => {
        commentInpEl.value = '';
        commentInpEl.focus();
    };

    document.getElementById('bt-toggle-presets').onclick = e => {
        e.stopPropagation();
        commentMenuEl.style.display = commentMenuEl.style.display === 'block' ? 'none' : 'block';
    };

    commentMenuEl.querySelectorAll('.bt-cm-item').forEach(item => {
        item.onclick = () => {
            commentInpEl.value = item.innerText;
            commentMenuEl.style.display = 'none';
        };
    });

	modalEl.addEventListener('click', e => {
        if (!e.target.closest('#bt-comment-menu') && !e.target.closest('#bt-toggle-presets')) {
            commentMenuEl.style.display = 'none';
        }
        const ccMenu = document.getElementById('bt-cc-menu');
        if (ccMenu && !e.target.closest('#bt-cc-wrap')) {
            ccMenu.style.display = 'none';
        }
    });

    mainCopyBtn.onclick = buildAndCopyRow;
    document.getElementById('bt-close-x').onclick = () => { modalEl.style.display = 'none'; };
    document.getElementById('bt-fab-btn').onclick = () => openModal();

    (function initDrag() {
        const handle = document.getElementById('bt-header-drag');
        let dragging = false, sx = 0, sy = 0, it = 0, ir = 0;
        handle.addEventListener('mousedown', e => {
            if (e.target.closest('button')) return;
            dragging = true; sx = e.clientX; sy = e.clientY;
            const rect = modalEl.getBoundingClientRect();
            it = rect.top; ir = window.innerWidth - rect.right;
            e.preventDefault();
        });
        window.addEventListener('mousemove', e => {
            if (!dragging) return;
            modalEl.style.top = Math.max(10, it + (e.clientY - sy)) + 'px';
            modalEl.style.right = Math.max(10, ir - (e.clientX - sx)) + 'px';
        });
        window.addEventListener('mouseup', () => {
            if (!dragging) return;
            dragging = false;
            saveState({
                posTop: parseInt(modalEl.style.top, 10) || 20,
                posRight: parseInt(modalEl.style.right, 10) || 20
            });
        });
    })();

    function openModal() {
        lastChartSelectedId = null;
        if (!scanChartTrades(false)) return;
        manualSessionOverride = null;
        if (commentInpEl) commentInpEl.value = '';
        if (commentMenuEl) commentMenuEl.style.display = 'none';
        toggleSettingsView(false);
        renderMainFields();
        updateHeaderUI();
        if (document.activeElement && typeof document.activeElement.blur === 'function') {
            document.activeElement.blur();
        }
        modalEl.style.display = 'block';
        modalEl.focus();
    }

    window.addEventListener('paste', e => {
        if (modalEl.style.display !== 'block') return;
        const tag = document.activeElement?.tagName;
        if (tag === 'INPUT' || tag === 'TEXTAREA') return;
        const pasted = (e.clipboardData || window.clipboardData)?.getData('text') || '';
        if (/^https?:\/\//i.test(pasted.trim())) {
            e.preventDefault();
            screenshotUrl = pasted.trim();
            updateScreenshotPillUI();
            showToast('📸 Скриншот прикреплен ✓');
        }
    });

    function isOurInputFocused(target) {
        if (!target) return false;
        return modalEl.contains(target) && ['INPUT', 'TEXTAREA', 'SELECT'].includes(target.tagName);
    }

    window.addEventListener('keypress', e => {
        if (modalEl.style.display === 'block' && !isOurInputFocused(e.target)) {
            e.preventDefault();
            e.stopPropagation();
            e.stopImmediatePropagation();
        }
    }, true);

    window.addEventListener('keydown', e => {
        const isOpen = modalEl.style.display === 'block';
        const inOurInput = isOurInputFocused(e.target);

        const isQKey = e.code === 'KeyQ' || ['q', 'й', 'œ', 'ј'].includes((e.key || '').toLowerCase());
        const isOpenShortcut = e.key === 'F2' || e.code === 'F2' || e.keyCode === 113 || (isQKey && (e.ctrlKey || e.altKey) && !e.metaKey);

        if (isOpenShortcut && !inOurInput) {
            e.preventDefault();
            e.stopPropagation();
            e.stopImmediatePropagation();
            openModal();
            return;
        }

        if (!isOpen) return;

        if (e.key === 'Escape') {
            e.preventDefault();
            e.stopPropagation();
            e.stopImmediatePropagation();
            if (commentMenuEl && commentMenuEl.style.display === 'block') {
                commentMenuEl.style.display = 'none';
            } else if (inOurInput) {
                e.target.blur();
                modalEl.focus();
            } else if (isSettingsOpen) {
                toggleSettingsView(false);
            } else {
                modalEl.style.display = 'none';
            }
            return;
        }

        if (e.key === 'Enter' && !isSettingsOpen) {
            e.preventDefault();
            e.stopPropagation();
            e.stopImmediatePropagation();
            buildAndCopyRow();
            return;
        }

        if (e.key === 'Tab' && !inOurInput && !isSettingsOpen) {
            e.preventDefault();
            e.stopPropagation();
            e.stopImmediatePropagation();
            commentInpEl?.focus();
            return;
        }

        if (!inOurInput && !isSettingsOpen && !e.ctrlKey && !e.altKey && !e.metaKey) {
            if (e.key === 'ArrowLeft') {
                e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
                switchChartTrade(-1);
                return;
            }
            if (e.key === 'ArrowRight') {
                e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
                switchChartTrade(1);
                return;
            }

            const zoneMap = {
                '1': 'Ордер-блок',
                '2': 'Имбаланс',
                '3': 'Материнская свеча',
                '4': 'Зона ликвидности',
                '5': 'Правильный откат'
            };
            const poiMap = { '1': 'OB', '2': 'Имбаланс', '3': '50% МС' };
            const formMap = { '1': 'SCOB', '2': 'Свип', '3': 'Пин-бар (IFS)', '4': 'Поглощение', '5': 'Правильный откат' };

            const resMap = {
                'q': 'Тейк', 'й': 'Тейк',
                'w': 'Стоп', 'ц': 'Стоп',
                'e': 'Безубыток', 'у': 'Безубыток',
                'r': 'Пропуск', 'к': 'Пропуск'
            };
            const res30Map = {
                'q': 'цель', 'й': 'цель',
                'w': 'стоп', 'ц': 'стоп',
                'e': 'безубыток', 'у': 'безубыток',
                'r': 'закрыл руками', 'к': 'закрыл руками'
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

            const kl = (e.key || '').toLowerCase();

            if (zoneMap[e.key] || poiMap[e.key] || formMap[e.key]) {
                e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
                if (document.getElementById('f-zone') && zoneMap[e.key]) {
                    setChipVal('f-zone', zoneMap[e.key]);
                    saveState({ zone: zoneMap[e.key] });
                } else if (document.getElementById('f-poi-type') && poiMap[e.key]) {
                    setChipVal('f-poi-type', poiMap[e.key]);
                } else if (document.getElementById('f-form-m1') && formMap[e.key]) {
                    setChipVal('f-form-m1', formMap[e.key]);
                }
            } else if (resMap[kl]) {
                e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
                if (document.getElementById('f-res30')) {
                    const v30 = res30Map[kl];
                    setChipVal('f-res30', v30);
                    const manWrap = document.getElementById('bt-manual-r-wrap');
                    if (manWrap) manWrap.style.display = (v30 === 'закрыл руками') ? 'flex' : 'none';
                } else {
                    setChipVal('f-res', resMap[kl]);
                    if (resMap[kl] === 'Пропуск' && getChipVal('f-liq') === 'нет') {
                        const confEl = document.getElementById('f-conf');
                        if (confEl) confEl.value = 'Без подтверждения';
                    }
                    autoSuggestError();
                }
            } else if (trendMap[kl]) {
                e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
                setChipVal('f-trend', trendMap[kl]);
                saveState({ trend: trendMap[kl] });
            } else if (confMap[kl]) {
                e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
                const confEl = document.getElementById('f-conf');
                if (confEl) {
                    confEl.value = confMap[kl];
                    saveState({ conf: confMap[kl] });
                    autoSuggestError();
                }
            } else if (e.code === 'Space') {
                e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
                const nextLiq = getChipVal('f-liq', 'да') === 'да' ? 'нет' : 'да';
                setChipVal('f-liq', nextLiq);
                saveState({ liq: nextLiq });
                if (nextLiq === 'нет' && getChipVal('f-res') === 'Пропуск') {
                    const confEl = document.getElementById('f-conf');
                    if (confEl) confEl.value = 'Без подтверждения';
                }
                autoSuggestError();
            }
        }
    }, true);

})();
