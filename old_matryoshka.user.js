// ==UserScript==
// @name         Matryoshka Backtest Auto-Logger
// @namespace    http://tampermonkey.net/
// @version      3.8
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
  const BOLT = `<span class="bt-bolt" title="Заполнено автоматически по данным графика">⚡</span>`;

  const CANDLE_COLORS = {
    'Лонг 3': { hex: '#72f238', hint: 'салатовая' },
    'Лонг 2': { hex: '#4bb8d4', hint: 'голубая' },
    'Лонг 1': { hex: '#3b7dff', hint: 'синяя' },
    'Шорт 3': { hex: '#f7e84b', hint: 'жёлтая' },
    'Шорт 2': { hex: '#eb3b4b', hint: 'красная' },
    'Шорт 1': { hex: '#8a2bbd', hint: 'фиолетовая' }
  };

  const PRESET_SCHEMAS = {
    bt_main: {
      name: 'Бэктест «Матрёшка»', pasteCol: 'B', startRow: 6, sessionStyle: 'std', tfStyle: 'lat',
      columns: ['date', 'session', 'trend', 'tf', 'zone', 'liq', 'conf', 'entry', 'stop', 'target', 'risk_pts', 'pot_pts', 'rr', 'result', 'r_out', 'cum_r', 'dd_r', 'error', 'comment', 'screenshot']
    },
    bt_30: {
      name: 'Бэктест 30 сделок', pasteCol: 'B', startRow: 4, sessionStyle: 'eu_us', tfStyle: 'cyr',
      columns: ['date', 'time_msk', 'session', 'tf', 'dir_lower', 'flow_d', 'flow_4h', 'flow_1h', 'flow_15m', 'struct_30', 'zone_30', 'fuel_30', 'touch_30', 'against_flow', 'swept_30', 'swept_level_30', 'candle_col', 'entry', 'stop', 'target', 'risk_pts', 'rr', 'lots_30', 'check_30', 'res_30', 'manual_r_30', 'r_out', 'profit_usd', 'hit_1r', 'hit_3r', 'screenshot', 'comment_30']
    },
    dz_poi: {
      name: 'ДЗ: Журнал POI', pasteCol: 'B', startRow: 6, sessionStyle: 'std', tfStyle: 'lat',
      columns: ['aoi_num', 'date', 'poi_type', 'liq', 'react', 'rr', 'result', 'r_out', 'cum_r', 'error_poi', 'comment', 'screenshot']
    },
    dz_aoi: {
      name: 'ДЗ: Таблица AOI', pasteCol: 'B', startRow: 6, sessionStyle: 'std', tfStyle: 'lat',
      columns: ['aoi_num', 'date', 'trend_h4', 'f_ret', 'f_ob', 'f_fvg', 'f_liq', 'factors_cnt', 'valid_aoi', 'upper_bound', 'lower_bound', 'reached', 'aoi_res', 'comment', 'screenshot']
    },
    office: {
      name: 'Торговый офис (Недели 1–4)', pasteCol: 'B', startRow: 16, sessionStyle: 'std', tfStyle: 'lat',
      columns: ['time_slot', 'dir_title', 'tf_d1', 'tf_h4', 'tf_h1', 'tf_m15', 'form_m1', 'lots_off', 'entry', 'stop', 'target', 'safe_pct', 'safe_rr', 'risk_pct_col', 'rr', 'res_off', 'r_out_off', 'prof_pct_off', 'move_pts', 'tf_ok', 'emotion', 'screenshot', 'comment']
    },
    prices_only: {
      name: 'Только 3 цены (Вход/Стоп/Цель)', pasteCol: 'B', startRow: 2, sessionStyle: 'std', tfStyle: 'lat',
      columns: ['entry', 'stop', 'target']
    }
  };

  const TOKEN_LABELS = {
    symbol: 'Инструмент (Тикер)', date: 'Дата входа', time_msk: 'Время входа (строго МСК)', time_slot: 'Время входа (МСК / Лок.)',
    session: 'Торговая сессия', tf: 'Таймфрейм', dir_title: 'Направление (Лонг/Шорт)', dir_lower: 'Сторона (лонг/шорт)',
    trend: 'Фаза рынка', trend_h4: 'Тренд H4', flow_d: 'Поток D1 (вверх/вниз)', flow_4h: 'Поток H4 (вверх/вниз)',
    flow_1h: 'Поток H1 (вверх/вниз)', flow_15m: 'Поток M15 (вверх/вниз)', tf_d1: 'D1 по направлению (да/нет)',
    tf_h4: 'H4 по направлению (да/нет)', tf_h1: 'H1 по направлению (да/нет)', tf_m15: 'M15 по направлению (да/нет)',
    struct_30: 'Структура перед входом', zone: 'Зона входа', zone_30: 'Вид зоны', poi_type: 'Тип POI (OB / Имб / МС)',
    form_m1: 'Формация входа M1', candle_col: 'Свеча входа (цвет)', liq: 'Снятие ликвидности (да/нет)',
    swept_30: 'Что снято перед входом', swept_level_30: 'Уровень снятия (цена)', fuel_30: 'Топливо (да/нет)',
    touch_30: 'Касание зоны (1/2/3+)', against_flow: 'Против потока (да/нет)', conf: 'Подтверждение входа',
    react: 'Реакция от зоны (да/нет)', entry: 'Цена входа', stop: 'Цена стопа', target: 'Цена цели (тейк)',
    upper_bound: 'Верхняя граница зоны', lower_bound: 'Нижняя граница зоны', risk_pts: 'Риск в пунктах / $',
    pot_pts: 'Потенциал в пунктах', rr: 'Соотношение R:R', lots_30: 'Объём (лот)', lots_off: 'Объём (лот от % риска)',
    check_30: 'Чек-лист 7/7 (да/нет)', result: 'Результат (Тейк/Стоп/БУ)', res_30: 'Исход сделки',
    res_off: 'Результат (Торговый офис)', manual_r_30: 'Фактический R (вручную)', r_out: 'Итог сделки, R',
    r_out_off: 'Итог R (с сейфом)', profit_usd: 'Финансовый итог, $', prof_pct_off: 'Прибыль / Убыток, %',
    move_pts: 'Взятое движение, п.', cum_r: 'Накопительный итог, R', dd_r: 'Просадка, R',
    hit_1r: 'Дошла до 1:1 (да/нет)', hit_3r: 'Дошла до 1:3 (да/нет)', error: 'Причина ошибки',
    error_poi: 'Ошибка выбора POI', emotion: 'Эмоциональное состояние', comment: 'Комментарий',
    comment_30: 'Комментарий (30 сд.)', screenshot: 'Скриншот (ссылка)', aoi_num: '№ зоны / области',
    tf_ok: 'По системе (да/нет)', safe_pct: 'Доля сейфа (0.5)', safe_rr: 'Уровень сейфа RR (3)',
    risk_pct_col: 'Риск на сделку % (0.003)', f_ret: 'Фактор: Откат 50–62%', f_ob: 'Фактор: OB H4',
    f_fvg: 'Фактор: Имбаланс H4', f_liq: 'Фактор: Ликвидность рядом', factors_cnt: 'Количество факторов (формула)',
    valid_aoi: 'Валидность зоны (формула)', reached: 'Цена дошла до зоны', aoi_res: 'Реакция области',
    empty: 'Пустая ячейка (пропуск)'
  };

  const SHORT_COL_NAMES = {
    symbol: 'Тикер', date: 'Дата', time_msk: 'Время МСК', time_slot: 'Время', session: 'Сессия', tf: 'ТФ',
    dir_title: 'Направл.', dir_lower: 'Сторона', trend: 'Фаза', trend_h4: 'Тренд H4', flow_d: 'Поток Д',
    flow_4h: 'Поток 4Ч', flow_1h: 'Поток 1Ч', flow_15m: 'Поток 15м', tf_d1: 'D1', tf_h4: 'H4', tf_h1: 'H1',
    tf_m15: 'M15', struct_30: 'Структура', zone: 'Зона', zone_30: 'Вид зоны', poi_type: 'POI', form_m1: 'Формация',
    candle_col: 'Свеча', liq: 'Ликв.', swept_30: 'Снято', swept_level_30: 'Ур. снятия', fuel_30: 'Топливо',
    touch_30: 'Касание', against_flow: 'Против пот.', conf: 'Подтв.', react: 'Реакция', entry: 'Вход', stop: 'Стоп',
    target: 'Тейк', upper_bound: 'Верх', lower_bound: 'Низ', risk_pts: 'Риск п.', pot_pts: 'Потенц.', rr: 'R:R',
    lots_30: 'Лоты', lots_off: 'Объём', check_30: 'Чек 7/7', result: 'Результат', res_30: 'Исход', res_off: 'Результат',
    manual_r_30: 'R ручн.', r_out: 'Итог R', r_out_off: 'Итог RR', profit_usd: 'Итог $', prof_pct_off: 'Профит %',
    move_pts: 'Ход цены', cum_r: 'Накоп. R', dd_r: 'Просадка', hit_1r: '1:1', hit_3r: '1:3', error: 'Ошибка',
    error_poi: 'Ошибка', emotion: 'Эмоция', comment: 'Разбор', comment_30: 'Коммент.', screenshot: 'Скрин',
    aoi_num: '№ AOI', tf_ok: 'По сист.', safe_pct: 'Сейф %', safe_rr: 'Сейф RR', risk_pct_col: 'Риск %',
    f_ret: 'Откат', f_ob: 'OB H4', f_fvg: 'Имб H4', f_liq: 'Ликв', factors_cnt: 'Факторы', valid_aoi: 'Валидна',
    reached: 'Дошла', aoi_res: 'Итог AOI', empty: 'Пусто'
  };

  const defaultState = {
    presetKey: 'bt_main', customColumns: null, customPasteCol: 'B', customStartRow: 2,
    outputProfile: 'gs_ru', decimalSep: ',', timeMode: 'msk', trend: 'Тренд вверх',
    zone: 'Ордер-блок', liq: 'да', conf: 'Слом структуры', struct30: '', emotion: 'Спокоен',
    aoiNum: 1, flowD: 'вверх', flow4H: 'вверх', flow1H: 'вверх', flow15M: 'вверх',
    perTpl: {}, minimized: false, posTop: 20, posRight: 20
  };

  function loadState() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      const parsed = raw ? { ...defaultState, ...JSON.parse(raw) } : { ...defaultState };
      if (!['gs_ru', 'ru_form', 'en_form'].includes(parsed.outputProfile)) parsed.outputProfile = 'gs_ru';
      return parsed;
    } catch (e) { return { ...defaultState }; }
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
    if (!state.perTpl[key]) state.perTpl[key] = { batchQueue: [], cumR: 0, maxCumR: 0, lastSavedSig: '' };
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
  let isFixHelpOpen = false;
  let isInternalCopy = false;

  const hookedDelegates = new WeakSet();
  const hookedPrimitiveMaps = new WeakSet();
  const hookedTools = new WeakSet();
  const hookedStudies = new WeakSet();
  const hookedModels = new WeakSet();
  let liveSyncTimer = null;

  function scheduleLiveSync(delay = 30) {
    clearTimeout(liveSyncTimer);
    liveSyncTimer = setTimeout(() => {
      if (modalEl && modalEl.style.display === 'block') syncLiveChartChanges();
    }, delay);
  }

  function subscribeTvDelegate(del, cb) {
    if (!del || typeof del !== 'object' || typeof del.subscribe !== 'function' || hookedDelegates.has(del)) return;
    try {
      if (Array.isArray(del._listeners)) del.subscribe(null, cb);
      else del.subscribe(cb);
      hookedDelegates.add(del);
    } catch (e) {}
  }

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
        if (targetTime) {
          for (let d = 1; d <= 12; d++) {
            const rA = getPlotRow(plotsObj, hintIdx - d);
            if (rA && rA[0] === targetTime) return { index: hintIdx - d, val: rA };
            const rB = getPlotRow(plotsObj, hintIdx + d);
            if (rB && rB[0] === targetTime) return { index: hintIdx + d, val: rB };
          }
        } else if (r0) {
          return { index: hintIdx, val: r0 };
        }
      }
      if (!targetTime) return null;

      const rawFirst = typeof plotsObj.firstIndex === 'function' ? plotsObj.firstIndex() : null;
      const rawLast = typeof plotsObj.lastIndex === 'function' ? plotsObj.lastIndex() : null;
      if (rawFirst == null || rawLast == null) return null;

      const lastRow = getPlotRow(plotsObj, rawLast);
      const prevRow = getPlotRow(plotsObj, rawLast - 1);
      if (lastRow && prevRow && lastRow[0] > prevRow[0]) {
        const stepSec = lastRow[0] - prevRow[0];
        const estIdx = rawLast - Math.round((lastRow[0] - targetTime) / stepSec);
        const rEst = getPlotRow(plotsObj, estIdx);
        if (rEst && rEst[0] === targetTime) return { index: estIdx, val: rEst };
        for (let d = 1; d <= 8; d++) {
          const r1 = getPlotRow(plotsObj, estIdx - d);
          if (r1 && r1[0] === targetTime) return { index: estIdx - d, val: r1 };
          const r2 = getPlotRow(plotsObj, estIdx + d);
          if (r2 && r2[0] === targetTime) return { index: estIdx + d, val: r2 };
        }
      }

      const searchSegment = (loInit, hiInit) => {
        let lo = loInit, hi = hiInit;
        while (lo <= hi) {
          const mid = Math.floor((lo + hi) / 2);
          const row = getPlotRow(plotsObj, mid);
          if (!row || row[0] == null) break;
          if (row[0] === targetTime) return { index: mid, val: row };
          if (row[0] < targetTime) lo = mid + 1;
          else hi = mid - 1;
        }
        return null;
      };

      if (rawFirst < -500000 && rawLast >= 0) {
        const hitMain = searchSegment(0, rawLast);
        if (hitMain) return hitMain;
        return searchSegment(rawFirst, -999000);
      }
      return searchSegment(rawFirst, rawLast);
    } catch (e) {}
    return null;
  }

  function getPlotRowAtBar(plotsObj, bars, barIdx, plotIdxOffset = 0) {
    if (!plotsObj || barIdx == null) return null;
    const barRow = bars ? getPlotRow(bars, barIdx) : null;
    const barTime = barRow?.[0] || null;
    const direct = getPlotRow(plotsObj, barIdx + plotIdxOffset);
    if (direct && (!barTime || direct[0] === barTime)) return direct;
    if (barTime) {
      const found = findPlotRowByIndexOrTime(plotsObj, barTime, barIdx + plotIdxOffset);
      if (found?.val) return found.val;
    }
    return direct;
  }

  function extractDwgPrimitives(graphicsObj, getterName, onMapFound = null) {
    const out = [];
    if (!graphicsObj || typeof graphicsObj[getterName] !== 'function') return out;
    try {
      const visited = new Set();
      const unpack = (obj) => {
        if (!obj || typeof obj !== 'object' || visited.has(obj)) return;
        visited.add(obj);
        if (obj._primitivesDataById instanceof Map) {
          if (typeof onMapFound === 'function') onMapFound(obj._primitivesDataById);
          for (const v of obj._primitivesDataById.values()) if (v) out.push(v);
        } else if (obj instanceof Map || obj instanceof Set || Array.isArray(obj)) {
          const iter = obj instanceof Map ? obj.values() : obj;
          for (const v of iter) {
            if (v && typeof v === 'object' && (v._primitivesDataById || v instanceof Map || v instanceof Set || Array.isArray(v))) unpack(v);
            else if (v && typeof v === 'object') out.push(v);
          }
        }
      };
      unpack(graphicsObj[getterName]());
    } catch (e) {}
    return out;
  }

  const memoryBySym = {};
  function getSymMemory(symKey) {
    if (!memoryBySym[symKey]) memoryBySym[symKey] = { zones: new Map(), levels: new Map(), lqLabels: new Map() };
    return memoryBySym[symKey];
  }

  function trimMapFIFO(mapObj, maxEntries = 300) {
    while (mapObj.size > maxEntries) {
      const oldestKey = mapObj.keys().next().value;
      mapObj.delete(oldestKey);
    }
  }

  function parseZoneLabelText(rawText, boxColor = 0, boxTva = '', boxBgColor = 0) {
    const t = String(rawText || '').trim();
    const low = t.toLowerCase();
    if (low.includes('держит')) return null;
    if (!t && ((Number(boxBgColor || 0) >>> 24) === 0 || boxColor === boxBgColor)) return null;

    let zoneType = null, htfRank = 0;
    if (low.includes('разворотная')) zoneType = 'Разворотная зона';
    else if (low.includes('сильный покупатель')) zoneType = 'Сильный покупатель';
    else if (low.includes('средний покупатель')) zoneType = 'Средний покупатель';
    else if (low.includes('красивый покупатель')) zoneType = 'Красивый покупатель';
    else if (low.includes('сильный продавец')) zoneType = 'Сильный продавец';
    else if (low.includes('средний продавец')) zoneType = 'Средний продавец';
    else if (low.includes('красивый продавец')) zoneType = 'Красивый продавец';
    else if (low.includes('зона входа')) zoneType = 'Зона входа';
    else if (low.includes('проторговка')) zoneType = 'Проторговка';
    else if (/(^|\s)д\s*0[.,]5/i.test(t)) { zoneType = 'Зона Д 0.5'; htfRank = 4; }
    else if (/4ч\s*0[.,]5/i.test(t)) { zoneType = 'Зона 4Ч 0.5'; htfRank = 3; }
    else if (/1ч\s*0[.,]5/i.test(t)) { zoneType = 'Зона 1Ч 0.5'; htfRank = 2; }
    else if (/15м\s*0[.,]5/i.test(t)) { zoneType = 'Зона 15м 0.5'; htfRank = 1; }

    if (!zoneType && !t) {
      if (boxColor === 4536050 && boxTva === 'top') zoneType = 'Сильный продавец';
      else if (boxColor === 8493320 && boxTva === 'top') zoneType = 'Сильный покупатель';
      else if (boxColor === 39167 && boxTva === 'bottom') zoneType = 'Проторговка';
      else if (boxColor === 16736809 && boxTva === 'bottom') {
        return {
          zoneType: 'Проторговка', fuel: 'нет', touch: '1', maxTouchNum: 1, against: 'да',
          side: 'ANY', isProtor: true, htfRank: 0, isDeadFallback: true,
          rawText: 'Проторговка · против потока (отработана)'
        };
      }
    }
    if (!zoneType) return null;

    const fuel = low.includes('топливо') ? 'да' : 'нет';
    const against = (low.includes('против потока') || boxColor === 16736809) ? 'да' : 'нет';
    let touch = '1', maxTouchNum = 1;
    const touchMatch = low.match(/касание\s*(\d+)/);
    if (touchMatch) {
      maxTouchNum = parseInt(touchMatch[1], 10) || 1;
      if (maxTouchNum === 2) touch = '2';
      else if (maxTouchNum >= 3) touch = '3 и больше';
    }

    const side = zoneType.includes('покупатель') ? 'LONG' : (zoneType.includes('продавец') ? 'SHORT' : 'ANY');
    const isProtor = zoneType === 'Проторговка';
    return {
      zoneType, fuel, touch, maxTouchNum, against, side, isProtor, htfRank,
      isDeadFallback: !t, rawText: t || `${zoneType} (отработана)`
    };
  }

  function parseLiquidityLabelText(rawText, fallbackY) {
    const t = String(rawText || '').trim();
    if (!t) return null;
    const low = t.toLowerCase();

    if (low.includes('lq') && low.includes('снят')) {
      let sweptType = 'уровень 15м';
      if (low.includes('15m') || low.includes('15м')) sweptType = 'уровень 15м';
      else if (low.includes('1h') || low.includes('1ч')) sweptType = 'уровень 1Ч';
      else if (low.includes('4h') || low.includes('4ч')) sweptType = 'уровень 4Ч';
      else if (low.includes('d') || low.includes('д')) sweptType = 'уровень Д';
      return { sweptType, price: Number(fallbackY || 0), isDirectSweepLabel: true, rawText: t };
    }

    let sweptType = null;
    if (low.includes('приманка')) sweptType = 'приманка';
    else if (low === 'полка' || low.startsWith('полка ')) sweptType = 'полка';
    else if (low.includes('15м/1ч')) sweptType = 'склейка 15м/1Ч';
    else if (/д\s*[·.]\s*4ч/i.test(low)) sweptType = 'склейка Д · 4Ч';
    else if (/15м\s+(макс|мин)/i.test(low)) sweptType = 'уровень 15м';
    else if (/1ч\s+(макс|мин)/i.test(low)) sweptType = 'уровень 1Ч';
    else if (/4ч\s+(макс|мин)/i.test(low)) sweptType = 'уровень 4Ч';
    else if (/(^|\s)д\s+(макс|мин)/i.test(low)) sweptType = 'уровень Д';
    if (!sweptType) return null;

    let price = Number(fallbackY || 0);
    const priceMatch = t.match(/(\d{3,6}[.,]\d+)/);
    if (priceMatch) {
      const parsedP = parseFloat(priceMatch[1].replace(',', '.'));
      if (!isNaN(parsedP) && parsedP > 0) price = parsedP;
    }
    if (!price || isNaN(price)) return null;
    return { sweptType, price, isDirectSweepLabel: false, rawText: t };
  }

  function computeTouchAtEntryBar(bars, barX1, triggerIdx, bottom, top, maxTouchNum) {
    if (!bars || maxTouchNum <= 1 || barX1 < -100000 || triggerIdx == null || triggerIdx <= barX1) return '1';
    let touches = 0, leftZone = false;
    for (let i = barX1 + 1; i <= triggerIdx; i++) {
      const b = getPlotRow(bars, i);
      if (!b) continue;
      const hi = Number(b[2]), lo = Number(b[3]);
      if (isNaN(hi) || isNaN(lo)) continue;
      const inZone = (hi >= bottom && lo <= top);
      if (!inZone) leftZone = true;
      else if (leftZone) { touches++; leftZone = false; }
    }
    const actual = Math.max(1, Math.min(touches || 1, maxTouchNum));
    return actual === 1 ? '1' : (actual === 2 ? '2' : '3 и больше');
  }

  function getRightmostBarFromG(g) {
    const idxArr = Array.isArray(g?._indexes) ? g._indexes : null;
    if (!idxArr || !idxArr.length) return 999999;
    for (let i = idxArr.length - 1; i >= 0; i--) {
      const v = idxArr[i];
      if (v != null && v > -1000000) return v;
    }
    return 999999;
  }

  function hookPrimitiveMapInstance(primMap, kind, g, symKey) {
    if (!primMap || !(primMap instanceof Map)) return;
    primMap._btCtx = { kind, g, symKey };
    if (hookedPrimitiveMaps.has(primMap)) return;
    hookedPrimitiveMaps.add(primMap);

    const resolveBar = (ctxG, x, fb) => {
      const idxArr = Array.isArray(ctxG?._indexes) ? ctxG._indexes : null;
      if (!idxArr || x == null || x < 0 || x >= idxArr.length) return fb;
      const v = idxArr[x];
      return (v != null && v > -1000000) ? v : fb;
    };

    const origSet = primMap.set;
    const origDelete = primMap.delete;

    primMap.set = function(id, val) {
      const res = origSet.call(this, id, val);
      try {
        const ctx = this._btCtx;
        if (!ctx || !val || typeof val !== 'object') return res;
        const mem = getSymMemory(ctx.symKey);
        const curBar = getRightmostBarFromG(ctx.g);

        if (ctx.kind === 'dwgboxes') {
          const text = String(val.t ?? '').trim();
          if (!text) return res;
          const top = Math.max(Number(val.y1 || 0), Number(val.y2 || 0));
          const bottom = Math.min(Number(val.y1 || 0), Number(val.y2 || 0));
          const barX1 = resolveBar(ctx.g, Math.min(val.x1, val.x2), -999999);
          const barX2 = resolveBar(ctx.g, Math.max(val.x1, val.x2), curBar);
          if (!top || !bottom || barX1 < -100000) return res;

          if (text === 'Полка') {
            mem.levels.set(`полка_${bottom.toFixed(2)}`, { sweptType: 'полка', price: bottom, barX1, lastSeenBar: curBar, rawText: `Полка ${bottom}` });
            trimMapFIFO(mem.levels);
          } else {
            const zInfo = parseZoneLabelText(text, val.c, val.tva, val.bc);
            if (zInfo && !zInfo.isDeadFallback) {
              mem.zones.set(`${bottom.toFixed(2)}_${top.toFixed(2)}`, { barX1, barX2, bottom, top, zInfo });
              trimMapFIFO(mem.zones);
            }
          }
        } else if (ctx.kind === 'dwglabels') {
          const lInfo = parseLiquidityLabelText(val.t, val.y);
          if (!lInfo) return res;
          const barX = resolveBar(ctx.g, val.x, curBar);
          if (lInfo.isDirectSweepLabel) {
            mem.lqLabels.set(`lq_${barX}_${lInfo.sweptType}`, { barX, sweptType: lInfo.sweptType, price: lInfo.price, rawText: lInfo.rawText });
            trimMapFIFO(mem.lqLabels);
          } else {
            const key = `${lInfo.sweptType}_${lInfo.price.toFixed(2)}`;
            const prev = mem.levels.get(key);
            mem.levels.set(key, {
              sweptType: lInfo.sweptType, price: lInfo.price,
              barX1: prev ? prev.barX1 : (barX - 20), lastSeenBar: curBar, rawText: lInfo.rawText
            });
            trimMapFIFO(mem.levels);
          }
        }
      } catch (e) {}
      return res;
    };

    primMap.delete = function(id) {
      const oldVal = this.get(id);
      const res = origDelete.call(this, id);
      try {
        const ctx = this._btCtx;
        if (!ctx || !oldVal || typeof oldVal !== 'object') return res;
        const mem = getSymMemory(ctx.symKey);
        const curBar = getRightmostBarFromG(ctx.g);

        if (ctx.kind === 'dwglabels') {
          const lInfo = parseLiquidityLabelText(oldVal.t, oldVal.y);
          if (lInfo && !lInfo.isDirectSweepLabel) {
            const key = `${lInfo.sweptType}_${lInfo.price.toFixed(2)}`;
            const existing = mem.levels.get(key);
            if (existing) {
              existing.sweptAtBar = curBar;
              existing.lastSeenBar = curBar;
            } else {
              mem.levels.set(key, {
                sweptType: lInfo.sweptType, price: lInfo.price,
                barX1: curBar - 25, lastSeenBar: curBar, sweptAtBar: curBar, rawText: lInfo.rawText
              });
              trimMapFIFO(mem.levels);
            }
          }
        } else if (ctx.kind === 'dwgboxes') {
          const text = String(oldVal.t ?? '').trim();
          if (text === 'Полка') {
            const bottom = Math.min(Number(oldVal.y1 || 0), Number(oldVal.y2 || 0));
            const ex = mem.levels.get(`полка_${bottom.toFixed(2)}`);
            if (ex) { ex.sweptAtBar = curBar; ex.lastSeenBar = curBar; }
          }
        }
      } catch (e) {}
      return res;
    };
  }

  function syncGraphicsMemoryO1(model, symKey) {
    if (!model || typeof model.dataSources !== 'function') return null;
    const mem = getSymMemory(symKey);
    let mainG = null;

    try {
      for (const s of model.dataSources()) {
        const meta = typeof s.metaInfo === 'function' ? s.metaInfo() : s._metaInfo;
        const title = meta?.description || meta?.name || '';
        if (!title.includes('Матрёшка')) continue;

        const g = typeof s.graphics === 'function' ? s.graphics() : s._graphics;
        if (!g) continue;
        if (title.includes('Зоны интереса')) mainG = g;

        const idxArr = Array.isArray(g._indexes) ? g._indexes : null;
        const toBar = (x, fb) => {
          if (!idxArr || x == null || x < 0 || x >= idxArr.length) return fb;
          const v = idxArr[x];
          return (v != null && v > -1000000) ? v : fb;
        };

        const boxes = extractDwgPrimitives(g, 'dwgboxes', m => hookPrimitiveMapInstance(m, 'dwgboxes', g, symKey));
        const labels = extractDwgPrimitives(g, 'dwglabels', m => hookPrimitiveMapInstance(m, 'dwglabels', g, symKey));
        const lines = extractDwgPrimitives(g, 'dwglines', m => hookPrimitiveMapInstance(m, 'dwglines', g, symKey));

        let maxActivePineX = 0;
        for (const b of boxes) if (b.t && b.x2 > maxActivePineX) maxActivePineX = b.x2;
        const currentRightBar = maxActivePineX > 0 ? toBar(maxActivePineX, getRightmostBarFromG(g)) : getRightmostBarFromG(g);

        for (const b of boxes) {
          const text = String(b.t ?? '').trim();
          if (!text) continue;
          const top = Math.max(Number(b.y1 || 0), Number(b.y2 || 0));
          const bottom = Math.min(Number(b.y1 || 0), Number(b.y2 || 0));
          const barX1 = toBar(Math.min(b.x1, b.x2), -999999);
          const barX2 = toBar(Math.max(b.x1, b.x2), 999999);
          if (!top || !bottom || barX1 < -100000) continue;

          if (text === 'Полка') {
            mem.levels.set(`полка_${bottom.toFixed(2)}`, { sweptType: 'полка', price: bottom, barX1, lastSeenBar: currentRightBar, rawText: `Полка ${bottom}` });
            trimMapFIFO(mem.levels);
            continue;
          }

          const zInfo = parseZoneLabelText(text, b.c, b.tva, b.bc);
          if (zInfo && !zInfo.isDeadFallback) {
            mem.zones.set(`${bottom.toFixed(2)}_${top.toFixed(2)}`, { barX1, barX2, bottom, top, zInfo });
            trimMapFIFO(mem.zones);
          }
        }

        for (const lb of labels) {
          const lInfo = parseLiquidityLabelText(lb.t, lb.y);
          if (!lInfo) continue;
          const barX = toBar(lb.x, currentRightBar);

          if (lInfo.isDirectSweepLabel) {
            mem.lqLabels.set(`lq_${barX}_${lInfo.sweptType}`, { barX, sweptType: lInfo.sweptType, price: lInfo.price, rawText: lInfo.rawText });
            trimMapFIFO(mem.lqLabels);
          } else {
            const matchLine = lines.find(ln => Math.abs(Number(ln.y1 || 0) - lInfo.price) < 0.05);
            const birthBar = matchLine ? toBar(Math.min(matchLine.x1, matchLine.x2), barX - 20) : (barX - 20);
            const key = `${lInfo.sweptType}_${lInfo.price.toFixed(2)}`;
            const prev = mem.levels.get(key);
            mem.levels.set(key, {
              sweptType: lInfo.sweptType, price: lInfo.price,
              barX1: prev ? Math.min(prev.barX1, birthBar) : birthBar,
              lastSeenBar: currentRightBar, sweptAtBar: prev?.sweptAtBar || null, rawText: lInfo.rawText
            });
            trimMapFIFO(mem.levels);
          }
        }
      }
    } catch (e) {}
    return mainG;
  }

  
  function isValidStructureOnBars(bars, bx1, bx2, price, dir) {
    if (bx1 == null || bx2 == null || bx1 < -100000 || bx2 >= 900000) return false;
    if (bx2 <= bx1 || bx2 - bx1 > 120) return false;
    if (!price || isNaN(price) || price <= 0) return false;
    if (!bars) return true;

    const r2 = getPlotRow(bars, bx2);
    if (r2) {
      const hi2 = Number(r2[2]), lo2 = Number(r2[3]);
      if (!isNaN(hi2) && !isNaN(lo2)) {
        const rangeTol = Math.max(0.25, (hi2 - lo2) * 1.5);
        if (price < lo2 - rangeTol || price > hi2 + rangeTol) return false;
      }
    }

    const r1 = getPlotRow(bars, bx1);
    if (r1) {
      const hi1 = Number(r1[2]), lo1 = Number(r1[3]);
      if (!isNaN(hi1) && !isNaN(lo1)) {
        const startTol = Math.max(0.15, (hi1 - lo1) * 0.5);
        if (dir === '↑' && Math.abs(price - hi1) > startTol && (price < lo1 - startTol || price > hi1 + startTol)) return false;
        if (dir === '↓' && Math.abs(price - lo1) > startTol && (price < lo1 - startTol || price > hi1 + startTol)) return false;
      }
    }
    return true;
  }

  function analyzeChartGraphicsO1(g, plotsObj, plotIdxOffset, symKey, triggerIdx, isLong, entryPrice, stopPrice, entryExtreme, bars) {
    const out = { matchedZone: null, matchedSweep: null, detectedStructure: null, isHistorical: false };
    if (!g || !entryPrice || !stopPrice) return out;

    const mem = getSymMemory(symKey);
    const idxArr = Array.isArray(g._indexes) ? g._indexes : null;
    const toBar = (x, fallback) => {
      if (!idxArr || x == null || x < 0 || x >= idxArr.length) return fallback;
      const v = idxArr[x];
      return (v != null && v > -1000000) ? v : fallback;
    };

    const boxes = extractDwgPrimitives(g, 'dwgboxes');
    const labels = extractDwgPrimitives(g, 'dwglabels');
    const lines = extractDwgPrimitives(g, 'dwglines');

    let maxActivePineX = 0;
    for (const b of boxes) if (b.t && b.x2 > maxActivePineX) maxActivePineX = b.x2;
    const lastBarFallback = typeof bars?.lastIndex === 'function' ? bars.lastIndex() : getRightmostBarFromG(g);
    const currentRightBar = maxActivePineX > 0 ? toBar(maxActivePineX, lastBarFallback) : lastBarFallback;

    if (triggerIdx != null && currentRightBar < 900000 && triggerIdx < currentRightBar - 15) {
      out.isHistorical = true;
    }

    const structByBar = new Map();
    const upsertStruct = (bx1, bx2, type, dir, inner, price) => {
      if (bx2 >= currentRightBar) return;
      if (!isValidStructureOnBars(bars, bx1, bx2, price, dir)) return;
      const existing = structByBar.get(bx2);
      if (!existing) {
        structByBar.set(bx2, { barX1: bx1, barX2: bx2, type, dir, inner, price });
      } else if (existing.inner && !inner) {
        structByBar.set(bx2, { barX1: bx1, barX2: bx2, type, dir, inner: false, price });
      }
    };

    for (const l of lines) {
      const y1 = Number(l.y1 ?? 0), y2 = Number(l.y2 ?? 0);
      if (!y1 || Math.abs(y1 - y2) > 0.001) continue;
      const bx1 = toBar(Math.min(l.x1, l.x2), -999999);
      const bx2 = toBar(Math.max(l.x1, l.x2), 999999);

      if (l.w === 1 && (l.st === 'sol' || l.st === 'dsh') && (l.ci === 4282726130 || l.ci === 4286683400)) {
        const dir = l.ci === 4282726130 ? '↓' : '↑';
        const type = l.st === 'sol' ? 'пробой' : 'слом';
        upsertStruct(bx1, bx2, type, dir, false, y1);
      } else if (l.w === 1 && l.st === 'dot' && (l.ci === 2353346290 || l.ci === 2357303560)) {
        const dir = l.ci === 2353346290 ? '↓' : '↑';
        upsertStruct(bx1, bx2, 'слом', dir, true, y1);
      }
    }

    for (const b of boxes) {
      const text = String(b.t ?? '').trim().toLowerCase();
      if (!text.includes('слом') && !text.includes('пробой')) continue;
      const bx1 = toBar(Math.min(b.x1, b.x2), -999999);
      const bx2 = toBar(Math.max(b.x1, b.x2), 999999);
      const yPrice = Number(b.y1 ?? b.y2 ?? 0);
      const dir = text.includes('↓') ? '↓' : '↑';
      const type = text.includes('слом') ? 'слом' : 'пробой';
      const isInner = b.ts === 'tiny' || text.includes('внутр');
      upsertStruct(bx1, bx2, type, dir, isInner, yPrice);
    }

    const tradeDirArrow = isLong ? '↑' : '↓';
    const allValidStructs = Array.from(structByBar.values()).sort((a, b) => a.barX2 - b.barX2);
    let pastStruct = allValidStructs.filter(e => triggerIdx == null || e.barX2 <= triggerIdx);

    if (triggerIdx != null && !pastStruct.some(e => !e.inner && e.dir === tradeDirArrow && triggerIdx - e.barX2 <= 15)) {
      const nextBarBreak = allValidStructs.find(e => e.barX2 === triggerIdx + 1 && !e.inner && e.dir === tradeDirArrow);
      if (nextBarBreak) pastStruct.push(nextBarBreak);
    }

    if (pastStruct.length > 0) {
      const lastEv = pastStruct[pastStruct.length - 1];
      const major = pastStruct.filter(e => !e.inner);
      const mLast = major.length > 0 ? major[major.length - 1] : null;

      const useInner = lastEv.inner &&
        (triggerIdx == null || triggerIdx - lastEv.barX2 <= 6) &&
        (lastEv.dir === tradeDirArrow || !mLast);

      if (useInner) {
        out.detectedStructure = `внутр. слом ${lastEv.dir}`;
      } else if (mLast) {
        let targetMajorIdx = major.length - 1;
        if (mLast.dir !== tradeDirArrow && major.length >= 2) {
          const prevMajor = major[major.length - 2];
          if (prevMajor.dir === tradeDirArrow && (triggerIdx == null || triggerIdx - prevMajor.barX2 <= 12)) {
            targetMajorIdx = major.length - 2;
          }
        }
        const activeMajor = major[targetMajorIdx];

        if (activeMajor.type === 'слом') {
          const oppDir = activeMajor.dir === '↑' ? '↓' : '↑';
          let hadFakeout = false;
          if (plotsObj && activeMajor.barX2 != null) {
            for (let bi = activeMajor.barX2 - 8; bi <= activeMajor.barX2; bi++) {
              const pr = getPlotRowAtBar(plotsObj, bars, bi, plotIdxOffset);
              if (pr && pr[7] === 1) { hadFakeout = true; break; }
            }
          }
          let cntOpp = 0;
          for (let i = targetMajorIdx - 1; i >= 0; i--) {
            if (major[i].type === 'пробой' && major[i].dir === oppDir) cntOpp++;
            else break;
          }
          if (hadFakeout) out.detectedStructure = `ложный пробой ${oppDir} · слом ${activeMajor.dir}`;
          else if (cntOpp >= 2) out.detectedStructure = `пробой ${oppDir} · пробой ${oppDir} · слом ${activeMajor.dir}`;
          else out.detectedStructure = `слом ${activeMajor.dir}`;
        } else {
          const sameDir = activeMajor.dir;
          let cntSame = 0;
          for (let i = targetMajorIdx; i >= 0; i--) {
            if (major[i].type === 'пробой' && major[i].dir === sameDir) cntSame++;
            else break;
          }
          out.detectedStructure = cntSame >= 2 ? `пробой ${sameDir} · пробой ${sameDir}` : `пробой ${sameDir}`;
        }
      } else {
        out.detectedStructure = `внутр. слом ${lastEv.dir}`;
      }
    }

    const tradeLow = Math.min(entryPrice, stopPrice);
    const tradeHigh = Math.max(entryPrice, stopPrice);
    const minP = Math.min(tradeLow, entryExtreme !== '' ? Number(entryExtreme) : tradeLow);
    const maxP = Math.max(tradeHigh, entryExtreme !== '' ? Number(entryExtreme) : tradeHigh);
    const tol = Math.max(0.08, Math.abs(entryPrice - stopPrice) * 0.08);

    const zoneCandidates = [];
    for (const b of boxes) {
      const top = Math.max(Number(b.y1 || 0), Number(b.y2 || 0));
      const bottom = Math.min(Number(b.y1 || 0), Number(b.y2 || 0));
      if (!top || !bottom) continue;

      const barX1 = toBar(Math.min(b.x1, b.x2), -999999);
      const barX2 = toBar(Math.max(b.x1, b.x2), 999999);

      const memKey = `${bottom.toFixed(2)}_${top.toFixed(2)}`;
      const zInfo = (!b.t && mem.zones.has(memKey))
        ? mem.zones.get(memKey).zInfo
        : parseZoneLabelText(b.t, b.c, b.tva, b.bc);
      if (!zInfo) continue;

      const isDead2Bar = zInfo.isDeadFallback && (barX2 - barX1 <= 3);
      if (triggerIdx != null) {
        if (barX1 > triggerIdx + 1) continue;
        if (!isDead2Bar && barX2 < triggerIdx - 1) continue;
        if (isDead2Bar && barX1 > -100000 && triggerIdx - barX1 > 120) continue;
      }

      const entryInside = (entryPrice >= bottom - tol && entryPrice <= top + tol);
      const extremeInside = (entryExtreme !== '' && Number(entryExtreme) >= bottom - tol && Number(entryExtreme) <= top + tol);
      const overlapHeight = Math.min(tradeHigh, top) - Math.max(tradeLow, bottom);
      if (!entryInside && !extremeInside && overlapHeight <= 0) continue;

      const sideMatch = (zInfo.side === (isLong ? 'LONG' : 'SHORT')) || (zInfo.side === 'ANY' && !zInfo.isProtor);
      let rank = 30;
      if (zInfo.isProtor) rank = entryInside ? 40 : 50;
      else if (sideMatch && entryInside) rank = 0;
      else if (sideMatch && extremeInside) rank = 5;
      else if (sideMatch) rank = 12;
      else if (entryInside) rank = 20;
      if (zInfo.isDeadFallback) rank += 4;

      const entryDist = entryPrice < bottom ? (bottom - entryPrice) : (entryPrice > top ? (entryPrice - top) : 0);
      const barDist = (triggerIdx != null && barX1 > -100000) ? Math.abs(triggerIdx - barX1) : 500;
      zoneCandidates.push({ rank, entryDist, barDist, top, bottom, barX1, zInfo });
    }

    if (zoneCandidates.length > 0) {
      zoneCandidates.sort((a, b) => (a.rank - b.rank) || (a.entryDist - b.entryDist) || (a.barDist - b.barDist));
      const best = zoneCandidates[0];
      const touchAtEntry = (triggerIdx != null && triggerIdx < currentRightBar - 2)
        ? computeTouchAtEntryBar(bars, best.barX1, triggerIdx, best.bottom, best.top, best.zInfo.maxTouchNum)
        : best.zInfo.touch;

      out.matchedZone = {
        zoneType: best.zInfo.zoneType, fuel: best.zInfo.fuel,
        touch: touchAtEntry, against: best.zInfo.against, rawText: best.zInfo.rawText
      };
    } else {
      const htfMatches = [];
      for (const lb of labels) {
        const z05 = parseZoneLabelText(lb.t);
        if (!z05 || !z05.htfRank) continue;
        const midY = Number(lb.y || 0);
        if (midY >= minP - tol && midY <= maxP + tol) {
          htfMatches.push({ htfRank: z05.htfRank, midDist: Math.abs(entryPrice - midY), z05 });
        }
      }
      if (htfMatches.length > 0) {
        htfMatches.sort((a, b) => (b.htfRank - a.htfRank) || (a.midDist - b.midDist));
        const bestHtf = htfMatches[0].z05;
        out.matchedZone = {
          zoneType: bestHtf.zoneType, fuel: bestHtf.fuel,
          touch: '1', against: bestHtf.against, rawText: bestHtf.rawText
        };
      }
    }

    let sweepHigh = maxP, sweepLow = minP;
    if (bars && triggerIdx != null) {
      for (let i = triggerIdx - 5; i <= triggerIdx; i++) {
        const b = getPlotRow(bars, i);
        if (!b) continue;
        const hi = Number(b[2]), lo = Number(b[3]);
        if (!isNaN(hi)) sweepHigh = Math.max(sweepHigh, hi);
        if (!isNaN(lo)) sweepLow = Math.min(sweepLow, lo);
      }
    }

    const sweepCandidates = [];
    for (const lqItem of mem.lqLabels.values()) {
      if (triggerIdx != null && (lqItem.barX < triggerIdx - 5 || lqItem.barX > triggerIdx + 1)) continue;
      const barDiff = triggerIdx != null ? Math.abs(triggerIdx - lqItem.barX) : 0;
      const swPrice = Number(entryExtreme !== '' ? entryExtreme : lqItem.price);
      sweepCandidates.push({ priority: 0, barDiff, priceDiff: 0, sType: lqItem.sweptType, price: swPrice, rawText: lqItem.rawText });
    }

    for (const l of lines) {
      const ly = Number(l.y1 ?? 0);
      if (!ly || Math.abs(ly - Number(l.y2 ?? 0)) > 0.001) continue;

      const isSweptBait = l.st === 'dot' && l.w === 1 && (l.ci === 1505670578 || l.ci === 1027311288);
      const isSweptShelf = l.st === 'dot' && l.w === 2 && (l.ci === 1497708274 || l.ci === 1501665544);
      if (!isSweptBait && !isSweptShelf) continue;

      const barX1 = toBar(Math.min(l.x1, l.x2), -999999);
      const barX2 = toBar(Math.max(l.x1, l.x2), 999999);
      if (barX2 >= currentRightBar || barX1 < -100000) continue;

      if (triggerIdx != null) {
        if (barX1 >= triggerIdx) continue;
        if (barX2 < triggerIdx - 6 || barX2 > triggerIdx + 1) continue;
      }

      const inPriceRange = isLong
        ? (ly >= sweepLow - tol && ly <= entryPrice + tol)
        : (ly <= sweepHigh + tol && ly >= entryPrice - tol);
      if (!inPriceRange) continue;

      const sType = isSweptShelf ? 'полка' : 'приманка';
      const barDiff = triggerIdx != null ? Math.abs(triggerIdx - barX2) : 0;
      const priceDiff = Math.abs(Number(entryExtreme !== '' ? entryExtreme : entryPrice) - ly);
      sweepCandidates.push({ priority: 1, barDiff, priceDiff, sType, price: ly, rawText: `${sType} ${ly}` });
    }

    for (const lvl of mem.levels.values()) {
      if (triggerIdx != null && lvl.barX1 >= triggerIdx) continue;
      if (!lvl.sweptAtBar && lvl.lastSeenBar >= currentRightBar) continue;

      const inPriceRange = isLong
        ? (lvl.price >= sweepLow - tol && lvl.price <= entryPrice + tol)
        : (lvl.price <= sweepHigh + tol && lvl.price >= entryPrice - tol);
      if (!inPriceRange) continue;

      const refBar = lvl.sweptAtBar || lvl.lastSeenBar;
      const barDiff = triggerIdx != null ? Math.abs(triggerIdx - refBar) : 2;
      if (barDiff > 8) continue;
      const priceDiff = Math.abs(Number(entryExtreme !== '' ? entryExtreme : entryPrice) - lvl.price);
      sweepCandidates.push({ priority: 1, barDiff, priceDiff, sType: lvl.sweptType, price: lvl.price, rawText: lvl.rawText });
    }

    if (plotsObj && triggerIdx != null) {
      for (let bi = triggerIdx; bi >= triggerIdx - 5; bi--) {
        const pr = getPlotRowAtBar(plotsObj, bars, bi, plotIdxOffset);
        if (!pr) continue;
        const prevPr = getPlotRowAtBar(plotsObj, bars, bi - 1, plotIdxOffset) || pr;
        const barRow = getPlotRow(bars, bi);
        const barHi = Number(barRow?.[2] ?? sweepHigh);
        const barLo = Number(barRow?.[3] ?? sweepLow);

        const h4Lvl = Number(prevPr[isLong ? 2 : 1] || pr[isLong ? 2 : 1] || 0);
        const d1Lvl = Number(prevPr[isLong ? 4 : 3] || pr[isLong ? 4 : 3] || 0);
        const directHitD1 = d1Lvl > 0 && (isLong ? (barLo <= d1Lvl + tol && entryPrice >= d1Lvl - tol) : (barHi >= d1Lvl - tol && entryPrice <= d1Lvl + tol));
        const directHitH4 = h4Lvl > 0 && (isLong ? (barLo <= h4Lvl + tol && entryPrice >= h4Lvl - tol) : (barHi >= h4Lvl - tol && entryPrice <= h4Lvl + tol));

        const sweptHtf = (isLong ? (pr[15] === 1) : (pr[14] === 1)) || directHitD1 || directHitH4;
        const swept15m = (pr[18] === 1) || (isLong ? (pr[12] === 1) : (pr[13] === 1));

        if (sweptHtf) {
          let sType = 'уровень 1Ч';
          let swPrice = Number(entryExtreme !== '' ? entryExtreme : entryPrice);

          if (directHitD1 && directHitH4 && Math.abs(d1Lvl - h4Lvl) < 0.1) { sType = 'склейка Д · 4Ч'; swPrice = d1Lvl; }
          else if (directHitD1) { sType = 'уровень Д'; swPrice = d1Lvl; }
          else if (directHitH4) { sType = 'уровень 4Ч'; swPrice = h4Lvl; }
          sweepCandidates.push({ priority: 0, barDiff: triggerIdx - bi, priceDiff: 0, sType, price: swPrice, rawText: `${sType} ${swPrice}` });
          break;
        } else if (swept15m && sweepCandidates.length === 0) {
          const swPrice = Number(entryExtreme !== '' ? entryExtreme : entryPrice);
          sweepCandidates.push({ priority: 2, barDiff: triggerIdx - bi, priceDiff: 0, sType: 'уровень 15м', price: swPrice, rawText: `уровень 15м ${swPrice}` });
        }
      }
    }

    if (sweepCandidates.length > 0) {
      sweepCandidates.sort((a, b) => (a.priority - b.priority) || (a.barDiff - b.barDiff) || (a.priceDiff - b.priceDiff));
      const bestSw = sweepCandidates[0];
      out.matchedSweep = { sweptType: bestSw.sType, price: Number(bestSw.price), rawText: bestSw.rawText };
    }

    return out;
  }

  
  function resolveMvPlotIndices(s, sampleRow) {
    let colorIdx = 3, modelIdx = 4;
    try {
      const meta = typeof s.metaInfo === 'function' ? s.metaInfo() : s._metaInfo;
      const plotsMeta = meta?.plots;
      const styles = meta?.styles || meta?.defaults?.styles;
      if (Array.isArray(plotsMeta) && styles) {
        plotsMeta.forEach((p, i) => {
          const stTitle = String(styles[p.id]?.title || p.title || p.id || '').toLowerCase();
          if (stTitle.includes('цвет')) colorIdx = i + 1;
          else if (stTitle.includes('модел')) modelIdx = i + 1;
        });
      }
    } catch (e) {}

    if (Array.isArray(sampleRow)) {
      if (sampleRow.length <= 3) {
        colorIdx = 1;
        modelIdx = 2;
      } else if (sampleRow[colorIdx] === undefined && sampleRow[1] !== undefined && Math.abs(Number(sampleRow[1])) <= 3) {
        colorIdx = 1;
        modelIdx = 2;
      }
    }
    return { colorIdx, modelIdx };
  }

  function mapMvModelToFormM1(modelCode, isContinuation) {
    if (isContinuation || modelCode === 18) return 'Правильный откат';
    if (modelCode === 12) return 'Поглощение';
    if (modelCode === 5) return 'Пин-бар (IFS)';
    if (modelCode === 2) return 'Свип';
    return 'SCOB';
  }

  function readMatryoshkaIndicators(model, symKey, entryTime, triggerIdx, startIdx, endIdx, isLong, entryPrice, stopPrice, entryExtreme, bars) {
    const result = { flows: null, flowsHistorical: false, candleColor: null, modelCode: 0, isContinuation: false, matchedZone: null, matchedSweep: null, detectedStructure: null };
    if (!model || typeof model.dataSources !== 'function') return result;

    try {
      syncGraphicsMemoryO1(model, symKey);
      const sources = model.dataSources();
      for (const s of sources) {
        const meta = typeof s.metaInfo === 'function' ? s.metaInfo() : s._metaInfo;
        const title = meta?.description || meta?.name || '';

        if (title.includes('Матрёшка Зоны интереса')) {
          const g = typeof s.graphics === 'function' ? s.graphics() : s._graphics;
          const plotsObj = typeof s.plots === 'function' ? s.plots() : s._data?.bars?.();
          const tcCells = extractDwgPrimitives(g, 'dwgtablecells');
          if (tcCells.length > 0) {
            const rowsByTid = {};
            for (const cell of tcCells) {
              if (!cell || cell.tid == null) continue;
              const key = `${cell.tid}_${cell.row}`;
              if (!rowsByTid[key]) rowsByTid[key] = {};
              rowsByTid[key][cell.col] = String(cell.t || '').trim();
            }
            const flows = {};
            for (const rObj of Object.values(rowsByTid)) {
              const label = rObj[0], sym = rObj[1];
              if (!label || !sym) continue;
              const dirWord = sym === '↑' ? 'вверх' : (sym === '↓' ? 'вниз' : '—');
              if (label === 'Д' || label === 'D1') flows.d = dirWord;
              else if (label === '4Ч' || label === 'H4') flows.h4 = dirWord;
              else if (label === '1Ч' || label === 'H1') flows.h1 = dirWord;
              else if (label === '15м' || label === 'M15') flows.m15 = dirWord;
            }
            if (Object.keys(flows).length > 0) result.flows = flows;
          }

          if (g && entryPrice && stopPrice) {
            const ziPlotBar = findPlotRowByIndexOrTime(plotsObj, entryTime, triggerIdx);
            const plotIdxOffset = (ziPlotBar && triggerIdx != null) ? (ziPlotBar.index - triggerIdx) : 0;
            const direct = analyzeChartGraphicsO1(g, plotsObj, plotIdxOffset, symKey, triggerIdx, isLong, entryPrice, stopPrice, entryExtreme, bars);
            result.matchedZone = direct.matchedZone;
            result.matchedSweep = direct.matchedSweep;
            result.detectedStructure = direct.detectedStructure;
            result.flowsHistorical = direct.isHistorical;
          }
        }

        if (!result.candleColor && title.includes('Матрёшка Модели входа') && !title.includes('Price Action')) {
          const plots = typeof s.plots === 'function' ? s.plots() : s._data?.bars?.();
          const foundBar = findPlotRowByIndexOrTime(plots, entryTime, triggerIdx);
          if (foundBar && plots) {
            const { colorIdx, modelIdx } = resolveMvPlotIndices(s, foundBar.val);
            const candidates = [foundBar.index, foundBar.index - 1];
            let matchedCode = 0, matchedModel = 0, isCont = false;
            for (const cIdx of candidates) {
              const r = getPlotRow(plots, cIdx);
              if (!r) continue;
              let code = Number(r[colorIdx] ?? 0);
              let mCode = Number(r[modelIdx] ?? 0);
              if (code === 0 && colorIdx !== 1 && r[1] != null && Math.abs(Number(r[1])) <= 3) {
                code = Number(r[1] ?? 0);
                mCode = Number(r[2] ?? 0);
              }
              if ((isLong && code > 0 && code <= 3) || (!isLong && code < 0 && code >= -3)) {
                matchedCode = code;
                matchedModel = mCode;
                isCont = (mCode === 18) || (isLong ? (r?.[7] === 1) : (r?.[8] === 1));
                break;
              }
            }
            const colorMap = { '3': 'Лонг 3', '2': 'Лонг 2', '1': 'Лонг 1', '-1': 'Шорт 1', '-2': 'Шорт 2', '-3': 'Шорт 3' };
            if (colorMap[String(matchedCode)]) {
              result.candleColor = colorMap[String(matchedCode)];
              result.modelCode = matchedModel;
              result.isContinuation = isCont;
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
    } catch (e) { return ''; }
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
    } catch (e) { return null; }
  }

  function readWatched(val) {
    try {
      if (val != null && typeof val === 'object' && typeof val.value === 'function') return val.value();
    } catch (e) {}
    return val;
  }

  function getChartTimezone(model) {
    try {
      const props = typeof model?.properties === 'function' ? model.properties() : model?._properties;
      const tzObj = props?.timezone || props?.childs?.()?.timezone;
      const tz = readWatched(tzObj);
      if (typeof tz === 'string' && tz.includes('/')) return tz;
    } catch (e) {}
    return Intl.DateTimeFormat().resolvedOptions().timeZone || 'Europe/Moscow';
  }

  function attachToolSubscriptions(tool) {
    if (!tool || hookedTools.has(tool)) return;
    hookedTools.add(tool);
    try {
      const onToolMutate = () => scheduleLiveSync(25);
      subscribeTvDelegate(tool._pointChanged, onToolMutate);
      subscribeTvDelegate(tool._fixedPointsChanged, onToolMutate);
      subscribeTvDelegate(tool._normalizedPointsChanged, onToolMutate);
      const rawProps = typeof tool.properties === 'function' ? tool.properties() : tool._properties;
      subscribeTvDelegate(rawProps?._changed, onToolMutate);
      const ch = typeof rawProps?.childs === 'function' ? rawProps.childs() : rawProps?._childs;
      if (ch) {
        subscribeTvDelegate(ch.stopLevel?._changed, onToolMutate);
        subscribeTvDelegate(ch.profitLevel?._changed, onToolMutate);
      }
    } catch (e) {}
  }

  function attachModelAndStudySubscriptions(model) {
    if (!model) return;
    try {
      const c = pageWindow._exposed_chartWidgetCollection;
      if (c?.activeChartWidget && typeof c.activeChartWidget.subscribe === 'function') {
        subscribeTvDelegate(c.activeChartWidget, () => scheduleLiveSync(40));
      }

      if (!hookedModels.has(model)) {
        hookedModels.add(model);
        const sel = typeof model.selection === 'function' ? model.selection() : model._selection;
        subscribeTvDelegate(sel?._changed, () => {
          const selId = detectCurrentlySelectedToolId(model);
          if (selId && selId !== lastChartSelectedId) {
            lastChartSelectedId = selId;
            lastSelectedToolId = selId;
            scheduleLiveSync(20);
          }
        });
      }

      const ms = typeof model.mainSeries === 'function' ? model.mainSeries() : model.mainSeries;
      const symInfo = (typeof ms?.symbolInfo === 'function' ? ms.symbolInfo() : ms?._symbolInfo) || {};
      const cleanSymbol = String(symInfo.name || 'XAUUSD').split(':').pop();
      const rawInterval = typeof ms?.interval === 'function' ? ms.interval() : (ms?._interval || '1');
      const symKey = `${cleanSymbol}_${rawInterval}`;

      const sources = typeof model.dataSources === 'function' ? model.dataSources() : [];
      for (const s of sources) {
        if (hookedStudies.has(s)) continue;
        const meta = typeof s.metaInfo === 'function' ? s.metaInfo() : s._metaInfo;
        const title = meta?.description || meta?.name || '';
        if (!title.includes('Матрёшка')) continue;

        hookedStudies.add(s);
        subscribeTvDelegate(s._dataUpdated, () => {
          syncGraphicsMemoryO1(model, symKey);
          scheduleLiveSync(50);
        });
        syncGraphicsMemoryO1(model, symKey);
      }

      const tools = typeof model.allLineTools === 'function' ? model.allLineTools() : model.allLineTools;
      const list = Array.isArray(tools) ? tools : Object.values(tools || {});
      for (const t of list) if (isSupportedTool(t)) attachToolSubscriptions(t);
    } catch (e) {}
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
      if (typeof model.previousSelectedDrawingOrGroupIds === 'function') {
        const prevIds = model.previousSelectedDrawingOrGroupIds();
        if (Array.isArray(prevIds) && prevIds.length > 0) {
          const tools = typeof model.allLineTools === 'function' ? model.allLineTools() : [];
          const list = Array.isArray(tools) ? tools : Object.values(tools || {});
          const match = list.find(t => isSupportedTool(t) && prevIds.includes(getToolId(t)));
          if (match) return getToolId(match);
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
    const std = mskH >= 15.0 ? 'Нью-Йорк' : (mskH >= 9.0 ? 'Лондон' : 'Азия');
    const eu_us = mskH >= 15.5 ? 'Америка' : (mskH >= 9.0 ? 'Европа' : 'Азия');
    return { std, eu_us };
  }

  function formatMskTime(timestamp) {
    const d = new Date(timestamp * 1000);
    return `${String((d.getUTCHours() + 3) % 24).padStart(2, '0')}:${String(d.getUTCMinutes()).padStart(2, '0')}`;
  }

  function formatLocalTime(timestamp, chartTz) {
    const d = new Date(timestamp * 1000);
    try {
      return new Intl.DateTimeFormat('ru-RU', { timeZone: chartTz || undefined, hour: '2-digit', minute: '2-digit', hour12: false }).format(d);
    } catch (e) {
      return `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
    }
  }

  function formatDate(timestamp, timeZone = 'Europe/Moscow') {
    const d = new Date(timestamp * 1000);
    try {
      return new Intl.DateTimeFormat('ru-RU', { timeZone, day: '2-digit', month: '2-digit', year: 'numeric' }).format(d);
    } catch (e) {
      return `${String(d.getDate()).padStart(2, '0')}.${String(d.getMonth() + 1).padStart(2, '0')}.${d.getFullYear()}`;
    }
  }

  function formatShortDate(timestamp, timeZone = 'Europe/Moscow') {
    return formatDate(timestamp, timeZone).slice(0, 5);
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

    const pt0 = tool?._points?.[0] || points?.[0];
    const rawPtTime = normalizeUnixSeconds(pt0?.time_t || pt0?.time?.getTime?.() || pt0?.time);
    if (rawPtTime) return rawPtTime;

    return Math.floor(Date.now() / 1000);
  }

  function analyzeBars(bars, startIndex, endIndex, isLong, entryPrice, stopPrice, targetPrice, points) {
    const risk = Math.abs(entryPrice - stopPrice);
    const price1R = isLong ? entryPrice + risk : entryPrice - risk;
    const price3R = isLong ? entryPrice + (risk * 3) : entryPrice - (risk * 3);

    if (!bars || startIndex == null) {
      return { outcome: 'Тейк', reached1R: true, reached3R: true, entryCandleExtreme: null, triggerIdx: startIndex ?? 0 };
    }

    const from = Math.min(startIndex, endIndex ?? startIndex);
    const to = Math.max(startIndex, endIndex ?? startIndex);

    
    let triggerIdx = null;
    const searchHead = Math.min(to, from + 3);
    let bestDiff = Infinity;
    const matchTol = Math.max(0.05, risk * 0.04);
    for (let i = from; i <= searchHead; i++) {
      const b = getPlotRow(bars, i);
      if (!b) continue;
      const op = Number(b[1]), hi = Number(b[2]), lo = Number(b[3]), cl = Number(b[4]);
      if (isNaN(hi) || isNaN(lo) || isNaN(cl)) continue;
      const cDiff = Math.abs(cl - entryPrice);
      const oDiff = !isNaN(op) ? Math.abs(op - entryPrice) : Infinity;
      const minDiff = Math.min(cDiff, oDiff);
      if (minDiff <= matchTol && minDiff < bestDiff) {
        bestDiff = minDiff;
        triggerIdx = i;
      }
    }

    
    if (triggerIdx === null && points?.[2]?.index != null) {
      const p2Idx = Number(points[2].index);
      const b2 = getPlotRow(bars, p2Idx);
      if (b2 && Number(b2[3]) <= entryPrice && Number(b2[2]) >= entryPrice) {
        triggerIdx = p2Idx;
      }
    }

    
    if (triggerIdx === null) {
      for (let i = from; i <= to; i++) {
        const b = getPlotRow(bars, i);
        if (!b) continue;
        const hi = Number(b[2]), lo = Number(b[3]);
        if (!isNaN(hi) && !isNaN(lo) && lo <= entryPrice && hi >= entryPrice) {
          triggerIdx = i;
          break;
        }
      }
    }
    if (triggerIdx === null) {
      for (let i = from - 1; i >= from - 3; i--) {
        const b = getPlotRow(bars, i);
        if (!b) continue;
        const hi = Number(b[2]), lo = Number(b[3]);
        if (!isNaN(hi) && !isNaN(lo) && lo <= entryPrice && hi >= entryPrice) {
          triggerIdx = i;
          break;
        }
      }
    }
    if (triggerIdx === null) triggerIdx = from;

    let sweepExtreme = null;
    for (let i = triggerIdx - 2; i <= triggerIdx; i++) {
      const b = getPlotRow(bars, i);
      if (!b) continue;
      const hi = Number(b[2]), lo = Number(b[3]);
      if (isNaN(hi) || isNaN(lo)) continue;
      sweepExtreme = sweepExtreme === null
        ? (isLong ? lo : hi)
        : (isLong ? Math.min(sweepExtreme, lo) : Math.max(sweepExtreme, hi));
    }

    let reached1R = false, reached3R = false, outcome = 'Безубыток', outcomeDecided = false;

    for (let i = triggerIdx; i <= to; i++) {
      const bar = getPlotRow(bars, i);
      if (!bar) continue;
      const high = Number(bar[2]), low = Number(bar[3]), close = Number(bar[4]);
      if (isNaN(high) || isNaN(low)) continue;

      const checkHigh = (i === triggerIdx) ? Math.max(entryPrice, close) : high;
      const checkLow = (i === triggerIdx) ? Math.min(entryPrice, close) : low;

      if (isLong ? (checkHigh >= price1R) : (checkLow <= price1R)) reached1R = true;
      if (isLong ? (checkHigh >= price3R) : (checkLow <= price3R)) reached3R = true;

      if (!outcomeDecided && (i > triggerIdx || Math.abs(close - entryPrice) > risk * 0.25)) {
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

    if (!outcomeDecided && points?.[3]?.price != null) {
      const exitP = Number(points[3].price);
      const tol = Math.max(0.05, risk * 0.05);
      if (Math.abs(exitP - targetPrice) <= tol) {
        outcome = 'Тейк';
        if (Math.abs(targetPrice - entryPrice) >= risk * 0.98) reached1R = true;
        if (Math.abs(targetPrice - entryPrice) >= risk * 2.95) reached3R = true;
      } else if (Math.abs(exitP - stopPrice) <= tol) {
        outcome = 'Стоп';
      }
    }

    return { outcome, reached1R, reached3R, entryCandleExtreme: sweepExtreme, triggerIdx };
  }

  function parseSingleTool(tool, model, ms, bars, symInfo, priceStep) {
    attachToolSubscriptions(tool);

    const rawProps = typeof tool.properties === 'function' ? tool.properties() : (tool.properties || tool._properties);
    let props = rawProps;
    try {
      if (rawProps && typeof rawProps.state === 'function') props = rawProps.state();
    } catch (e) {}

    const getPropVal = (key) => {
      const v1 = readWatched(props?.[key]);
      if (v1 != null) return v1;
      const v2 = readWatched(rawProps?.[key]);
      if (v2 != null) return v2;
      const ch = typeof rawProps?.childs === 'function' ? rawProps.childs() : rawProps?._childs;
      return readWatched(ch?.[key]);
    };

    let points = [];
    try {
      const rawPts = typeof tool.points === 'function' ? tool.points() : (tool._points || tool.points || []);
      if (Array.isArray(rawPts)) points = rawPts;
    } catch (e) {
      points = Array.isArray(tool._points) ? tool._points : [];
    }

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
      const stopLevel = Number(getPropVal('stopLevel') ?? 0);
      const profitLevel = Number(getPropVal('profitLevel') ?? 0);
      stop = isLong ? entry - (stopLevel * priceStep) : entry + (stopLevel * priceStep);
      target = isLong ? entry + (profitLevel * priceStep) : entry - (profitLevel * priceStep);
      rrNum = (stopLevel > 0) ? (profitLevel / stopLevel) : 0;
      barAnalysis = analyzeBars(bars, startIdx, endIdx, isLong, entry, stop, target, points);
    }

    const effectiveIdx = barAnalysis.triggerIdx ?? startIdx;
    const time = resolveBarTimestamp(model, bars, effectiveIdx, tool, points);
    const chartTz = getChartTimezone(model);

    const rawSymbol = getPropVal('symbol') || symInfo.name || 'XAUUSD';
    const cleanSymbol = String(rawSymbol).split(':').pop();
    const rawInterval = getPropVal('interval') || (typeof ms?.interval === 'function' ? ms.interval() : '15');
    const symKey = `${cleanSymbol}_${rawInterval}`;

    const decimals = symInfo.pricescale
      ? Math.max(0, Math.round(Math.log10(symInfo.pricescale)))
      : (priceStep < 1 ? Math.min(8, Number(priceStep).toFixed(8).replace(/0+$/, '').split('.')[1]?.length || 2) : 2);

    const entryR = Number(entry.toFixed(decimals));
    const stopR = Number(stop.toFixed(decimals));
    const targetR = Number(target.toFixed(decimals));
    const extremeR = barAnalysis.entryCandleExtreme != null ? Number(Number(barAnalysis.entryCandleExtreme).toFixed(decimals)) : '';
    const sess = detectSession(time);
    const id = getToolId(tool);

    const ind = readMatryoshkaIndicators(model, symKey, time, effectiveIdx, startIdx, endIdx, isLong, entryR, stopR, extremeR, bars);
    if (ind.matchedSweep && ind.matchedSweep.price != null && !isNaN(Number(ind.matchedSweep.price))) {
      ind.matchedSweep.price = Number(Number(ind.matchedSweep.price).toFixed(decimals));
    }

    const signature = `${id}_${startIdx}_${endIdx}_${entryR}_${stopR}_${targetR}`;

    return {
      id, signature, serverTime: Number(tool._serverUpdateTime || 0), time,
      symbol: cleanSymbol, timeframeLat: formatTf(rawInterval, 'lat'), timeframeCyr: formatTf(rawInterval, 'cyr'),
      direction: isLong ? 'LONG' : 'SHORT', entryPrice: entryR, stopPrice: stopR, targetPrice: targetR,
      entryExtreme: extremeR, upperBound: Math.max(entryR, stopR), lowerBound: Math.min(entryR, stopR),
      riskPts: Number(Math.abs(entryR - stopR).toFixed(decimals)),
      potPts: Number(Math.abs(targetR - entryR).toFixed(decimals)),
      riskReward: Number(rrNum.toFixed(2)),
      candleDateMsk: formatDate(time, 'Europe/Moscow'), candleDateLocal: formatDate(time, chartTz),
      shortDateMsk: formatShortDate(time, 'Europe/Moscow'), shortDateLocal: formatShortDate(time, chartTz),
      mskTime: formatMskTime(time), localTime: formatLocalTime(time, chartTz),
      sessionStd: sess.std, sessionEuUs: sess.eu_us,
      autoResult: barAnalysis.outcome, reached1R: barAnalysis.reached1R, reached3R: barAnalysis.reached3R, ind
    };
  }

  function scanChartTrades(silent = false, preferChartSelection = false) {
    try {
      const model = getTvModel();
      if (!model) {
        if (!silent) showToast('График ещё не загружен', 'error');
        return false;
      }

      attachModelAndStudySubscriptions(model);

      const ms = typeof model.mainSeries === 'function' ? model.mainSeries() : model.mainSeries;
      const bars = typeof ms?.bars === 'function' ? ms.bars() : ms?.bars;
      const tools = typeof model.allLineTools === 'function' ? model.allLineTools() : model.allLineTools;
      const list = Array.isArray(tools) ? tools : Object.values(tools || {});

      const matchingTools = list.filter(isSupportedTool);
      if (!matchingTools.length) {
        if (!silent) showToast('Нарисуйте на графике инструмент «Длинная» или «Короткая позиция»', 'error');
        return false;
      }

      let symInfo = {};
      try { symInfo = (typeof ms?.symbolInfo === 'function' ? ms.symbolInfo() : ms?._symbolInfo) || {}; } catch (e) {}

      const rawStep = typeof ms?.priceStep === 'function' ? ms.priceStep() : ms?.priceStep;
      const priceStep = (typeof rawStep === 'number' && !isNaN(rawStep) && rawStep > 0)
        ? rawStep
        : (symInfo.minmov && symInfo.pricescale ? symInfo.minmov / symInfo.pricescale : 0.01);

      const parsed = [];
      for (const t of matchingTools) {
        try { parsed.push(parseSingleTool(t, model, ms, bars, symInfo, priceStep)); } catch (toolErr) {}
      }
      if (!parsed.length) {
        if (!silent) showToast('Не удалось прочитать координаты позиции на графике', 'error');
        return false;
      }

      parsed.sort((a, b) => a.time - b.time);
      allChartTrades = parsed;

      const chartSelId = detectCurrentlySelectedToolId(model);
      if (preferChartSelection && chartSelId) {
        lastChartSelectedId = chartSelId;
        lastSelectedToolId = chartSelId;
      } else if (chartSelId && chartSelId !== lastChartSelectedId) {
        lastChartSelectedId = chartSelId;
        lastSelectedToolId = chartSelId;
      }

      let targetId = (lastSelectedToolId && parsed.some(p => p.id === lastSelectedToolId)) ? lastSelectedToolId : chartSelId;
      if (!targetId && lastHoveredToolId && parsed.some(p => p.id === lastHoveredToolId)) targetId = lastHoveredToolId;
      if (!targetId) {
        let newest = parsed[parsed.length - 1];
        for (const p of parsed) if (p.serverTime > newest.serverTime) newest = p;
        targetId = newest.id;
      }

      const idx = parsed.findIndex(p => p.id === targetId);
      activeTradeIdx = idx >= 0 ? idx : (parsed.length - 1);
      currentTrade = allChartTrades[activeTradeIdx];

      if (currentTrade?.ind?.flows) {
        const fl = currentTrade.ind.flows;
        saveState({
          flowD: fl.d || state.flowD, flow4H: fl.h4 || state.flow4H,
          flow1H: fl.h1 || state.flow1H, flow15M: fl.m15 || state.flow15M
        });
      }
      return true;
    } catch (e) {
      if (!silent) showToast('Не удалось прочитать инструмент на графике', 'error');
      return false;
    }
  }

  function getActiveSep() { return state.outputProfile === 'en_form' ? '.' : ','; }
  function numStr(val) {
    if (val == null || val === '') return '';
    const s = String(val);
    return getActiveSep() === ',' ? s.replace(/\./g, ',') : s.replace(/,/g, '.');
  }
  function cleanText(s) { return String(s || '').replace(/[\t\r\n]+/g, ' ').trim(); }

  function getActiveColumns() {
    if (state.presetKey === 'custom' && Array.isArray(state.customColumns) && state.customColumns.length) {
      return [...state.customColumns];
    }
    return [...(PRESET_SCHEMAS[state.presetKey] || PRESET_SCHEMAS.bt_main).columns];
  }
  function getActivePasteCol() {
    return state.presetKey === 'custom' ? (state.customPasteCol || 'B') : (PRESET_SCHEMAS[state.presetKey]?.pasteCol || 'B');
  }
  function getActiveStartRow() {
    return state.presetKey === 'custom' ? (Number(state.customStartRow) || 2) : (PRESET_SCHEMAS[state.presetKey]?.startRow || 6);
  }
  function getActiveSession() {
    if (manualSessionOverride) return manualSessionOverride;
    if (!currentTrade) return 'Лондон';
    return (PRESET_SCHEMAS[state.presetKey]?.sessionStyle === 'eu_us') ? currentTrade.sessionEuUs : currentTrade.sessionStd;
  }

  function colLetterToIndex(letter) {
    const clean = String(letter || 'B').toUpperCase().replace(/[^A-Z]/g, '') || 'B';
    let idx = 0;
    for (let i = 0; i < clean.length; i++) idx = idx * 26 + (clean.charCodeAt(i) - 64);
    return idx - 1;
  }
  function colIndexToLetter(idx) {
    let col = '', n = idx;
    while (n >= 0) { col = String.fromCharCode((n % 26) + 65) + col; n = Math.floor(n / 26) - 1; }
    return col;
  }

  function buildDynamicFormula(tok, cols, pasteCol, startRow, fallbackVal) {
    const prof = state.outputProfile;
    const isRuNames = (prof === 'ru_form');
    const s = (prof === 'en_form') ? ',' : ';';

    const IF = isRuNames ? 'ЕСЛИ' : 'IF';
    const OR = isRuNames ? 'ИЛИ' : 'OR';
    const IFERROR = isRuNames ? 'ЕСЛИОШИБКА' : 'IFERROR';
    const COUNT = isRuNames ? 'СЧЁТ' : 'COUNT';
    const COUNTIF = isRuNames ? 'СЧЁТЕСЛИ' : 'COUNTIF';
    const SUM = isRuNames ? 'СУММ' : 'SUM';
    const MAX = isRuNames ? 'МАКС' : 'MAX';
    const INT = isRuNames ? 'ЦЕЛОЕ' : 'INT';
    const ROUND = isRuNames ? 'ОКРУГЛ' : 'ROUND';

    const baseOffset = colLetterToIndex(pasteCol);
    const findCol = (...names) => {
      for (const name of names) {
        const i = cols.indexOf(name);
        if (i !== -1) return colIndexToLetter(baseOffset + i);
      }
      return null;
    };
    const ref = (colLet) => isRuNames ? `ИНДЕКС(${colLet}:${colLet}${s}СТРОКА())` : `INDEX(${colLet}:${colLet}${s}ROW())`;
    const vRef = (...names) => { const colLet = findCol(...names); return colLet ? ref(colLet) : null; };

    if (tok === 'risk_pts') {
      const en = vRef('entry'), st = vRef('stop');
      if (!en || !st) return fallbackVal;
      return state.presetKey === 'bt_30' ? `=${IF}(${OR}(${en}=""${s}${st}="")${s}""${s}ABS(${en}-${st}))` : `=${IFERROR}(ABS(${en}-${st})${s}"")`;
    }
    if (tok === 'pot_pts') {
      const en = vRef('entry'), tg = vRef('target');
      return (!en || !tg) ? fallbackVal : `=${IFERROR}(ABS(${tg}-${en})${s}"")`;
    }
    if (tok === 'rr') {
      const pot = vRef('pot_pts'), rsk = vRef('risk_pts'), en = vRef('entry'), st = vRef('stop'), tg = vRef('target');
      if (state.presetKey === 'bt_30' && en && tg && rsk) return `=${IF}(${OR}(${rsk}=""${s}${rsk}=0${s}${tg}="")${s}""${s}ABS(${tg}-${en})/${rsk})`;
      if (pot && rsk) return `=${IFERROR}(${pot}/${rsk}${s}"")`;
      if (en && st && tg) return `=${IFERROR}(ABS(${tg}-${en})/ABS(${en}-${st})${s}"")`;
      return fallbackVal;
    }
    if (tok === 'lots_30') {
      const rsk = vRef('risk_pts');
      return (!rsk || state.presetKey !== 'bt_30') ? fallbackVal : `=${IF}(${OR}(${rsk}=""${s}${rsk}=0)${s}""${s}${INT}(('Параметры'!$B$5/${rsk}/'Параметры'!$B$6)/'Параметры'!$B$7)*'Параметры'!$B$7)`;
    }
    if (tok === 'r_out') {
      const rrRef = vRef('rr'), res30Ref = vRef('res_30'), manRRef = vRef('manual_r_30');
      if (res30Ref && rrRef) return `=${IF}(${res30Ref}=""${s}""${s}${IF}(${res30Ref}="цель"${s}${rrRef}${s}${IF}(${res30Ref}="стоп"${s}-1${s}${IF}(${res30Ref}="безубыток"${s}0${s}${manRRef || '""'}))))`;
      const resRef = vRef('result');
      return (resRef && rrRef) ? `=${IF}(${resRef}="Тейк"${s}${rrRef}${s}${IF}(${resRef}="Стоп"${s}-1${s}${IF}(${resRef}="Безубыток"${s}0${s}"")))` : fallbackVal;
    }
    if (tok === 'profit_usd') {
      const rOutRef = vRef('r_out'), lotsRef = vRef('lots_30'), rskRef = vRef('risk_pts');
      return (!rOutRef || !lotsRef || !rskRef || state.presetKey !== 'bt_30') ? fallbackVal : `=${IF}(${OR}(${rOutRef}=""${s}${lotsRef}="")${s}""${s}${rOutRef}*${lotsRef}*'Параметры'!$B$6*${rskRef})`;
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
      return (!c1 || !c4) ? fallbackVal : `=${COUNTIF}(${ref(c1)}:${ref(c4)}${s}"да")`;
    }
    if (tok === 'valid_aoi') {
      const c1 = findCol('f_ret'), c4 = findCol('f_liq');
      return (!c1 || !c4) ? fallbackVal : `=${IF}(${COUNTIF}(${ref(c1)}:${ref(c4)}${s}"да")>=2${s}"ДА"${s}"нет")`;
    }
    if (tok === 'lots_off') {
      const en = vRef('entry'), st = vRef('stop'), rskPct = vRef('risk_pct_col');
      return (!en || !st || !rskPct || state.presetKey !== 'office') ? fallbackVal : `=${IF}(${OR}(${en}=""${s}${st}=""${s}${en}=${st})${s}""${s}${ROUND}('Итоги месяца'!$B$2*${rskPct}/(ABS(${en}-${st})*100*'Итоги месяца'!$B$3)${s}2))`;
    }
    if (tok === 'r_out_off') {
      const res = vRef('res_off'), sfPct = vRef('safe_pct'), sfRr = vRef('safe_rr'), rr = vRef('rr');
      return (!res || !sfPct || !sfRr || !rr) ? fallbackVal : `=${IF}(${res}=""${s}""${s}${IF}(${res}="Стоп"${s}-1${s}${IF}(${res}="Тейк"${s}${IF}(${sfPct}=""${s}0${s}${sfPct})*${IF}(${sfRr}=""${s}0${s}${sfRr})+(1-${IF}(${sfPct}=""${s}0${s}${sfPct}))*${rr}${s}${IF}(${res}="Сейф/БУ"${s}${IF}(${sfPct}=""${s}0${s}${sfPct})*${IF}(${sfRr}=""${s}0${s}${sfRr})${s}${IF}(${res}="Бу"${s}0${s}"")))))`;
    }
    if (tok === 'prof_pct_off') {
      const rOutOff = vRef('r_out_off'), rskPct = vRef('risk_pct_col');
      return (!rOutOff || !rskPct) ? fallbackVal : `=${IF}(${rOutOff}=""${s}""${s}${rOutOff}*${rskPct})`;
    }
    if (tok === 'move_pts') {
      const res = vRef('res_off', 'result'), dir = vRef('dir_title'), en = vRef('entry'), st = vRef('stop'), tg = vRef('target');
      return (!res || !dir || !en || !st || !tg) ? fallbackVal : `=${IF}(${res}=""${s}""${s}${IF}(${dir}="Боковик"${s}0${s}${IF}(${res}="Тейк"${s}${IF}(${dir}="Лонг"${s}${tg}-${en}${s}${en}-${tg})${s}${IF}(${res}="Стоп"${s}${IF}(${dir}="Лонг"${s}${st}-${en}${s}${en}-${st})${s}0))))`;
    }
    return fallbackVal;
  }

  const addTokenOptionsHtml = Object.entries(TOKEN_LABELS).map(([k, v]) => `<option value="${k}">+ Добавить поле: ${v}</option>`).join('');

  const uiHtml = `
  <div id="bt-toast" style="display:none;position:fixed;bottom:24px;right:24px;z-index:10000000;background:#141417;color:#f4f4f5;padding:11px 16px;border-radius:8px;border-left:4px solid #089981;box-shadow:0 12px 32px rgba(0,0,0,0.9);font-family:-apple-system,BlinkMacSystemFont,'Inter','Segoe UI',Roboto,sans-serif;font-size:12px;font-weight:500;max-width:370px;"></div>
  <button type="button" id="bt-fab-btn" style="position:fixed;bottom:18px;left:58px;z-index:9999990;display:flex;align-items:center;gap:6px;background:#141417;color:#e5c158;border:1px solid rgba(229,193,88,0.45);border-radius:7px;padding:6px 12px;font-family:-apple-system,BlinkMacSystemFont,'Inter',sans-serif;font-size:12px;font-weight:600;cursor:pointer;box-shadow:0 4px 16px rgba(0,0,0,0.75);transition:0.15s;">
    ${LOGO_SVG}<span>Матрёшка</span><span style="color:#9ca3af;font-weight:400;font-size:11px;">(${isMac ? '⌃Q' : 'F2'})</span>
  </button>
  <div id="bt-modal" class="bt-no-scroll" tabindex="-1" style="display:none;outline:none;position:fixed;top:${state.posTop}px;right:${state.posRight}px;z-index:9999999;background:#0d0d10;color:#f4f4f5;padding:13px;border-radius:12px;box-shadow:0 22px 55px rgba(0,0,0,0.92),0 0 0 1px rgba(229,193,88,0.25);font-family:-apple-system,BlinkMacSystemFont,'Inter','Segoe UI',Roboto,sans-serif;width:394px;max-height:95vh;overflow-y:auto;box-sizing:border-box;user-select:none;">
    <div id="bt-header-drag" style="display:flex;justify-content:space-between;align-items:center;padding-bottom:9px;margin-bottom:9px;border-bottom:1px solid #222228;cursor:move;">
      <div style="display:flex;align-items:center;gap:7px;">${LOGO_SVG}<span id="bt-active-tpl-title" style="font-size:13px;font-weight:700;color:#f4f4f5;"></span></div>
      <div style="display:flex;align-items:center;gap:4px;">
        <div style="display:flex;align-items:center;background:#16161a;border:1px solid #2a2a32;border-radius:5px;padding:1px 3px;" title="Переключение между сделками на графике (стрелки ← и →)">
          <button type="button" id="bt-prev-trade" class="bt-icon-btn">◀</button>
          <span id="bt-trade-counter" style="font-size:11px;font-weight:700;color:#e5c158;padding:0 5px;">1/1</span>
          <button type="button" id="bt-next-trade" class="bt-icon-btn">▶</button>
        </div>
        <button type="button" id="bt-open-settings" class="bt-top-btn" title="Настройки журнала">⚙️</button>
        <button type="button" id="bt-min-btn" class="bt-top-btn" title="Свернуть или развернуть окно">—</button>
        <button type="button" id="bt-close-x" class="bt-top-btn" title="Закрыть окно (Esc)">✕</button>
      </div>
    </div>
    <div id="bt-trade-card" style="background:#141418;border:1px solid #26262e;border-radius:8px;padding:8px 10px;margin-bottom:9px;transition:border-color 0.2s ease;">
      <div style="display:flex;justify-content:space-between;align-items:center;gap:4px;margin-bottom:6px;white-space:nowrap;">
        <div style="display:flex;align-items:center;gap:6px;font-size:11.5px;min-width:0;overflow:hidden;">
          <b id="bt-info-sym" style="color:#ffffff;white-space:nowrap;">XAUUSD · M1</b>
          <button type="button" id="bt-session-pill" title="Торговая сессия по времени входа. Нажмите, чтобы изменить" style="background:rgba(229,193,88,0.14);color:#f3d26b;border:1px solid rgba(229,193,88,0.4);border-radius:4px;padding:1px 6px;font-size:10.5px;font-weight:600;cursor:pointer;white-space:nowrap;">⚡ Европа</button>
          <button type="button" id="bt-info-date" title="Время свечи фактического входа. Нажмите для переключения МСК (UTC+3) / пояса графика" style="background:transparent;border:1px solid transparent;border-radius:4px;color:#9ca3af;font-size:10.5px;padding:1px 4px;cursor:pointer;white-space:nowrap;transition:0.12s;"></button>
        </div>
        <span id="bt-info-badge" style="background:#089981;color:#fff;font-size:10.5px;padding:2px 7px;border-radius:4px;font-weight:700;white-space:nowrap;flex-shrink:0;"></span>
      </div>
      <div id="bt-quick-copy-prices" title="Нажмите, чтобы скопировать только цены: Вход, Стоп и Цель" style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:4px;background:#0e0e11;padding:6px 8px;border-radius:5px;cursor:pointer;border:1px solid #222228;font-size:11.5px;white-space:nowrap;transition:0.12s;">
        <div><span style="color:#9ca3af;font-size:10px;">Вход </span><b id="bt-info-en" style="color:#ffffff;"></b></div>
        <div style="text-align:center;"><span style="color:#9ca3af;font-size:10px;">Стоп </span><b id="bt-info-st" style="color:#ffffff;"></b></div>
        <div style="text-align:right;"><span style="color:#9ca3af;font-size:10px;">Цель </span><b id="bt-info-tg" style="color:#f3d26b;"></b></div>
      </div>
    </div>
    <div id="bt-collapsible-body">
      <div id="bt-settings-drawer" style="display:none;background:#141418;border:1px solid #3d3522;border-radius:8px;padding:11px;">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;padding-bottom:7px;border-bottom:1px solid #26262e;">
          <span style="font-size:11.5px;font-weight:700;color:#f3d26b;">⚙️ НАСТРОЙКИ ЖУРНАЛА</span>
          <button type="button" id="bt-close-settings-btn" style="background:#e5c158;color:#09090b;border:none;border-radius:5px;padding:4px 10px;font-size:11px;font-weight:700;cursor:pointer;">← Готово</button>
        </div>
        <div class="bt-lbl">ШАБЛОН ТАБЛИЦЫ</div>
        <select id="bt-preset-select" class="bt-sel" style="border-color:#e5c158;color:#f3d26b;font-weight:600;margin-bottom:10px;">
          <option value="bt_main">📘 1. Бэктест «Матрёшка»</option>
          <option value="bt_30">📊 2. Бэктест 30 сделок (30_trades.xlsx)</option>
          <option value="dz_poi">🎯 3. ДЗ: Журнал POI</option>
          <option value="dz_aoi">🗺️ 4. ДЗ: Таблица AOI</option>
          <option value="office">🏢 5. Торговый офис (Недели 1–4)</option>
          <option value="prices_only">⚡ 6. Только 3 цены (Вход / Стоп / Цель)</option>
          <option value="custom">✨ 7. Свой набор колонок</option>
        </select>
        <button type="button" id="bt-toggle-cols-btn" class="bt-mini-pill" style="width:100%;text-align:center;padding:5px;">🔧 Настроить колонки таблицы...</button>
        <div id="bt-cols-editor-wrap" style="display:none;margin-top:8px;padding-top:8px;border-top:1px solid #26262e;">
          <div id="bt-custom-anchors-grid" style="display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:7px;">
            <div>
              <div class="bt-lbl">ПЕРВЫЙ СТОЛБЕЦ ВСТАВКИ</div>
              <select id="bt-custom-paste-col" class="bt-sel" style="margin-bottom:0;height:26px;">
                <option value="A">Столбец A (с 1-й кол.)</option>
                <option value="B">Столбец B (после №)</option>
                <option value="C">Столбец C</option>
                <option value="D">Столбец D</option>
              </select>
            </div>
            <div id="bt-start-row-box">
              <div class="bt-lbl" title="Нужна только для формул Накопительного R и Просадки ($P$6)">СТРОКА НАЧАЛА СЧЁТА R</div>
              <input type="number" id="bt-custom-start-row" min="1" max="500" value="${getActiveStartRow()}" style="width:100%;height:26px;box-sizing:border-box;background:#16161b;color:#f4f4f5;border:1px solid #282830;border-radius:5px;padding:0 6px;font-size:11px;outline:none;">
            </div>
          </div>
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:5px;white-space:nowrap;">
            <span style="font-size:10px;font-weight:700;color:#9ca3af;">ПОРЯДОК КОЛОНОК (ПЕРЕТАЩИТЕ МЫШКОЙ):</span>
            <button type="button" id="bt-reset-schema" class="bt-mini-pill" style="padding:2px 7px;font-size:10px;white-space:nowrap;">↺ По умолчанию</button>
          </div>
          <div id="bt-schema-pills" class="bt-no-scroll" style="display:flex;flex-wrap:wrap;gap:4px;max-height:135px;overflow-y:auto;padding:6px;background:#0b0b0e;border-radius:6px;border:1px solid #27272e;margin-bottom:6px;"></div>
          <select id="bt-add-token-select" class="bt-sel" style="margin-bottom:0;font-size:11px;color:#f3d26b;">
            <option value="">+ Добавить поле в таблицу...</option>${addTokenOptionsHtml}
          </select>
        </div>
      </div>
      <div id="bt-main-view">
        <div id="bt-dynamic-fields"></div>
        <div id="bt-extra-inputs" style="position:relative;display:flex;gap:6px;margin-top:4px;margin-bottom:7px;">
          <div style="position:relative;flex:1;display:flex;align-items:center;">
            <input type="text" id="bt-inp-comment" autocomplete="off" placeholder="Комментарий к сделке (Tab)..." style="width:100%;box-sizing:border-box;background:#141418;border:1px solid #2a2a32;color:#ffffff;padding:7px 44px 7px 9px;border-radius:6px;font-size:11.5px;outline:none;">
            <button type="button" id="bt-clear-comment" title="Очистить комментарий" style="position:absolute;right:22px;background:transparent;border:none;color:#9ca3af;cursor:pointer;font-size:11px;padding:2px 4px;">✕</button>
            <button type="button" id="bt-toggle-presets" title="Выбрать частый комментарий" style="position:absolute;right:4px;background:transparent;border:none;color:#e5c158;cursor:pointer;font-size:11px;padding:2px 5px;">▾</button>
          </div>
          <div id="bt-comment-menu" class="bt-no-scroll" style="display:none;position:absolute;bottom:36px;left:0;width:285px;max-height:185px;overflow-y:auto;background:#18181e;border:1px solid #3d3522;border-radius:6px;box-shadow:0 10px 25px rgba(0,0,0,0.85);z-index:100;padding:4px;">
            <div class="bt-cm-item" style="padding:5px 8px;font-size:11px;color:#e4e4e7;cursor:pointer;border-radius:4px;">7 из 7 — по закрытой свече</div>
            <div class="bt-cm-item" style="padding:5px 8px;font-size:11px;color:#e4e4e7;cursor:pointer;border-radius:4px;">рано — приманка не снята</div>
            <div class="bt-cm-item" style="padding:5px 8px;font-size:11px;color:#e4e4e7;cursor:pointer;border-radius:4px;">зона против потока</div>
            <div class="bt-cm-item" style="padding:5px 8px;font-size:11px;color:#e4e4e7;cursor:pointer;border-radius:4px;">стоп за тень — не за зону</div>
            <div class="bt-cm-item" style="padding:5px 8px;font-size:11px;color:#e4e4e7;cursor:pointer;border-radius:4px;">цель дальше ближайшей 0.5</div>
            <div class="bt-cm-item" style="padding:5px 8px;font-size:11px;color:#e4e4e7;cursor:pointer;border-radius:4px;">догонял движение</div>
            <div class="bt-cm-item" style="padding:5px 8px;font-size:11px;color:#e4e4e7;cursor:pointer;border-radius:4px;">внутри проторговки</div>
            <div class="bt-cm-item" style="padding:5px 8px;font-size:11px;color:#e4e4e7;cursor:pointer;border-radius:4px;">по бледной свече</div>
            <div class="bt-cm-item" style="padding:5px 8px;font-size:11px;color:#e4e4e7;cursor:pointer;border-radius:4px;">зона с касанием 2</div>
            <div class="bt-cm-item" style="padding:5px 8px;font-size:11px;color:#e4e4e7;cursor:pointer;border-radius:4px;">по правилам — рынок не пошёл</div>
          </div>
          <button type="button" id="bt-screen-pill" title="Нажмите Alt+S на графике — ссылка прикрепится автоматически" style="background:#16161a;color:#d4d4d8;border:1px solid #2a2a32;border-radius:6px;padding:0 10px;font-size:11.5px;font-weight:500;cursor:pointer;white-space:nowrap;transition:0.15s;">📸 Скрин</button>
        </div>
        <div style="background:#111115;border:1px solid #22222a;border-radius:6px;padding:5px 6px;margin-bottom:6px;">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;">
            <span id="bt-preview-hint" style="font-size:9.5px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:0.3px;">СТРОКА ДЛЯ ВСТАВКИ</span>
            <button type="button" id="bt-toggle-fix-help" style="background:transparent;border:none;color:#e5c158;font-size:10px;font-weight:600;cursor:pointer;padding:0;">? Ошибка в таблице?</button>
          </div>
          <div id="bt-row-preview" class="bt-hide-scroll" style="display:flex;gap:4px;overflow-x:auto;"></div>
        </div>
        <div id="bt-fix-help-box" style="display:none;background:#181612;border:1px solid #4a3f23;border-radius:6px;padding:7px 8px;margin-bottom:7px;font-size:10.5px;color:#d4d4d8;">
          <div style="font-weight:700;color:#f3d26b;margin-bottom:5px;">Если после вставки в таблице появилась ошибка:</div>
          <div style="display:flex;flex-direction:column;gap:4px;">
            <button type="button" class="bt-fix-action" data-fix="gs_ru" style="text-align:left;background:#121216;color:#e4e4e7;border:1px solid #2e2e38;border-radius:4px;padding:5px 7px;font-size:10px;cursor:pointer;">🔴 Ошибка <b style="color:#f87171;">#ERROR!</b> в Google Таблице → <u style="color:#f3d26b;">Переключить на Google Docs RU и скопировать</u></button>
            <button type="button" class="bt-fix-action" data-fix="ru_form" style="text-align:left;background:#121216;color:#e4e4e7;border:1px solid #2e2e38;border-radius:4px;padding:5px 7px;font-size:10px;cursor:pointer;">🔴 Ошибка <b style="color:#f87171;">#ERROR!</b> в программе Excel → <u style="color:#f3d26b;">Переключить на Excel RU и скопировать</u></button>
            <button type="button" class="bt-fix-action" data-fix="en_form" style="text-align:left;background:#121216;color:#e4e4e7;border:1px solid #2e2e38;border-radius:4px;padding:5px 7px;font-size:10px;cursor:pointer;">🔴 Шести-семи-значные числа, формулы не принимаются из-за ";" в программе Excel → <u style="color:#f3d26b;">Переключить на EN и скопировать</u></button>
            <div style="font-size:9.5px;color:#9ca3af;padding:2px 2px 0;">💡 Если пишет <b>#ЗНАЧ! (#VALUE!)</b> — выделите ячейку в столбце <b id="bt-fix-col-lbl" style="color:#f3d26b;">B</b> и нажмите Ctrl+V.</div>
          </div>
        </div>
        <div style="display:grid;grid-template-columns:1.1fr 1fr 0.9fr;gap:3px;margin-bottom:7px;background:#09090c;padding:2px;border-radius:6px;border:1px solid #22222a;">
          <button type="button" class="bt-prof-btn" data-prof="gs_ru" title="Для Google Таблиц: числа через запятую, функции IF / INDEX">📊 Google Docs RU</button>
          <button type="button" class="bt-prof-btn" data-prof="ru_form" title="Для программы Excel на русском языке: числа через запятую, функции ЕСЛИ / ИНДЕКС">📗 Excel RU</button>
          <button type="button" class="bt-prof-btn" data-prof="en_form" title="Числа через точку, функции IF / INDEX">🌐 EN</button>
        </div>
        <div style="display:flex;gap:6px;align-items:stretch;">
          <button type="button" id="bt-main-copy-btn" style="flex:1;background:linear-gradient(180deg,#ebd074 0%,#cfa538 100%);color:#09090b;border:none;padding:9px 10px;border-radius:6px;cursor:pointer;font-weight:700;font-size:12px;display:flex;justify-content:center;align-items:center;gap:6px;box-shadow:0 3px 12px rgba(229,193,88,0.28);white-space:nowrap;">
            <span id="bt-main-btn-text">Скопировать строку (в стлб. B)</span>
            <kbd style="background:rgba(0,0,0,0.18);color:#09090b;padding:1px 5px;border-radius:3px;font-size:10.5px;font-family:inherit;">↵</kbd>
          </button>
          <div style="display:flex;align-items:center;background:#141418;border:1px solid #2a2a32;border-radius:6px;overflow:hidden;">
            <button type="button" id="bt-copy-batch-btn" title="Скопировать все сохранённые сделки одним списком" style="background:transparent;color:#e4e4e7;border:none;padding:0 9px;height:100%;font-size:11px;font-weight:600;cursor:pointer;display:flex;align-items:center;gap:4px;white-space:nowrap;">📦 <span id="bt-batch-count" style="color:#f3d26b;">0</span></button>
            <button type="button" id="bt-clear-batch-btn" title="Очистить список сохранённых сделок" style="background:rgba(242,54,69,0.14);color:#f23645;border:none;border-left:1px solid #2a2a32;padding:0 7px;height:100%;font-size:11px;cursor:pointer;">✕</button>
          </div>
        </div>
      </div>
    </div>
  </div>
  <style>
    .bt-no-scroll::-webkit-scrollbar{width:4px;height:4px}.bt-no-scroll::-webkit-scrollbar-track{background:transparent}.bt-no-scroll::-webkit-scrollbar-thumb{background:#2e2e38;border-radius:4px}
    .bt-hide-scroll{scrollbar-width:none;-ms-overflow-style:none}.bt-hide-scroll::-webkit-scrollbar{display:none;width:0;height:0}
    .bt-section-box{background:#121216;border:1px solid #22222a;border-radius:7px;padding:7px 8px;margin-bottom:7px}
    .bt-lbl{font-size:9.5px;color:#9ca3af;height:13px;line-height:13px;margin-bottom:4px;font-weight:700;text-transform:uppercase;letter-spacing:0.35px;display:flex;align-items:center;gap:3px;white-space:nowrap;overflow:hidden}
    .bt-bolt{color:#e5c158;font-size:10px;line-height:1;flex-shrink:0}
    .bt-seg-track{background:#09090c;border:1px solid #252530;border-radius:6px;padding:2px;margin-bottom:6px;box-sizing:border-box}
    .bt-grid-1{display:grid;grid-template-columns:1fr;gap:3px}.bt-grid-2{display:grid;grid-template-columns:1fr 1fr;gap:3px}.bt-grid-3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:3px}.bt-grid-4{display:grid;grid-template-columns:repeat(4,1fr);gap:3px}.bt-grid-5{display:grid;grid-template-columns:repeat(5,1fr);gap:2px}
    .bt-chip{position:relative;box-sizing:border-box;width:100%;height:24px;background:transparent;color:#9ca3af;border:1px solid transparent;padding:0 6px;font-size:11px;font-weight:500;border-radius:4px;cursor:pointer;display:flex;justify-content:space-between;align-items:center;transition:0.12s ease;white-space:nowrap}
    .bt-chip.bt-chip-centered{justify-content:center;text-align:center;padding:0 4px}.bt-chip:hover{color:#f4f4f5;background:rgba(255,255,255,0.04)}
    .bt-chip.active{background:var(--ac-bg,#242014);color:var(--ac-fg,#f3d26b);border-color:var(--ac-bd,rgba(229,193,88,0.58));box-shadow:0 2px 6px rgba(0,0,0,0.5);font-weight:600}
    .bt-chip kbd{font-size:9px;font-family:inherit;color:#8a8a93;background:rgba(255,255,255,0.05);border:1px solid rgba(255,255,255,0.1);padding:0 3.5px;border-radius:3px;margin-left:3px;font-weight:600}
    .bt-chip.active kbd{background:rgba(0,0,0,0.28);border-color:transparent;color:inherit}
    .bt-switch{width:100%;height:28px;box-sizing:border-box;border-radius:6px;padding:0 8px;margin-bottom:6px;background:#14141a;border:1px solid #282834;color:#9ca3af;display:flex;align-items:center;justify-content:space-between;gap:5px;cursor:pointer;transition:all 0.15s ease;outline:none}
    .bt-switch:hover{border-color:#454554;color:#e4e4e7;background:#181820}.bt-sw-txt{font-size:11px;font-weight:600;line-height:1;white-space:nowrap;transition:color 0.15s}
    .bt-sw-track{width:26px;height:14px;border-radius:99px;background:#252530;border:1px solid #3a3a48;position:relative;flex-shrink:0;transition:all 0.18s ease;box-sizing:border-box}
    .bt-sw-thumb{width:10px;height:10px;border-radius:50%;background:#9ca3af;position:absolute;top:1px;left:1px;transition:transform 0.18s cubic-bezier(0.4,0,0.2,1),background 0.18s ease;box-shadow:0 1px 2px rgba(0,0,0,0.6)}
    .bt-switch[data-val="да"] .bt-sw-thumb{transform:translateX(12px);background:#ffffff}
    .bt-switch[data-val="да"][data-mode="gold"]{background:rgba(229,193,88,0.12);border-color:rgba(229,193,88,0.5);color:#f3d26b}.bt-switch[data-val="да"][data-mode="gold"] .bt-sw-track{background:#c69a31;border-color:#e5c158}
    .bt-switch[data-val="да"][data-mode="check"]{background:rgba(16,185,129,0.13);border-color:rgba(16,185,129,0.5);color:#34d399}.bt-switch[data-val="да"][data-mode="check"] .bt-sw-track{background:#059669;border-color:#10b981}
    .bt-switch[data-val="нет"][data-mode="check"]{background:rgba(242,54,69,0.12);border-color:rgba(242,54,69,0.45);color:#f87171}.bt-switch[data-val="нет"][data-mode="check"] .bt-sw-track{background:#3a181e;border-color:#f23645}.bt-switch[data-val="нет"][data-mode="check"] .bt-sw-thumb{background:#f87171}
    .bt-switch[data-val="да"][data-mode="danger"]{background:rgba(242,54,69,0.14);border-color:rgba(242,54,69,0.55);color:#f87171}.bt-switch[data-val="да"][data-mode="danger"] .bt-sw-track{background:#dc2626;border-color:#f87171}
    .bt-flow-pill{width:100%;height:28px;border-radius:6px;font-size:11px;font-weight:600;cursor:pointer;display:flex;justify-content:space-between;align-items:center;padding:0 7px;box-sizing:border-box;transition:0.12s ease;border:1px solid #282830;background:#16161b;color:#d4d4d8}
    .bt-flow-pill:hover{filter:brightness(1.12)}.bt-flow-pill[data-dir="вверх"]{background:rgba(16,185,129,0.14);color:#34d399;border-color:rgba(16,185,129,0.45)}.bt-flow-pill[data-dir="вниз"]{background:rgba(242,54,69,0.14);color:#f87171;border-color:rgba(242,54,69,0.45)}.bt-flow-pill[data-dir="—"]{background:#14141a;color:#9ca3af;border-color:#282834}
    .bt-prof-btn{background:transparent;color:#9ca3af;border:1px solid transparent;border-radius:4px;padding:4px;font-size:10.5px;font-weight:600;cursor:pointer;transition:0.12s;white-space:nowrap}.bt-prof-btn:hover{color:#f4f4f5}.bt-prof-btn.active{background:#242014;color:#f3d26b;border-color:rgba(229,193,88,0.55);box-shadow:0 1px 4px rgba(0,0,0,0.5)}
    .bt-struct-preset{background:#15151b;color:#9ca3af;border:1px solid #282834;border-radius:4px;padding:1.5px 5px;font-size:9.5px;cursor:pointer;transition:0.12s;white-space:nowrap}.bt-struct-preset:hover{color:#f3d26b;border-color:#52525b}.bt-struct-preset.active{background:rgba(229,193,88,0.16);color:#f3d26b;border-color:rgba(229,193,88,0.55);font-weight:600}
    .bt-sel{width:100%;height:28px;box-sizing:border-box;background:#16161b;color:#f4f4f5;border:1px solid #282830;padding:0 6px;border-radius:6px;font-size:11px;font-weight:500;outline:none;margin-bottom:6px;cursor:pointer}.bt-sel:focus,#bt-inp-comment:focus,#f-swept-lvl:focus{border-color:#e5c158}
    .bt-mini-pill{background:#1c1c22;color:#d4d4d8;border:1px solid #2e2e38;padding:3px 8px;font-size:10.5px;border-radius:4px;cursor:pointer;transition:0.12s}.bt-mini-pill.active{background:rgba(229,193,88,0.2);color:#f3d26b;border-color:#e5c158;font-weight:700}
    .bt-top-btn{background:#16161a;color:#d4d4d8;border:1px solid #2a2a32;border-radius:5px;padding:2px 7px;font-size:11.5px;cursor:pointer;transition:0.12s}.bt-top-btn:hover,.bt-top-btn.active{color:#f3d26b;border-color:#e5c158;background:rgba(229,193,88,0.14)}
    .bt-icon-btn{background:transparent;color:#d4d4d8;border:none;cursor:pointer;font-size:10.5px;padding:2px 5px}.bt-icon-btn:hover{color:#f3d26b}
    .bt-drag-pill{background:#1a1a20;color:#f4f4f5;border:1px solid #383844;border-radius:5px;padding:3px 7px;font-size:10.5px;cursor:grab;display:inline-flex;align-items:center;gap:5px;transition:transform 0.1s ease,border-color 0.1s ease}
    .bt-drag-pill.dragging-active{background:#2a2515;border-color:#f3d26b;color:#f3d26b;transform:scale(1.05);z-index:10;cursor:grabbing;box-shadow:0 4px 12px rgba(0,0,0,0.6)}
    .bt-del-col{color:#f23645;font-weight:bold;cursor:pointer;padding:0 2px;font-size:11px}
    .bt-cm-item:hover,.bt-cc-item:hover,.bt-fix-action:hover{background:rgba(229,193,88,0.18)!important;border-color:#e5c158!important;color:#f3d26b!important}
    #bt-cc-btn:hover{border-color:#52525b}#bt-info-date:hover{border-color:#3a3a48;color:#f3d26b!important;background:#1c1c24!important}#bt-quick-copy-prices:hover{border-color:rgba(229,193,88,0.55)}
    .bt-prev-cell{background:#16161c;border:1px solid #262630;border-radius:4px;padding:2px 6px;flex-shrink:0;display:flex;flex-direction:column;min-width:48px;max-width:130px}.bt-prev-cell.is-first{border-color:rgba(229,193,88,0.55);background:rgba(229,193,88,0.08)}
  </style>`;

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
  const dateBtnEl = document.getElementById('bt-info-date');
  const commentInpEl = document.getElementById('bt-inp-comment');
  const commentMenuEl = document.getElementById('bt-comment-menu');
  const rowPreviewEl = document.getElementById('bt-row-preview');
  const fixHelpBoxEl = document.getElementById('bt-fix-help-box');
  let toastTimer = null;

  function tryCaptureTvScreenshotUrl(rawText) {
    if (isInternalCopy) return false;
    const s = String(rawText || '').trim();
    const m = s.match(/^https?:\/\/(?:[a-z0-9-]+\.)?tradingview\.com\/x\/[a-zA-Z0-9]+\/?$/i);
    if (m) {
      screenshotUrl = m[0];
      updateScreenshotPillUI();
      if (modalEl && modalEl.style.display === 'block') {
        updatePreviewBar();
        showToast('📸 Скриншот (Alt+S) прикреплён автоматически ✓');
      }
      return true;
    }
    return false;
  }

  (function installClipboardHooks() {
    try {
      const clip = pageWindow.navigator?.clipboard || navigator.clipboard;
      if (clip && typeof clip.writeText === 'function' && !clip.writeText._btHooked) {
        const origWrite = clip.writeText.bind(clip);
        const wrappedWrite = function(text) {
          try { tryCaptureTvScreenshotUrl(text); } catch (e) {}
          return origWrite(text);
        };
        wrappedWrite._btHooked = true;
        clip.writeText = wrappedWrite;
      }
    } catch (e) {}

    window.addEventListener('copy', e => {
      try {
        const copied = e.clipboardData?.getData('text/plain') || '';
        if (copied) tryCaptureTvScreenshotUrl(copied);
      } catch (err) {}
    }, true);
  })();

  rowPreviewEl.addEventListener('wheel', e => {
    if (e.deltaY !== 0) { e.preventDefault(); rowPreviewEl.scrollLeft += e.deltaY * 0.8; }
  }, { passive: false });

  function showToast(msg, type = 'success') {
    toastEl.innerText = msg;
    toastEl.style.borderLeftColor = type === 'error' ? '#f23645' : '#089981';
    toastEl.style.display = 'block';
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { toastEl.style.display = 'none'; }, 3200);
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
      if (o.tone === 'green') styleVars = `style="--ac-bg:#102620;--ac-fg:#34d399;--ac-bd:rgba(16,185,129,0.58);"`;
      else if (o.tone === 'red') styleVars = `style="--ac-bg:#2a1418;--ac-fg:#f87171;--ac-bd:rgba(242,54,69,0.58);"`;
      return `<button type="button" class="bt-chip ${centeredClass} ${val === activeVal ? 'active' : ''}" data-val="${val}" ${styleVars}><span>${label}</span>${kbd}</button>`;
    }).join('');
    return `<div class="bt-group bt-seg-track bt-grid-${cols}" id="${id}">${htmlOpts}</div>`;
  }

  function makeSwitch(id, activeVal = 'да', mode = 'gold', mb = '6px') {
    return `<button type="button" class="bt-switch" id="${id}" data-val="${activeVal}" data-mode="${mode}" style="margin-bottom:${mb};"><span class="bt-sw-txt">${activeVal === 'да' ? 'Да' : 'Нет'}</span><span class="bt-sw-track"><span class="bt-sw-thumb"></span></span></button>`;
  }

  function updateSwitchUI(el, nextVal) {
    if (!el) return;
    el.dataset.val = nextVal;
    const txt = el.querySelector('.bt-sw-txt');
    if (txt) txt.innerText = nextVal === 'да' ? 'Да' : 'Нет';
  }

  function makeCandleColorSelect(activeVal) {
    const info = CANDLE_COLORS[activeVal] || CANDLE_COLORS['Лонг 3'];
    const itemsHtml = Object.entries(CANDLE_COLORS).map(([k, v]) => `
      <div class="bt-cc-item" data-val="${k}" style="display:flex;align-items:center;justify-content:space-between;padding:5px 8px;border-radius:4px;cursor:pointer;font-size:11px;color:#f4f4f5;">
        <div style="display:flex;align-items:center;gap:7px;"><span style="width:13px;height:13px;border-radius:3px;background:${v.hex};display:inline-block;box-shadow:0 0 0 1px rgba(255,255,255,0.2);"></span><span style="font-weight:600;">${k}</span></div>
        <span style="color:#9ca3af;font-size:10px;">${v.hint}</span>
      </div>`).join('');
    return `
      <div id="bt-cc-wrap" style="position:relative;margin-bottom:6px;">
        <input type="hidden" id="f-candle-col" value="${activeVal}">
        <button type="button" id="bt-cc-btn" style="width:100%;height:28px;box-sizing:border-box;background:#16161b;color:#f4f4f5;border:1px solid #282830;padding:0 7px;border-radius:6px;font-size:11px;font-weight:500;cursor:pointer;display:flex;align-items:center;justify-content:space-between;outline:none;transition:0.12s;">
          <span style="display:flex;align-items:center;gap:6px;min-width:0;overflow:hidden;">
            <span id="bt-cc-dot" style="width:12px;height:12px;border-radius:3px;background:${info.hex};flex-shrink:0;box-shadow:0 0 0 1px rgba(255,255,255,0.22);"></span>
            <span id="bt-cc-label" style="font-weight:600;">${activeVal}</span>
            <span id="bt-cc-hint" style="color:#9ca3af;font-size:10px;">${info.hint}</span>
          </span><span style="color:#9ca3af;font-size:10px;">▾</span>
        </button>
        <div id="bt-cc-menu" style="display:none;position:absolute;top:31px;left:0;width:100%;background:#18181e;border:1px solid #3d3522;border-radius:6px;box-shadow:0 10px 25px rgba(0,0,0,0.92);z-index:120;padding:3px;box-sizing:border-box;">${itemsHtml}</div>
      </div>`;
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
    const el = document.getElementById(id);
    if (!el) return fallback;
    if (el.classList.contains('bt-switch')) return el.dataset.val || fallback;
    return el.querySelector('.active')?.dataset.val || fallback;
  }

  function setChipVal(id, val) {
    const el = document.getElementById(id);
    if (!el) return;
    if (el.classList.contains('bt-switch')) { updateSwitchUI(el, val); return; }
    el.querySelectorAll('.bt-chip').forEach(b => b.classList.toggle('active', b.dataset.val === val));
  }

  function makeFlowPill(id, tfLabel, dirVal) {
    const arrowText = dirVal === 'вверх' ? '↑ вверх' : (dirVal === 'вниз' ? '↓ вниз' : '—');
    return `<button type="button" class="bt-flow-pill" id="${id}" data-tf="${tfLabel}" data-dir="${dirVal}" title="Нажмите, чтобы изменить направление потока ${tfLabel}"><span style="color:#9ca3af;font-weight:500;">${tfLabel}</span><span class="bt-flow-val">${arrowText}</span></button>`;
  }

  function getFlowVal(id, fallback = 'вверх') { return document.getElementById(id)?.dataset.dir || fallback; }

  function syncPresetHighlights() {
    const stVal = document.getElementById('f-struct')?.value;
    dynFieldsEl.querySelectorAll('.bt-struct-preset').forEach(b => b.classList.toggle('active', b.dataset.st === stVal));
  }

  function updateScreenshotPillUI() {
    if (screenshotUrl) {
      screenPillBtn.innerText = '📸 Скрин ✓';
      screenPillBtn.style.background = 'rgba(16,185,129,0.18)';
      screenPillBtn.style.color = '#34d399';
      screenPillBtn.style.borderColor = '#10b981';
      screenPillBtn.title = `Прикреплён: ${screenshotUrl} (нажмите, чтобы удалить)`;
    } else {
      screenPillBtn.innerText = '📸 Скрин';
      screenPillBtn.style.background = '#16161a';
      screenPillBtn.style.color = '#d4d4d8';
      screenPillBtn.style.borderColor = '#2a2a32';
      screenPillBtn.title = 'Нажмите Alt+S на графике — ссылка прикрепится автоматически';
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

  function recalculateQueueTotals(queue) {
    let runningCum = 0, runningMax = 0;
    for (const snap of queue) {
      if (!snap || typeof snap !== 'object') continue;
      snap.cumRBefore = runningCum;
      if (typeof snap.rOut === 'number' && !isNaN(snap.rOut)) runningCum = Number((runningCum + snap.rOut).toFixed(2));
      if (runningCum > runningMax) runningMax = runningCum;
      snap.newCum = runningCum;
      snap.newMax = runningMax;
      snap.dd = Number(Math.max(0, runningMax - runningCum).toFixed(2));
    }
    return { cumR: runningCum, maxCumR: runningMax };
  }

  function collectCurrentSnapshot() {
    if (!currentTrade) return null;
    const cols = getActiveColumns();
    const t = currentTrade;
    const isLong = t.direction === 'LONG';
    const hasRes30 = cols.includes('res_30');
    const hasResOff = cols.includes('res_off');

    const defOffRes = t.autoResult === 'Безубыток' ? 'Сейф/БУ' : t.autoResult;
    const resStd = getChipVal('f-res', hasResOff ? defOffRes : t.autoResult);
    const res30Val = getChipVal('f-res30', t.autoResult === 'Тейк' ? 'цель' : (t.autoResult === 'Стоп' ? 'стоп' : 'безубыток'));
    const activeResult = hasRes30 ? res30Val : resStd;

    const bucket = getTplBucket();
    const existingSnap = bucket.batchQueue.find(item => item.tradeId === t.id || item.signature === t.signature);
    const baseCumR = existingSnap ? (existingSnap.cumRBefore ?? 0) : bucket.cumR;

    const manualRStr = cleanText(document.getElementById('f-manual-r30')?.value || '');
    const manualRNum = manualRStr ? parseFloat(manualRStr.replace(',', '.')) : NaN;

    let rOut = '', newCum = baseCumR, newMax = bucket.maxCumR;
    if (activeResult === 'Тейк' || activeResult === 'цель') {
      rOut = t.riskReward;
      newCum = Number((baseCumR + rOut).toFixed(2));
    } else if (activeResult === 'Стоп' || activeResult === 'стоп') {
      rOut = -1;
      newCum = Number((baseCumR - 1).toFixed(2));
    } else if (activeResult === 'Безубыток' || activeResult === 'безубыток' || activeResult === 'Бу' || activeResult === 'Сейф/БУ') {
      rOut = 0;
    } else if (!isNaN(manualRNum)) {
      rOut = manualRNum;
      newCum = Number((baseCumR + rOut).toFixed(2));
    }

    if (newCum > newMax) newMax = newCum;
    const dd = Number(Math.max(0, newMax - newCum).toFixed(2));

    const aoiInp = document.getElementById('f-aoi-num');
    const aoiNum = aoiInp ? (parseInt(aoiInp.value, 10) || 1) : state.aoiNum;

    const flD = getFlowVal('f-fl-d', state.flowD);
    const fl4H = getFlowVal('f-fl-4h', state.flow4H);
    const fl1H = getFlowVal('f-fl-1h', state.flow1H);
    const fl15M = getFlowVal('f-fl-15m', state.flow15M);

    const matchDir = (fl) => fl === 'вверх' ? (isLong ? 'да' : 'нет') : (fl === 'вниз' ? (!isLong ? 'да' : 'нет') : 'нет');
    const tfD1 = getChipVal('f-tf-d1', matchDir(flD));
    const tfH4 = getChipVal('f-tf-h4', matchDir(fl4H));
    const tfH1 = getChipVal('f-tf-h1', matchDir(fl1H));
    const tfM15 = getChipVal('f-tf-m15', matchDir(fl15M));
    const tfOk = getChipVal('f-tf-ok', 'да');
    const emotionVal = document.getElementById('f-emotion')?.value || state.emotion || 'Спокоен';

    const f1 = getChipVal('f-f1', 'да'), f2 = getChipVal('f-f2', 'да'), f3 = getChipVal('f-f3', 'нет'), f4 = getChipVal('f-f4', 'да');
    const factorsCnt = [f1, f2, f3, f4].filter(x => x === 'да').length;

    const stopUsd = Math.max(0.01, t.riskPts);
    const lots30 = Math.floor((30 / stopUsd / 100) * 100) / 100;
    const profitUsd = rOut === '' ? '' : Number((Number(rOut) * lots30 * 100 * stopUsd).toFixed(2));
    const lotsOff = Number(((10000 * 0.003) / (stopUsd * 100)).toFixed(2));
    const rOutOff = resStd === 'Стоп' ? -1 : (resStd === 'Тейк' ? Number((0.5 * 3 + 0.5 * t.riskReward).toFixed(2)) : (resStd === 'Сейф/БУ' ? 1.5 : 0));

    const comment = cleanText(commentInpEl?.value || '');
    const chk30 = getChipVal('f-check30', 'да');
    const img = cleanText(screenshotUrl);
    const sweptVal = document.getElementById('f-swept')?.value || 'ничего не снято';
    const sweptLvlRaw = cleanText(document.getElementById('f-swept-lvl')?.value || '').replace(',', '.');
    const useLocal = state.timeMode === 'local';

    return {
      tradeId: t.id, signature: t.signature, symbol: t.symbol,
      candleDate: useLocal ? t.candleDateLocal : t.candleDateMsk,
      mskTime: t.mskTime, officeTimeSlot: useLocal ? t.localTime : t.mskTime,
      session: getActiveSession(), isLong, timeframeCyr: t.timeframeCyr, timeframeLat: t.timeframeLat,
      trend: getChipVal('f-trend', isLong ? 'Тренд вверх' : 'Тренд вниз'),
      trend_h4: getChipVal('f-h4-trend', isLong ? 'Вверх' : 'Вниз'),
      zone: getChipVal('f-zone', state.zone),
      zone_30: document.getElementById('f-zone30')?.value || 'Разворотная зона',
      poi_type: getChipVal('f-poi-type', 'OB'),
      form_m1: getChipVal('f-form-m1', mapMvModelToFormM1(t.ind?.modelCode, t.ind?.isContinuation)),
      liq: getChipVal('f-liq', state.liq), react: getChipVal('f-react', 'да'),
      conf: document.getElementById('f-conf')?.value || state.conf,
      entryPrice: t.entryPrice, stopPrice: t.stopPrice, targetPrice: t.targetPrice,
      upperBound: t.upperBound, lowerBound: t.lowerBound, riskPts: t.riskPts, potPts: t.potPts, riskReward: t.riskReward,
      lots30, lotsOff, resStd, res30Val, activeResult, manualRNum, rOut, rOutOff, profitUsd,
      cumRBefore: baseCumR, newCum, newMax, dd,
      hit_1r: getChipVal('f-hit1r', t.reached1R ? 'да' : 'нет'),
      hit_3r: getChipVal('f-hit3r', t.reached3R ? 'да' : 'нет'),
      errVal: document.getElementById('f-err')?.value || '—',
      comment, img, aoiNum, flD, fl4H, fl1H, fl15M, tfD1, tfH4, tfH1, tfM15, tfOk, emotionVal,
      struct_30: document.getElementById('f-struct')?.value || (isLong ? 'слом ↑' : 'слом ↓'),
      sweptVal, sweptLvlRaw,
      candle_col: document.getElementById('f-candle-col')?.value || (isLong ? 'Лонг 3' : 'Шорт 3'),
      against_flow: getChipVal('f-against30', 'нет'),
      fuel_30: getChipVal('f-fuel30', 'нет'),
      touch_30: getChipVal('f-touch30', '1'),
      chk30, f1, f2, f3, f4, factorsCnt,
      reached: getChipVal('f-reached', 'да'),
      aoi_res: getChipVal('f-aoi-res', 'Отбой')
    };
  }

  function renderSnapshotToRow(snap) {
    const cols = getActiveColumns();
    const pasteCol = getActivePasteCol();
    const startRow = getActiveStartRow();
    const schemaStyle = PRESET_SCHEMAS[state.presetKey] || PRESET_SCHEMAS.bt_main;
    const f = (tok, fallback) => buildDynamicFormula(tok, cols, pasteCol, startRow, fallback);
    const sweptLvlFormatted = (snap.sweptVal === 'ничего не снято' || !snap.sweptLvlRaw) ? '' : numStr(snap.sweptLvlRaw);

    const tokenValues = {
      symbol: snap.symbol, date: snap.candleDate, time_msk: snap.mskTime, time_slot: snap.officeTimeSlot,
      session: snap.session, trend: snap.trend, trend_h4: snap.trend_h4,
      tf: schemaStyle.tfStyle === 'cyr' ? snap.timeframeCyr : snap.timeframeLat,
      dir_lower: snap.isLong ? 'лонг' : 'шорт', dir_title: snap.isLong ? 'Лонг' : 'Шорт',
      zone: snap.zone, zone_30: snap.zone_30, poi_type: snap.poi_type, form_m1: snap.form_m1,
      liq: snap.liq, react: snap.react, conf: snap.conf,
      entry: numStr(snap.entryPrice), stop: numStr(snap.stopPrice), target: numStr(snap.targetPrice),
      upper_bound: numStr(snap.upperBound), lower_bound: numStr(snap.lowerBound),
      risk_pts: f('risk_pts', numStr(snap.riskPts)), pot_pts: f('pot_pts', numStr(snap.potPts)),
      rr: f('rr', numStr(snap.riskReward)), lots_30: f('lots_30', numStr(snap.lots30)),
      lots_off: f('lots_off', numStr(snap.lotsOff)), result: snap.resStd, res_30: snap.res30Val,
      manual_r_30: (snap.res30Val === 'частичная фиксация' || snap.res30Val === 'закрыл руками') && !isNaN(snap.manualRNum) ? numStr(snap.manualRNum) : '',
      res_off: snap.resStd === 'Безубыток' ? 'Сейф/БУ' : snap.resStd,
      r_out: f('r_out', snap.rOut === '' ? '' : numStr(snap.rOut)),
      r_out_off: f('r_out_off', numStr(snap.rOutOff)),
      profit_usd: f('profit_usd', snap.profitUsd === '' ? '' : numStr(snap.profitUsd)),
      prof_pct_off: f('prof_pct_off', numStr(Number((snap.rOutOff * 0.003).toFixed(4)))),
      move_pts: f('move_pts', numStr(snap.resStd === 'Тейк' ? snap.potPts : (snap.resStd === 'Стоп' ? -snap.riskPts : 0))),
      cum_r: f('cum_r', snap.resStd === 'Пропуск' ? numStr(snap.cumRBefore) : numStr(snap.newCum)),
      dd_r: f('dd_r', snap.resStd === 'Пропуск' ? numStr(0) : numStr(snap.dd)),
      hit_1r: snap.hit_1r, hit_3r: snap.hit_3r, error: snap.errVal, error_poi: snap.errVal,
      comment: snap.comment, comment_30: snap.comment, screenshot: snap.img, aoi_num: snap.aoiNum,
      flow_d: snap.flD, flow_4h: snap.fl4H, flow_1h: snap.fl1H, flow_15m: snap.fl15M,
      tf_d1: snap.tfD1 ?? 'да', tf_h4: snap.tfH4 ?? 'да', tf_h1: snap.tfH1 ?? 'да', tf_m15: snap.tfM15 ?? 'да',
      struct_30: snap.struct_30, swept_30: snap.sweptVal, swept_level_30: sweptLvlFormatted,
      candle_col: snap.candle_col, against_flow: snap.against_flow, fuel_30: snap.fuel_30,
      touch_30: snap.touch_30, check_30: snap.chk30, tf_ok: snap.tfOk ?? 'да',
      emotion: snap.emotionVal || 'Спокоен', safe_pct: numStr(0.5), safe_rr: numStr(3), risk_pct_col: numStr(0.003),
      f_ret: snap.f1, f_ob: snap.f2, f_fvg: snap.f3, f_liq: snap.f4,
      factors_cnt: f('factors_cnt', snap.factorsCnt), valid_aoi: f('valid_aoi', snap.factorsCnt >= 2 ? 'ДА' : 'нет'),
      reached: snap.reached, aoi_res: snap.aoi_res, empty: ''
    };

    const baseIdx = colLetterToIndex(pasteCol);
    const cells = cols.map((tok, i) => ({
      colLet: colIndexToLetter(baseIdx + i), tok,
      shortLabel: SHORT_COL_NAMES[tok] || tok, val: String(tokenValues[tok] ?? '')
    }));
    return { pasteCol, cells, tsv: cells.map(c => c.val).join('\t') };
  }

  function updatePreviewBar() {
    if (!rowPreviewEl || !currentTrade) return;
    const snap = collectCurrentSnapshot();
    if (!snap) return;
    const rendered = renderSnapshotToRow(snap);
    const firstCell = rendered.cells[0];

    const hintEl = document.getElementById('bt-preview-hint');
    if (hintEl && firstCell) {
      hintEl.innerHTML = `СТРОКА ДЛЯ ВСТАВКИ (НАЧАТЬ СО СТОЛБЦА <span style="color:#f3d26b;">${firstCell.colLet} · ${firstCell.shortLabel}</span>):`;
    }
    const fixColLbl = document.getElementById('bt-fix-col-lbl');
    if (fixColLbl && firstCell) fixColLbl.innerText = `${firstCell.colLet} (${firstCell.shortLabel})`;

    rowPreviewEl.innerHTML = rendered.cells.map((c, idx) => {
      const isFormula = c.val.startsWith('=');
      const displayVal = isFormula ? 'fx Формула' : (c.val || '—');
      const valColor = isFormula ? '#60a5fa' : (c.val ? '#f4f4f5' : '#52525b');
      const safeTitle = c.val.replace(/"/g, '&quot;');
      return `<div class="bt-prev-cell ${idx === 0 ? 'is-first' : ''}" title="Столбец ${c.colLet} (${TOKEN_LABELS[c.tok] || c.tok}): ${safeTitle}"><span style="font-size:8.5px;color:#9ca3af;white-space:nowrap;"><b style="color:#e5c158;">${c.colLet}</b> · ${c.shortLabel}</span><span style="font-size:10px;font-weight:600;color:${valColor};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${displayVal}</span></div>`;
    }).join('');
  }

  function restoreSavedTradeState() {
    if (!currentTrade) return;
    const bucket = getTplBucket();
    const saved = bucket.batchQueue.find(item => item.signature === currentTrade.signature);
    if (!saved) return;

    screenshotUrl = saved.img || '';
    manualSessionOverride = saved.session || null;
    if (commentInpEl) commentInpEl.value = saved.comment || '';

    const setVal = (id, val) => { const el = document.getElementById(id); if (el && val != null) el.value = val; };
    setVal('f-aoi-num', saved.aoiNum);
    setVal('f-struct', saved.struct_30);
    setVal('f-zone30', saved.zone_30);
    setVal('f-swept', saved.sweptVal);
    setVal('f-conf', saved.conf);
    setVal('f-err', saved.errVal);
    setVal('f-emotion', saved.emotionVal);

    if (saved.candle_col) setCandleColVal(saved.candle_col);

    const setFlow = (id, dir) => {
      const p = document.getElementById(id);
      if (!p || !dir) return;
      p.dataset.dir = dir;
      const v = p.querySelector('.bt-flow-val');
      if (v) v.innerText = dir === 'вверх' ? '↑ вверх' : (dir === 'вниз' ? '↓ вниз' : '—');
    };
    setFlow('f-fl-d', saved.flD);
    setFlow('f-fl-4h', saved.fl4H);
    setFlow('f-fl-1h', saved.fl1H);
    setFlow('f-fl-15m', saved.fl15M);

    if (saved.trend) setChipVal('f-trend', saved.trend);
    if (saved.trend_h4) setChipVal('f-h4-trend', saved.trend_h4);
    if (saved.zone) setChipVal('f-zone', saved.zone);
    if (saved.poi_type) setChipVal('f-poi-type', saved.poi_type);
    if (saved.form_m1) setChipVal('f-form-m1', saved.form_m1);
    if (saved.fuel_30) setChipVal('f-fuel30', saved.fuel_30);
    if (saved.touch_30) setChipVal('f-touch30', saved.touch_30);
    if (saved.against_flow) setChipVal('f-against30', saved.against_flow);
    if (saved.chk30) setChipVal('f-check30', saved.chk30);
    if (saved.liq) setChipVal('f-liq', saved.liq);
    if (saved.react) setChipVal('f-react', saved.react);
    if (saved.f1) setChipVal('f-f1', saved.f1);
    if (saved.f2) setChipVal('f-f2', saved.f2);
    if (saved.f3) setChipVal('f-f3', saved.f3);
    if (saved.f4) setChipVal('f-f4', saved.f4);
    if (saved.tfD1) setChipVal('f-tf-d1', saved.tfD1);
    if (saved.tfH4) setChipVal('f-tf-h4', saved.tfH4);
    if (saved.tfH1) setChipVal('f-tf-h1', saved.tfH1);
    if (saved.tfM15) setChipVal('f-tf-m15', saved.tfM15);
    if (saved.tfOk) setChipVal('f-tf-ok', saved.tfOk);
    if (saved.reached) setChipVal('f-reached', saved.reached);
    if (saved.aoi_res) setChipVal('f-aoi-res', saved.aoi_res);

    setVal('f-swept-lvl', saved.sweptVal === 'ничего не снято' ? '' : numStr(saved.sweptLvlRaw));
    if (saved.hit_1r) setChipVal('f-hit1r', saved.hit_1r);
    if (saved.hit_3r) setChipVal('f-hit3r', saved.hit_3r);
    if (saved.resStd) setChipVal('f-res', saved.resStd);
    if (saved.res30Val) {
      setChipVal('f-res30', saved.res30Val);
      const manWrap = document.getElementById('bt-manual-r-wrap');
      if (manWrap) manWrap.style.display = (saved.res30Val === 'частичная фиксация' || saved.res30Val === 'закрыл руками') ? 'flex' : 'none';
      if (!isNaN(saved.manualRNum)) setVal('f-manual-r30', numStr(saved.manualRNum));
    }
  }

  function renderMainFields() {
    const cols = getActiveColumns();
    const has = tok => cols.includes(tok);
    const t = currentTrade || {
      direction: 'LONG', sessionStd: 'Лондон', sessionEuUs: 'Европа', autoResult: 'Тейк', reached1R: true, reached3R: true, entryExtreme: '', ind: {}
    };
    const isLong = t.direction === 'LONG';
    const ind = t.ind || {};
    const flowBolt = (ind.flows && !ind.flowsHistorical)
      ? BOLT
      : (ind.flowsHistorical ? `<span class="bt-bolt" style="color:#f59e0b;" title="Последние значения потока могли поменяться от таковых с точки входа">⚠️</span>` : '');

    let html = '';

    if (has('aoi_num')) {
      html += `<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:7px;"><span class="bt-lbl" style="margin:0;">№ ОБЛАСТИ / ЗОНЫ</span><input type="number" id="f-aoi-num" value="${state.aoiNum}" min="1" max="999" style="width:54px;background:#151519;color:#f3d26b;border:1px solid #282830;padding:3px;border-radius:4px;font-weight:700;text-align:center;"></div>`;
    }

    if (has('trend')) {
      const autoTrend = (state.trend === 'Тренд вверх' || state.trend === 'Тренд вниз') ? (isLong ? 'Тренд вверх' : 'Тренд вниз') : state.trend;
      html += `<div class="bt-lbl">ФАЗА РЫНКА ${BOLT}</div>` + makeChips('f-trend', [
        { val: 'Тренд вверх', label: 'Тренд вверх', kbd: 'A' },
        { val: 'Тренд вниз', label: 'Тренд вниз', kbd: 'S' },
        { val: 'Боковик', label: 'Боковик', kbd: 'D' },
        { val: 'Импульс/новости', label: 'Импульс / новости', kbd: 'F' }
      ], autoTrend, 2);
    }

    if (has('trend_h4') || has('f_ret') || has('f_ob') || has('f_fvg') || has('f_liq')) {
      const h4Def = ind.flows?.h4 ? (ind.flows.h4 === 'вверх' ? 'Вверх' : 'Вниз') : (isLong ? 'Вверх' : 'Вниз');
      html += `<div class="bt-lbl">ТРЕНД H4 ${flowBolt} И ФАКТОРЫ ОБЛАСТИ</div>` +
        makeChips('f-h4-trend', ['Вверх', 'Вниз'], h4Def, 2) +
        `<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:5px;margin-bottom:2px;">
          <div><div style="font-size:10px;color:#9ca3af;text-align:center;margin-bottom:2px;">Откат</div>${makeSwitch('f-f1', 'да', 'gold')}</div>
          <div><div style="font-size:10px;color:#9ca3af;text-align:center;margin-bottom:2px;">OB H4</div>${makeSwitch('f-f2', 'да', 'gold')}</div>
          <div><div style="font-size:10px;color:#9ca3af;text-align:center;margin-bottom:2px;">Имб.</div>${makeSwitch('f-f3', 'нет', 'gold')}</div>
          <div><div style="font-size:10px;color:#9ca3af;text-align:center;margin-bottom:2px;">Ликв.</div>${makeSwitch('f-f4', 'да', 'gold')}</div>
        </div>
        <div style="display:grid;grid-template-columns:1fr 2fr;gap:6px;">
          <div><div class="bt-lbl">ЦЕНА ДОШЛА</div>${makeSwitch('f-reached', 'да', 'gold')}</div>
          <div><div class="bt-lbl">РЕЗУЛЬТАТ ОБЛАСТИ</div>${makeChips('f-aoi-res', ['Отбой', 'Пробита', 'Ждёт'], 'Отбой', 3)}</div>
        </div>`;
    }

    if (has('zone')) {
      html += `<div class="bt-lbl">ЗОНА ВХОДА</div>` + makeChips('f-zone', [
        { val: 'Ордер-блок', label: 'Ордер-блок', kbd: '1' },
        { val: 'Имбаланс', label: 'Имбаланс', kbd: '2' },
        { val: 'Материнская свеча', label: 'Материнская свеча', kbd: '3' },
        { val: 'Зона ликвидности', label: 'Зона ликвидности', kbd: '4' },
        { val: 'Правильный откат', label: 'Правильный откат', kbd: '5' }
      ], state.zone, 2);
    }

    if (has('poi_type')) {
      html += `<div class="bt-lbl">ТИП ТОЧКИ ВХОДА (POI)</div>` + makeChips('f-poi-type', [
        { val: 'OB', label: 'OB', kbd: '1' },
        { val: 'Имбаланс', label: 'Имбаланс', kbd: '2' },
        { val: '50% МС', label: '50% МС', kbd: '3' }
      ], 'OB', 3);
    }

    if (has('tf_d1') || has('tf_h4') || has('tf_h1') || has('tf_m15')) {
      const matchDir = (fl) => fl === 'вверх' ? (isLong ? 'да' : 'нет') : (fl === 'вниз' ? (!isLong ? 'да' : 'нет') : 'нет');
      const arrow = (fl) => fl === 'вверх' ? '↑' : (fl === 'вниз' ? '↓' : '—');
      html += `<div class="bt-lbl">ПОДТВЕРЖДЕНИЕ ТФ ПО НАПРАВЛЕНИЮ СДЕЛКИ ${flowBolt}</div>
        <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:5px;margin-bottom:2px;">
          <div><div style="font-size:10px;color:#9ca3af;text-align:center;margin-bottom:2px;">D1 (${arrow(state.flowD)})</div>${makeSwitch('f-tf-d1', matchDir(state.flowD), 'gold')}</div>
          <div><div style="font-size:10px;color:#9ca3af;text-align:center;margin-bottom:2px;">H4 (${arrow(state.flow4H)})</div>${makeSwitch('f-tf-h4', matchDir(state.flow4H), 'gold')}</div>
          <div><div style="font-size:10px;color:#9ca3af;text-align:center;margin-bottom:2px;">H1 (${arrow(state.flow1H)})</div>${makeSwitch('f-tf-h1', matchDir(state.flow1H), 'gold')}</div>
          <div><div style="font-size:10px;color:#9ca3af;text-align:center;margin-bottom:2px;">M15 (${arrow(state.flow15M)})</div>${makeSwitch('f-tf-m15', matchDir(state.flow15M), 'gold')}</div>
        </div>`;
    }

    if (has('form_m1')) {
      const defForm = mapMvModelToFormM1(ind.modelCode, ind.isContinuation);
      html += `<div class="bt-lbl">ФОРМАЦИЯ ВХОДА НА M1 ${ind.modelCode ? BOLT : ''}</div>` + makeChips('f-form-m1', [
        { val: 'SCOB', label: 'SCOB', kbd: '1' },
        { val: 'Свип', label: 'Свип', kbd: '2' },
        { val: 'Пин-бар (IFS)', label: 'Пин-бар', kbd: '3' },
        { val: 'Поглощение', label: 'Поглощение', kbd: '4' },
        { val: 'Правильный откат', label: 'Правильный откат', kbd: '5' }
      ], defForm, 2);
    }

    if (has('tf_ok') || has('emotion')) {
      html += `<div style="display:grid;grid-template-columns:1fr 1.6fr;gap:7px;">`;
      if (has('tf_ok')) html += `<div><div class="bt-lbl">ПО СИСТЕМЕ</div>${makeSwitch('f-tf-ok', 'да', 'check')}</div>`;
      if (has('emotion')) html += `<div><div class="bt-lbl">ЭМОЦИЯ</div>${makeSelect('f-emotion', ['Спокоен', 'Уверен', 'Страх', 'Жадность', 'FOMO', 'Тильт'], state.emotion || 'Спокоен')}</div>`;
      html += `</div>`;
    }

    const hasAnyFlowOr30 = has('flow_d') || has('flow_4h') || has('flow_1h') || has('flow_15m') ||
      has('struct_30') || has('zone_30') || has('candle_col') || has('swept_30') ||
      has('swept_level_30') || has('fuel_30') || has('touch_30') || has('against_flow') ||
      has('check_30') || has('hit_1r') || has('hit_3r');

    if (hasAnyFlowOr30) {
      const defCandleCol = ind.candleColor || (isLong ? 'Лонг 3' : 'Шорт 3');
      const stSlom = isLong ? 'слом ↑' : 'слом ↓';
      const st2Prob = isLong ? 'пробой ↓ · пробой ↓ · слом ↑' : 'пробой ↑ · пробой ↑ · слом ↓';
      const stFake = isLong ? 'ложный пробой ↓ · слом ↑' : 'ложный пробой ↑ · слом ↓';
      const stCont = isLong ? 'пробой ↑ · пробой ↑' : 'пробой ↓ · пробой ↓';
      const defStruct = ind.detectedStructure || state.struct30 || st2Prob;

      const mz = ind.matchedZone;
      const msweep = ind.matchedSweep;
      const defZone30 = mz?.zoneType || 'Разворотная зона';
      const defFuel30 = mz ? mz.fuel : 'нет';
      const defTouch30 = mz ? mz.touch : '1';
      const defAgainst30 = mz ? mz.against : 'нет';
      const defSwept30 = msweep ? msweep.sweptType : 'ничего не снято';
      const defSweptPrice = msweep?.price != null ? msweep.price : '';
      const zoneBolt = mz ? `<span class="bt-bolt" title="Распознано по зоне на графике: «${mz.rawText}»">⚡</span>` : '';
      const sweepBolt = msweep ? `<span class="bt-bolt" title="Распознан снятый уровень: «${msweep.rawText}»">⚡</span>` : '';

      html += `
        <div class="bt-section-box">
          <div class="bt-lbl">ПОТОКИ В МОМЕНТ ВХОДА ${flowBolt}</div>
          <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:5px;margin-bottom:7px;">
            ${makeFlowPill('f-fl-d', 'Д', state.flowD)}
            ${makeFlowPill('f-fl-4h', '4Ч', state.flow4H)}
            ${makeFlowPill('f-fl-1h', '1Ч', state.flow1H)}
            ${makeFlowPill('f-fl-15m', '15м', state.flow15M)}
          </div>
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:3px;gap:4px;">
            <div class="bt-lbl" style="margin:0;flex-shrink:0;">СТРУКТУРА ${ind.detectedStructure ? BOLT : ''}</div>
            <div style="display:flex;gap:3px;flex-wrap:nowrap;">
              <button type="button" class="bt-struct-preset" data-st="${stSlom}">Слом</button>
              <button type="button" class="bt-struct-preset" data-st="${st2Prob}">2 проб.+слом</button>
              <button type="button" class="bt-struct-preset" data-st="${stFake}">Ложн. пробой</button>
              <button type="button" class="bt-struct-preset" data-st="${stCont}">2 пробоя</button>
            </div>
          </div>
          ${makeSelect('f-struct', ['пробой ↑', 'пробой ↑ · пробой ↑', 'пробой ↑ · пробой ↑ · слом ↓', 'пробой ↓', 'пробой ↓ · пробой ↓', 'пробой ↓ · пробой ↓ · слом ↑', 'слом ↑', 'слом ↓', 'ложный пробой ↑ · слом ↓', 'ложный пробой ↓ · слом ↑', 'внутр. слом ↑', 'внутр. слом ↓', 'смена ↑ на 4Ч', 'смена ↓ на 4Ч'], defStruct, '0px')}
        </div>
        <div class="bt-section-box">
          <div style="display:grid;grid-template-columns:1.05fr 0.95fr;gap:6px;">
            <div><div class="bt-lbl">ВИД ЗОНЫ ${zoneBolt}</div>${makeSelect('f-zone30', ['Разворотная зона', 'Сильный покупатель', 'Средний покупатель', 'Красивый покупатель', 'Сильный продавец', 'Средний продавец', 'Красивый продавец', 'Зона входа', 'Проторговка', 'Зона 15м 0.5', 'Зона 1Ч 0.5', 'Зона 4Ч 0.5', 'Зона Д 0.5'], defZone30)}</div>
            <div><div class="bt-lbl">СВЕЧА ВХОДА ${ind.candleColor ? BOLT : ''}</div>${makeCandleColorSelect(defCandleCol)}</div>
          </div>
          <div style="display:grid;grid-template-columns:0.95fr 1.25fr 0.95fr;gap:6px;align-items:start;">
            <div><div class="bt-lbl">ТОПЛИВО ${zoneBolt}</div>${makeSwitch('f-fuel30', defFuel30, 'gold')}</div>
            <div><div class="bt-lbl">КАСАНИЕ ЗОНЫ ${zoneBolt}</div>${makeChips('f-touch30', [{val:'1', label:'1'}, {val:'2', label:'2'}, {val:'3 и больше', label:'3+'}], defTouch30, 3)}</div>
            <div><div class="bt-lbl">ПРОТИВ ПОТОКА ${zoneBolt}</div>${makeSwitch('f-against30', defAgainst30, 'danger')}</div>
          </div>
          <div style="display:grid;grid-template-columns:1.15fr 1fr;gap:6px;align-items:start;">
            <div><div class="bt-lbl">ЧТО СНЯТО ПЕРЕД ВХОДОМ ${sweepBolt}</div>${makeSelect('f-swept', ['приманка', 'полка', 'уровень 15м', 'уровень 1Ч', 'уровень 4Ч', 'уровень Д', 'склейка 15м/1Ч', 'склейка Д · 4Ч', 'ничего не снято'], defSwept30, '0px')}</div>
            <div>
              <div class="bt-lbl">ЦЕНА УРОВНЯ СНЯТИЯ ${sweepBolt}</div>
              <div style="display:flex;gap:3px;height:28px;">
                <input type="text" id="f-swept-lvl" value="${defSweptPrice !== '' ? numStr(defSweptPrice) : ''}" placeholder="Цена уровня..." style="width:100%;height:28px;box-sizing:border-box;background:#16161b;color:#f4f4f5;border:1px solid #282830;padding:0 6px;border-radius:6px;font-size:11px;outline:none;">
                <button type="button" id="bt-fill-swept-lvl" title="Подставить расчётную цену уровня / экстремум свечи (${msweep?.price ?? t.entryExtreme})" style="height:28px;background:#18181d;color:#f3d26b;border:1px solid #282830;border-radius:6px;padding:0 6px;cursor:pointer;font-size:11px;">↺</button>
              </div>
            </div>
          </div>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px;">
          <div><div class="bt-lbl">ЧЕК-ЛИСТ 7/7</div>${makeSwitch('f-check30', 'да', 'check')}</div>
          <div><div class="bt-lbl">ДОШЛА ДО 1:1 ${BOLT}</div>${makeSwitch('f-hit1r', t.reached1R ? 'да' : 'нет', 'gold')}</div>
          <div><div class="bt-lbl">ДОШЛА ДО 1:3 ${BOLT}</div>${makeSwitch('f-hit3r', t.reached3R ? 'да' : 'нет', 'gold')}</div>
        </div>`;
    }

    if (has('liq') || has('conf') || has('react')) {
      const autoLiq = ind.matchedSweep ? 'да' : state.liq;
      html += `<div style="display:grid;grid-template-columns:1.1fr 1.9fr;gap:7px;">`;
      if (has('liq')) html += `<div><div class="bt-lbl">СНЯТИЕ ЛИКВ. ${ind.matchedSweep ? BOLT : ''}</div>${makeSwitch('f-liq', autoLiq, 'gold')}</div>`;
      if (has('conf')) {
        html += `<div><div class="bt-lbl">ПОДТВЕРЖДЕНИЕ <span style="color:#9ca3af;">(Z–B)</span></div>${makeSelect('f-conf', [
          { val: 'Слом структуры', label: 'Слом структуры [Z]' },
          { val: 'Реакция от зоны', label: 'Реакция от зоны [X]' },
          { val: 'Поглощение', label: 'Поглощение [C]' },
          { val: 'Снятие ликвидности', label: 'Снятие ликвидности [V]' },
          { val: 'Без подтверждения', label: 'Без подтверждения [B]' }
        ], state.conf)}</div>`;
      } else if (has('react')) {
        html += `<div><div class="bt-lbl">РЕАКЦИЯ ОТ ЗОНЫ</div>${makeSwitch('f-react', 'да', 'gold')}</div>`;
      }
      html += `</div>`;
    }

    if (has('res_30')) {
      const defRes30 = t.autoResult === 'Тейк' ? 'цель' : (t.autoResult === 'Стоп' ? 'стоп' : 'безубыток');
      html += `<div class="bt-lbl">ИСХОД СДЕЛКИ ${BOLT} <span style="color:#9ca3af;font-weight:500;">(Q / W / E / R)</span></div>` +
        makeChips('f-res30', [
          { val: 'цель', label: 'Цель', kbd: 'Q', tone: 'green' },
          { val: 'стоп', label: 'Стоп', kbd: 'W', tone: 'red' },
          { val: 'безубыток', label: 'БУ', kbd: 'E' },
          { val: 'частичная фиксация', label: 'Частич.' },
          { val: 'закрыл руками', label: 'Руками', kbd: 'R' }
        ], defRes30, 5) +
        `<div id="bt-manual-r-wrap" style="display:none;justify-content:space-between;align-items:center;background:#141418;padding:5px 8px;border-radius:5px;border:1px solid #3d3522;margin-bottom:6px;">
          <span style="font-size:10.5px;color:#f3d26b;">Фактический итог в R (например 1,4):</span>
          <input type="text" id="f-manual-r30" placeholder="0,0" style="width:60px;background:#0d0d10;color:#ffffff;border:1px solid #4a3f23;border-radius:4px;padding:3px 6px;font-size:11px;text-align:center;outline:none;">
        </div>`;
    } else if (has('res_off')) {
      const defOff = t.autoResult === 'Безубыток' ? 'Сейф/БУ' : t.autoResult;
      html += `<div class="bt-lbl">РЕЗУЛЬТАТ СДЕЛКИ ${BOLT}</div>` + makeChips('f-res', [
        { val: 'Тейк', label: 'Тейк', kbd: 'Q', tone: 'green' },
        { val: 'Стоп', label: 'Стоп', kbd: 'W', tone: 'red' },
        { val: 'Сейф/БУ', label: 'Сейф/БУ', kbd: 'E' },
        { val: 'Бу', label: 'Бу', kbd: 'R' }
      ], defOff, 4);
    } else if (has('result') || has('r_out')) {
      html += `<div class="bt-lbl">РЕЗУЛЬТАТ СДЕЛКИ ${BOLT}</div>` + makeChips('f-res', [
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
      html += `<div style="display:flex;justify-content:space-between;align-items:center;"><div class="bt-lbl">ПРИЧИНА ОШИБКИ ${BOLT}</div><span id="bt-err-reason" style="font-size:10px;color:#f3d26b;font-weight:600;margin-bottom:4px;"></span></div>${makeSelect('f-err', errList, '—')}`;
    }

    dynFieldsEl.innerHTML = html;

    const zone5Btn = dynFieldsEl.querySelector('#f-zone .bt-chip:nth-child(5)');
    if (zone5Btn) zone5Btn.style.gridColumn = 'span 2';
    const form5Btn = dynFieldsEl.querySelector('#f-form-m1 .bt-chip:nth-child(5)');
    if (form5Btn) form5Btn.style.gridColumn = 'span 2';

    document.getElementById('bt-extra-inputs').style.display = (has('comment') || has('comment_30') || has('screenshot')) ? 'flex' : 'none';

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
        updatePreviewBar();
      };
    });

    dynFieldsEl.querySelectorAll('.bt-switch').forEach(sw => {
      sw.onclick = () => {
        const next = sw.dataset.val === 'да' ? 'нет' : 'да';
        updateSwitchUI(sw, next);
        if (sw.id === 'f-liq') {
          saveState({ liq: next });
          if (next === 'нет' && getChipVal('f-res') === 'Пропуск') {
            const confEl = document.getElementById('f-conf');
            if (confEl) confEl.value = 'Без подтверждения';
          }
          autoSuggestError();
        }
        updatePreviewBar();
      };
    });

    dynFieldsEl.querySelectorAll('.bt-struct-preset').forEach(btn => {
      btn.onclick = () => {
        const sel = document.getElementById('f-struct');
        if (sel) { sel.value = btn.dataset.st; saveState({ struct30: sel.value }); }
        syncPresetHighlights();
        updatePreviewBar();
      };
    });

    const ccBtn = document.getElementById('bt-cc-btn');
    const ccMenu = document.getElementById('bt-cc-menu');
    if (ccBtn && ccMenu) {
      ccBtn.onclick = e => { e.stopPropagation(); ccMenu.style.display = ccMenu.style.display === 'block' ? 'none' : 'block'; };
      ccMenu.querySelectorAll('.bt-cc-item').forEach(item => {
        item.onclick = () => { setCandleColVal(item.dataset.val); ccMenu.style.display = 'none'; updatePreviewBar(); };
      });
    }

    const fillSweptBtn = document.getElementById('bt-fill-swept-lvl');
    if (fillSweptBtn) {
      fillSweptBtn.onclick = () => {
        const inp = document.getElementById('f-swept-lvl');
        const fallbackPrice = currentTrade?.ind?.matchedSweep?.price ?? currentTrade?.entryExtreme ?? '';
        if (inp && fallbackPrice !== '') { inp.value = numStr(fallbackPrice); updatePreviewBar(); }
      };
    }

    dynFieldsEl.querySelectorAll('select, input').forEach(el => {
      el.addEventListener('change', () => {
        if (el.id === 'f-swept') {
          const lvlInp = document.getElementById('f-swept-lvl');
          if (lvlInp) {
            const fallbackPrice = currentTrade?.ind?.matchedSweep?.price ?? currentTrade?.entryExtreme ?? '';
            lvlInp.value = el.value === 'ничего не снято' ? '' : (fallbackPrice !== '' ? numStr(fallbackPrice) : '');
          }
        }
        if (el.id === 'f-struct') { saveState({ struct30: el.value }); syncPresetHighlights(); }
        if (el.id === 'f-emotion') saveState({ emotion: el.value });
        if (el.id === 'f-conf') { saveState({ conf: el.value }); autoSuggestError(); }
        updatePreviewBar();
      });
      el.addEventListener('input', () => updatePreviewBar());
    });

    dynFieldsEl.querySelectorAll('.bt-group').forEach(grp => {
      grp.addEventListener('click', e => {
        const chip = e.target.closest('.bt-chip');
        if (!chip) return;
        grp.querySelectorAll('.bt-chip').forEach(b => b.classList.remove('active'));
        chip.classList.add('active');

        const val = chip.dataset.val;
        if (grp.id === 'f-trend') saveState({ trend: val });
        if (grp.id === 'f-zone') saveState({ zone: val });
        if (grp.id === 'f-res30') {
          const manWrap = document.getElementById('bt-manual-r-wrap');
          if (manWrap) manWrap.style.display = (val === 'частичная фиксация' || val === 'закрыл руками') ? 'flex' : 'none';
        }
        if (grp.id === 'f-res') {
          if (val === 'Пропуск' && getChipVal('f-liq') === 'нет') {
            const confEl = document.getElementById('f-conf');
            if (confEl) confEl.value = 'Без подтверждения';
          }
          autoSuggestError();
        }
        updatePreviewBar();
      });
    });

    autoSuggestError();
    restoreSavedTradeState();
    syncPresetHighlights();
    updatePreviewBar();
  }

  function syncLiveChartChanges() {
    if (modalEl.style.display !== 'block') return;
    const prevId = currentTrade?.id;
    const prevSig = currentTrade?.signature;

    if (!scanChartTrades(true, false) || !currentTrade) return;

    if (currentTrade.id !== prevId) {
      manualSessionOverride = null;
      if (commentInpEl) commentInpEl.value = '';
      renderMainFields();
      updateHeaderUI();
      pulseTradeCard('#089981');
      return;
    }

    if (currentTrade.signature !== prevSig) {
      const savedComment = commentInpEl ? commentInpEl.value : '';
      renderMainFields();
      if (commentInpEl && savedComment && !commentInpEl.value) commentInpEl.value = savedComment;
      updateHeaderUI();
      pulseTradeCard('#e5c158');
    }
  }

  let lastObservedBarKey = '';

  function syncRadarOnBarClose(model) {
    if (!model) return;
    try {
      attachModelAndStudySubscriptions(model);
      const ms = typeof model.mainSeries === 'function' ? model.mainSeries() : model.mainSeries;
      const bars = typeof ms?.bars === 'function' ? ms.bars() : ms?.bars;
      if (!bars || typeof bars.lastIndex !== 'function') return;

      const lastIdx = bars.lastIndex();
      if (lastIdx == null) return;
      const lastBar = getPlotRow(bars, lastIdx);
      const lastTime = lastBar?.[0] || 0;
      const barKey = `${lastIdx}_${lastTime}`;

      if (barKey === lastObservedBarKey) return;
      lastObservedBarKey = barKey;

      const symInfo = (typeof ms.symbolInfo === 'function' ? ms.symbolInfo() : ms._symbolInfo) || {};
      const cleanSymbol = String(symInfo.name || 'XAUUSD').split(':').pop();
      const rawInterval = typeof ms.interval === 'function' ? ms.interval() : (ms._interval || '1');
      syncGraphicsMemoryO1(model, `${cleanSymbol}_${rawInterval}`);
    } catch (e) {}
  }

  setInterval(() => {
    const model = getTvModel();
    if (!model) return;
    syncRadarOnBarClose(model);

    const selId = detectCurrentlySelectedToolId(model);
    if (selId && selId !== lastChartSelectedId) {
      lastChartSelectedId = selId;
      lastSelectedToolId = selId;
      if (modalEl.style.display === 'block') syncLiveChartChanges();
    }

    try {
      const hov = typeof model.hoveredSource === 'function' ? model.hoveredSource() : model._hoveredSource;
      if (hov && isSupportedTool(hov)) lastHoveredToolId = getToolId(hov);
    } catch (e) {}
  }, 500);

  window.addEventListener('pointerup', e => {
    if (modalEl.style.display !== 'block' || modalEl.contains(e.target)) return;
    scheduleLiveSync(30);
  }, true);

  function renderSettingsDrawer() {
    document.getElementById('bt-preset-select').value = state.presetKey;
    document.getElementById('bt-custom-paste-col').value = getActivePasteCol();
    document.getElementById('bt-custom-start-row').value = getActiveStartRow();

    const cols = getActiveColumns();
    const needsStartRow = cols.includes('cum_r') || cols.includes('dd_r');
    const startRowBox = document.getElementById('bt-start-row-box');
    const anchorsGrid = document.getElementById('bt-custom-anchors-grid');
    if (startRowBox && anchorsGrid) {
      startRowBox.style.display = needsStartRow ? 'block' : 'none';
      anchorsGrid.style.gridTemplateColumns = needsStartRow ? '1fr 1fr' : '1fr';
    }

    const colsWrap = document.getElementById('bt-cols-editor-wrap');
    const toggleColsBtn = document.getElementById('bt-toggle-cols-btn');
    const shouldShowCols = showColumnEditor || state.presetKey === 'custom';
    colsWrap.style.display = shouldShowCols ? 'block' : 'none';
    toggleColsBtn.innerText = shouldShowCols ? '▲ Скрыть список колонок' : `🔧 Настроить колонки таблицы (${cols.length} шт.)`;

    const baseColIdx = colLetterToIndex(getActivePasteCol());
    const pillsContainer = document.getElementById('bt-schema-pills');

    pillsContainer.innerHTML = cols.map((tok, idx) => {
      const colLet = colIndexToLetter(baseColIdx + idx);
      return `<div class="bt-drag-pill" data-idx="${idx}" data-tok="${tok}" title="Столбец ${colLet}. Перетащите мышкой для смены порядка"><span class="bt-pill-num" style="color:#e5c158;font-weight:700;">${colLet}</span><span>${TOKEN_LABELS[tok] || tok}</span><span class="bt-del-col" data-del="${idx}" title="Удалить колонку">✕</span></div>`;
    }).join('');

    pillsContainer.querySelectorAll('.bt-del-col').forEach(delBtn => {
      delBtn.addEventListener('mousedown', e => e.stopPropagation());
      delBtn.addEventListener('click', e => {
        e.stopPropagation();
        const i = parseInt(delBtn.dataset.del, 10);
        const updated = getActiveColumns();
        updated.splice(i, 1);
        saveState({ presetKey: 'custom', customColumns: updated, customPasteCol: getActivePasteCol(), customStartRow: getActiveStartRow() });
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
            if (activeIdx < overIdx) pillsContainer.insertBefore(activePill, overPill.nextSibling);
            else pillsContainer.insertBefore(activePill, overPill);
            updateColLetters();
          }
        };

        const onMouseUp = () => {
          window.removeEventListener('mousemove', onMouseMove, true);
          window.removeEventListener('mouseup', onMouseUp, true);
          activePill.classList.remove('dragging-active');
          const newOrder = Array.from(pillsContainer.querySelectorAll('.bt-drag-pill')).map(el => el.dataset.tok);
          saveState({ presetKey: 'custom', customColumns: newOrder, customPasteCol: getActivePasteCol(), customStartRow: getActiveStartRow() });
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
      const res = item.activeResult || item.result;
      const rrVal = Number(item.riskReward ?? item.rr ?? 0);
      if (res === 'Тейк' || res === 'цель') { wins++; totalR += rrVal; }
      else if (res === 'Стоп' || res === 'стоп') { losses++; totalR -= 1; }
      else if (typeof item.rOut === 'number' && !isNaN(item.rOut)) totalR += item.rOut;
    }
    const closed = wins + losses;
    const wr = closed > 0 ? Math.round((wins / closed) * 100) : 0;
    const sign = totalR > 0 ? '+' : '';
    return {
      count: bucket.batchQueue.length, totalR,
      shortText: bucket.batchQueue.length > 0 ? `${bucket.batchQueue.length} · ${sign}${totalR.toFixed(1)}R` : '0',
      tooltip: `Сохранено сделок: ${bucket.batchQueue.length} | Винрейт: ${wr}% (${wins} плюс / ${losses} стоп) | Итог: ${sign}${totalR.toFixed(2)} R`
    };
  }

  function updateHeaderUI() {
    const schemaObj = PRESET_SCHEMAS[state.presetKey];
    const bucket = getTplBucket();
    const cols = getActiveColumns();

    document.getElementById('bt-active-tpl-title').innerText = schemaObj ? schemaObj.name : `Свой журнал (${cols.length} кол.)`;
    collapsibleBodyEl.style.display = state.minimized ? 'none' : 'block';

    document.querySelectorAll('.bt-prof-btn').forEach(btn => btn.classList.toggle('active', btn.dataset.prof === state.outputProfile));

    const isReEdit = !!(currentTrade && bucket.batchQueue.some(item => item.tradeId === currentTrade.id || item.signature === currentTrade.signature));
    const pasteCol = getActivePasteCol();
    const firstColShort = SHORT_COL_NAMES[cols[0]] || cols[0] || '';

    mainBtnText.innerText = isReEdit
      ? `Обновить строку (в стлб. ${pasteCol} · ${firstColShort})`
      : `Скопировать строку (в стлб. ${pasteCol} · ${firstColShort})`;

    const bStats = computeBatchStats();
    const batchCntEl = document.getElementById('bt-batch-count');
    batchCntEl.innerText = bStats.shortText;
    batchCntEl.style.color = bStats.totalR > 0 ? '#34d399' : (bStats.totalR < 0 ? '#f87171' : '#f3d26b');
    document.getElementById('bt-copy-batch-btn').title = `${bStats.tooltip}\nНажмите, чтобы скопировать все сохранённые сделки сразу (для вставки в столбец ${pasteCol})`;

    updateScreenshotPillUI();
    if (!currentTrade) return;

    const total = allChartTrades.length || 1;
    const counterEl = document.getElementById('bt-trade-counter');
    counterEl.innerText = `${isReEdit ? '✓ ' : ''}${activeTradeIdx + 1}/${total}`;
    counterEl.style.color = isReEdit ? '#34d399' : '#e5c158';

    document.getElementById('bt-info-sym').innerText = `${currentTrade.symbol} · ${currentTrade.timeframeLat}`;
    sessionPillBtn.innerText = (manualSessionOverride ? '' : '⚡ ') + getActiveSession();

    const useLocal = state.timeMode === 'local';
    dateBtnEl.innerText = `${useLocal ? currentTrade.shortDateLocal : currentTrade.shortDateMsk} · ${useLocal ? currentTrade.localTime : currentTrade.mskTime} ${useLocal ? 'Лок.' : 'МСК'}`;

    const badge = document.getElementById('bt-info-badge');
    badge.innerText = `${currentTrade.direction} · ${numStr(currentTrade.riskReward)}R`;
    badge.style.background = currentTrade.direction === 'LONG' ? '#089981' : '#f23645';

    document.getElementById('bt-info-en').innerText = numStr(currentTrade.entryPrice);
    document.getElementById('bt-info-st').innerText = numStr(currentTrade.stopPrice);
    document.getElementById('bt-info-tg').innerText = numStr(currentTrade.targetPrice);

    updatePreviewBar();
  }

  function buildAndCopyRow(stayOpen = false) {
    if (!currentTrade) return;
    const snap = collectCurrentSnapshot();
    if (!snap) return;

    const bucket = getTplBucket();
    const existingIdx = bucket.batchQueue.findIndex(item => item.tradeId === snap.tradeId || item.signature === snap.signature);
    const newBatch = [...bucket.batchQueue];
    if (existingIdx >= 0) newBatch[existingIdx] = snap;
    else newBatch.push(snap);

    const totals = recalculateQueueTotals(newBatch);
    const rendered = renderSnapshotToRow(snap);
    const cols = getActiveColumns();
    const firstColShort = SHORT_COL_NAMES[cols[0]] || cols[0] || '';

    isInternalCopy = true;
    navigator.clipboard.writeText(rendered.tsv).then(() => {
      isInternalCopy = false;
      saveTplBucket({ batchQueue: newBatch, lastSavedSig: snap.signature, cumR: totals.cumR, maxCumR: totals.maxCumR });
      saveState({
        aoiNum: (cols.includes('valid_aoi') && existingIdx === -1) ? snap.aoiNum + 1 : snap.aoiNum,
        flowD: snap.flD, flow4H: snap.fl4H, flow1H: snap.fl1H, flow15M: snap.fl15M,
        struct30: snap.struct_30, emotion: snap.emotionVal
      });
      if (!stayOpen) modalEl.style.display = 'none';
      updateHeaderUI();
      showToast(`✓ Скопировано! Вставьте (Ctrl+V) в столбец ${rendered.pasteCol} (${firstColShort})`);
    }).catch(() => {
      isInternalCopy = false;
      showToast('Не удалось скопировать в буфер обмена', 'error');
    });
  }

  function applyProfile(prof, copyImmediately = false) {
    saveState({ outputProfile: prof, decimalSep: prof === 'en_form' ? '.' : ',' });
    const sweptLvlEl = document.getElementById('f-swept-lvl');
    if (sweptLvlEl && sweptLvlEl.value) sweptLvlEl.value = numStr(sweptLvlEl.value);
    updateHeaderUI();
    const labels = { gs_ru: '📊 Google Docs RU', ru_form: '📗 Excel RU', en_form: '🌐 EN' };
    if (copyImmediately && currentTrade) buildAndCopyRow(true);
    else showToast(`Выбран формат: ${labels[prof]}`);
  }

  document.querySelectorAll('.bt-prof-btn').forEach(btn => { btn.onclick = () => applyProfile(btn.dataset.prof, false); });
  document.getElementById('bt-toggle-fix-help').onclick = () => {
    isFixHelpOpen = !isFixHelpOpen;
    fixHelpBoxEl.style.display = isFixHelpOpen ? 'block' : 'none';
  };
  document.querySelectorAll('.bt-fix-action').forEach(btn => {
    btn.onclick = () => { isFixHelpOpen = false; fixHelpBoxEl.style.display = 'none'; applyProfile(btn.dataset.fix, true); };
  });

  document.getElementById('bt-custom-paste-col').onchange = e => {
    saveState({ presetKey: 'custom', customColumns: getActiveColumns(), customPasteCol: e.target.value });
    renderSettingsDrawer();
    updateHeaderUI();
  };
  document.getElementById('bt-custom-start-row').onchange = e => {
    saveState({ presetKey: 'custom', customColumns: getActiveColumns(), customStartRow: Math.max(1, parseInt(e.target.value, 10) || 2) });
    renderSettingsDrawer();
    updatePreviewBar();
  };

  sessionPillBtn.onclick = () => {
    const isEuUs = (PRESET_SCHEMAS[state.presetKey]?.sessionStyle === 'eu_us');
    const list = isEuUs ? ['Азия', 'Европа', 'Америка'] : ['Азия', 'Лондон', 'Нью-Йорк'];
    const curr = getActiveSession();
    manualSessionOverride = list[(list.indexOf(curr) + 1) % list.length];
    updateHeaderUI();
  };

  dateBtnEl.onclick = () => {
    const nextMode = state.timeMode === 'local' ? 'msk' : 'local';
    saveState({ timeMode: nextMode });
    updateHeaderUI();
    showToast(nextMode === 'local' ? 'Включено время по часовому поясу графика' : 'Включено московское время (МСК, UTC+3)');
  };

  screenPillBtn.onclick = () => {
    if (screenshotUrl) {
      screenshotUrl = '';
      updateScreenshotPillUI();
      updatePreviewBar();
      showToast('Ссылка на скриншот удалена');
    } else {
      showToast('Нажмите Alt+S на графике — ссылка прикрепится автоматически!');
    }
  };

  document.getElementById('bt-min-btn').onclick = () => { saveState({ minimized: !state.minimized }); updateHeaderUI(); };
  openSettingsBtn.onclick = () => toggleSettingsView();
  document.getElementById('bt-close-settings-btn').onclick = () => toggleSettingsView(false);
  document.getElementById('bt-toggle-cols-btn').onclick = () => { showColumnEditor = !showColumnEditor; renderSettingsDrawer(); };

  document.getElementById('bt-preset-select').onchange = e => {
    const val = e.target.value;
    showColumnEditor = (val === 'custom');
    saveState({ presetKey: val, customColumns: val === 'custom' ? getActiveColumns() : null });
    manualSessionOverride = null;
    renderMainFields();
    renderSettingsDrawer();
    updateHeaderUI();
    if (val !== 'custom') toggleSettingsView(false);
  };

  document.getElementById('bt-reset-schema').onclick = () => {
    const baseKey = state.presetKey === 'custom' ? 'bt_main' : state.presetKey;
    showColumnEditor = false;
    saveState({ presetKey: baseKey, customColumns: null });
    renderMainFields();
    renderSettingsDrawer();
    updateHeaderUI();
    showToast('Колонки возвращены к стандартному шаблону');
  };

  document.getElementById('bt-add-token-select').onchange = e => {
    const tok = e.target.value;
    if (!tok) return;
    saveState({ presetKey: 'custom', customColumns: [...getActiveColumns(), tok], customPasteCol: getActivePasteCol(), customStartRow: getActiveStartRow() });
    e.target.value = '';
    renderMainFields();
    renderSettingsDrawer();
    updateHeaderUI();
  };

  document.getElementById('bt-copy-batch-btn').onclick = () => {
    const bucket = getTplBucket();
    if (!bucket.batchQueue.length) { showToast('Список сохранённых сделок пока пуст', 'error'); return; }
    const allTsv = bucket.batchQueue.map(item => (item && item.candleDate) ? renderSnapshotToRow(item).tsv : (item?.tsv || '')).filter(Boolean).join('\n');
    const pasteCol = getActivePasteCol();
    isInternalCopy = true;
    navigator.clipboard.writeText(allTsv).then(() => {
      isInternalCopy = false;
      showToast(`✓ Скопировано сделок: ${bucket.batchQueue.length}! Вставьте в столбец ${pasteCol}`);
    }).catch(() => { isInternalCopy = false; });
  };

  document.getElementById('bt-clear-batch-btn').onclick = () => {
    saveTplBucket({ batchQueue: [], cumR: 0, maxCumR: 0, lastSavedSig: '' });
    screenshotUrl = '';
    manualSessionOverride = null;
    if (commentInpEl) commentInpEl.value = '';
    renderMainFields();
    updateHeaderUI();
    showToast('Список сохранённых сделок очищен');
  };

  function switchChartTrade(delta) {
    if (!allChartTrades.length) return;
    activeTradeIdx = (activeTradeIdx + delta + allChartTrades.length) % allChartTrades.length;
    currentTrade = allChartTrades[activeTradeIdx];
    lastSelectedToolId = currentTrade.id;
    lastChartSelectedId = detectCurrentlySelectedToolId(getTvModel());
    screenshotUrl = '';
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
    isInternalCopy = true;
    navigator.clipboard.writeText(tsv).then(() => {
      isInternalCopy = false;
      showToast('✓ Скопированы 3 цены: Вход, Стоп и Цель');
    }).catch(() => { isInternalCopy = false; });
  };

  commentInpEl.addEventListener('input', () => updatePreviewBar());
  document.getElementById('bt-clear-comment').onclick = () => { commentInpEl.value = ''; updatePreviewBar(); commentInpEl.focus(); };
  document.getElementById('bt-toggle-presets').onclick = e => {
    e.stopPropagation();
    commentMenuEl.style.display = commentMenuEl.style.display === 'block' ? 'none' : 'block';
  };
  commentMenuEl.querySelectorAll('.bt-cm-item').forEach(item => {
    item.onclick = () => { commentInpEl.value = item.innerText; commentMenuEl.style.display = 'none'; updatePreviewBar(); };
  });

  modalEl.addEventListener('click', e => {
    if (!e.target.closest('#bt-comment-menu') && !e.target.closest('#bt-toggle-presets')) commentMenuEl.style.display = 'none';
    const ccMenu = document.getElementById('bt-cc-menu');
    if (ccMenu && !e.target.closest('#bt-cc-wrap')) ccMenu.style.display = 'none';
  });

  mainCopyBtn.onclick = () => buildAndCopyRow(false);
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
      saveState({ posTop: parseInt(modalEl.style.top, 10) || 20, posRight: parseInt(modalEl.style.right, 10) || 20 });
    });
  })();

  function openModal() {
    const prevId = currentTrade?.id;
    if (!scanChartTrades(false, true)) return;
    if (currentTrade?.id !== prevId) {
      manualSessionOverride = null;
      if (commentInpEl) commentInpEl.value = '';
    }
    if (commentMenuEl) commentMenuEl.style.display = 'none';
    toggleSettingsView(false);
    renderMainFields();
    updateHeaderUI();
    if (document.activeElement && typeof document.activeElement.blur === 'function') document.activeElement.blur();
    modalEl.style.display = 'block';
    modalEl.focus();
  }

  window.addEventListener('paste', e => {
    if (modalEl.style.display !== 'block') return;
    const pasted = (e.clipboardData || window.clipboardData)?.getData('text') || '';
    const urlMatch = pasted.trim().match(/https?:\/\/[^\s]+/i);
    if (urlMatch) {
      e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
      screenshotUrl = urlMatch[0];
      updateScreenshotPillUI();
      updatePreviewBar();
      showToast('📸 Скриншот прикреплён ✓');
      return;
    }
    if (!isOurInputFocused(e.target)) { e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation(); }
  }, true);

  function isOurInputFocused(target) {
    return !!(target && modalEl.contains(target) && ['INPUT', 'TEXTAREA', 'SELECT'].includes(target.tagName));
  }

  window.addEventListener('keypress', e => {
    if (modalEl.style.display === 'block' && !isOurInputFocused(e.target)) {
      e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
    }
  }, true);

  window.addEventListener('keydown', e => {
    const isOpen = modalEl.style.display === 'block';
    const inOurInput = isOurInputFocused(e.target);
    const isQKey = e.code === 'KeyQ' || ['q', 'й', 'œ', 'ј'].includes((e.key || '').toLowerCase());
    const isOpenShortcut = e.key === 'F2' || e.code === 'F2' || e.keyCode === 113 || (isQKey && (e.ctrlKey || e.altKey) && !e.metaKey);

    if (isOpenShortcut && !inOurInput) {
      e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
      openModal();
      return;
    }
    if (!isOpen) return;

    if (e.key === 'Escape') {
      e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
      const ccMenu = document.getElementById('bt-cc-menu');
      if (commentMenuEl && commentMenuEl.style.display === 'block') commentMenuEl.style.display = 'none';
      else if (ccMenu && ccMenu.style.display === 'block') ccMenu.style.display = 'none';
      else if (inOurInput) { e.target.blur(); modalEl.focus(); }
      else if (isSettingsOpen) toggleSettingsView(false);
      else modalEl.style.display = 'none';
      return;
    }

    if (e.key === 'Enter' && !isSettingsOpen) {
      e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
      buildAndCopyRow(false);
      return;
    }

    if (e.key === 'Tab' && !inOurInput && !isSettingsOpen) {
      e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
      commentInpEl?.focus();
      return;
    }

    if (!inOurInput && !isSettingsOpen && !e.ctrlKey && !e.altKey && !e.metaKey) {
      if (e.key === 'ArrowLeft') { e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation(); switchChartTrade(-1); return; }
      if (e.key === 'ArrowRight') { e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation(); switchChartTrade(1); return; }

      const zoneMap = { '1': 'Ордер-блок', '2': 'Имбаланс', '3': 'Материнская свеча', '4': 'Зона ликвидности', '5': 'Правильный откат' };
      const poiMap = { '1': 'OB', '2': 'Имбаланс', '3': '50% МС' };
      const formMap = { '1': 'SCOB', '2': 'Свип', '3': 'Пин-бар (IFS)', '4': 'Поглощение', '5': 'Правильный откат' };

      const hasResOff = getActiveColumns().includes('res_off');
      const resMap = {
        'q': 'Тейк', 'й': 'Тейк', 'w': 'Стоп', 'ц': 'Стоп',
        'e': hasResOff ? 'Сейф/БУ' : 'Безубыток', 'у': hasResOff ? 'Сейф/БУ' : 'Безубыток',
        'r': hasResOff ? 'Бу' : 'Пропуск', 'к': hasResOff ? 'Бу' : 'Пропуск'
      };
      const res30Map = { 'q': 'цель', 'й': 'цель', 'w': 'стоп', 'ц': 'стоп', 'e': 'безубыток', 'у': 'безубыток', 'r': 'закрыл руками', 'к': 'закрыл руками' };
      const trendMap = { 'a': 'Тренд вверх', 'ф': 'Тренд вверх', 's': 'Тренд вниз', 'ы': 'Тренд вниз', 'd': 'Боковик', 'в': 'Боковик', 'f': 'Импульс/новости', 'а': 'Импульс/новости' };
      const confMap = { 'z': 'Слом структуры', 'я': 'Слом структуры', 'x': 'Реакция от зоны', 'ч': 'Реакция от зоны', 'c': 'Поглощение', 'с': 'Поглощение', 'v': 'Снятие ликвидности', 'м': 'Снятие ликвидности', 'b': 'Без подтверждения', 'и': 'Без подтверждения' };

      const kl = (e.key || '').toLowerCase();

      if (zoneMap[e.key] || poiMap[e.key] || formMap[e.key]) {
        e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
        if (document.getElementById('f-zone') && zoneMap[e.key]) { setChipVal('f-zone', zoneMap[e.key]); saveState({ zone: zoneMap[e.key] }); }
        else if (document.getElementById('f-poi-type') && poiMap[e.key]) setChipVal('f-poi-type', poiMap[e.key]);
        else if (document.getElementById('f-form-m1') && formMap[e.key]) setChipVal('f-form-m1', formMap[e.key]);
        updatePreviewBar();
      } else if (resMap[kl]) {
        e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
        if (document.getElementById('f-res30')) {
          const v30 = res30Map[kl];
          setChipVal('f-res30', v30);
          const manWrap = document.getElementById('bt-manual-r-wrap');
          if (manWrap) manWrap.style.display = (v30 === 'частичная фиксация' || v30 === 'закрыл руками') ? 'flex' : 'none';
        } else {
          setChipVal('f-res', resMap[kl]);
          if (resMap[kl] === 'Пропуск' && getChipVal('f-liq') === 'нет') {
            const confEl = document.getElementById('f-conf');
            if (confEl) confEl.value = 'Без подтверждения';
          }
          autoSuggestError();
        }
        updatePreviewBar();
      } else if (trendMap[kl]) {
        e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
        setChipVal('f-trend', trendMap[kl]);
        saveState({ trend: trendMap[kl] });
        updatePreviewBar();
      } else if (confMap[kl]) {
        e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
        const confEl = document.getElementById('f-conf');
        if (confEl) { confEl.value = confMap[kl]; saveState({ conf: confMap[kl] }); autoSuggestError(); }
        updatePreviewBar();
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
        updatePreviewBar();
      }
    }
  }, true);

})();
