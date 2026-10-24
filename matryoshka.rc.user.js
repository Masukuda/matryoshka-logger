// ==UserScript==
// @name         Matryoshka Backtest Auto-Logger
// @namespace    http://tampermonkey.net/
// @version      4.0
// @description  Экспорт сделок из TradingView в Excel и Google Таблицы
// @match        https://*.tradingview.com/*
// @grant        unsafeWindow
// @grant        GM_xmlhttpRequest
// @connect      raw.githubusercontent.com
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

  const SCRIPT_VERSION = '4.0';
  const SCRIPT_RAW_URL = 'https://raw.githubusercontent.com/Masukuda/matryoshka-logger/main/matryoshka.user.js';
  const MIN_ZI_VERSION = 3.3;
  const MIN_MV_PLOTS = 23;

  const LOGO_SVG = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" style="flex-shrink:0;"><circle cx="12" cy="7" r="4.5" stroke="#e5c158" stroke-width="2"/><path d="M5.5 19.5C5.5 15.634 8.41015 12.5 12 12.5C15.5899 12.5 18.5 15.634 18.5 19.5C18.5 20.8807 17.3807 22 16 22H8C6.61929 22 5.5 20.8807 5.5 19.5Z" stroke="#e5c158" stroke-width="2"/><circle cx="12" cy="17.5" r="1.5" fill="#e5c158"/></svg>`;
  const NATIVE_TOOLBAR_SVG = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 28 28" width="28" height="28" fill="none"><circle cx="14" cy="9.5" r="4.5" stroke="#e5c158" stroke-width="1.8"/><path d="M7.5 21C7.5 17.134 10.4101 14 14 14C17.5899 14 20.5 17.134 20.5 21C20.5 22.3807 19.3807 23.5 18 23.5H10C8.61929 23.5 7.5 22.3807 7.5 21Z" stroke="#e5c158" stroke-width="1.8"/><circle cx="14" cy="19" r="1.6" fill="#e5c158"/></svg>`;
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
    bt_30: {
      name: 'Бэктест 30 сделок', pasteCol: 'B', startRow: 4, sessionStyle: 'eu_us', tfStyle: 'cyr',
      columns: ['date', 'time_msk', 'session', 'tf', 'dir_lower', 'flow_d', 'flow_4h', 'flow_1h', 'flow_15m', 'struct_30', 'zone_30', 'fuel_30', 'touch_30', 'against_flow', 'swept_30', 'swept_level_30', 'candle_col', 'entry', 'stop', 'target', 'risk_pts', 'rr', 'lots_30', 'check_30', 'res_30', 'manual_r_30', 'r_out', 'profit_usd', 'hit_1r', 'hit_3r', 'screenshot', 'comment_30']
    },
    bt_main: {
      name: 'Бэктест «Матрёшка»', pasteCol: 'B', startRow: 6, sessionStyle: 'std', tfStyle: 'lat',
      columns: ['date', 'session', 'trend', 'tf', 'zone', 'liq', 'conf', 'entry', 'stop', 'target', 'risk_pts', 'pot_pts', 'rr', 'result', 'r_out', 'cum_r', 'dd_r', 'error', 'comment', 'screenshot']
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
    presetKey: 'bt_30', migratedV4Bt30: true, customColumns: null, customPasteCol: 'B', customStartRow: 2,
    outputProfile: 'gs_ru', decimalSep: ',', timeMode: 'msk', trend: 'Тренд вверх',
    zone: 'Ордер-блок', liq: 'да', conf: 'Слом структуры', struct30: '', emotion: 'Спокоен',
    aoiNum: 1, flowD: 'вверх', flow4H: 'вверх', flow1H: 'вверх', flow15M: 'вверх',
    safeExportMode: '1row', safeShare: 0.5, autoStripOpen: false,
    badgeMode: 'pct', riskPctPer1R: 0.3,
    perTpl: {}, minimized: false, posTop: 56, posRight: 20
  };

  function isMatryoshkaZiTitle(title) {
    const t = String(title || '').toLowerCase();
    return (t.includes('матрёшка') || t.includes('матрешка')) &&
      (t.includes('зоны интереса') || t.includes(' зи') || t.startsWith('зи '));
  }

  function loadState() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return { ...defaultState };
      const parsed = { ...defaultState, ...JSON.parse(raw) };
      if (!['gs_ru', 'ru_form', 'en_form'].includes(parsed.outputProfile)) parsed.outputProfile = 'gs_ru';
      
      if (!parsed.migratedV4Bt30) {
        if (parsed.presetKey === 'bt_main') parsed.presetKey = 'bt_30';
        if (!parsed.posTop || parsed.posTop < 48) parsed.posTop = 56;
        parsed.migratedV4Bt30 = true;
        try { localStorage.setItem(STORAGE_KEY, JSON.stringify(parsed)); } catch (e) {}
      }
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
  let remoteUpdateVersion = null;

  
  let activeSafeState = {
    enabled: false,
    mode: 'off', // 'off' | 'chart' | '2r' | '3r' | 'mid' | 'custom'
    price: '',
    share: state.safeShare || 0.5,
    exportMode: state.safeExportMode || '1row', // '1row' | '2rows'
    remOutcome: 'auto' // 'auto' | 'цель' | 'безубыток' | 'стоп'
  };

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

  function handleRemoteScriptText(txt) {
    if (!txt) return;
    const m = txt.match(/\/\/\s*@version\s+([0-9.]+)/i);
    if (m && m[1]) {
      const remoteVer = m[1].trim();
      sessionStorage.setItem('bt_remote_ver', remoteVer);
      if (remoteVer !== SCRIPT_VERSION && parseFloat(remoteVer) > parseFloat(SCRIPT_VERSION)) {
        remoteUpdateVersion = remoteVer;
        if (typeof updateVersionBannerUI === 'function') updateVersionBannerUI();
      }
    }
  }

  function checkGithubScriptUpdate() {
    try {
      const cached = sessionStorage.getItem('bt_remote_ver');
      if (cached) {
        if (cached !== SCRIPT_VERSION && parseFloat(cached) > parseFloat(SCRIPT_VERSION)) {
          remoteUpdateVersion = cached;
          if (typeof updateVersionBannerUI === 'function') updateVersionBannerUI();
        }
        return;
      }

      if (typeof GM_xmlhttpRequest !== 'undefined') {
        GM_xmlhttpRequest({
          method: 'GET',
          url: `${SCRIPT_RAW_URL}?t=${Date.now()}`,
          headers: { 'Cache-Control': 'no-cache' },
          onload: function(res) {
            if (res.status === 200 && res.responseText) handleRemoteScriptText(res.responseText);
          },
          onerror: function() {}
        });
      } else {
        fetch(`${SCRIPT_RAW_URL}?t=${Date.now()}`, { cache: 'no-store' })
          .then(r => r.ok ? r.text() : '')
          .then(txt => handleRemoteScriptText(txt))
          .catch(() => {});
      }
    } catch (e) {}
  }

  function inspectIndicatorsHealth(ziStudy, mvStudy) {
    const warnings = [];
    let isAllLatest = true;

    if (!ziStudy) {
      isAllLatest = false;
      warnings.push('На графике нет индикатора <b>«Матрёшка ЗИ 3.3+»</b>.');
    } else {
      const meta = typeof ziStudy.metaInfo === 'function' ? ziStudy.metaInfo() : ziStudy._metaInfo;
      const title = String(meta?.description || meta?.name || '');
      const isVisible = readWatched(ziStudy.properties?.()?.visible) !== false;
      const verMatch = title.match(/(\d+\.\d+)\+?/);
      const verNum = verMatch ? parseFloat(verMatch[1]) : 0;

      if (!isVisible) {
        isAllLatest = false;
        warnings.push('Индикатор <b>«Матрёшка ЗИ»</b> скрыт (включите «глазик» в списке индикаторов).');
      }
      if (verNum < MIN_ZI_VERSION) {
        isAllLatest = false;
        warnings.push(`На графике старая версия <b>«${title}»</b> вместо <b>3.3+</b>. Удалите индикатор и добавьте заново.`);
      }
    }

    if (!mvStudy) {
      isAllLatest = false;
      warnings.push('На графике нет индикатора <b>«Матрёшка Модели входа»</b>.');
    } else {
      const meta = typeof mvStudy.metaInfo === 'function' ? mvStudy.metaInfo() : mvStudy._metaInfo;
      const plotsCount = Array.isArray(meta?.plots) ? meta.plots.length : 0;
      const isVisible = readWatched(mvStudy.properties?.()?.visible) !== false;

      if (!isVisible) {
        isAllLatest = false;
        warnings.push('Индикатор <b>«Модели входа»</b> скрыт (включите «глазик» в списке индикаторов).');
      }
      if (plotsCount < MIN_MV_PLOTS) {
        isAllLatest = false;
        warnings.push('На графике старая версия <b>«Матрёшка Модели входа»</b> (обновите до <b>6.3+</b> передобавлением на график).');
      }
    }

    return { isAllLatest, warnings };
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
          for (let d = 1; d <= 20; d++) {
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
        for (let d = 1; d <= 15; d++) {
          const r1 = getPlotRow(plotsObj, estIdx - d);
          if (r1 && r1[0] === targetTime) return { index: estIdx - d, val: r1 };
          const r2 = getPlotRow(plotsObj, estIdx + d);
          if (r2 && r2[0] === targetTime) return { index: estIdx + d, val: r2 };
        }
      }

      const searchSegment = (loInit, hiInit) => {
        let lo = loInit, hi = hiInit;
        let guard = 0;
        while (lo <= hi && guard++ < 60) {
          const mid = Math.floor((lo + hi) / 2);
          const row = getPlotRow(plotsObj, mid);
          if (!row || row[0] == null || isNaN(row[0])) {
            
            lo = mid + 1;
            continue;
          }
          if (row[0] === targetTime) return { index: mid, val: row };
          if (row[0] < targetTime) lo = mid + 1;
          else hi = mid - 1;
        }
        
        for (let k = lo - 4; k <= lo + 4; k++) {
          const rNear = getPlotRow(plotsObj, k);
          if (rNear && rNear[0] === targetTime) return { index: k, val: rNear };
        }
        return null;
      };

      if (rawFirst < -500000 && rawLast >= 0) {
        const hitMain = searchSegment(0, rawLast);
        if (hitMain) return hitMain;
        return searchSegment(rawFirst, -999000);
      }
      return searchSegment(Math.max(0, rawFirst), rawLast) || searchSegment(rawFirst, rawLast);
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

  function parseZoneLabelText(rawText, boxColor = 0, boxTva = '', boxBgColor = 0, boxTha = '') {
    const t = String(rawText || '').trim();
    const low = t.toLowerCase();
    if (low.includes('держит') || low.includes('слом') || low.includes('пробой') || low.startsWith('полка')) return null;

    const rgb = Number(boxColor || 0) & 0xFFFFFF;
    const bgAlpha = (Number(boxBgColor || 0) >>> 24) & 0xFF;

    if (!t) {
      if (boxTva === 'center') return null;
      if (boxTva === 'bottom' && boxTha === 'center') return null;
      if (rgb !== 4536050 && rgb !== 8493320 && rgb !== 39167 && rgb !== 16736809) return null;
    }

    const gradeMatch = t.match(/(?:^|[\s·])(A\+|A|B)(?=$|[\s·])/i);
    const grade = gradeMatch ? gradeMatch[1].toUpperCase() : '';

    let zoneType = null, htfRank = 0;
    if (low.includes('разворотная')) zoneType = 'Разворотная зона';
    else if (low.includes('сильный покупатель')) zoneType = `Сильный покупатель · ${grade || 'A+'}`;
    else if (low.includes('средний покупатель')) zoneType = `Средний покупатель · ${grade || 'A'}`;
    else if (low.includes('красивый покупатель')) zoneType = `Красивый покупатель · ${grade || 'B'}`;
    else if (low.includes('сильный продавец')) zoneType = `Сильный продавец · ${grade || 'A+'}`;
    else if (low.includes('средний продавец')) zoneType = `Средний продавец · ${grade || 'A'}`;
    else if (low.includes('красивый продавец')) zoneType = `Красивый продавец · ${grade || 'B'}`;
    else if (low.includes('покупатель')) {
      zoneType = grade === 'A+' ? 'Сильный покупатель · A+' : (grade === 'B' ? 'Красивый покупатель · B' : 'Средний покупатель · A');
    } else if (low.includes('продавец')) {
      zoneType = grade === 'A+' ? 'Сильный продавец · A+' : (grade === 'B' ? 'Красивый продавец · B' : 'Средний продавец · A');
    } else if (low.includes('зона входа')) {
      zoneType = `Зона входа · ${grade || 'A+'}`;
    } else if (low.includes('проторговка')) {
      zoneType = low.includes('спит') ? 'Проторговка · спит' : 'Проторговка';
    } else if (/(^|\s|·)д\s*(·\s*)?0[.,]5/i.test(low)) {
      zoneType = grade ? `Зона Д 0.5 · ${grade}` : 'Зона Д 0.5'; htfRank = 4;
    } else if (/4ч\s*(·\s*)?0[.,]5/i.test(low)) {
      zoneType = grade ? `Зона 4Ч 0.5 · ${grade}` : 'Зона 4Ч 0.5'; htfRank = 3;
    } else if (/1ч\s*(·\s*)?0[.,]5/i.test(low)) {
      zoneType = grade ? `Зона 1Ч 0.5 · ${grade}` : 'Зона 1Ч 0.5'; htfRank = 2;
    } else if (/15м\s*(·\s*)?0[.,]5/i.test(low)) {
      zoneType = grade ? `Зона 15м 0.5 · ${grade}` : 'Зона 15м 0.5'; htfRank = 1;
    }

    if (!zoneType && !t) {
      if (rgb === 4536050 && boxTva === 'top') zoneType = 'Сильный продавец · A+';
      else if (rgb === 8493320 && boxTva === 'top') zoneType = 'Сильный покупатель · A+';
      else if (rgb === 39167 && boxTva === 'bottom') zoneType = 'Проторговка';
      else if (rgb === 16736809 && boxTva === 'bottom') {
        return {
          zoneType: 'Проторговка', fuel: 'нет', touch: '1', maxTouchNum: 1, against: 'да',
          side: 'ANY', isProtor: true, htfRank: 0, isDeadFallback: true, isHiddenAlpha0: bgAlpha === 0,
          rawText: 'Проторговка · против потока (скрыта/отработана)'
        };
      }
    }
    if (!zoneType) return null;

    const fuel = low.includes('топливо') ? 'да' : 'нет';
    const against = (low.includes('против потока') || rgb === 16736809) ? 'да' : 'нет';
    let touch = '1', maxTouchNum = 1;
    const touchMatch = low.match(/касание\s*(\d+)/);
    if (touchMatch) {
      maxTouchNum = parseInt(touchMatch[1], 10) || 1;
      if (maxTouchNum === 2) touch = '2';
      else if (maxTouchNum >= 3) touch = '3 и больше';
    }

    let side = 'ANY';
    if (zoneType.includes('покупатель') || low.includes('лонг') || (zoneType.startsWith('Зона входа') && rgb === 8493320)) side = 'LONG';
    else if (zoneType.includes('продавец') || low.includes('шорт') || (zoneType.startsWith('Зона входа') && rgb === 4536050)) side = 'SHORT';

    const isProtor = zoneType.startsWith('Проторговка');
    return {
      zoneType, fuel, touch, maxTouchNum, against, side, isProtor, htfRank,
      isDeadFallback: !t, isHiddenAlpha0: !t && bgAlpha === 0,
      rawText: t.replace(/\s*\n\s*/g, ' / ') || `${zoneType} (скрыта/отработана)`
    };
  }

  
  const SWEEP_STRENGTH = {
    'склейка Д · 4Ч': 90,
    'уровень Д': 80,
    'уровень 4Ч': 70,
    'склейка 15м/1Ч': 60,
    'уровень 1Ч': 50,
    'уровень 15м': 40,
    'полка': 30,
    'приманка': 20
  };

  
  
  function parseLiquidityLabelText(rawText, fallbackY) {
    const t = String(rawText || '').trim();
    if (!t) return null;
    const low = t.toLowerCase();
    if (low.startsWith('стопы')) return null;

    let price = Number(fallbackY || 0);
    const priceMatch = t.match(/(\d{3,6}[.,]\d+)/);
    if (priceMatch) {
      const parsedP = parseFloat(priceMatch[1].replace(',', '.'));
      if (!isNaN(parsedP) && parsedP > 0) price = parsedP;
    }

    const detectedTypes = [];
    const addType = (st) => { if (st && !detectedTypes.includes(st)) detectedTypes.push(st); };

    if (/(м\/н\/д|н\/д|(^|[\s·./])д)\s*[·./]\s*4ч/i.test(low)) addType('склейка Д · 4Ч');
    if (low.includes('15м/1ч') || low.includes('15m/1h') || /(1ч|1h)\s*[·./]\s*(15м|15m)/i.test(low)) addType('склейка 15м/1Ч');
    if (/(^|[\s·./:])(д|d|d1|н|м)(\s+(макс|мин)|$|[\s·./])/i.test(low)) addType('уровень Д');
    if (/(^|[\s·./:])(4ч|4h|h4)(\s+(макс|мин)|$|[\s·./])/i.test(low)) addType('уровень 4Ч');
    if (/(^|[\s·./:])(1ч|1h|h1)(\s+(макс|мин)|$|[\s·./])/i.test(low)) addType('уровень 1Ч');
    if (/(^|[\s·./:])(15м|15m|m15)(\s+(макс|мин)|$|[\s·./])/i.test(low)) addType('уровень 15м');
    if (low === 'полка' || low.startsWith('полка ') || low.includes('· полка')) addType('полка');
    if (low.includes('приманка')) addType('приманка');

    const isDirectSweep = low.includes('lq') && low.includes('снят');
    if (isDirectSweep && detectedTypes.length === 0) addType('уровень 15м');
    if (detectedTypes.length === 0) return null;

    detectedTypes.sort((a, b) => (SWEEP_STRENGTH[b] || 0) - (SWEEP_STRENGTH[a] || 0));
    const sweptType = detectedTypes[0];

    if (!isDirectSweep && (!price || isNaN(price))) return null;
    return {
      sweptType,
      allTypes: detectedTypes,
      price: Number(price || 0),
      isDirectSweepLabel: isDirectSweep,
      rawText: t
    };
  }

  function computeTouchAtEntryBar(bars, barX1, triggerIdx, bottom, top, maxTouchNum) {
    if (!bars || maxTouchNum <= 1 || barX1 < -100000 || triggerIdx == null || triggerIdx <= barX1) return '1';
    const zoneHeight = Math.max(0.05, top - bottom);
    const leaveTol = zoneHeight * 0.25;
    let touches = 1;
    let barsOutside = 0;
    let wasInZoneAfterBirth = true;

    for (let i = barX1 + 1; i <= triggerIdx; i++) {
      const b = getPlotRow(bars, i);
      if (!b) continue;
      const hi = Number(b[2]), lo = Number(b[3]);
      if (isNaN(hi) || isNaN(lo)) continue;
      const inZone = (hi >= bottom && lo <= top);
      const clearlyOutside = (lo > top + leaveTol) || (hi < bottom - leaveTol);

      if (!inZone) {
        if (clearlyOutside || barsOutside > 0) barsOutside++;
      } else {
        
        if (!wasInZoneAfterBirth && barsOutside >= 2) {
          touches++;
        }
        barsOutside = 0;
        wasInZoneAfterBirth = true;
      }
      if (barsOutside >= 2) wasInZoneAfterBirth = false;
    }
    const actual = Math.max(1, Math.min(touches, maxTouchNum));
    return actual === 1 ? '1' : (actual === 2 ? '2' : '3 и больше');
  }

  function findZoneBreakBar(bars, barX1, bottom, top, boxColor) {
    if (!bars || barX1 == null || barX1 < 0) return null;
    const rgb = Number(boxColor || 0) & 0xFFFFFF;
    const isBullishZone = (rgb === 8493320 || rgb === 16736809);
    const lastIdx = typeof bars.lastIndex === 'function' ? bars.lastIndex() : barX1 + 300;
    const maxScan = Math.min(lastIdx ?? (barX1 + 300), barX1 + 300);
    for (let i = barX1 + 2; i <= maxScan; i++) {
      const b = getPlotRow(bars, i);
      if (!b) continue;
      const cl = Number(b[4]);
      if (isNaN(cl)) continue;
      if (isBullishZone && cl < bottom) return i;
      if (!isBullishZone && cl > top) return i;
    }
    return null;
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
          if (!top || !bottom) return res;

          if (text === 'Полка') {
            mem.levels.set(`полка_${bottom.toFixed(2)}`, { sweptType: 'полка', price: bottom, barX1, lastSeenBar: curBar, rawText: `Полка ${bottom}` });
            trimMapFIFO(mem.levels);
          } else {
            const zInfo = parseZoneLabelText(text, val.c, val.tva, val.bc, val.tha);
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
            for (const st of (lInfo.allTypes || [lInfo.sweptType])) {
              mem.lqLabels.set(`lq_${barX}_${st}`, { barX, sweptType: st, price: lInfo.price, rawText: lInfo.rawText });
            }
            trimMapFIFO(mem.lqLabels);
          } else {
            for (const st of (lInfo.allTypes || [lInfo.sweptType])) {
              const key = `${st}_${lInfo.price.toFixed(2)}`;
              const prev = mem.levels.get(key);
              mem.levels.set(key, {
                sweptType: st, price: lInfo.price,
                barX1: prev ? prev.barX1 : (barX - 20), lastSeenBar: curBar, rawText: lInfo.rawText
              });
            }
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
            for (const st of (lInfo.allTypes || [lInfo.sweptType])) {
              const key = `${st}_${lInfo.price.toFixed(2)}`;
              const existing = mem.levels.get(key);
              if (existing) {
                existing.sweptAtBar = curBar;
                existing.lastSeenBar = curBar;
              } else {
                mem.levels.set(key, {
                  sweptType: st, price: lInfo.price,
                  barX1: curBar - 25, lastSeenBar: curBar, sweptAtBar: curBar, rawText: lInfo.rawText
                });
              }
            }
            trimMapFIFO(mem.levels);
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
        if (!/матрёшка|матрешка/i.test(title)) continue;

        const g = typeof s.graphics === 'function' ? s.graphics() : s._graphics;
        if (!g) continue;
        if (isMatryoshkaZiTitle(title)) mainG = g;

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
          if (!top || !bottom) continue;

          if (text === 'Полка') {
            mem.levels.set(`полка_${bottom.toFixed(2)}`, { sweptType: 'полка', price: bottom, barX1, lastSeenBar: currentRightBar, rawText: `Полка ${bottom}` });
            trimMapFIFO(mem.levels);
            continue;
          }

          const zInfo = parseZoneLabelText(text, b.c, b.tva, b.bc, b.tha);
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
            for (const st of (lInfo.allTypes || [lInfo.sweptType])) {
              mem.lqLabels.set(`lq_${barX}_${st}`, { barX, sweptType: st, price: lInfo.price, rawText: lInfo.rawText });
            }
            trimMapFIFO(mem.lqLabels);
          } else {
            const matchLine = lines.find(ln => Math.abs(Number(ln.y1 || 0) - lInfo.price) < 0.05);
            const birthBar = matchLine ? toBar(Math.min(matchLine.x1, matchLine.x2), barX - 20) : (barX - 20);
            for (const st of (lInfo.allTypes || [lInfo.sweptType])) {
              const key = `${st}_${lInfo.price.toFixed(2)}`;
              const prev = mem.levels.get(key);
              mem.levels.set(key, {
                sweptType: st, price: lInfo.price,
                barX1: prev ? Math.min(prev.barX1, birthBar) : birthBar,
                lastSeenBar: currentRightBar, sweptAtBar: prev?.sweptAtBar || null, rawText: lInfo.rawText
              });
            }
            trimMapFIFO(mem.levels);
          }
        }
      }
    } catch (e) {}
    return mainG;
  }

  function resolveZiPlotIndices(s) {
    const idx = {
      h4Max: 1, h4Min: 2, dMax: 3, dMin: 4,
      cont: 5, bos: 6, fake: 7, inCont: 8, inBos: 9,
      swLow: 12, swHigh: 13, swHtfHigh: 14, swHtfLow: 15,
      swWeek: 16, swMonth: 17, sw15m: 18
    };
    try {
      const meta = typeof s?.metaInfo === 'function' ? s.metaInfo() : s?._metaInfo;
      const plotsMeta = meta?.plots;
      const styles = meta?.styles || meta?.defaults?.styles;
      if (Array.isArray(plotsMeta) && styles) {
        plotsMeta.forEach((p, i) => {
          const t = String(styles[p.id]?.title || p.title || '').toLowerCase();
          const col = i + 1;
          if (t.includes('4ч макс')) idx.h4Max = col;
          else if (t.includes('4ч мин')) idx.h4Min = col;
          else if (t.includes('макс') && (t.includes('д') || t.includes('н'))) idx.dMax = col;
          else if (t.includes('мин') && (t.includes('д') || t.includes('н'))) idx.dMin = col;
          else if (t === 'продолжение') idx.cont = col;
          else if (t.includes('смена потока')) idx.bos = col;
          else if (t.includes('ложный пробой')) idx.fake = col;
          else if (t.includes('внутреннее продолжение')) idx.inCont = col;
          else if (t.includes('внутренняя смена')) idx.inBos = col;
          else if (t.includes('снята нижняя')) idx.swLow = col;
          else if (t.includes('снята верхняя')) idx.swHigh = col;
          else if (t.includes('снят уровень сверху')) idx.swHtfHigh = col;
          else if (t.includes('снят уровень снизу')) idx.swHtfLow = col;
          else if (t.includes('недели')) idx.swWeek = col;
          else if (t.includes('месяца')) idx.swMonth = col;
          else if (t.includes('снятие или тест')) idx.sw15m = col;
        });
      }
    } catch (e) {}
    return idx;
  }

  function resolveMvPlotIndices(s, sampleRow) {
    let colorIdx = 3, modelIdx = 4, zoneDirIdx = 8, forbidIdx = 9, structIdx = 13, contLongIdx = 16, contShortIdx = 17;
    try {
      const meta = typeof s?.metaInfo === 'function' ? s.metaInfo() : s?._metaInfo;
      const plotsMeta = meta?.plots;
      const styles = meta?.styles || meta?.defaults?.styles;
      if (Array.isArray(plotsMeta) && styles) {
        plotsMeta.forEach((p, i) => {
          const stTitle = String(styles[p.id]?.title || p.title || p.id || '').toLowerCase();
          const col = i + 1;
          if (stTitle.includes('цвет')) colorIdx = col;
          else if (stTitle.includes('модел')) modelIdx = col;
          else if (stTitle.includes('мв зона')) zoneDirIdx = col;
          else if (stTitle.includes('мв запрет')) forbidIdx = col;
          else if (stTitle.includes('структура графика')) structIdx = col;
          else if (stTitle.includes('продолжение лонг')) contLongIdx = col;
          else if (stTitle.includes('продолжение шорт')) contShortIdx = col;
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
    return { colorIdx, modelIdx, zoneDirIdx, forbidIdx, structIdx, contLongIdx, contShortIdx };
  }

  function isValidStructureOnBars(bars, bx1, bx2, price, isFromBox = false) {
    if (bx2 == null || bx2 < -100000 || bx2 >= 900000) return false;
    if (bx1 != null && bx1 > -100000) {
      if (isFromBox ? (bx2 < bx1) : (bx2 <= bx1)) return false;
      if (bx2 - bx1 > 400) return false;
    }
    if (!price || isNaN(price) || price <= 0) return false;
    if (!bars || isFromBox) return true;

    const r2 = getPlotRow(bars, bx2);
    if (r2) {
      const hi2 = Number(r2[2]), lo2 = Number(r2[3]);
      if (!isNaN(hi2) && !isNaN(lo2)) {
        const rangeTol = Math.max(1.2, (hi2 - lo2) * 3.0);
        if (price < lo2 - rangeTol || price > hi2 + rangeTol) return false;
      }
    }
    return true;
  }

  function analyzeChartGraphicsO1(g, plotsObj, plotIdxOffset, ziIdx, mvPlotsObj, mvIdxOffset, mvMetaIdx, symKey, triggerIdx, isLong, entryPrice, stopPrice, entryExtreme, bars) {
    const out = {
      matchedZone: null, zoneOptions: [],
      matchedSweep: null, sweepOptions: [],
      detectedStructure: null, isHistorical: false
    };
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

    const rightEdgeLabelPrices = [];
    for (const lb of labels) {
      const ly = Number(lb.y || 0);
      if (ly > 0) rightEdgeLabelPrices.push(ly);
    }
    const isHtfRightEdgeLine = (bx2, yPrice) => {
      if (bx2 < currentRightBar) return false;
      return rightEdgeLabelPrices.some(p => Math.abs(p - yPrice) < 0.05);
    };

    const fakeoutBars = new Set();
    for (const l of lines) {
      if (l.st === 'dot' && l.w === 1 && (l.ci === 1027311288 || (Number(l.ci || 0) & 0xFFFFFF) === 3901112)) {
        const fBar = toBar(Math.max(l.x1, l.x2), -999999);
        if (fBar > -100000 && fBar <= currentRightBar) fakeoutBars.add(fBar);
      }
    }

    const structByBar = new Map();
    const upsertStruct = (bx1, bx2, type, dir, inner, price, isFromBox = false) => {
      if (bx2 > currentRightBar + 1 || isHtfRightEdgeLine(bx2, price)) return;
      if (!isValidStructureOnBars(bars, bx1, bx2, price, isFromBox)) return;

      if (isFromBox) {
        for (let d = -3; d <= 3; d++) {
          if (structByBar.has(bx2 + d)) return;
        }
      }

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
        upsertStruct(bx1, bx2, type, dir, false, y1, false);
      } else if (l.w === 1 && l.st === 'dot' && (l.ci === 2353346290 || l.ci === 2357303560)) {
        const dir = l.ci === 2353346290 ? '↓' : '↑';
        upsertStruct(bx1, bx2, 'слом', dir, true, y1, false);
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
      upsertStruct(bx1, bx2, type, dir, isInner, yPrice, true);
    }

    if (plotsObj && triggerIdx != null) {
      const scanFrom = Math.max(0, triggerIdx - 120);
      const scanTo = triggerIdx + 1;
      for (let bi = scanFrom; bi <= scanTo; bi++) {
        const pr = getPlotRowAtBar(plotsObj, bars, bi, plotIdxOffset);
        if (!pr) continue;
        if (pr[ziIdx.fake] === 1) fakeoutBars.add(bi);

        const isBos = pr[ziIdx.bos] === 1;
        const isCont = pr[ziIdx.cont] === 1;
        const isInBos = pr[ziIdx.inBos] === 1;
        if (!isBos && !isCont && !isInBos) continue;
        if (structByBar.has(bi)) continue;

        let dir = null;
        if (mvPlotsObj && mvMetaIdx?.structIdx) {
          const mvRow = getPlotRowAtBar(mvPlotsObj, bars, bi, mvIdxOffset);
          const stVal = Number(mvRow?.[mvMetaIdx.structIdx] || 0);
          if (stVal > 0) dir = '↑';
          else if (stVal < 0) dir = '↓';
        }
        const bRow = getPlotRow(bars, bi);
        if (!dir && bRow) {
          const op = Number(bRow[1]), cl = Number(bRow[4]);
          if (!isNaN(op) && !isNaN(cl)) dir = cl >= op ? '↑' : '↓';
        }
        if (!dir) dir = isLong ? '↑' : '↓';

        const type = (isBos || isInBos) ? 'слом' : 'пробой';
        const inner = !isBos && !isCont && isInBos;
        structByBar.set(bi, { barX1: bi - 5, barX2: bi, type, dir, inner, price: Number(bRow?.[4] || entryPrice) });
      }
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
          if (activeMajor.barX2 != null) {
            for (let bi = activeMajor.barX2 - 8; bi <= activeMajor.barX2; bi++) {
              if (fakeoutBars.has(bi)) { hadFakeout = true; break; }
              const pr = getPlotRowAtBar(plotsObj, bars, bi, plotIdxOffset);
              if (pr && pr[ziIdx.fake] === 1) { hadFakeout = true; break; }
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
    } else if (mvPlotsObj && mvMetaIdx?.structIdx && triggerIdx != null) {
      const mvRow = getPlotRowAtBar(mvPlotsObj, bars, triggerIdx, mvIdxOffset);
      const stDir = Number(mvRow?.[mvMetaIdx.structIdx] || 0);
      if (stDir !== 0) {
        const prevMvRow = getPlotRowAtBar(mvPlotsObj, bars, triggerIdx - 3, mvIdxOffset);
        const prevDir = Number(prevMvRow?.[mvMetaIdx.structIdx] || stDir);
        const arrow = stDir > 0 ? '↑' : '↓';
        out.detectedStructure = (prevDir !== 0 && prevDir !== stDir) ? `слом ${arrow}` : `пробой ${arrow}`;
      }
    }

    const tradeLow = Math.min(entryPrice, stopPrice);
    const tradeHigh = Math.max(entryPrice, stopPrice);
    const minP = Math.min(tradeLow, entryExtreme !== '' ? Number(entryExtreme) : tradeLow);
    const maxP = Math.max(tradeHigh, entryExtreme !== '' ? Number(entryExtreme) : tradeHigh);
    const tol = Math.max(0.08, Math.abs(entryPrice - stopPrice) * 0.08);

    const fuelLines = [];
    for (const l of lines) {
      if (l.st === 'dsh' && l.w === 1 && (l.ci === 11544476 || l.ci === 2796562332 || (Number(l.ci || 0) & 0xFFFFFF) === 11544476)) {
        const fy = Number(l.y1 ?? 0);
        const fbx1 = toBar(Math.min(l.x1, l.x2), -999999);
        if (fy > 0) fuelLines.push({ y: fy, barX1: fbx1 });
      }
    }

    const zoneCandidates = [];
    for (const b of boxes) {
      const top = Math.max(Number(b.y1 || 0), Number(b.y2 || 0));
      const bottom = Math.min(Number(b.y1 || 0), Number(b.y2 || 0));
      if (!top || !bottom) continue;

      const barX1 = toBar(Math.min(b.x1, b.x2), -999999);
      const barX2 = toBar(Math.max(b.x1, b.x2), 999999);

      const memKey = `${bottom.toFixed(2)}_${top.toFixed(2)}`;
      const zInfo = (!b.t && mem.zones.has(memKey))
        ? { ...mem.zones.get(memKey).zInfo }
        : parseZoneLabelText(b.t, b.c, b.tva, b.bc, b.tha);
      if (!zInfo) continue;

      const isDead2Bar = zInfo.isDeadFallback && (barX2 - barX1 <= 3);
      if (triggerIdx != null) {
        if (barX1 > triggerIdx) continue;
        if (!zInfo.isDeadFallback && barX2 < triggerIdx - 1) continue;
        if (zInfo.isDeadFallback && barX1 > -100000) {
          if (isDead2Bar && triggerIdx - barX1 > 120) continue;
          const breakBar = findZoneBreakBar(bars, barX1, bottom, top, b.c);
          if (breakBar != null && triggerIdx > breakBar + 1) continue;
        }
      }

      if (zInfo.isDeadFallback && zInfo.fuel === 'нет') {
        const hasFuelLine = fuelLines.some(fl => fl.y >= bottom - 0.05 && fl.y <= top + 0.05 && (barX1 < -100000 || fl.barX1 < -100000 || Math.abs(fl.barX1 - barX1) <= 5));
        if (hasFuelLine) zInfo.fuel = 'да';
      }

      const entryInside = (entryPrice >= bottom - tol && entryPrice <= top + tol);
      const extremeInside = (entryExtreme !== '' && Number(entryExtreme) >= bottom - tol && Number(entryExtreme) <= top + tol);
      const overlapHeight = Math.min(tradeHigh, top) - Math.max(tradeLow, bottom);
      if (!entryInside && !extremeInside && overlapHeight <= 0) continue;

      const sideMatch = (zInfo.side === (isLong ? 'LONG' : 'SHORT')) || (zInfo.side === 'ANY' && !zInfo.isProtor);
      const riskHeight = Math.max(0.05, tradeHigh - tradeLow);
      const inRiskBox = overlapHeight >= Math.min(riskHeight * 0.15, (top - bottom) * 0.25);
      const bornOnEntryBar = (triggerIdx != null && barX1 >= triggerIdx - 1);

      let rank = 30;
      if (zInfo.isProtor) rank = entryInside ? 40 : 50;
      else if (sideMatch && entryInside && inRiskBox) rank = 0;
      else if (sideMatch && inRiskBox) rank = 3;
      else if (sideMatch && extremeInside) rank = 6;
      else if (sideMatch) rank = 12;
      else if (entryInside) rank = 22;

      if (bornOnEntryBar) rank += 15;
      if (zInfo.isDeadFallback) rank += 2;

      const entryDist = entryPrice < bottom ? (bottom - entryPrice) : (entryPrice > top ? (entryPrice - top) : 0);
      const barDist = (triggerIdx != null && barX1 > -100000) ? Math.abs(triggerIdx - barX1) : 500;
      zoneCandidates.push({ rank, overlapHeight, entryDist, barDist, top, bottom, barX1, zInfo });
    }

    for (const lb of labels) {
      const z05 = parseZoneLabelText(lb.t);
      if (!z05 || !z05.htfRank) continue;
      const midY = Number(lb.y || 0);
      if (!midY) continue;

      const hitMid = (midY >= minP - tol && midY <= maxP + tol);
      if (!hitMid) continue;

      const sideMatch = (z05.side === 'ANY' || z05.side === (isLong ? 'LONG' : 'SHORT'));
      const htfRankScore = sideMatch ? (14 - z05.htfRank) : (24 - z05.htfRank);
      zoneCandidates.push({
        rank: htfRankScore,
        overlapHeight: tol * 2,
        entryDist: Math.abs(entryPrice - midY),
        barDist: 0,
        top: midY + tol,
        bottom: midY - tol,
        barX1: -999999,
        zInfo: z05
      });
    }

    if (zoneCandidates.length > 0) {
      zoneCandidates.sort((a, b) => (a.rank - b.rank) || (b.overlapHeight - a.overlapHeight) || (a.entryDist - b.entryDist) || (a.barDist - b.barDist));

      const resolveZoneCandidateObj = (cand) => {
        const touchAtEntry = (!cand.zInfo.htfRank && triggerIdx != null && (triggerIdx < currentRightBar - 2 || cand.zInfo.isDeadFallback))
          ? computeTouchAtEntryBar(bars, cand.barX1, triggerIdx, cand.bottom, cand.top, cand.zInfo.isDeadFallback ? 5 : cand.zInfo.maxTouchNum)
          : cand.zInfo.touch;

        let againstVal = cand.zInfo.against;
        if (cand.zInfo.isDeadFallback && mvPlotsObj && mvMetaIdx?.forbidIdx && triggerIdx != null) {
          const mvRow = getPlotRowAtBar(mvPlotsObj, bars, triggerIdx, mvIdxOffset);
          const forb = Number(mvRow?.[mvMetaIdx.forbidIdx] || 0);
          if ((isLong && forb === 1) || (!isLong && forb === -1)) againstVal = 'да';
        }
        return {
          zoneType: cand.zInfo.zoneType,
          fuel: cand.zInfo.fuel,
          touch: touchAtEntry,
          against: againstVal,
          rawText: cand.zInfo.rawText
        };
      };

      const seenZoneTypes = new Set();
      for (const cand of zoneCandidates) {
        const resolved = resolveZoneCandidateObj(cand);
        if (!seenZoneTypes.has(resolved.zoneType)) {
          seenZoneTypes.add(resolved.zoneType);
          out.zoneOptions.push(resolved);
          if (out.zoneOptions.length >= 3) break;
        }
      }
      out.matchedZone = out.zoneOptions[0] || null;
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
      const strength = SWEEP_STRENGTH[lqItem.sweptType] || 10;
      sweepCandidates.push({ priority: 0, strength, barDiff, priceDiff: 0, sType: lqItem.sweptType, price: swPrice, rawText: lqItem.rawText });
    }

    for (const l of lines) {
      const ly = Number(l.y1 ?? 0);
      if (!ly || Math.abs(ly - Number(l.y2 ?? 0)) > 0.001) continue;

      const lRgb = Number(l.ci || 0) & 0xFFFFFF;
      const lAlpha = (Number(l.ci || 0) >>> 24) & 0xFF;
      const isSweptBait = l.st === 'dot' && l.w === 1 && (l.ci === 1505670578 || l.ci === 1027311288 || ((lRgb === 12498354 || lRgb === 3901112) && lAlpha < 120));
      const isSweptShelf = l.st === 'dot' && l.w === 2 && (l.ci === 1497708274 || l.ci === 1501665544 || ((lRgb === 4536050 || lRgb === 8493320) && lAlpha < 120));
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
      const strength = SWEEP_STRENGTH[sType] || 10;
      sweepCandidates.push({ priority: 1, strength, barDiff, priceDiff, sType, price: ly, rawText: `${sType} ${ly}` });
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
      const strength = SWEEP_STRENGTH[lvl.sweptType] || 10;
      sweepCandidates.push({ priority: 1, strength, barDiff, priceDiff, sType: lvl.sweptType, price: lvl.price, rawText: lvl.rawText });
    }

    if (plotsObj && triggerIdx != null) {
      for (let bi = triggerIdx; bi >= triggerIdx - 5; bi--) {
        const pr = getPlotRowAtBar(plotsObj, bars, bi, plotIdxOffset);
        if (!pr) continue;
        const prevPr = getPlotRowAtBar(plotsObj, bars, bi - 1, plotIdxOffset) || pr;
        const barRow = getPlotRow(bars, bi);
        const barHi = Number(barRow?.[2] ?? sweepHigh);
        const barLo = Number(barRow?.[3] ?? sweepLow);

        const h4Col = isLong ? ziIdx.h4Min : ziIdx.h4Max;
        const d1Col = isLong ? ziIdx.dMin : ziIdx.dMax;
        const h4Lvl = Number(prevPr[h4Col] || pr[h4Col] || 0);
        const d1Lvl = Number(prevPr[d1Col] || pr[d1Col] || 0);
        const directHitD1 = d1Lvl > 0 && (isLong ? (barLo <= d1Lvl + tol && entryPrice >= d1Lvl - tol) : (barHi >= d1Lvl - tol && entryPrice <= d1Lvl + tol));
        const directHitH4 = h4Lvl > 0 && (isLong ? (barLo <= h4Lvl + tol && entryPrice >= h4Lvl - tol) : (barHi >= h4Lvl - tol && entryPrice <= h4Lvl + tol));

        const sweptHtf = (isLong ? (pr[ziIdx.swHtfLow] === 1) : (pr[ziIdx.swHtfHigh] === 1)) || (pr[ziIdx.swWeek] === 1) || (pr[ziIdx.swMonth] === 1) || directHitD1 || directHitH4;
        const swept15m = (pr[ziIdx.sw15m] === 1) || (isLong ? (pr[ziIdx.swLow] === 1) : (pr[ziIdx.swHigh] === 1));

        if (sweptHtf) {
          let sType = 'уровень 1Ч';
          let swPrice = Number(entryExtreme !== '' ? entryExtreme : entryPrice);

          if (directHitD1 && directHitH4 && Math.abs(d1Lvl - h4Lvl) < 0.1) { sType = 'склейка Д · 4Ч'; swPrice = d1Lvl; }
          else if (directHitD1 || pr[ziIdx.swWeek] === 1 || pr[ziIdx.swMonth] === 1) { sType = 'уровень Д'; if (directHitD1) swPrice = d1Lvl; }
          else if (directHitH4) { sType = 'уровень 4Ч'; swPrice = h4Lvl; }
          sweepCandidates.push({ priority: 0, strength: SWEEP_STRENGTH[sType] || 50, barDiff: triggerIdx - bi, priceDiff: 0, sType, price: swPrice, rawText: `${sType} ${swPrice}` });
        }
        if (swept15m) {
          const swPrice = Number(entryExtreme !== '' ? entryExtreme : entryPrice);
          sweepCandidates.push({ priority: 2, strength: SWEEP_STRENGTH['уровень 15м'], barDiff: triggerIdx - bi, priceDiff: 0, sType: 'уровень 15м', price: swPrice, rawText: `уровень 15м ${swPrice}` });
        }
      }
    }

    if (sweepCandidates.length > 0) {
      
      sweepCandidates.sort((a, b) => (a.priority - b.priority) || (b.strength - a.strength) || (a.barDiff - b.barDiff) || (a.priceDiff - b.priceDiff));
      const seenSweepTypes = new Set();
      for (const sc of sweepCandidates) {
        if (!seenSweepTypes.has(sc.sType)) {
          seenSweepTypes.add(sc.sType);
          out.sweepOptions.push({ sweptType: sc.sType, price: Number(sc.price), rawText: sc.rawText });
          if (out.sweepOptions.length >= 4) break;
        }
      }
      out.matchedSweep = out.sweepOptions[0] || null;
    }

    return out;
  }

  function mapMvModelToFormM1(modelCode, isContinuation) {
    if (isContinuation || modelCode === 18) return 'Правильный откат';
    if (modelCode === 12) return 'Поглощение';
    if (modelCode === 5) return 'Пин-бар (IFS)';
    if (modelCode === 2) return 'Свип';
    return 'SCOB';
  }

  function readMatryoshkaIndicators(model, symKey, rawInterval, entryTime, triggerIdx, startIdx, endIdx, isLong, entryPrice, stopPrice, entryExtreme, bars) {
    const result = {
      flows: null, flowsHistorical: false, flowsMissing: [], chartFlowDir: null,
      candleColor: null, candleMatched: false, modelCode: 0, isContinuation: false,
      matchedZone: null, zoneOptions: [],
      matchedSweep: null, sweepOptions: [],
      detectedStructure: null, health: null, isEmptyData: true
    };
    if (!model || typeof model.dataSources !== 'function') return result;

    try {
      syncGraphicsMemoryO1(model, symKey);
      const sources = model.dataSources();
      let ziStudy = null, mvStudy = null;

      for (const s of sources) {
        const meta = typeof s.metaInfo === 'function' ? s.metaInfo() : s._metaInfo;
        const title = meta?.description || meta?.name || '';
        if (!ziStudy && isMatryoshkaZiTitle(title)) ziStudy = s;
        if (!mvStudy && /модели входа/i.test(title) && !/price action/i.test(title)) mvStudy = s;
      }

      result.health = inspectIndicatorsHealth(ziStudy, mvStudy);

      let mvPlotsObj = null, mvIdxOffset = 0, mvMetaIdx = null;
      if (mvStudy) {
        mvPlotsObj = typeof mvStudy.plots === 'function' ? mvStudy.plots() : mvStudy._data?.bars?.();
        const foundBar = findPlotRowByIndexOrTime(mvPlotsObj, entryTime, triggerIdx);
        if (foundBar && mvPlotsObj) {
          mvMetaIdx = resolveMvPlotIndices(mvStudy, foundBar.val);
          mvIdxOffset = triggerIdx != null ? (foundBar.index - triggerIdx) : 0;
          const { colorIdx, modelIdx, contLongIdx, contShortIdx } = mvMetaIdx;
          const baseIdx = foundBar.index;

          const candidates = [baseIdx, baseIdx - 1, baseIdx + 1, baseIdx - 2, baseIdx + 2];
          let matchedCode = 0, matchedModel = 0, isCont = false, fallbackCode = 0;

          for (const cIdx of candidates) {
            const r = getPlotRow(mvPlotsObj, cIdx);
            if (!r) continue;
            let code = Number(r[colorIdx] ?? 0);
            let mCode = Number(r[modelIdx] ?? 0);
            if ((isNaN(code) || code === 0) && colorIdx !== 1 && r[1] != null && Math.abs(Number(r[1])) <= 3) {
              code = Number(r[1] ?? 0);
              mCode = Number(r[2] ?? 0);
            }
            if (isNaN(code) || code === 0) continue;

            if (!fallbackCode) fallbackCode = code;
            if ((isLong && code > 0 && code <= 3) || (!isLong && code < 0 && code >= -3)) {
              matchedCode = code;
              matchedModel = mCode;
              isCont = (mCode === 18) || (isLong ? (r?.[contLongIdx] === 1) : (r?.[contShortIdx] === 1));
              break;
            }
          }

          const finalCode = matchedCode || fallbackCode;
          const colorMap = { '3': 'Лонг 3', '2': 'Лонг 2', '1': 'Лонг 1', '-1': 'Шорт 1', '-2': 'Шорт 2', '-3': 'Шорт 3' };
          if (colorMap[String(finalCode)]) {
            result.candleColor = colorMap[String(finalCode)];
            result.candleMatched = matchedCode !== 0;
            result.modelCode = matchedModel;
            result.isContinuation = isCont;
          }
        }
      }

      if (ziStudy) {
        const g = typeof ziStudy.graphics === 'function' ? ziStudy.graphics() : ziStudy._graphics;
        const plotsObj = typeof ziStudy.plots === 'function' ? ziStudy.plots() : ziStudy._data?.bars?.();
        const ziIdx = resolveZiPlotIndices(ziStudy);
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
          let chartFlowDir = null;
          for (const rObj of Object.values(rowsByTid)) {
            const label = String(rObj[0] || '').trim();
            const sym = String(rObj[1] || '').trim();
            if (!label || !sym) continue;
            const dirWord = sym.includes('↑') ? 'вверх' : (sym.includes('↓') ? 'вниз' : '—');
            if (label === 'Д' || label === 'D1' || label === 'D') flows.d = dirWord;
            else if (label === '4Ч' || label === 'H4' || label === '4H') flows.h4 = dirWord;
            else if (label === '1Ч' || label === 'H1' || label === '1H') flows.h1 = dirWord;
            else if (label === '15м' || label === 'M15' || label === '15m') flows.m15 = dirWord;
            else if (label.toLowerCase() === 'график') chartFlowDir = dirWord;
          }

          
          const ivStr = String(rawInterval || '').toUpperCase().trim();
          if (chartFlowDir) {
            if ((ivStr === '15' || ivStr === '15M') && !flows.m15) flows.m15 = chartFlowDir;
            else if ((ivStr === '60' || ivStr === '1H') && !flows.h1) flows.h1 = chartFlowDir;
            else if ((ivStr === '240' || ivStr === '4H') && !flows.h4) flows.h4 = chartFlowDir;
            else if ((ivStr === 'D' || ivStr === '1D') && !flows.d) flows.d = chartFlowDir;
          }

          if (Object.keys(flows).length > 0 || chartFlowDir) {
            const missing = [];
            if (!flows.d) missing.push('Д');
            if (!flows.h4) missing.push('4Ч');
            if (!flows.h1) missing.push('1Ч');
            if (!flows.m15) missing.push('15м');
            result.flows = flows;
            result.flowsMissing = missing;
            result.chartFlowDir = chartFlowDir;
          }
        }

        if (g && entryPrice && stopPrice) {
          const ziPlotBar = findPlotRowByIndexOrTime(plotsObj, entryTime, triggerIdx);
          const plotIdxOffset = (ziPlotBar && triggerIdx != null) ? (ziPlotBar.index - triggerIdx) : 0;
          const direct = analyzeChartGraphicsO1(g, plotsObj, plotIdxOffset, ziIdx, mvPlotsObj, mvIdxOffset, mvMetaIdx, symKey, triggerIdx, isLong, entryPrice, stopPrice, entryExtreme, bars);
          result.matchedZone = direct.matchedZone;
          result.zoneOptions = direct.zoneOptions || [];
          result.matchedSweep = direct.matchedSweep;
          result.sweepOptions = direct.sweepOptions || [];
          result.detectedStructure = direct.detectedStructure;
          result.flowsHistorical = direct.isHistorical;
        }
      }

      result.isEmptyData = !(result.flows || result.matchedZone || result.candleColor || result.detectedStructure || result.matchedSweep);
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

  
  
  
  function detectChartSafeLevel(model, bars, mainToolId, startIdx, endIdx, triggerIdx, isLong, entryPrice, stopPrice, targetPrice, decimals) {
    if (!model || !entryPrice || !targetPrice || entryPrice === targetPrice) return null;
    try {
      const tools = typeof model.allLineTools === 'function' ? model.allLineTools() : model.allLineTools;
      const list = Array.isArray(tools) ? tools : Object.values(tools || {});
      const minProf = Math.min(entryPrice, targetPrice);
      const maxProf = Math.max(entryPrice, targetPrice);
      const profRange = maxProf - minProf;
      const margin = profRange * 0.03;

      const posLeftBar = Math.min(startIdx ?? 0, endIdx ?? 0);
      const posRightBar = Math.max(startIdx ?? 0, endIdx ?? 0);

      const candidates = [];
      for (const t of list) {
        const tid = getToolId(t);
        if (!t || tid === mainToolId) continue;
        const name = getToolName(t);
        if (name.includes('RiskReward')) continue;

        let pts = [];
        try {
          const rawPts = typeof t.points === 'function' ? t.points() : (t._points || t.points || []);
          if (Array.isArray(rawPts)) pts = rawPts;
        } catch (e) {}
        if (!pts.length) continue;

        const p0 = Number(pts[0]?.price ?? 0);
        const p1 = Number(pts[1]?.price ?? p0);
        if (!p0 || isNaN(p0)) continue;

        const isHorizontal = pts.length === 1 || Math.abs(p0 - p1) <= Math.max(0.05, profRange * 0.025);
        if (!isHorizontal) continue;

        attachToolSubscriptions(t);

        if (p0 <= minProf + margin || p0 >= maxProf - margin) continue;

        const b0 = pts[0]?.index ?? null;
        const b1 = pts[1]?.index ?? b0;
        if (b0 == null) continue;

        const lineLeft = Math.min(b0, b1 ?? b0);
        const lineRight = Math.max(b0, b1 ?? b0);
        const lineMid = (lineLeft + lineRight) / 2;

        const rawProps = typeof t.properties === 'function' ? t.properties() : t._properties;
        const ch = typeof rawProps?.childs === 'function' ? rawProps.childs() : rawProps?._childs;
        const txt = String(
          readWatched(ch?.text) || readWatched(ch?.title) || readWatched(rawProps?.text) || ''
        ).trim();
        const hasSafeWord = /сейф|safe|сф|1\/2|50%|фикс|частич/i.test(txt);

        
        
        const overlapsPos = (lineRight >= posLeftBar + 1) && (lineLeft <= posRightBar);
        
        const isStrictlyInsidePos = overlapsPos && (lineMid >= posLeftBar) && (lineMid <= posRightBar + 1) && (lineRight - lineLeft <= Math.max(80, (posRightBar - posLeftBar) * 2));

        if (!hasSafeWord && !isStrictlyInsidePos) continue;
        if (hasSafeWord && !overlapsPos) continue;

        let rawSafeP = p0;
        
        if (pts.length === 1) {
          const numInTxt = txt.match(/(\d{3,6}[.,]\d+)/);
          if (numInTxt) {
            const parsedNum = parseFloat(numInTxt[1].replace(',', '.'));
            if (!isNaN(parsedNum) && parsedNum > minProf && parsedNum < maxProf) rawSafeP = parsedNum;
          } else {
            const vAlign = String(readWatched(ch?.vertAlign) || 'top').toLowerCase();
            const fSize = Number(readWatched(ch?.fontSize) || 14);
            const halfBoxPx = Math.round(fSize * 1.05);
            try {
              const ms = typeof model.mainSeries === 'function' ? model.mainSeries() : model.mainSeries;
              const ps = typeof ms?.priceScale === 'function' ? ms.priceScale() : ms?._priceScale;
              const fv = typeof ms?.firstValue === 'function' ? ms.firstValue() : null;
              if (ps && typeof ps.priceToCoordinate === 'function' && typeof ps.coordinateToPrice === 'function' && fv != null) {
                const yCoord = ps.priceToCoordinate(p0, fv);
                const centerCoord = vAlign === 'top' ? (yCoord + halfBoxPx) : (vAlign === 'bottom' ? (yCoord - halfBoxPx) : yCoord);
                const centeredP = ps.coordinateToPrice(centerCoord, fv);
                if (centeredP && !isNaN(centeredP) && centeredP > minProf && centeredP < maxProf) rawSafeP = centeredP;
              }
            } catch (e) {}
          }
        }

        const safePrice = Number(rawSafeP.toFixed(decimals));
        const risk = Math.max(0.01, Math.abs(entryPrice - stopPrice));
        const safeRR = Number((Math.abs(safePrice - entryPrice) / risk).toFixed(2));
        if (safeRR < 0.25) continue;

        candidates.push({
          priority: hasSafeWord ? 0 : 1,
          price: safePrice,
          rr: safeRR,
          label: txt || `Сейф ${safeRR}R`
        });
      }

      if (!candidates.length) return null;
      
      candidates.sort((a, b) => (a.priority - b.priority) || (a.rr - b.rr));
      const uniqueByPrice = [];
      for (const c of candidates) {
        if (!uniqueByPrice.some(u => Math.abs(u.price - c.price) < Math.max(0.05, profRange * 0.03))) {
          uniqueByPrice.push(c);
        }
      }

      const primary = uniqueByPrice[0];
      const secondary = uniqueByPrice.length > 1 ? uniqueByPrice[1] : null;

      
      let reachedSafe1 = true;
      let reachedSafe2 = !!secondary;
      let remOutcome = 'цель';

      if (bars && triggerIdx != null) {
        reachedSafe1 = false;
        reachedSafe2 = false;
        let hitStopBeforeSafe1 = false;
        const scanEnd = Math.max(triggerIdx, posRightBar);

        for (let i = triggerIdx; i <= scanEnd; i++) {
          const b = getPlotRow(bars, i);
          if (!b) continue;
          const hi = Number(b[2]), lo = Number(b[3]), cl = Number(b[4]);
          const cHi = (i === triggerIdx) ? Math.max(entryPrice, cl) : hi;
          const cLo = (i === triggerIdx) ? Math.min(entryPrice, cl) : lo;

          if (!reachedSafe1) {
            const hitSt = isLong ? (cLo <= stopPrice) : (cHi >= stopPrice);
            const hitSf1 = isLong ? (cHi >= primary.price) : (cLo <= primary.price);
            if (hitSt && !hitSf1) { hitStopBeforeSafe1 = true; break; }
            if (hitSf1) {
              reachedSafe1 = true;
              if (secondary && (isLong ? (cHi >= secondary.price) : (cLo <= secondary.price))) {
                reachedSafe2 = true;
              }
            }
          } else {
            if (secondary && !reachedSafe2) {
              const hitSf2 = isLong ? (cHi >= secondary.price) : (cLo <= secondary.price);
              if (hitSf2) reachedSafe2 = true;
            }
            const hitTg = isLong ? (cHi >= targetPrice) : (cLo <= targetPrice);
            const hitSt = isLong ? (cLo <= stopPrice) : (cHi >= stopPrice);
            if (hitTg) { remOutcome = 'цель'; break; }
            if (hitSt) { remOutcome = 'стоп'; break; }
            remOutcome = 'безубыток';
          }
        }
        if (hitStopBeforeSafe1) reachedSafe1 = false;
      }

      return {
        ...primary,
        reachedSafe: reachedSafe1,
        remOutcome,
        secondSafe: (secondary && reachedSafe2) ? { price: secondary.price, rr: secondary.rr, reachedSafe: true } : null
      };
    } catch (e) { return null; }
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
        if (!/матрёшка|матрешка/i.test(title)) continue;

        hookedStudies.add(s);
        subscribeTvDelegate(s._dataUpdated, () => {
          syncGraphicsMemoryO1(model, symKey);
          scheduleLiveSync(50);
        });
        syncGraphicsMemoryO1(model, symKey);
      }

      const tools = typeof model.allLineTools === 'function' ? model.allLineTools() : model.allLineTools;
      const list = Array.isArray(tools) ? tools : Object.values(tools || {});
      for (const t of list) attachToolSubscriptions(t);
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

  const lastKnownFlowsBySym = {};

  function formatTf(rawInterval, style) {
    const str = String(rawInterval || '1').toUpperCase().trim();
    if (style === 'cyr') {
      if (str === '1') return 'М1';
      if (str === '3') return '3м';
      if (str === '5') return '5м';
      if (str === '15') return '15м';
      if (str === '30') return '30м';
      if (str === '45') return '45м';
      if (str === '60' || str === '1H') return '1Ч';
      if (str === '240' || str === '4H') return '4Ч';
      if (str === 'D' || str === '1D') return 'Д';
      if (/^\d+$/.test(str)) {
        const m = parseInt(str, 10);
        return (m >= 60 && m % 60 === 0) ? (m / 60) + 'Ч' : m + 'м';
      }
      return str;
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
    const pt0 = points?.[0] || tool?._points?.[0];
    const rawPtTime = normalizeUnixSeconds(pt0?.time_t || pt0?.time?.getTime?.() || pt0?.time);
    if (rawPtTime) return rawPtTime;

    try {
      const ts = typeof model.timeScale === 'function' ? model.timeScale() : model._timeScale;
      if (ts && barIdx != null && typeof ts.indexToTimePoint === 'function') {
        const tp = normalizeUnixSeconds(ts.indexToTimePoint(barIdx));
        if (tp) return tp;
      }
    } catch (e) {}

    return Math.floor(Date.now() / 1000);
  }

  function analyzeBars(bars, startIndex, endIndex, isLong, entryPrice, stopPrice, targetPrice, points) {
    const risk = Math.abs(entryPrice - stopPrice);
    const price1R = isLong ? entryPrice + risk : entryPrice - risk;
    const price3R = isLong ? entryPrice + (risk * 3) : entryPrice - (risk * 3);

    if (!bars || startIndex == null) {
      return { outcome: 'Тейк', reached1R: true, reached3R: true, entryCandleExtreme: null, triggerIdx: startIndex ?? 0, barsLoaded: false };
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
    let validBarsCount = 0;

    for (let i = triggerIdx; i <= to; i++) {
      const bar = getPlotRow(bars, i);
      if (!bar) continue;
      const high = Number(bar[2]), low = Number(bar[3]), close = Number(bar[4]);
      if (isNaN(high) || isNaN(low)) continue;
      validBarsCount++;

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

    return { outcome, reached1R, reached3R, entryCandleExtreme: sweepExtreme, triggerIdx, barsLoaded: validBarsCount > 0 };
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
    let barAnalysis = { outcome: 'Тейк', reached1R: true, reached3R: true, entryCandleExtreme: null, triggerIdx: startIdx, barsLoaded: true };

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
    const rawInterval = (typeof ms?.interval === 'function' ? ms.interval() : ms?._interval) || getPropVal('interval') || '1';
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

    const ind = readMatryoshkaIndicators(model, symKey, rawInterval, time, effectiveIdx, startIdx, endIdx, isLong, entryR, stopR, extremeR, bars);
    if (!barAnalysis.barsLoaded) ind.barsOffscreen = true;

    if (ind.matchedSweep && ind.matchedSweep.price != null && !isNaN(Number(ind.matchedSweep.price))) {
      ind.matchedSweep.price = Number(Number(ind.matchedSweep.price).toFixed(decimals));
    }
    if (Array.isArray(ind.sweepOptions)) {
      for (const swOpt of ind.sweepOptions) {
        if (swOpt.price != null && !isNaN(Number(swOpt.price))) {
          swOpt.price = Number(Number(swOpt.price).toFixed(decimals));
        }
      }
    }

    const chartSafe = isRect ? null : detectChartSafeLevel(model, bars, id, startIdx, endIdx, effectiveIdx, isLong, entryR, stopR, targetR, decimals);

    const lastBarIdx = typeof bars?.lastIndex === 'function' ? bars.lastIndex() : null;
    const tradeRightEdge = Math.max(effectiveIdx ?? 0, endIdx ?? 0);
    ind.flowsHistorical = (lastBarIdx != null && lastBarIdx > 0 && tradeRightEdge < lastBarIdx - 90);

    const safeSig = chartSafe ? `${chartSafe.price}_${chartSafe.secondSafe?.price || ''}_${chartSafe.remOutcome}_${chartSafe.reachedSafe ? 1 : 0}_${chartSafe.secondSafe?.reachedSafe ? 1 : 0}` : 'nosafe';
    const flSig = ind.flows ? `${ind.flows.d || ''}_${ind.flows.h4 || ''}_${ind.flows.h1 || ''}_${ind.flows.m15 || ''}_${ind.chartFlowDir || ''}` : 'nofl';
    const signature = `${id}_${rawInterval}_${barAnalysis.barsLoaded ? 1 : 0}_${barAnalysis.outcome}_${startIdx}_${endIdx}_${entryR}_${stopR}_${targetR}_${safeSig}_${flSig}`;

    return {
      id, signature, serverTime: Number(tool._serverUpdateTime || 0), time, decimals,
      barsLoaded: barAnalysis.barsLoaded,
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
      autoResult: barAnalysis.outcome, reached1R: barAnalysis.reached1R, reached3R: barAnalysis.reached3R,
      chartSafe, ind
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
        const sym = currentTrade.symbol;
        const fl = currentTrade.ind.flows;
        const chartFl = currentTrade.ind.chartFlowDir;
        if (!lastKnownFlowsBySym[sym]) lastKnownFlowsBySym[sym] = {};
        const cache = lastKnownFlowsBySym[sym];

        if (fl.d && fl.d !== '—') cache.d = fl.d;
        if (fl.h4 && fl.h4 !== '—') cache.h4 = fl.h4;
        if (fl.h1 && fl.h1 !== '—') cache.h1 = fl.h1;
        if (fl.m15 && fl.m15 !== '—') cache.m15 = fl.m15;

        const fallbackDir = chartFl || fl.h1 || (currentTrade.direction === 'LONG' ? 'вверх' : 'вниз');
        const pickFlow = (directVal, cachedVal, curStateVal) => {
          if (directVal) return directVal;
          if (preferChartSelection || curStateVal === '—' || !curStateVal) return cachedVal || fallbackDir;
          return curStateVal;
        };

        saveState({
          flowD: pickFlow(fl.d, cache.d, state.flowD),
          flow4H: pickFlow(fl.h4, cache.h4, state.flow4H),
          flow1H: pickFlow(fl.h1, cache.h1, state.flow1H),
          flow15M: pickFlow(fl.m15, cache.m15, state.flow15M)
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
    return [...(PRESET_SCHEMAS[state.presetKey] || PRESET_SCHEMAS.bt_30).columns];
  }
  function getActivePasteCol() {
    return state.presetKey === 'custom' ? (state.customPasteCol || 'B') : (PRESET_SCHEMAS[state.presetKey]?.pasteCol || 'B');
  }
  function getActiveStartRow() {
    return state.presetKey === 'custom' ? (Number(state.customStartRow) || 2) : (PRESET_SCHEMAS[state.presetKey]?.startRow || 4);
  }
  function getActiveSession() {
    if (manualSessionOverride) return manualSessionOverride;
    if (!currentTrade) return 'Европа';
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

  function buildDynamicFormula(tok, cols, pasteCol, startRow, fallbackVal, lotLegSpec = null) {
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
      if (!rsk || state.presetKey !== 'bt_30') return fallbackVal;
      const fullLotExpr = `${INT}(('Параметры'!$B$5/${rsk}/'Параметры'!$B$6)/'Параметры'!$B$7)*'Параметры'!$B$7`;
      if (lotLegSpec && (lotLegSpec.leg === 'safe' || lotLegSpec.leg === 'safe2')) {
        const shStr = numStr(lotLegSpec.share || 0.5);
        const safeLotExpr = `${INT}(((${fullLotExpr})*${shStr})/'Параметры'!$B$7)*'Параметры'!$B$7`;
        return `=${IF}(${OR}(${rsk}=""${s}${rsk}=0)${s}""${s}${safeLotExpr})`;
      }
      if (lotLegSpec && lotLegSpec.leg === 'rem') {
        const sh1Str = numStr(lotLegSpec.share1 || lotLegSpec.share || 0.5);
        const safe1Expr = `${INT}(((${fullLotExpr})*${sh1Str})/'Параметры'!$B$7)*'Параметры'!$B$7`;
        if (lotLegSpec.isTwoSafes) {
          const sh2Str = numStr(lotLegSpec.share2 || 0.33);
          const safe2Expr = `${INT}(((${fullLotExpr})*${sh2Str})/'Параметры'!$B$7)*'Параметры'!$B$7`;
          return `=${IF}(${OR}(${rsk}=""${s}${rsk}=0)${s}""${s}(${fullLotExpr})-(${safe1Expr})-(${safe2Expr}))`;
        }
        return `=${IF}(${OR}(${rsk}=""${s}${rsk}=0)${s}""${s}(${fullLotExpr})-(${safe1Expr}))`;
      }
      return `=${IF}(${OR}(${rsk}=""${s}${rsk}=0)${s}""${s}${fullLotExpr})`;
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

  function computeSafeMetricsForTrade(t) {
    if (!t || !activeSafeState.enabled) return null;
    const rawP = String(activeSafeState.price ?? '').replace(',', '.').trim();
    const safePrice = parseFloat(rawP);
    if (!safePrice || isNaN(safePrice) || safePrice <= 0) return null;

    const risk = Math.max(0.01, t.riskPts);
    const dec = t.decimals || 2;
    const safePotPts = Number(Math.abs(safePrice - t.entryPrice).toFixed(dec));
    const safeRR = Number((safePotPts / risk).toFixed(2));

    const rawP2 = String(activeSafeState.price2 ?? '').replace(',', '.').trim();
    const safePrice2 = parseFloat(rawP2);
    const hasSecondSafe = !isNaN(safePrice2) && safePrice2 > 0 && Math.abs(safePrice2 - safePrice) > 0.01;
    const safe2PotPts = hasSecondSafe ? Number(Math.abs(safePrice2 - t.entryPrice).toFixed(dec)) : 0;
    const safe2RR = hasSecondSafe ? Number((safe2PotPts / risk).toFixed(2)) : 0;

    
    
    
    const rawShare = Number(activeSafeState.share || 0.5);
    const share1 = hasSecondSafe ? (rawShare > 0.4 ? 0.5 : (1 / 3)) : rawShare;
    const share2 = hasSecondSafe ? (rawShare > 0.4 ? 0.25 : (1 / 3)) : 0;
    const remShare = Math.max(0, 1 - share1 - share2);

    let remOut = activeSafeState.remOutcome;
    if (remOut === 'auto') {
      if (t.chartSafe && Math.abs(t.chartSafe.price - safePrice) < 0.05) {
        remOut = t.chartSafe.remOutcome || (t.autoResult === 'Тейк' ? 'цель' : (t.autoResult === 'Стоп' ? 'стоп' : 'безубыток'));
      } else {
        remOut = (t.autoResult === 'Тейк') ? 'цель' : (t.autoResult === 'Стоп' ? 'стоп' : 'безубыток');
      }
    }

    const remRR = remOut === 'цель' ? t.riskReward : (remOut === 'стоп' ? -1 : 0);
    const safeWeightedR = Number((share1 * safeRR).toFixed(2));
    const safe2WeightedR = hasSecondSafe ? Number((share2 * safe2RR).toFixed(2)) : 0;
    const remWeightedR = Number((remShare * remRR).toFixed(2));
    const blendedR = Number((safeWeightedR + safe2WeightedR + remWeightedR).toFixed(2));

    return {
      safePrice: Number(safePrice.toFixed(dec)),
      safePotPts,
      safeRR,
      hasSecondSafe,
      safePrice2: hasSecondSafe ? Number(safePrice2.toFixed(dec)) : null,
      safe2PotPts,
      safe2RR,
      safe2WeightedR,
      share: Number(share1.toFixed(2)),
      share1: Number(share1.toFixed(2)),
      share2: Number(share2.toFixed(2)),
      remShare: Number(remShare.toFixed(2)),
      remOut,
      remRR,
      safeWeightedR,
      remWeightedR,
      blendedR,
      exportMode: activeSafeState.exportMode || '1row'
    };
  }

  function collectCurrentSnapshot() {
    if (!currentTrade) return null;
    const cols = getActiveColumns();
    const t = currentTrade;
    const isLong = t.direction === 'LONG';
    const hasRes30 = cols.includes('res_30');
    const hasResOff = cols.includes('res_off');

    const safeInfo = computeSafeMetricsForTrade(t);

    const defOffRes = t.autoResult === 'Безубыток' ? 'Сейф/БУ' : t.autoResult;
    const resStd = getChipVal('f-res', hasResOff ? defOffRes : t.autoResult);
    let res30Val = getChipVal('f-res30', t.autoResult === 'Тейк' ? 'цель' : (t.autoResult === 'Стоп' ? 'стоп' : 'безубыток'));

    const manualRStr = cleanText(document.getElementById('f-manual-r30')?.value || '');
    let manualRNum = manualRStr ? parseFloat(manualRStr.replace(',', '.')) : NaN;

    if (safeInfo && hasRes30 && safeInfo.exportMode === '1row') {
      res30Val = 'частичная фиксация';
      if (isNaN(manualRNum)) manualRNum = safeInfo.blendedR;
    }

    const activeResult = hasRes30 ? res30Val : resStd;

    const bucket = getTplBucket();
    const existingSnap = bucket.batchQueue.find(item => item.tradeId === t.id || item.signature === t.signature);
    const baseCumR = existingSnap ? (existingSnap.cumRBefore ?? 0) : bucket.cumR;

    let rOut = '', newCum = baseCumR, newMax = bucket.maxCumR;
    if (safeInfo && hasRes30 && safeInfo.exportMode === '2rows') {
      rOut = safeInfo.blendedR;
      newCum = Number((baseCumR + rOut).toFixed(2));
    } else if (activeResult === 'Тейк' || activeResult === 'цель') {
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
      aoi_res: getChipVal('f-aoi-res', 'Отбой'),
      safeInfo,
      safeStateCopy: { ...activeSafeState }
    };
  }

  function buildSafeAutoComment(si) {
    if (!si) return '';
    if (si.hasSecondSafe) {
      const is5025 = Math.abs(si.share1 - 0.5) < 0.03;
      const s1Lbl = is5025 ? '50%' : '1/3';
      const s2Lbl = is5025 ? '25%' : '1/3';
      return `Сейф ${s1Lbl} на ${numStr(si.safePrice)} (+${numStr(si.safeRR)}R), ${s2Lbl} на ${numStr(si.safePrice2)} (+${numStr(si.safe2RR)}R)`;
    }
    const sLbl = Math.abs(si.share1 - 0.5) < 0.03 ? '1/2' : '1/3';
    return `Сейф ${sLbl} на ${numStr(si.safePrice)} (+${numStr(si.safeRR)}R)`;
  }

  function renderSingleRowFromSnap(snap) {
    const cols = getActiveColumns();
    const pasteCol = getActivePasteCol();
    const startRow = getActiveStartRow();
    const schemaStyle = PRESET_SCHEMAS[state.presetKey] || PRESET_SCHEMAS.bt_30;
    const f = (tok, fallback) => buildDynamicFormula(tok, cols, pasteCol, startRow, fallback, null);
    const sweptLvlFormatted = (snap.sweptVal === 'ничего не снято' || !snap.sweptLvlRaw) ? '' : numStr(snap.sweptLvlRaw);

    const targetPrice = snap.targetPrice;
    const potPts = snap.potPts;
    const rrVal = snap.riskReward;
    const res30Val = snap.res30Val;
    const resStdVal = snap.resStd;
    const manualRStr = (snap.res30Val === 'частичная фиксация' || snap.res30Val === 'закрыл руками') && !isNaN(snap.manualRNum) ? numStr(snap.manualRNum) : '';
    const rOutVal = snap.rOut;
    const lots30Val = snap.lots30;

    let commentVal = snap.comment || '';
    if (snap.safeInfo) {
      const safeNote = buildSafeAutoComment(snap.safeInfo);
      if (!commentVal) commentVal = safeNote;
      else if (!/сейф/i.test(commentVal)) commentVal = `${safeNote} · ${commentVal}`;
    }

    const stopUsd = Math.max(0.01, snap.riskPts);
    const profitUsdVal = rOutVal === '' ? '' : Number((Number(rOutVal) * lots30Val * 100 * stopUsd).toFixed(2));

    const tokenValues = {
      symbol: snap.symbol, date: snap.candleDate, time_msk: snap.mskTime, time_slot: snap.officeTimeSlot,
      session: snap.session, trend: snap.trend, trend_h4: snap.trend_h4,
      tf: schemaStyle.tfStyle === 'cyr' ? snap.timeframeCyr : snap.timeframeLat,
      dir_lower: snap.isLong ? 'лонг' : 'шорт', dir_title: snap.isLong ? 'Лонг' : 'Шорт',
      zone: snap.zone, zone_30: snap.zone_30, poi_type: snap.poi_type, form_m1: snap.form_m1,
      liq: snap.liq, react: snap.react, conf: snap.conf,
      entry: numStr(snap.entryPrice), stop: numStr(snap.stopPrice), target: numStr(targetPrice),
      upper_bound: numStr(snap.upperBound), lower_bound: numStr(snap.lowerBound),
      risk_pts: f('risk_pts', numStr(snap.riskPts)), pot_pts: f('pot_pts', numStr(potPts)),
      rr: f('rr', numStr(rrVal)), lots_30: f('lots_30', numStr(lots30Val)),
      lots_off: f('lots_off', numStr(snap.lotsOff)), result: resStdVal, res_30: res30Val,
      manual_r_30: manualRStr,
      res_off: resStdVal === 'Безубыток' ? 'Сейф/БУ' : resStdVal,
      r_out: f('r_out', rOutVal === '' ? '' : numStr(rOutVal)),
      r_out_off: f('r_out_off', numStr(snap.rOutOff)),
      profit_usd: f('profit_usd', profitUsdVal === '' ? '' : numStr(profitUsdVal)),
      prof_pct_off: f('prof_pct_off', numStr(Number((snap.rOutOff * 0.003).toFixed(4)))),
      move_pts: f('move_pts', numStr(resStdVal === 'Тейк' ? potPts : (resStdVal === 'Стоп' ? -snap.riskPts : 0))),
      cum_r: f('cum_r', resStdVal === 'Пропуск' ? numStr(snap.cumRBefore) : numStr(snap.newCum)),
      dd_r: f('dd_r', resStdVal === 'Пропуск' ? numStr(0) : numStr(snap.dd)),
      hit_1r: snap.hit_1r, hit_3r: snap.hit_3r, error: snap.errVal, error_poi: snap.errVal,
      comment: commentVal, comment_30: commentVal, screenshot: snap.img, aoi_num: snap.aoiNum,
      flow_d: snap.flD, flow_4h: snap.fl4H, flow_1h: snap.fl1H, flow_15m: snap.fl15M,
      tf_d1: snap.tfD1 ?? 'да', tf_h4: snap.tfH4 ?? 'да', tf_h1: snap.tfH1 ?? 'да', tf_m15: snap.tfM15 ?? 'да',
      struct_30: snap.struct_30, swept_30: snap.sweptVal, swept_level_30: sweptLvlFormatted,
      candle_col: snap.candle_col, against_flow: snap.against_flow, fuel_30: snap.fuel_30,
      touch_30: snap.touch_30, check_30: snap.chk30, tf_ok: snap.tfOk ?? 'да',
      emotion: snap.emotionVal || 'Спокоен',
      safe_pct: numStr(snap.safeInfo ? snap.safeInfo.share1 : 0.5),
      safe_rr: numStr(snap.safeInfo ? snap.safeInfo.safeRR : 3),
      risk_pct_col: numStr(0.003),
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

  function renderSnapshotToRow(snap) {
    return renderSingleRowFromSnap(snap);
  }

  const addTokenOptionsHtml = Object.entries(TOKEN_LABELS).map(([k, v]) => `<option value="${k}">+ Добавить поле: ${v}</option>`).join('');

  const uiHtml = `
  <div id="bt-toast" style="display:none;position:fixed;bottom:24px;right:24px;z-index:10000000;background:#141417;color:#f4f4f5;padding:11px 16px;border-radius:8px;border-left:4px solid #089981;box-shadow:0 12px 32px rgba(0,0,0,0.9);font-family:-apple-system,BlinkMacSystemFont,'Inter','Segoe UI',Roboto,sans-serif;font-size:12px;font-weight:500;max-width:390px;"></div>
  <button type="button" id="bt-fab-btn" style="position:fixed;bottom:18px;left:58px;z-index:9999990;display:none;align-items:center;gap:6px;background:#141417;color:#e5c158;border:1px solid rgba(229,193,88,0.45);border-radius:7px;padding:6px 12px;font-family:-apple-system,BlinkMacSystemFont,'Inter',sans-serif;font-size:12px;font-weight:600;cursor:pointer;box-shadow:0 4px 16px rgba(0,0,0,0.75);transition:0.15s;">
    ${LOGO_SVG}<span>Бэктест Хэлпер</span><span style="color:#9ca3af;font-weight:400;font-size:11px;">(${isMac ? '⌃Q' : 'F2'})</span>
  </button>
  <div id="bt-modal" class="bt-no-scroll" tabindex="-1" style="display:none;outline:none;position:fixed;top:${Math.max(48, state.posTop)}px;right:${Math.max(12, state.posRight)}px;z-index:9999999;background:#0d0d10;color:#f4f4f5;padding:12px;border-radius:12px;box-shadow:0 22px 55px rgba(0,0,0,0.92),0 0 0 1px rgba(229,193,88,0.28);font-family:-apple-system,BlinkMacSystemFont,'Inter','Segoe UI',Roboto,sans-serif;width:402px;max-height:calc(100vh - 64px);overflow-y:auto;box-sizing:border-box;user-select:none;">
    <div id="bt-header-drag" style="position:relative;display:flex;justify-content:space-between;align-items:center;padding-bottom:8px;margin-bottom:8px;border-bottom:1px solid #222228;cursor:move;">
      <div style="display:flex;align-items:center;gap:6px;min-width:0;">
        ${LOGO_SVG}
        <button type="button" id="bt-quick-tpl-btn" title="Нажмите, чтобы быстро сменить шаблон таблицы" style="background:#16161c;border:1px solid rgba(229,193,88,0.38);border-radius:6px;padding:3px 8px;display:flex;align-items:center;gap:5px;cursor:pointer;color:#f4f4f5;font-size:12px;font-weight:700;transition:0.12s;max-width:215px;">
          <span id="bt-active-tpl-title" style="white-space:nowrap;overflow:hidden;text-overflow:ellipsis;"></span>
          <span style="color:#e5c158;font-size:10px;">▾</span>
        </button>
      </div>
      <div id="bt-quick-tpl-menu" style="display:none;position:absolute;top:34px;left:0;width:270px;background:#18181f;border:1px solid #e5c158;border-radius:8px;box-shadow:0 14px 32px rgba(0,0,0,0.95);z-index:200;padding:4px;">
        <div style="font-size:9.5px;font-weight:700;color:#9ca3af;padding:4px 8px 3px;text-transform:uppercase;">Выберите вашу таблицу:</div>
        <div class="bt-qtpl-item" data-tpl="bt_30" style="padding:6px 8px;font-size:11.5px;color:#f4f4f5;cursor:pointer;border-radius:5px;font-weight:600;">📊 Бэктест 30 сделок (основной)</div>
        <div class="bt-qtpl-item" data-tpl="bt_main" style="padding:6px 8px;font-size:11.5px;color:#f4f4f5;cursor:pointer;border-radius:5px;">📘 Бэктест «Матрёшка»</div>
        <div class="bt-qtpl-item" data-tpl="office" style="padding:6px 8px;font-size:11.5px;color:#f4f4f5;cursor:pointer;border-radius:5px;">🏢 Торговый офис (Недели 1–4)</div>
        <div class="bt-qtpl-item" data-tpl="dz_poi" style="padding:6px 8px;font-size:11.5px;color:#f4f4f5;cursor:pointer;border-radius:5px;">🎯 ДЗ: Журнал POI</div>
        <div class="bt-qtpl-item" data-tpl="dz_aoi" style="padding:6px 8px;font-size:11.5px;color:#f4f4f5;cursor:pointer;border-radius:5px;">🗺️ ДЗ: Таблица AOI</div>
        <div class="bt-qtpl-item" data-tpl="prices_only" style="padding:6px 8px;font-size:11.5px;color:#f4f4f5;cursor:pointer;border-radius:5px;">⚡ Только 3 цены (Вход/Стоп/Цель)</div>
        <div class="bt-qtpl-item" data-tpl="custom" style="padding:6px 8px;font-size:11.5px;color:#f3d26b;cursor:pointer;border-radius:5px;border-top:1px solid #282832;margin-top:2px;">✨ Свой набор колонок...</div>
      </div>
      <div style="display:flex;align-items:center;gap:4px;flex-shrink:0;">
        <div style="display:flex;align-items:center;background:#16161a;border:1px solid #2a2a32;border-radius:5px;padding:1px 3px;" title="Переключение между сделками на графике (стрелки ← и →)">
          <button type="button" id="bt-prev-trade" class="bt-icon-btn">◀</button>
          <span id="bt-trade-counter" style="font-size:11px;font-weight:700;color:#e5c158;padding:0 4px;">1/1</span>
          <button type="button" id="bt-next-trade" class="bt-icon-btn">▶</button>
        </div>
        <button type="button" id="bt-open-settings" class="bt-top-btn" title="Настройки колонок и шаблона">⚙️</button>
        <button type="button" id="bt-min-btn" class="bt-top-btn" title="Свернуть или развернуть окно">—</button>
        <button type="button" id="bt-close-x" class="bt-top-btn" title="Закрыть окно (Esc)">✕</button>
      </div>
    </div>
    <div id="bt-trade-card" style="background:#141418;border:1px solid #26262e;border-radius:8px;padding:7px 9px;margin-bottom:8px;transition:border-color 0.2s ease;">
      <div id="bt-version-banner" style="display:none;margin-bottom:7px;"></div>
      <div style="display:flex;justify-content:space-between;align-items:center;gap:4px;margin-bottom:5px;white-space:nowrap;">
        <div style="display:flex;align-items:center;gap:6px;font-size:11.5px;min-width:0;overflow:hidden;">
          <b id="bt-info-sym" style="color:#ffffff;white-space:nowrap;">XAUUSD · M1</b>
          <button type="button" id="bt-session-pill" title="Торговая сессия по времени входа. Нажмите, чтобы изменить" style="background:rgba(229,193,88,0.14);color:#f3d26b;border:1px solid rgba(229,193,88,0.4);border-radius:4px;padding:1px 6px;font-size:10.5px;font-weight:600;cursor:pointer;white-space:nowrap;">⚡ Европа</button>
          <button type="button" id="bt-info-date" title="Время свечи фактического входа. Нажмите для переключения МСК (UTC+3) / пояса графика" style="background:transparent;border:1px solid transparent;border-radius:4px;color:#9ca3af;font-size:10.5px;padding:1px 4px;cursor:pointer;white-space:nowrap;transition:0.12s;"></button>
        </div>
        <span id="bt-info-badge" style="background:#089981;color:#fff;font-size:10.5px;padding:2px 7px;border-radius:4px;font-weight:700;white-space:nowrap;flex-shrink:0;"></span>
      </div>
      <div id="bt-quick-copy-prices" title="Нажмите, чтобы скопировать только цены: Вход, Стоп и Цель" style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:4px;background:#0e0e11;padding:5px 8px;border-radius:5px;cursor:pointer;border:1px solid #222228;font-size:11.5px;white-space:nowrap;transition:0.12s;">
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
        <div class="bt-lbl" style="display:none;">ШАБЛОН ТАБЛИЦЫ</div>
        <select id="bt-preset-select" style="display:none;">
          <option value="bt_30">📊 1. Бэктест 30 сделок (30_trades.xlsx)</option>
          <option value="bt_main">📘 2. Бэктест «Матрёшка»</option>
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
            <input type="text" id="bt-inp-comment" autocomplete="off" placeholder="Комментарий к сделке (Tab)..." style="width:100%;box-sizing:border-box;background:#141418;border:1px solid #2a2a32;color:#ffffff;padding:6px 44px 6px 9px;border-radius:6px;font-size:11.5px;outline:none;">
            <button type="button" id="bt-clear-comment" title="Очистить комментарий" style="position:absolute;right:22px;background:transparent;border:none;color:#9ca3af;cursor:pointer;font-size:11px;padding:2px 4px;">✕</button>
            <button type="button" id="bt-toggle-presets" title="Выбрать частый комментарий" style="position:absolute;right:4px;background:transparent;border:none;color:#e5c158;cursor:pointer;font-size:11px;padding:2px 5px;">▾</button>
          </div>
          <div id="bt-comment-menu" class="bt-no-scroll" style="display:none;position:absolute;bottom:34px;left:0;width:285px;max-height:185px;overflow-y:auto;background:#18181e;border:1px solid #3d3522;border-radius:6px;box-shadow:0 10px 25px rgba(0,0,0,0.85);z-index:100;padding:4px;">
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
        <div style="padding:3px 2px 6px;margin-bottom:4px;">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;gap:4px;">
            <span id="bt-preview-hint" style="font-size:9px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:0.3px;">СТРОКА ДЛЯ ВСТАВКИ</span>
            <div style="display:flex;align-items:center;gap:4px;">
              <div style="display:inline-flex;background:#080a0f;border:1px solid #1e2330;border-radius:4px;padding:1px;">
                <button type="button" class="bt-prof-btn" data-prof="gs_ru" style="font-size:9px;padding:1px 5px;border-radius:3px;" title="Формат для Google Таблиц (RU)">Google</button>
                <button type="button" class="bt-prof-btn" data-prof="ru_form" style="font-size:9px;padding:1px 5px;border-radius:3px;" title="Формат для программы Excel (RU)">Excel RU</button>
                <button type="button" class="bt-prof-btn" data-prof="en_form" style="font-size:9px;padding:1px 5px;border-radius:3px;" title="Международный формат (EN, через точку)">EN</button>
              </div>
              <button type="button" id="bt-toggle-fix-help" title="Что делать, если в таблице ошибка #ERROR или #ЗНАЧ" style="background:#141722;border:1px solid #282e42;border-radius:4px;color:#e5c158;font-size:9.5px;font-weight:700;cursor:pointer;padding:1px 5px;">?</button>
            </div>
          </div>
          <div id="bt-row-preview" class="bt-hide-scroll" style="display:flex;gap:4px;overflow-x:auto;"></div>
        </div>
        <div id="bt-fix-help-box" style="display:none;background:#181612;border:1px solid #4a3f23;border-radius:6px;padding:7px 8px;margin-bottom:7px;font-size:10.5px;color:#d4d4d8;">
          <div style="font-weight:700;color:#f3d26b;margin-bottom:5px;">Если после вставки в таблице появилась ошибка:</div>
          <div style="display:flex;flex-direction:column;gap:4px;">
            <button type="button" class="bt-fix-action" data-fix="gs_ru" style="text-align:left;background:#121216;color:#e4e4e7;border:1px solid #2e2e38;border-radius:4px;padding:5px 7px;font-size:10px;cursor:pointer;">🔴 Ошибка <b style="color:#f87171;">#ERROR!</b> в Google Таблице → <u style="color:#f3d26b;">Переключить на Google и скопировать</u></button>
            <button type="button" class="bt-fix-action" data-fix="ru_form" style="text-align:left;background:#121216;color:#e4e4e7;border:1px solid #2e2e38;border-radius:4px;padding:5px 7px;font-size:10px;cursor:pointer;">🔴 Ошибка <b style="color:#f87171;">#ERROR!</b> в программе Excel → <u style="color:#f3d26b;">Переключить на Excel RU и скопировать</u></button>
            <button type="button" class="bt-fix-action" data-fix="en_form" style="text-align:left;background:#121216;color:#e4e4e7;border:1px solid #2e2e38;border-radius:4px;padding:5px 7px;font-size:10px;cursor:pointer;">🔴 Числа без запятой или ошибка ";" в Excel → <u style="color:#f3d26b;">Переключить на EN и скопировать</u></button>
            <div style="font-size:9.5px;color:#9ca3af;padding:2px 2px 0;">💡 Если пишет <b>#ЗНАЧ! (#VALUE!)</b> — выделите ячейку в столбце <b id="bt-fix-col-lbl" style="color:#f3d26b;">B</b> и нажмите Ctrl+V.</div>
          </div>
        </div>
        <div style="position:relative;display:flex;gap:5px;align-items:stretch;">
          <div id="bt-batch-popover" style="display:none;position:absolute;bottom:44px;right:0;width:356px;background:#141722;border:1px solid #e5c158;border-radius:10px;box-shadow:0 18px 45px rgba(0,0,0,0.95);z-index:350;padding:10px;box-sizing:border-box;">
            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:7px;padding-bottom:6px;border-bottom:1px solid #262c3f;">
              <span style="font-size:11px;font-weight:700;color:#f3d26b;">📊 ОТЧЁТ И СПИСОК СДЕЛОК</span>
              <div style="display:flex;align-items:center;gap:5px;">
                <span style="font-size:9.5px;color:#94a3b8;">1R=</span>
                <input type="text" id="bt-risk-pct-inp" value="${state.riskPctPer1R || 0.3}" title="Риск на 1 стоп в % от депозита" style="width:32px;height:20px;background:#090b10;color:#f3d26b;border:1px solid #333a52;border-radius:4px;font-size:10px;font-weight:700;text-align:center;outline:none;">
                <span style="font-size:9.5px;color:#94a3b8;">%</span>
                <div style="display:flex;background:#090b10;border:1px solid #262c3f;border-radius:4px;padding:1px;">
                  <button type="button" class="bt-bmode-btn" data-bm="pct" style="background:transparent;border:none;color:#94a3b8;font-size:9.5px;font-weight:700;padding:1px 5px;border-radius:3px;cursor:pointer;" title="Показывать в процентах (%)">%</button>
                  <button type="button" class="bt-bmode-btn" data-bm="r" style="background:transparent;border:none;color:#94a3b8;font-size:9.5px;font-weight:700;padding:1px 5px;border-radius:3px;cursor:pointer;" title="Показывать в рисках (R)">R</button>
                  <button type="button" class="bt-bmode-btn" data-bm="cnt" style="background:transparent;border:none;color:#94a3b8;font-size:9.5px;font-weight:700;padding:1px 5px;border-radius:3px;cursor:pointer;" title="Показывать количество сделок">№</button>
                </div>
                <button type="button" id="bt-close-batch-pop" title="Закрыть список сделок (Esc)" style="background:#1f2433;color:#e2e8f0;border:1px solid #333a52;border-radius:5px;width:22px;height:20px;font-size:11px;line-height:1;cursor:pointer;display:flex;align-items:center;justify-content:center;margin-left:2px;">✕</button>
              </div>
            </div>
            <div style="background:#0d0d11;border:1px solid #262632;border-radius:6px;padding:6px 8px;margin-bottom:6px;display:flex;justify-content:space-between;align-items:center;gap:6px;">
              <pre id="bt-daily-report-text" style="margin:0;font-family:inherit;font-size:10.5px;line-height:1.4;color:#e4e4e7;white-space:pre-wrap;"></pre>
              <button type="button" id="bt-copy-report-btn" style="background:#e5c158;color:#09090b;border:none;border-radius:5px;padding:5px 8px;font-size:10px;font-weight:700;cursor:pointer;white-space:nowrap;flex-shrink:0;">📋</button>
            </div>
            <div style="font-size:9px;font-weight:700;color:#9ca3af;margin-bottom:3px;text-transform:uppercase;">Сохранённые сделки (нажмите ✕ для удаления):</div>
            <div id="bt-batch-items-list" class="bt-no-scroll" style="max-height:135px;overflow-y:auto;display:flex;flex-direction:column;gap:3px;"></div>
          </div>

          <button type="button" id="bt-main-copy-btn" style="flex:1;min-width:0;background:linear-gradient(180deg,#ebd074 0%,#cfa538 100%);color:#09090b;border:none;padding:9px 8px;border-radius:6px;cursor:pointer;font-weight:700;font-size:11.5px;display:flex;justify-content:center;align-items:center;gap:5px;box-shadow:0 3px 12px rgba(229,193,88,0.28);white-space:nowrap;overflow:hidden;">
            <span id="bt-main-btn-text" style="overflow:hidden;text-overflow:ellipsis;">Скопировать строку</span>
            <kbd style="background:rgba(0,0,0,0.18);color:#09090b;padding:1px 4px;border-radius:3px;font-size:10px;font-family:inherit;flex-shrink:0;">↵</kbd>
          </button>
          <div style="display:flex;align-items:center;background:#141418;border:1px solid #2a2a32;border-radius:6px;overflow:hidden;flex-shrink:0;">
            <button type="button" id="bt-copy-batch-btn" title="ЛКМ — скопировать все сделки для таблицы | ПКМ — открыть список сделок" style="background:transparent;color:#e4e4e7;border:none;padding:0 7px;height:100%;font-size:11px;font-weight:600;cursor:pointer;display:flex;align-items:center;gap:4px;white-space:nowrap;">📦 <span id="bt-batch-count" style="color:#f3d26b;">0</span></button>
            <button type="button" id="bt-open-report-btn" title="Открыть список сделок (или ПКМ по коробке)" style="background:rgba(229,193,88,0.1);color:#f3d26b;border:none;border-left:1px solid #2a2a32;padding:0 6px;height:100%;font-size:11px;cursor:pointer;">📋</button>
            <button type="button" id="bt-clear-batch-btn" title="Очистить список сохранённых сделок" style="background:rgba(242,54,69,0.16);color:#f87171;border:none;border-left:1px solid #2a2a32;padding:0 7px;height:100%;font-size:11px;font-weight:700;cursor:pointer;">✕</button>
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
    .bt-switch{width:100%;height:27px;box-sizing:border-box;border-radius:6px;padding:0 8px;margin-bottom:6px;background:#14141a;border:1px solid #282834;color:#9ca3af;display:flex;align-items:center;justify-content:space-between;gap:5px;cursor:pointer;transition:all 0.15s ease;outline:none}
    .bt-switch:hover{border-color:#454554;color:#e4e4e7;background:#181820}.bt-sw-txt{font-size:11px;font-weight:600;line-height:1;white-space:nowrap;transition:color 0.15s}
    .bt-sw-track{width:26px;height:14px;border-radius:99px;background:#252530;border:1px solid #3a3a48;position:relative;flex-shrink:0;transition:all 0.18s ease;box-sizing:border-box}
    .bt-sw-thumb{width:10px;height:10px;border-radius:50%;background:#9ca3af;position:absolute;top:1px;left:1px;transition:transform 0.18s cubic-bezier(0.4,0,0.2,1),background 0.18s ease;box-shadow:0 1px 2px rgba(0,0,0,0.6)}
    .bt-switch[data-val="да"] .bt-sw-thumb{transform:translateX(12px);background:#ffffff}
    .bt-switch[data-val="да"][data-mode="gold"]{background:rgba(229,193,88,0.12);border-color:rgba(229,193,88,0.5);color:#f3d26b}.bt-switch[data-val="да"][data-mode="gold"] .bt-sw-track{background:#c69a31;border-color:#e5c158}
    .bt-switch[data-val="да"][data-mode="check"]{background:rgba(16,185,129,0.13);border-color:rgba(16,185,129,0.5);color:#34d399}.bt-switch[data-val="да"][data-mode="check"] .bt-sw-track{background:#059669;border-color:#10b981}
    .bt-switch[data-val="нет"][data-mode="check"]{background:rgba(242,54,69,0.12);border-color:rgba(242,54,69,0.45);color:#f87171}.bt-switch[data-val="нет"][data-mode="check"] .bt-sw-track{background:#3a181e;border-color:#f23645}.bt-switch[data-val="нет"][data-mode="check"] .bt-sw-thumb{background:#f87171}
    .bt-switch[data-val="да"][data-mode="danger"]{background:rgba(242,54,69,0.14);border-color:rgba(242,54,69,0.55);color:#f87171}.bt-switch[data-val="да"][data-mode="danger"] .bt-sw-track{background:#dc2626;border-color:#f87171}
    .bt-flow-pill{width:100%;height:26px;border-radius:6px;font-size:11px;font-weight:600;cursor:pointer;display:flex;justify-content:space-between;align-items:center;padding:0 7px;box-sizing:border-box;transition:0.12s ease;border:1px solid #282830;background:#16161b;color:#d4d4d8}
    .bt-flow-pill:hover{filter:brightness(1.12)}.bt-flow-pill[data-dir="вверх"]{background:rgba(16,185,129,0.14);color:#34d399;border-color:rgba(16,185,129,0.45)}.bt-flow-pill[data-dir="вниз"]{background:rgba(242,54,69,0.14);color:#f87171;border-color:rgba(242,54,69,0.45)}.bt-flow-pill[data-dir="—"]{background:#14141a;color:#9ca3af;border-color:#282834}
    .bt-prof-btn{background:transparent;color:#9ca3af;border:1px solid transparent;border-radius:4px;padding:4px;font-size:10.5px;font-weight:600;cursor:pointer;transition:0.12s;white-space:nowrap}.bt-prof-btn:hover{color:#f4f4f5}.bt-prof-btn.active{background:#242014;color:#f3d26b;border-color:rgba(229,193,88,0.55);box-shadow:0 1px 4px rgba(0,0,0,0.5)}
    .bt-struct-preset,.bt-cand-pill{background:#15151b;color:#9ca3af;border:1px solid #282834;border-radius:4px;padding:1.5px 5px;font-size:9.5px;cursor:pointer;transition:0.12s;white-space:nowrap}.bt-struct-preset:hover,.bt-cand-pill:hover{color:#f3d26b;border-color:#52525b}.bt-struct-preset.active,.bt-cand-pill.active{background:rgba(229,193,88,0.16);color:#f3d26b;border-color:rgba(229,193,88,0.55);font-weight:600}
    .bt-sel{width:100%;height:27px;box-sizing:border-box;background:#16161b;color:#f4f4f5;border:1px solid #282830;padding:0 6px;border-radius:6px;font-size:11px;font-weight:500;outline:none;margin-bottom:6px;cursor:pointer}.bt-sel:focus,#bt-inp-comment:focus,#f-swept-lvl:focus,#bt-safe-price-inp:focus{border-color:#e5c158}
    .bt-mini-pill{background:#1c1c22;color:#d4d4d8;border:1px solid #2e2e38;padding:3px 8px;font-size:10.5px;border-radius:4px;cursor:pointer;transition:0.12s}.bt-mini-pill.active{background:rgba(229,193,88,0.2);color:#f3d26b;border-color:#e5c158;font-weight:700}
    .bt-top-btn{background:#16161a;color:#d4d4d8;border:1px solid #2a2a32;border-radius:5px;padding:2px 7px;font-size:11.5px;cursor:pointer;transition:0.12s}.bt-top-btn:hover,.bt-top-btn.active,#bt-quick-tpl-btn:hover{color:#f3d26b;border-color:#e5c158;background:rgba(229,193,88,0.14)}
    .bt-icon-btn{background:transparent;color:#d4d4d8;border:none;cursor:pointer;font-size:10.5px;padding:2px 5px}.bt-icon-btn:hover{color:#f3d26b}
    .bt-drag-pill{background:#1a1a20;color:#f4f4f5;border:1px solid #383844;border-radius:5px;padding:3px 7px;font-size:10.5px;cursor:grab;display:inline-flex;align-items:center;gap:5px;transition:transform 0.1s ease,border-color 0.1s ease}
    .bt-drag-pill.dragging-active{background:#2a2515;border-color:#f3d26b;color:#f3d26b;transform:scale(1.05);z-index:10;cursor:grabbing;box-shadow:0 4px 12px rgba(0,0,0,0.6)}
    .bt-del-col{color:#f23645;font-weight:bold;cursor:pointer;padding:0 2px;font-size:11px}
    .bt-cm-item:hover,.bt-cc-item:hover,.bt-fix-action:hover,.bt-qtpl-item:hover{background:rgba(229,193,88,0.18)!important;border-color:#e5c158!important;color:#f3d26b!important}
    #bt-cc-btn:hover{border-color:#52525b}#bt-info-date:hover{border-color:#3a3a48;color:#f3d26b!important;background:#1c1c24!important}#bt-quick-copy-prices:hover{border-color:rgba(229,193,88,0.55)}
    .bt-prev-cell{background:#16161c;border:1px solid #262630;border-radius:4px;padding:2px 6px;flex-shrink:0;display:flex;flex-direction:column;min-width:48px;max-width:130px}.bt-prev-cell.is-first{border-color:rgba(229,193,88,0.55);background:rgba(229,193,88,0.08)}
    .bt-safe-btn{background:#15151c;color:#a1a1aa;border:1px solid #282834;border-radius:5px;padding:3px 6px;font-size:10px;font-weight:600;cursor:pointer;transition:0.12s;white-space:nowrap}
    .bt-safe-btn:hover{color:#f4f4f5;border-color:#454554}
    .bt-safe-btn.active{background:rgba(16,185,129,0.16);color:#34d399;border-color:rgba(16,185,129,0.55)}
    .bt-safe-btn.has-chart-safe{border-color:rgba(229,193,88,0.55);color:#f3d26b}
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
  const quickTplBtn = document.getElementById('bt-quick-tpl-btn');
  const quickTplMenu = document.getElementById('bt-quick-tpl-menu');
  let toastTimer = null;

  
  function clampModalPosition() {
    if (!modalEl) return;
    const minTop = 48; 
    const maxTop = Math.max(minTop, window.innerHeight - 120);
    const minRight = 10;
    const maxRight = Math.max(minRight, window.innerWidth - 415);

    const clampedTop = Math.min(maxTop, Math.max(minTop, state.posTop || 56));
    const clampedRight = Math.min(maxRight, Math.max(minRight, state.posRight || 20));
    modalEl.style.top = `${clampedTop}px`;
    modalEl.style.right = `${clampedRight}px`;
    modalEl.style.maxHeight = `calc(100vh - ${clampedTop + 14}px)`;
  }

  window.addEventListener('resize', clampModalPosition, { passive: true });

  function mountNativeToolbarButton() {
    const fabBtn = document.getElementById('bt-fab-btn');
    if (fabBtn) fabBtn.style.display = 'none';

    let nativeBtn = document.getElementById('bt-native-toolbar-btn');
    if (!nativeBtn) {
      const tipText = `Бэктест Хэлпер (${isMac ? '⌃Q' : 'F2'})`;
      nativeBtn = document.createElement('button');
      nativeBtn.id = 'bt-native-toolbar-btn';
      nativeBtn.type = 'button';
      nativeBtn.tabIndex = -1;
      nativeBtn.title = tipText;
      nativeBtn.setAttribute('aria-label', tipText);
      nativeBtn.style.cssText = 'position:fixed;z-index:9999990;width:34px;height:34px;padding:0;border:none;border-radius:6px;background:transparent;cursor:pointer;display:none;align-items:center;justify-content:center;transition:background 0.12s ease;';
      nativeBtn.onmouseenter = () => { nativeBtn.style.background = 'rgba(255,255,255,0.08)'; };
      nativeBtn.onmouseleave = () => { nativeBtn.style.background = 'transparent'; };
      nativeBtn.innerHTML = `
        <span role="img" aria-hidden="true" style="display:flex;align-items:center;justify-content:center;pointer-events:none;">
          ${NATIVE_TOOLBAR_SVG}
        </span>
        <span id="bt-native-badge" style="display:none;position:absolute;top:2px;right:2px;min-width:13px;height:13px;padding:0 3px;border-radius:99px;background:#e5c158;color:#09090b;font-size:9px;font-weight:800;line-height:13px;text-align:center;box-sizing:border-box;pointer-events:none;"></span>
      `;
      nativeBtn.onclick = e => {
        e.preventDefault();
        e.stopPropagation();
        if (modalEl.style.display === 'block') modalEl.style.display = 'none';
        else openModal();
      };
      document.body.appendChild(nativeBtn);
    }

    const favBtn = document.getElementById('drawing-toolbar-favorite-drawings');
    const trashEl = document.querySelector('#drawing-toolbar [data-name="removeAllDrawingTools"]');
    const favRect = favBtn?.getBoundingClientRect();
    const trashRect = trashEl?.getBoundingClientRect();
    const freeLeftGap = (favRect && trashRect) ? (favRect.top - trashRect.bottom) : 0;

    if (favRect && favRect.width > 0 && freeLeftGap >= 36) {
      nativeBtn.style.left = `${Math.round(favRect.left + (favRect.width - 34) / 2)}px`;
      nativeBtn.style.top = `${Math.round(favRect.top - 36)}px`;
      nativeBtn.style.display = 'flex';
    } else {
      const undoRedoWrap = document.getElementById('header-toolbar-undo-redo') || document.getElementById('header-toolbar-replay');
      const urRect = undoRedoWrap?.getBoundingClientRect();
      if (urRect && urRect.width > 0) {
        nativeBtn.style.left = `${Math.round(urRect.right + 10)}px`;
        nativeBtn.style.top = `${Math.round(urRect.top + (urRect.height - 34) / 2)}px`;
        nativeBtn.style.display = 'flex';
      } else if (favRect && favRect.width > 0) {
        nativeBtn.style.left = `${Math.round(favRect.right + 6)}px`;
        nativeBtn.style.top = `${Math.round(favRect.top)}px`;
        nativeBtn.style.display = 'flex';
      }
    }

    const badgeEl = document.getElementById('bt-native-badge');
    if (badgeEl) {
      const bStats = computeBatchStats();
      if (bStats.count > 0) {
        badgeEl.innerText = bStats.toolbarBadgeText;
        badgeEl.style.background = bStats.totalR > 0 ? '#10b981' : (bStats.totalR < 0 ? '#f23645' : '#e5c158');
        badgeEl.style.color = bStats.totalR !== 0 ? '#ffffff' : '#09090b';
        badgeEl.style.display = 'block';
      } else {
        badgeEl.style.display = 'none';
      }
    }
  }

  let emptyDataRetryTimer = null;
  let emptyDataRetryCount = 0;

  function scheduleEmptyDataRetry() {
    clearTimeout(emptyDataRetryTimer);
    if (!currentTrade?.ind?.isEmptyData || modalEl.style.display !== 'block') {
      emptyDataRetryCount = 0;
      return;
    }
    if (emptyDataRetryCount >= 8) return;
    emptyDataRetryCount++;
    emptyDataRetryTimer = setTimeout(() => {
      if (modalEl.style.display !== 'block') return;
      if (scanChartTrades(true, false) && currentTrade && !currentTrade.ind?.isEmptyData) {
        emptyDataRetryCount = 0;
        renderMainFields();
        updateHeaderUI();
        pulseTradeCard('#089981');
        showToast('⚡ Индикаторы загрузились — данные подставлены автоматически!');
      } else {
        scheduleEmptyDataRetry();
      }
    }, 380);
  }

  function updateVersionBannerUI() {
    const bannerEl = document.getElementById('bt-version-banner');
    if (!bannerEl) return;

    const health = currentTrade?.ind?.health;
    const isEmpty = !!currentTrade?.ind?.isEmptyData;
    let html = '';

    if (remoteUpdateVersion) {
      html += `
        <div style="background:rgba(229,193,88,0.14);border:1px solid #e5c158;border-radius:7px;padding:6px 9px;margin-bottom:5px;display:flex;justify-content:space-between;align-items:center;gap:6px;">
          <span style="font-size:10.5px;color:#f3d26b;font-weight:600;">🔄 Доступен новый скрипт <b>v${remoteUpdateVersion}</b> (у вас v${SCRIPT_VERSION})</span>
          <button type="button" id="bt-oneclick-update" style="background:#e5c158;color:#09090b;border:none;border-radius:4px;padding:3px 8px;font-size:10px;font-weight:700;cursor:pointer;white-space:nowrap;">Обновить в 1 клик</button>
        </div>`;
    }

    if (isEmpty && (!health || health.isAllLatest)) {
      html += `
        <div style="background:rgba(245,158,11,0.16);border:1px solid #f59e0b;border-radius:7px;padding:6px 9px;margin-bottom:4px;display:flex;justify-content:space-between;align-items:center;gap:6px;">
          <span style="font-size:10.5px;color:#fbbf24;font-weight:600;">⏳ Индикаторы ещё грузятся (авто-данных 0) — ждём расчёта...</span>
          <button type="button" id="bt-rescan-now-btn" style="background:#f59e0b;color:#09090b;border:none;border-radius:4px;padding:2px 7px;font-size:10px;font-weight:700;cursor:pointer;white-space:nowrap;">↻ Перечитать</button>
        </div>`;
      scheduleEmptyDataRetry();
    }

    if (health && !health.isAllLatest && health.warnings.length > 0) {
      html += `
        <div style="background:rgba(242,54,69,0.14);border:1px solid rgba(242,54,69,0.55);border-radius:7px;padding:7px 9px;font-size:10px;color:#fca5a5;line-height:1.35;">
          <div style="font-weight:700;color:#f87171;margin-bottom:3px;">⚠️ Внимание: автозаполнение может ошибаться!</div>
          ${health.warnings.map(w => `<div style="margin-top:2px;">• ${w}</div>`).join('')}
        </div>`;
    }

    if (html) {
      bannerEl.innerHTML = html;
      bannerEl.style.display = 'block';
      const updBtn = document.getElementById('bt-oneclick-update');
      if (updBtn) updBtn.onclick = () => window.open(`${SCRIPT_RAW_URL}?t=${Date.now()}`, '_blank');
      const rescanBtn = document.getElementById('bt-rescan-now-btn');
      if (rescanBtn) {
        rescanBtn.onclick = () => {
          if (scanChartTrades(false, true)) {
            renderMainFields();
            updateHeaderUI();
            pulseTradeCard('#089981');
          }
        };
      }
    } else {
      bannerEl.style.display = 'none';
      bannerEl.innerHTML = '';
    }
  }

  function updateScreenshotPillUI() {
    if (screenshotUrl) {
      screenPillBtn.innerText = '📸 Скрин ✓';
      screenPillBtn.style.background = 'rgba(16,185,129,0.16)';
      screenPillBtn.style.color = '#34d399';
      screenPillBtn.style.borderColor = 'rgba(16,185,129,0.5)';
      screenPillBtn.title = `Прикреплён: ${screenshotUrl} (нажмите, чтобы удалить)`;
    } else {
      screenPillBtn.innerText = '📸 Скрин';
      screenPillBtn.style.background = '#0c0e15';
      screenPillBtn.style.color = '#94a3b8';
      screenPillBtn.style.borderColor = '#252b3d';
      screenPillBtn.title = 'Скопируйте ссылку на скриншот и нажмите Ctrl+V в открытом окне';
    }
  }

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
    if (isSettingsOpen) {showColumnEditor = true; renderSettingsDrawer();}
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
        <button type="button" id="bt-cc-btn" style="width:100%;height:27px;box-sizing:border-box;background:#16161b;color:#f4f4f5;border:1px solid #282830;padding:0 7px;border-radius:6px;font-size:11px;font-weight:500;cursor:pointer;display:flex;align-items:center;justify-content:space-between;outline:none;transition:0.12s;">
          <span style="display:flex;align-items:center;gap:6px;min-width:0;overflow:hidden;">
            <span id="bt-cc-dot" style="width:12px;height:12px;border-radius:3px;background:${info.hex};flex-shrink:0;box-shadow:0 0 0 1px rgba(255,255,255,0.22);"></span>
            <span id="bt-cc-label" style="font-weight:600;">${activeVal}</span>
            <span id="bt-cc-hint" style="color:#9ca3af;font-size:10px;">${info.hint}</span>
          </span><span style="color:#9ca3af;font-size:10px;">▾</span>
        </button>
        <div id="bt-cc-menu" style="display:none;position:absolute;top:30px;left:0;width:100%;background:#18181e;border:1px solid #3d3522;border-radius:6px;box-shadow:0 10px 25px rgba(0,0,0,0.92);z-index:120;padding:3px;box-sizing:border-box;">${itemsHtml}</div>
      </div>`;
  }

  const SHIELD_SVG = `<svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round" style="flex-shrink:0;"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>`;

  function setCandleColVal(val) {
    const info = CANDLE_COLORS[val];
    if (!info) return;
    const inp = document.getElementById('f-candle-col');
    const dot = document.getElementById('bt-cc-dot');
    const lbl = document.getElementById('bt-cc-label');
    const hnt = document.getElementById('bt-cc-hint');
    const boltEl = document.getElementById('bt-cc-bolt');
    const boxEl = document.getElementById('bt-cc-box');

    if (inp) inp.value = val;
    if (dot) dot.style.background = info.hex;
    if (lbl) lbl.innerText = val;
    if (hnt) hnt.innerText = info.hint;

    
    const ind = currentTrade?.ind || {};
    const autoCol = ind.candleColor || null;
    const isAutoMatch = !!(ind.candleMatched && autoCol === val);

    if (boltEl) {
      if (isAutoMatch) {
        boltEl.style.display = 'inline';
        boltEl.style.opacity = '1';
        boltEl.style.color = '#e5c158';
        boltEl.title = 'Определено автоматически по свече входа на графике';
      } else if (ind.candleMatched && autoCol) {
        boltEl.style.display = 'inline';
        boltEl.style.opacity = '0.3';
        boltEl.style.color = '#94a3b8';
        boltEl.title = `Изменено вручную (авто на графике: ${autoCol})`;
      } else {
        boltEl.style.display = 'none';
      }
    }
    if (boxEl) {
      boxEl.classList.toggle('is-auto-field', isAutoMatch);
    }
    const ccMenu = document.getElementById('bt-cc-menu');
    if (ccMenu) {
      ccMenu.querySelectorAll('.bt-cc-item').forEach(it => {
        it.classList.toggle('active-cc', it.dataset.val === val);
      });
    }
  }

  function makeSelect(id, options, activeVal, mb = '6px', extraCls = '') {
    const htmlOpts = options.map(o => {
      const val = typeof o === 'string' ? o : o.val;
      const label = typeof o === 'string' ? o : o.label;
      return `<option value="${val}" ${val === activeVal ? 'selected' : ''}>${label}</option>`;
    }).join('');
    return `<select class="bt-sel ${extraCls}" id="${id}" style="margin-bottom:${mb};">${htmlOpts}</select>`;
  }

  function getChipVal(id, fallback = '') {
    const el = document.getElementById(id);
    if (!el) return fallback;
    if (el.classList.contains('bt-switch')) return el.dataset.val || fallback;
    return el.querySelector('.active')?.dataset.val || fallback;
  }

  function updateOutcomeBadgeUI(val) {
    const btn = document.getElementById('bt-outcome-cycle-btn');
    if (!btn) return;
    
    const effVal = (activeSafeState.enabled && val === 'частичная фиксация')
      ? (computeSafeMetricsForTrade(currentTrade)?.remOut || 'цель')
      : val;

    const map = {
      'цель': { icon: '🎯', text: activeSafeState.enabled ? 'Остаток: Тейк' : 'Цель (Тейк)', bg: 'rgba(16,185,129,0.14)', fg: '#34d399', bd: 'rgba(16,185,129,0.42)' },
      'стоп': { icon: '🛑', text: activeSafeState.enabled ? 'Остаток: Стоп' : 'Стоп-лосс', bg: 'rgba(242,54,69,0.14)', fg: '#f87171', bd: 'rgba(242,54,69,0.42)' },
      'безубыток': { icon: '⚖️', text: activeSafeState.enabled ? 'Остаток: в БУ' : 'Безубыток (0R)', bg: 'rgba(148,163,184,0.12)', fg: '#e2e8f0', bd: '#333a52' },
      'частичная фиксация': { icon: SHIELD_SVG, text: 'Частич. фикс.', bg: 'rgba(229,193,88,0.14)', fg: '#f3d26b', bd: 'rgba(229,193,88,0.45)' },
      'закрыл руками': { icon: '✋', text: 'Закрыл руками', bg: 'rgba(229,193,88,0.14)', fg: '#f3d26b', bd: 'rgba(229,193,88,0.45)' }
    };
    const st = map[effVal] || map['цель'];
    btn.style.background = st.bg;
    btn.style.color = st.fg;
    btn.style.borderColor = st.bd;
    btn.innerHTML = `<span style="display:flex;align-items:center;gap:5px;overflow:hidden;text-overflow:ellipsis;"><span>${st.icon}</span><b>${st.text}</b></span><span style="font-size:9px;opacity:0.7;flex-shrink:0;">▾</span>`;
  }

  function setChipVal(id, val) {
    const el = document.getElementById(id);
    if (!el) return;
    if (el.classList.contains('bt-switch')) {
      if (el.classList.contains('bt-prop-pill')) updateCompactPropPillUI(el, val);
      else updateSwitchUI(el, val);
      return;
    }
    el.querySelectorAll('.bt-chip').forEach(b => b.classList.toggle('active', b.dataset.val === val));
    if (id === 'f-res30') updateOutcomeBadgeUI(val);
  }

  function makeFlowPill(id, tfLabel, dirVal) {
    const arrowText = dirVal === 'вверх' ? '↑ вверх' : (dirVal === 'вниз' ? '↓ вниз' : '—');
    return `<button type="button" class="bt-flow-pill" id="${id}" data-tf="${tfLabel}" data-dir="${dirVal}" title="Нажмите, чтобы изменить направление потока ${tfLabel}"><span style="color:#9ca3af;font-weight:500;">${tfLabel}</span><span class="bt-flow-val">${arrowText}</span></button>`;
  }

  function getFlowVal(id, fallback = 'вверх') { return document.getElementById(id)?.dataset.dir || fallback; }

  let lastSafeInitTradeId = null;
  let lastChartSafePriceSeen = null;
  let lastChartSafePrice2Seen = null;
  let lastSafeReachedSeen = false;
  let lastBarsLoadedSeen = false;

  function syncPresetHighlights() {
    const ind = currentTrade?.ind || {};
    const stEl = document.getElementById('f-struct');
    const stVal = stEl?.value;
    dynFieldsEl.querySelectorAll('.bt-struct-preset').forEach(b => b.classList.toggle('active', b.dataset.st === stVal));
    if (stEl) {
      const isStAuto = !!(ind.detectedStructure && stVal === ind.detectedStructure);
      stEl.classList.toggle('is-auto-field', isStAuto);
      const stBolt = document.getElementById('bt-struct-bolt');
      if (stBolt) {
        stBolt.style.opacity = isStAuto ? '1' : '0.3';
        stBolt.style.color = isStAuto ? '#e5c158' : '#94a3b8';
        stBolt.title = isStAuto ? 'Структура определена автоматически по графику' : `Изменено вручную (авто с графика: ${ind.detectedStructure}) — нажмите для возврата`;
      }
    }

    const zEl = document.getElementById('f-zone30');
    const zVal = zEl?.value;
    const zOpts = Array.isArray(ind.zoneOptions) ? ind.zoneOptions : [];
    const isZoneFromChart = zOpts.some(zo => zo.zoneType === zVal);
    dynFieldsEl.querySelectorAll('.bt-zone-cand').forEach(b => b.classList.toggle('active', b.dataset.zone === zVal));
    if (zEl) {
      zEl.classList.toggle('is-auto-field', isZoneFromChart);
      zEl.classList.toggle('is-manual-fallback', !zOpts.length);
      const zBolt = document.getElementById('bt-zone-bolt');
      if (zBolt) {
        zBolt.style.opacity = isZoneFromChart ? '1' : '0.3';
        zBolt.style.color = isZoneFromChart ? '#e5c158' : '#94a3b8';
      }
    }

    const swEl = document.getElementById('f-swept');
    const swLvlInp = document.getElementById('f-swept-lvl');
    const swPriceBolt = document.getElementById('bt-swept-price-bolt');
    const swVal = swEl?.value;
    const swOpts = Array.isArray(ind.sweepOptions) ? ind.sweepOptions : [];
    const matchedSwOpt = swOpts.find(so => so.sweptType === swVal);
    const isSweepFromChart = !!matchedSwOpt;

    dynFieldsEl.querySelectorAll('.bt-sweep-cand').forEach(b => b.classList.toggle('active', b.dataset.swept === swVal));
    if (swEl) {
      swEl.classList.toggle('is-auto-field', isSweepFromChart);
      swEl.classList.toggle('is-manual-fallback', swVal === 'ничего не снято' && !swOpts.length);
      const swBolt = document.getElementById('bt-sweep-bolt');
      if (swBolt) {
        swBolt.style.opacity = isSweepFromChart ? '1' : '0.3';
        swBolt.style.color = isSweepFromChart ? '#e5c158' : '#94a3b8';
      }
    }
    if (swLvlInp) {
      const curNum = parseFloat(String(swLvlInp.value || '').replace(',', '.'));
      const autoPrice = matchedSwOpt?.price ?? ind.matchedSweep?.price ?? null;
      const isPriceAuto = !!(autoPrice != null && !isNaN(curNum) && Math.abs(curNum - Number(autoPrice)) < 0.015);
      swLvlInp.classList.toggle('is-auto-field', isPriceAuto);
      if (swPriceBolt) {
        if (swVal === 'ничего не снято') {
          swPriceBolt.style.display = 'none';
        } else if (isPriceAuto) {
          swPriceBolt.style.display = 'inline';
          swPriceBolt.style.opacity = '1';
          swPriceBolt.style.color = '#e5c158';
          swPriceBolt.innerText = '⚡';
          swPriceBolt.title = 'Цена уровня подставлена автоматически с графика';
        } else if (autoPrice != null || currentTrade?.entryExtreme) {
          swPriceBolt.style.display = 'inline';
          swPriceBolt.style.opacity = '0.45';
          swPriceBolt.style.color = '#94a3b8';
          swPriceBolt.innerText = '↺';
          swPriceBolt.title = `Нажмите, чтобы вернуть расчётную цену уровня (${numStr(autoPrice ?? currentTrade?.entryExtreme)})`;
        } else {
          swPriceBolt.style.display = 'none';
        }
      }
    }

    dynFieldsEl.querySelectorAll('.bt-safe-quick-trig').forEach(b => {
      const m = b.dataset.smode;
      const isAct = (m === 'off') ? !activeSafeState.enabled : (activeSafeState.enabled && activeSafeState.mode === m);
      b.classList.toggle('active', isAct);
    });
  }

  function updateAutoStripSummaryUI() {
    const mapFlowBtn = (id) => {
      const el = document.getElementById(id);
      if (!el) return;
      const dir = el.dataset.dir || 'вверх';
      const tf = el.dataset.tf || '';
      const arrow = dir === 'вверх' ? '↑' : (dir === 'вниз' ? '↓' : '—');
      el.innerHTML = `<span style="opacity:0.75;">${tf}</span><b>${arrow}</b>`;
    };
    mapFlowBtn('f-fl-d');
    mapFlowBtn('f-fl-4h');
    mapFlowBtn('f-fl-1h');
    mapFlowBtn('f-fl-15m');

    const mapMiniToggle = (id, label) => {
      const el = document.getElementById(id);
      if (!el) return;
      const isYes = el.dataset.val === 'да';
      el.innerHTML = `<span>${label}</span><b>${isYes ? '✓' : '✕'}</b>`;
    };
    mapMiniToggle('f-hit1r', '1:1');
    mapMiniToggle('f-hit3r', '1:3');
    mapMiniToggle('f-check30', '7/7');
  }

  function updateScreenshotPillUI() {
    if (screenshotUrl) {
      screenPillBtn.innerText = '📸 Скрин ✓';
      screenPillBtn.style.background = 'rgba(16,185,129,0.16)';
      screenPillBtn.style.color = '#34d399';
      screenPillBtn.style.borderColor = 'rgba(16,185,129,0.5)';
      screenPillBtn.title = `Прикреплён: ${screenshotUrl} (нажмите, чтобы удалить)`;
    } else {
      screenPillBtn.innerText = '📸 Скрин';
      screenPillBtn.style.background = '#0c0e15';
      screenPillBtn.style.color = '#94a3b8';
      screenPillBtn.style.borderColor = '#252b3d';
      screenPillBtn.title = 'Нажмите, чтобы вставить ссылку на скриншот из буфера (или нажмите Alt+S на графике)';
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

  function updatePreviewBar() {
    if (!rowPreviewEl || !currentTrade) return;
    const snap = collectCurrentSnapshot();
    if (!snap) return;
    const rendered = renderSnapshotToRow(snap);
    const firstCell = rendered.cells[0];
    const rCnt = rendered.rowsCount || 2;

    const hintEl = document.getElementById('bt-preview-hint');
    if (hintEl && firstCell) {
      const twoRowsBadge = rendered.isTwoRows
        ? ` <span style="background:rgba(16,185,129,0.18);color:#34d399;padding:1px 5px;border-radius:3px;margin-left:4px;font-size:8.5px;">${rCnt} СТРОКИ</span>`
        : '';
      hintEl.innerHTML = `<span style="color:#64748b;">БУФЕР → СТЛБ.</span> <b style="color:#cbd5e1;">${firstCell.colLet} (${firstCell.shortLabel})</b>${twoRowsBadge}`;
    }
    const fixColLbl = document.getElementById('bt-fix-col-lbl');
    if (fixColLbl && firstCell) fixColLbl.innerText = `${firstCell.colLet} (${firstCell.shortLabel})`;

    if (mainBtnText && firstCell) {
      const bucket = getTplBucket();
      const isReEdit = !!(currentTrade && bucket.batchQueue.some(item => item.tradeId === currentTrade.id || item.signature === currentTrade.signature));
      if (rendered.isTwoRows) {
        mainBtnText.innerText = isReEdit
          ? `Обновить ${rCnt} стр. в ${rendered.pasteCol} (${firstCell.shortLabel})`
          : `Скопировать ${rCnt} стр. в ${rendered.pasteCol} (${firstCell.shortLabel})`;
      } else {
        mainBtnText.innerText = isReEdit
          ? `Обновить в стлб. ${rendered.pasteCol} (${firstCell.shortLabel})`
          : `Скопировать в стлб. ${rendered.pasteCol} (${firstCell.shortLabel})`;
      }
    }

    if (mainCopyBtn && mainBtnText) {
      if (currentTrade && !currentTrade.barsLoaded) {
        mainCopyBtn.style.background = 'linear-gradient(180deg,#f59e0b 0%,#d97706 100%)';
        mainBtnText.innerText = `🔭 Прокрутите график к ${currentTrade.mskTime}!`;
      } else if (currentTrade?.ind?.isEmptyData) {
        mainCopyBtn.style.background = 'linear-gradient(180deg,#f59e0b 0%,#d97706 100%)';
        mainBtnText.innerText = '⚠️ Индикаторы грузятся! Скопировать?';
      } else {
        mainCopyBtn.style.background = 'linear-gradient(180deg,#ebd074 0%,#cfa538 100%)';
      }
    }

    rowPreviewEl.innerHTML = `<div class="bt-excel-strip">` + rendered.cells.map((c, idx) => {
      const isFormula = c.val.startsWith('=');
      const displayVal = isFormula ? 'fx' : (c.val || '—');
      const valColor = isFormula ? '#60a5fa' : (c.val ? '#94a3b8' : '#334155');
      const safeTitle = c.val.replace(/"/g, '&quot;');
      return `<div class="bt-excel-col ${idx === 0 ? 'is-anchor' : ''}" title="Столбец ${c.colLet} (${TOKEN_LABELS[c.tok] || c.tok}): ${safeTitle}"><span class="bt-excel-th"><b>${c.colLet}</b> ${c.shortLabel}</span><span class="bt-excel-td" style="color:${valColor};">${displayVal}</span></div>`;
    }).join('') + `</div>`;
  }

  function initTradeSafeDefault(t) {
    if (!t) return;
    const curChartSafeP = t.chartSafe ? t.chartSafe.price : null;
    const curChartSafeP2 = t.chartSafe?.secondSafe ? t.chartSafe.secondSafe.price : '';
    const curReached = !!(t.chartSafe && t.chartSafe.reachedSafe);
    const curBarsLoaded = !!t.barsLoaded;

    const justLoadedBars = (t.id === lastSafeInitTradeId) && (!lastBarsLoadedSeen && curBarsLoaded);
    const justReachedSafe = (t.id === lastSafeInitTradeId) && (!lastSafeReachedSeen && curReached);

    if (t.id === lastSafeInitTradeId && !justLoadedBars && !justReachedSafe) {
      
      if (curChartSafeP != null && (curChartSafeP !== lastChartSafePriceSeen || curChartSafeP2 !== lastChartSafePrice2Seen)) {
        lastChartSafePriceSeen = curChartSafeP;
        lastChartSafePrice2Seen = curChartSafeP2;
        lastSafeReachedSeen = curReached;
        activeSafeState.enabled = true;
        activeSafeState.mode = 'chart';
        activeSafeState.price = curChartSafeP;
        activeSafeState.price2 = curChartSafeP2;
        if (curChartSafeP2) activeSafeState.share = 0.33;
        if (activeSafeState.remOutcome === 'auto') {
          const manInp = document.getElementById('f-manual-r30');
          if (manInp) manInp.value = '';
        }
      } else if (curChartSafeP == null && activeSafeState.mode === 'chart') {
        lastChartSafePriceSeen = null;
        lastChartSafePrice2Seen = null;
        lastSafeReachedSeen = false;
        activeSafeState.enabled = false;
        activeSafeState.mode = 'off';
        activeSafeState.price = '';
        activeSafeState.price2 = '';
        const defRes30 = t.autoResult === 'Тейк' ? 'цель' : (t.autoResult === 'Стоп' ? 'стоп' : 'безубыток');
        setChipVal('f-res30', defRes30);
        const manInp = document.getElementById('f-manual-r30');
        if (manInp) manInp.value = '';
      } else if (activeSafeState.mode === 'chart' && curChartSafeP != null) {
        activeSafeState.price = curChartSafeP;
        activeSafeState.price2 = curChartSafeP2;
      }
      return;
    }

    lastSafeInitTradeId = t.id;
    lastChartSafePriceSeen = curChartSafeP;
    lastChartSafePrice2Seen = curChartSafeP2;
    lastSafeReachedSeen = curReached;
    lastBarsLoadedSeen = curBarsLoaded;

    if (t.chartSafe && t.chartSafe.reachedSafe) {
      activeSafeState = {
        enabled: true,
        mode: 'chart',
        price: t.chartSafe.price,
        price2: curChartSafeP2,
        share: curChartSafeP2 ? 0.33 : (state.safeShare || 0.5),
        exportMode: state.safeExportMode || '1row',
        remOutcome: 'auto'
      };
    } else {
      activeSafeState = {
        enabled: false,
        mode: 'off',
        price: t.chartSafe ? t.chartSafe.price : '',
        price2: curChartSafeP2,
        share: state.safeShare || 0.5,
        exportMode: state.safeExportMode || '1row',
        remOutcome: 'auto'
      };
    }
  }

  function syncSafeUI() {
    const t = currentTrade;
    if (!t) return;
    const panel = document.getElementById('bt-safe-details');
    const manWrap = document.getElementById('bt-manual-r-wrap');
    const manInp = document.getElementById('f-manual-r30');
    const priceInp = document.getElementById('bt-safe-price-inp');
    const price2Wrap = document.getElementById('bt-safe-price2-wrap');
    const price2Inp = document.getElementById('bt-safe-price2-inp');
    const summarySpan = document.getElementById('bt-safe-calc-hint');
    const outcomeLblEl = document.getElementById('bt-outcome-main-lbl');
    const shHalfBtn = dynFieldsEl.querySelector('.bt-safe-share-btn[data-sshare="0.5"]');
    const shThirdBtn = dynFieldsEl.querySelector('.bt-safe-share-btn[data-sshare="0.33"]');

    if (outcomeLblEl) {
      outcomeLblEl.innerHTML = activeSafeState.enabled ? `ОСТАТОК ПОЗЫ ${BOLT}` : `ИСХОД СДЕЛКИ ${BOLT}`;
    }

    dynFieldsEl.querySelectorAll('.bt-safe-share-btn').forEach(b => {
      b.classList.toggle('active', Math.abs(parseFloat(b.dataset.sshare) - activeSafeState.share) < 0.05);
    });

    if (panel) panel.style.display = activeSafeState.enabled ? 'block' : 'none';
    if (priceInp && activeSafeState.price !== '' && document.activeElement !== priceInp) {
      priceInp.value = numStr(activeSafeState.price);
      priceInp.classList.toggle('is-auto-field', activeSafeState.mode === 'chart');
    }
    if (price2Inp && document.activeElement !== price2Inp) {
      price2Inp.value = activeSafeState.price2 ? numStr(activeSafeState.price2) : '';
    }

    const si = computeSafeMetricsForTrade(t);
    if (price2Wrap) {
      const show2nd = !!(si?.hasSecondSafe || activeSafeState.price2);
      price2Wrap.style.display = show2nd ? 'flex' : 'none';
      if (shHalfBtn && shThirdBtn) {
        shHalfBtn.innerText = show2nd ? '50/25' : '1/2';
        shHalfBtn.title = show2nd ? '50% на 1-м сейфе, 25% на 2-м сейфе, 25% остаток' : 'Закрыть 1/2 (50%) позиции на сейфе';
        shThirdBtn.innerText = show2nd ? 'По 1/3' : '1/3';
        shThirdBtn.title = show2nd ? 'По 1/3 (33%) на каждом сейфе и 1/3 на остаток' : 'Закрыть 1/3 (33%) позиции на сейфе';
      }
    }

    if (si && summarySpan) {
      const riskPct = Number(state.riskPctPer1R) || 0.3;
      const pctStr = numStr(Number((si.blendedR * riskPct).toFixed(2)));
      const sign = si.blendedR > 0 ? '+' : '';
      summarySpan.innerHTML = `<span style="color:#94a3b8;">Итог:</span> <b style="color:${si.blendedR >= 0 ? '#34d399' : '#f87171'};">${sign}${numStr(si.blendedR)}R (${sign}${pctStr}%)</b>`;
    } else if (summarySpan) {
      summarySpan.innerHTML = '';
    }

    if (commentInpEl) {
      commentInpEl.placeholder = si
        ? `Авто: ${buildSafeAutoComment(si)} (или допишите...)`
        : 'Комментарий к сделке (Tab)...';
    }

    if (activeSafeState.enabled) {
      activeSafeState.exportMode = '1row';
      setChipVal('f-res30', 'частичная фиксация');
      if (manWrap) manWrap.style.display = 'none';
      if (manInp && si) manInp.value = numStr(si.blendedR);
    } else {
      const curRes30 = getChipVal('f-res30', '');
      updateOutcomeBadgeUI(curRes30);
      if (manWrap) manWrap.style.display = (curRes30 === 'закрыл руками') ? 'flex' : 'none';
    }
    syncPresetHighlights();
  }

  function applySafePresetMode(mode) {
    if (!currentTrade) return;
    const t = currentTrade;
    const isLong = t.direction === 'LONG';
    const risk = t.riskPts;
    const dec = t.decimals || 2;

    if (activeSafeState.enabled && activeSafeState.mode === mode && mode !== 'custom' && mode !== 'off') {
      mode = 'off';
    }

    if (mode === 'off') {
      activeSafeState.enabled = false;
      activeSafeState.mode = 'off';
      const defRes30 = t.autoResult === 'Тейк' ? 'цель' : (t.autoResult === 'Стоп' ? 'стоп' : 'безубыток');
      setChipVal('f-res30', defRes30);
      const manWrap = document.getElementById('bt-manual-r-wrap');
      const manInp = document.getElementById('f-manual-r30');
      if (manWrap) manWrap.style.display = 'none';
      if (manInp) manInp.value = '';
    } else {
      activeSafeState.enabled = true;
      activeSafeState.mode = mode;
      if (mode === 'chart' && t.chartSafe) {
        activeSafeState.price = t.chartSafe.price;
        activeSafeState.price2 = t.chartSafe.secondSafe ? t.chartSafe.secondSafe.price : '';
      } else if (mode === '2r') {
        activeSafeState.price = Number((isLong ? t.entryPrice + risk * 2 : t.entryPrice - risk * 2).toFixed(dec));
        activeSafeState.price2 = '';
      } else if (mode === '3r') {
        activeSafeState.price = Number((isLong ? t.entryPrice + risk * 3 : t.entryPrice - risk * 3).toFixed(dec));
        activeSafeState.price2 = '';
      } else if (mode === 'mid') {
        activeSafeState.price = Number(((t.entryPrice + t.targetPrice) / 2).toFixed(dec));
        activeSafeState.price2 = '';
      } else if (!activeSafeState.price) {
        activeSafeState.price = t.chartSafe ? t.chartSafe.price : Number(((t.entryPrice + t.targetPrice) / 2).toFixed(dec));
      }
      const manInp = document.getElementById('f-manual-r30');
      if (manInp) manInp.value = '';
    }
    syncSafeUI();
    updatePreviewBar();
  }

  function updateCompactPropPillUI(el, val) {
    if (!el) return;
    el.dataset.val = val;
    const kind = el.dataset.prop;
    const isYes = val === 'да';
    if (kind === 'fuel') {
      el.innerHTML = isYes
        ? `<span style="color:#94a3b8;">Топливо:</span><b style="color:#e5c158;">Да</b>`
        : `<span style="opacity:0.6;">Топливо:</span><b style="opacity:0.6;font-weight:500;">нет</b>`;
    } else if (kind === 'against') {
      el.innerHTML = isYes
        ? `<span style="color:#94a3b8;">Против:</span><b style="color:#f87171;">Да</b>`
        : `<span style="opacity:0.6;">Против:</span><b style="opacity:0.6;font-weight:500;">нет</b>`;
    }
  }

  function restoreSavedTradeState() {
    if (!currentTrade) return;
    const bucket = getTplBucket();
    const saved = bucket.batchQueue.find(item => item.signature === currentTrade.signature);
    if (!saved) {
      screenshotUrl = '';
      updateScreenshotPillUI();
      return;
    }

    const ind = currentTrade.ind || {};
    screenshotUrl = saved.img || '';
    updateScreenshotPillUI();
    manualSessionOverride = saved.session || null;
    if (commentInpEl) commentInpEl.value = saved.comment || '';

    if (saved.safeStateCopy && !currentTrade.chartSafe) {
      if (saved.safeStateCopy.mode === 'chart') {
        activeSafeState.enabled = false;
        activeSafeState.mode = 'off';
      } else {
        activeSafeState = { ...saved.safeStateCopy };
      }
    }

    const setVal = (id, val) => { const el = document.getElementById(id); if (el && val != null) el.value = val; };
    setVal('f-aoi-num', saved.aoiNum);
    setVal('f-conf', saved.conf);
    setVal('f-err', saved.errVal);
    setVal('f-emotion', saved.emotionVal);

    if (!ind.detectedStructure && saved.struct_30) setVal('f-struct', saved.struct_30);
    if (!ind.candleColor && saved.candle_col) setCandleColVal(saved.candle_col);

    if (!ind.matchedZone || (saved.zone_30 && saved.zone_30 !== 'Разворотная зона')) {
      setVal('f-zone30', saved.zone_30);
      if (saved.fuel_30) {
        setChipVal('f-fuel30', saved.fuel_30);
        updateCompactPropPillUI(document.getElementById('f-fuel30'), saved.fuel_30);
      }
      if (saved.touch_30) setChipVal('f-touch30', saved.touch_30);
      if (saved.against_flow) {
        setChipVal('f-against30', saved.against_flow);
        updateCompactPropPillUI(document.getElementById('f-against30'), saved.against_flow);
      }
    }

    if (!ind.matchedSweep || (saved.sweptVal && saved.sweptVal !== 'ничего не снято')) {
      setVal('f-swept', saved.sweptVal);
      setVal('f-swept-lvl', saved.sweptVal === 'ничего не снято' ? '' : numStr(saved.sweptLvlRaw));
    }

    if (!ind.flows) {
      const setFlow = (id, dir) => {
        const p = document.getElementById(id);
        if (!p || !dir) return;
        p.dataset.dir = dir;
      };
      setFlow('f-fl-d', saved.flD);
      setFlow('f-fl-4h', saved.fl4H);
      setFlow('f-fl-1h', saved.fl1H);
      setFlow('f-fl-15m', saved.fl15M);
    }

    if (saved.trend) setChipVal('f-trend', saved.trend);
    if (saved.trend_h4) setChipVal('f-h4-trend', saved.trend_h4);
    if (saved.zone) setChipVal('f-zone', saved.zone);
    if (saved.poi_type) setChipVal('f-poi-type', saved.poi_type);
    if (!ind.modelCode && saved.form_m1) setChipVal('f-form-m1', saved.form_m1);
    if (saved.chk30) setChipVal('f-check30', saved.chk30);
    if (!ind.matchedSweep && saved.liq) setChipVal('f-liq', saved.liq);
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

    if (saved.hit_1r) setChipVal('f-hit1r', saved.hit_1r);
    if (saved.hit_3r) setChipVal('f-hit3r', saved.hit_3r);
    if (saved.resStd) setChipVal('f-res', saved.resStd);
    if (saved.res30Val && !currentTrade.chartSafe) {
      setChipVal('f-res30', saved.res30Val);
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
    const isHistFlows = !!ind.flowsHistorical;
    const missingFlows = Array.isArray(ind.flowsMissing) ? ind.flowsMissing : [];
    const hasFlowWarn = isHistFlows || missingFlows.length > 0;

    let flowWarnTip = '';
    if (missingFlows.length > 0) {
      flowWarnTip = `На ТФ ${t.timeframeLat || 'графика'} индикатор скрывает поток: ${missingFlows.join(', ')} (подставлено из кэша/графика). Проверьте вручную или на М15!`;
    } else if (isHistFlows) {
      flowWarnTip = 'Сделка далеко в истории (>90 свечей назад): проверьте потоки кликом по кнопкам Д / 4Ч / 1Ч / 15м';
    }

    const flowBolt = (ind.flows && !hasFlowWarn)
      ? BOLT
      : (hasFlowWarn ? `<span class="bt-bolt" style="color:#f59e0b;cursor:help;" title="${flowWarnTip}">⚠️</span>` : '');

    let styleEl = document.getElementById('bt-v4-compact-styles');
    if (!styleEl) {
      styleEl = document.createElement('style');
      styleEl.id = 'bt-v4-compact-styles';
      document.head.appendChild(styleEl);
    }
    styleEl.textContent = `
      #bt-modal{background:#0b0d13!important;border:1px solid rgba(229,193,88,0.26)!important;overflow-x:hidden!important}
      #bt-trade-card{background:#131621!important;border:1px solid #242a3d!important;box-shadow:inset 0 1px 0 rgba(255,255,255,0.03);margin-bottom:6px!important;padding:7px 9px!important}
      #bt-quick-copy-prices{background:#090b10!important;border-color:#1c202e!important}
      .bt-section-box{background:#141722!important;border:1px solid #23293b!important;box-shadow:0 4px 14px rgba(0,0,0,0.3),inset 0 1px 0 rgba(255,255,255,0.03);border-radius:8px!important;padding:7px 8px!important;margin-bottom:6px!important;box-sizing:border-box;max-width:100%;position:relative}
      .bt-sel,.bt-price-inp{background:#0b0d14!important;border:1px solid #262c3f!important;height:27px!important;border-radius:6px!important;transition:all 0.14s ease;color:#e2e8f0!important;box-sizing:border-box}
      .bt-sel:hover,.bt-price-inp:hover{border-color:#3f4763!important;background:#0f121b!important}
      .bt-sel.is-auto-field,.bt-price-inp.is-auto-field,#bt-cc-box.is-auto-field{border-color:rgba(229,193,88,0.44)!important;background:linear-gradient(180deg,rgba(229,193,88,0.05) 0%,rgba(11,13,20,0.96) 100%)!important;color:#ffffff!important}
      .bt-sel.is-manual-fallback{border-style:dashed!important;border-color:#333a52!important;color:#94a3b8!important}
      .bt-fallback-tag{font-size:8.5px;color:#64748b;background:#0b0d14;border:1px dashed #2e354b;padding:0 4px;border-radius:3px;font-weight:600;text-transform:none;letter-spacing:0}
      .bt-mflow{height:20px;padding:0 5px;border-radius:4px;font-size:9.5px;cursor:pointer;display:inline-flex;align-items:center;gap:2px;border:1px solid #252a3a;background:#0b0d14;color:#94a3b8;transition:0.12s}
      .bt-mflow[data-dir="вверх"]{background:rgba(16,185,129,0.13);color:#34d399;border-color:rgba(16,185,129,0.38)}
      .bt-mflow[data-dir="вниз"]{background:rgba(242,54,69,0.13);color:#f87171;border-color:rgba(242,54,69,0.38)}
      .bt-mflow.is-missing-tf{border-style:dashed!important;border-color:#f59e0b!important}
      .bt-mtog{height:20px;padding:0 5px;border-radius:4px;font-size:9.5px;cursor:pointer;display:inline-flex;align-items:center;gap:2.5px;border:1px solid #252a3a;background:#0b0d14;color:#64748b;transition:0.12s}
      .bt-mtog[data-val="да"][data-mode="gold"]{background:rgba(229,193,88,0.12);color:#f3d26b;border-color:rgba(229,193,88,0.4)}
      .bt-mtog[data-val="да"][data-mode="check"]{background:rgba(16,185,129,0.13);color:#34d399;border-color:rgba(16,185,129,0.4)}
      .bt-mtog[data-val="нет"][data-mode="check"]{background:rgba(242,54,69,0.12);color:#f87171;border-color:rgba(242,54,69,0.38)}
      .bt-prop-pill{height:26px;padding:0 5px;border-radius:6px;font-size:9.5px;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:3px;border:1px solid #222736;background:#0b0d14;color:#94a3b8;transition:0.12s;white-space:nowrap;box-sizing:border-box;width:100%;min-width:0}
      .bt-prop-pill:hover{border-color:#384057;color:#cbd5e1}
      .bt-prop-pill[data-prop="fuel"][data-val="да"]{background:rgba(229,193,88,0.07);border-color:rgba(229,193,88,0.28)}
      .bt-prop-pill[data-prop="against"][data-val="да"]{background:rgba(242,54,69,0.08);border-color:rgba(242,54,69,0.3)}
      .bt-seg-well{height:27px;background:#080a0f;border:1px solid #222736;border-radius:6px;padding:2px;display:inline-flex;align-items:center;gap:1px;box-sizing:border-box;box-shadow:inset 0 1px 3px rgba(0,0,0,0.65);max-width:100%}
      .bt-seg-well .bt-chip,.bt-seg-well .bt-safe-btn,.bt-seg-well .bt-safe-quick-trig{height:21px!important;border:1px solid transparent!important;border-radius:4px!important;background:transparent!important;color:#64748b!important;font-size:10px!important;font-weight:600!important;padding:0 4px!important;cursor:pointer;transition:0.12s;display:inline-flex;align-items:center;justify-content:center;white-space:nowrap;margin:0!important}
      .bt-seg-well .bt-chip:hover,.bt-seg-well .bt-safe-btn:hover,.bt-seg-well .bt-safe-quick-trig:hover{color:#cbd5e1!important}
      .bt-seg-well .bt-chip.active,.bt-seg-well .bt-safe-btn.active,.bt-seg-well .bt-safe-quick-trig.active{background:#22273a!important;color:#f3d26b!important;border-color:rgba(229,193,88,0.42)!important;box-shadow:0 1px 3px rgba(0,0,0,0.6)!important}
      .bt-cand-pill{border-radius:4px!important;padding:1px 4.5px!important;font-size:8.5px!important;background:#0b0d14!important;border:1px solid #282e42!important;color:#94a3b8!important;line-height:12px!important}
      .bt-cand-pill:hover{border-color:#475569!important;color:#e2e8f0!important}
      .bt-cand-pill.active{background:rgba(229,193,88,0.16)!important;color:#f3d26b!important;border-color:rgba(229,193,88,0.55)!important;font-weight:600!important}
      #bt-extra-inputs{background:#141722;border:1px solid #23293b;border-top:1px solid #1e2333;border-radius:0 0 8px 8px;padding:6px 8px 7px;margin-top:-7px!important;margin-bottom:6px!important;position:relative;z-index:2}
      #bt-inp-comment{background:#0b0d14!important;border-color:#262c3f!important;height:27px!important;font-size:11px!important}
      #bt-row-preview{max-width:100%;overflow-x:auto;opacity:0.82;transition:opacity 0.15s}
      #bt-row-preview:hover{opacity:1}
      .bt-excel-strip{display:inline-flex;border:1px solid #1c202e;border-radius:5px;overflow:hidden;background:#080a0f;cursor:default;user-select:none}
      .bt-excel-col{display:flex;align-items:center;gap:4px;min-width:max-content;border-right:1px solid #161a26;padding:2.5px 6px}
      .bt-excel-col:last-child{border-right:none}
      .bt-excel-col.is-anchor{background:rgba(229,193,88,0.1)}
      .bt-excel-col.is-anchor .bt-excel-th b{color:#f3d26b;font-weight:800}
      .bt-excel-th{color:#64748b;font-size:8.5px;white-space:nowrap}
      .bt-excel-th b{color:#94a3b8}
      .bt-excel-td{font-family:ui-monospace,SFMono-Regular,Menlo,Monaco,Consolas,monospace;font-size:9.5px;font-weight:600;white-space:nowrap}
      .bt-out-item:hover,.bt-cc-item.active-cc{background:rgba(229,193,88,0.15)!important;color:#f3d26b!important}
    `;

    let html = '';

    if (currentTrade && !currentTrade.barsLoaded) {
      html += `
        <div style="background:rgba(245,158,11,0.14);border:1px solid rgba(245,158,11,0.5);border-radius:7px;padding:5px 8px;margin-bottom:6px;font-size:10px;color:#fbbf24;display:flex;align-items:center;gap:6px;">
          <span>🔭</span>
          <span>Сделка <b>(${currentTrade.mskTime})</b> за пределами экрана — прокрутите график к ней, чтобы подгрузить свечи и Сейф!</span>
        </div>`;
    }

    if (has('aoi_num')) {
      html += `<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;"><span class="bt-lbl" style="margin:0;">№ ОБЛАСТИ / ЗОНЫ</span><input type="number" id="f-aoi-num" value="${state.aoiNum}" min="1" max="999" style="width:54px;background:#151519;color:#f3d26b;border:1px solid #282830;padding:3px;border-radius:4px;font-weight:700;text-align:center;"></div>`;
    }

    if (has('trend')) {
      const autoTrend = (state.trend === 'Тренд вверх' || state.trend === 'Тренд вниз') ? (isLong ? 'Тренд вверх' : 'Тренд вниз') : state.trend;
      html += `<div class="bt-lbl">ФАЗА РЫНКА ${BOLT}</div>` + makeChips('f-trend', [
        { val: 'Тренд вверх', label: 'Тренд вверх' },
        { val: 'Тренд вниз', label: 'Тренд вниз' },
        { val: 'Боковик', label: 'Боковик' },
        { val: 'Импульс/новости', label: 'Импульс / новости' }
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
        { val: 'Ордер-блок', label: 'Ордер-блок' },
        { val: 'Имбаланс', label: 'Имбаланс' },
        { val: 'Материнская свеча', label: 'Материнская свеча' },
        { val: 'Зона ликвидности', label: 'Зона ликвидности' },
        { val: 'Правильный откат', label: 'Правильный откат' }
      ], state.zone, 2);
    }

    if (has('poi_type')) {
      html += `<div class="bt-lbl">ТИП ТОЧКИ ВХОДА (POI)</div>` + makeChips('f-poi-type', [
        { val: 'OB', label: 'OB' },
        { val: 'Имбаланс', label: 'Имбаланс' },
        { val: '50% МС', label: '50% МС' }
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
        { val: 'SCOB', label: 'SCOB' },
        { val: 'Свип', label: 'Свип' },
        { val: 'Пин-бар (IFS)', label: 'Пин-бар' },
        { val: 'Поглощение', label: 'Поглощение' },
        { val: 'Правильный откат', label: 'Правильный откат' }
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
      const autoCandleCol = ind.candleMatched ? ind.candleColor : null;
      const st2Prob = isLong ? 'пробой ↓ · пробой ↓ · слом ↑' : 'пробой ↑ · пробой ↑ · слом ↓';
      const defStruct = ind.detectedStructure || state.struct30 || st2Prob;

      const mz = ind.matchedZone;
      const zOpts = Array.isArray(ind.zoneOptions) ? ind.zoneOptions : [];
      const msweep = ind.matchedSweep;
      const swOpts = Array.isArray(ind.sweepOptions) ? ind.sweepOptions : [];

      const defZone30 = mz?.zoneType || 'Разворотная зона';
      const defFuel30 = mz ? mz.fuel : 'нет';
      const defTouch30 = mz ? mz.touch : '1';
      const defAgainst30 = mz ? mz.against : 'нет';
      const defSwept30 = msweep ? msweep.sweptType : 'ничего не снято';
      const defSweptPrice = msweep?.price != null ? msweep.price : '';

      const zoneStatusHtml = mz
        ? `<span id="bt-zone-bolt" class="bt-bolt" style="cursor:pointer;" title="Найдено на графике: «${mz.rawText}»">⚡</span>`
        : `<span class="bt-fallback-tag" title="Зона не обнаружена рядом с точкой входа — проверьте вручную">выберите</span>`;

      const structStatusHtml = ind.detectedStructure
        ? `<span id="bt-struct-bolt" class="bt-bolt" style="cursor:pointer;" title="Определено автоматически: ${ind.detectedStructure} (нажмите для возврата)">⚡</span>`
        : `<span class="bt-fallback-tag">вручную</span>`;

      const sweepStatusHtml = msweep
        ? `<span id="bt-sweep-bolt" class="bt-bolt" style="cursor:pointer;" title="Распознан снятый уровень: «${msweep.rawText}»">⚡</span>`
        : `<span class="bt-fallback-tag">не снято</span>`;

      const shortZName = (zt) => zt.replace('Сильный покупатель · ', 'Сил.Пок ').replace('Средний покупатель · ', 'Ср.Пок ').replace('Красивый покупатель · ', 'Кр.Пок ').replace('Сильный продавец · ', 'Сил.Прод ').replace('Средний продавец · ', 'Ср.Прод ').replace('Красивый продавец · ', 'Кр.Прод ').replace('Проторговка · спит', 'Проторг.сп').replace('Проторговка', 'Проторг.').replace('Разворотная зона', 'Развор.').replace('Зона ', '');
      const shortSwName = (st) => st.replace('уровень ', '').replace('склейка ', 'скл.');

      const zoneHeaderPills = zOpts.length > 1
        ? `<div style="display:flex;gap:2px;align-items:center;overflow:hidden;">${zOpts.map((zo, idx) => `<button type="button" class="bt-cand-pill bt-zone-cand ${idx === 0 ? 'active' : ''}" data-zidx="${idx}" data-zone="${zo.zoneType}" title="Вариант с графика #${idx + 1}: ${zo.rawText}">${shortZName(zo.zoneType)}</button>`).join('')}</div>`
        : '';

      const sweepHeaderPills = swOpts.length > 0
        ? `<div style="display:flex;gap:2px;align-items:center;overflow:hidden;">${swOpts.slice(0, 3).map((so, idx) => `<button type="button" class="bt-cand-pill bt-sweep-cand ${idx === 0 ? 'active' : ''}" data-sidx="${idx}" data-swept="${so.sweptType}" data-sprice="${so.price}" title="Найдено на графике: ${so.rawText} (${numStr(so.price)})">${shortSwName(so.sweptType)}</button>`).join('')}<button type="button" class="bt-cand-pill bt-sweep-cand" data-sidx="-1" data-swept="ничего не снято" data-sprice="" title="Очистить снятие">нет</button></div>`
        : '';

      const allZonesList = ['Разворотная зона', 'Сильный покупатель · A+', 'Средний покупатель · A', 'Красивый покупатель · B', 'Сильный продавец · A+', 'Средний продавец · A', 'Красивый продавец · B', 'Зона входа · A+', 'Зона входа · A', 'Зона входа · B', 'Проторговка', 'Проторговка · спит', 'Зона 15м 0.5 · A+', 'Зона 15м 0.5 · A', 'Зона 15м 0.5 · B', 'Зона 1Ч 0.5 · A+', 'Зона 1Ч 0.5 · A', 'Зона 1Ч 0.5 · B', 'Зона 4Ч 0.5 · A+', 'Зона 4Ч 0.5 · A', 'Зона 4Ч 0.5 · B', 'Зона Д 0.5 · A+', 'Зона Д 0.5 · A', 'Зона Д 0.5 · B', 'Зона 15м 0.5', 'Зона 1Ч 0.5', 'Зона 4Ч 0.5', 'Зона Д 0.5'];
      const detectedZoneSet = new Set(zOpts.map(z => z.zoneType));
      const zoneSelectOptions = [
        ...zOpts.map(z => ({ val: z.zoneType, label: `⚡ ${z.zoneType}${z.fuel === 'да' ? ' · топл.' : ''}` })),
        ...allZonesList.filter(z => !detectedZoneSet.has(z)).map(z => ({ val: z, label: z }))
      ];

      const allSweepsList = ['приманка', 'полка', 'уровень 15м', 'уровень 1Ч', 'уровень 4Ч', 'уровень Д', 'склейка 15м/1Ч', 'склейка Д · 4Ч', 'ничего не снято'];
      const detectedSweepSet = new Set(swOpts.map(s => s.sweptType));
      const sweepSelectOptions = [
        ...swOpts.map(s => ({ val: s.sweptType, label: `⚡ ${s.sweptType} (${numStr(s.price)})` })),
        ...allSweepsList.filter(s => !detectedSweepSet.has(s)).map(s => ({ val: s, label: s }))
      ];

      const cInfo = CANDLE_COLORS[defCandleCol] || CANDLE_COLORS['Лонг 3'];
      const ccMenuItemsHtml = Object.entries(CANDLE_COLORS).map(([k, v]) => {
        const isThisAuto = (autoCandleCol === k);
        const autoTag = isThisAuto ? `<span style="color:#e5c158;font-weight:700;font-size:9.5px;margin-left:4px;">⚡ авто</span>` : '';
        return `
        <div class="bt-cc-item ${k === defCandleCol ? 'active-cc' : ''}" data-val="${k}" style="display:flex;align-items:center;justify-content:space-between;padding:4px 7px;border-radius:4px;cursor:pointer;font-size:10.5px;color:#f4f4f5;">
          <div style="display:flex;align-items:center;gap:6px;"><span style="width:11px;height:11px;border-radius:3px;background:${v.hex};display:inline-block;"></span><span style="font-weight:600;">${k}</span>${autoTag}</div>
          <span style="color:#9ca3af;font-size:9.5px;">${v.hint}</span>
        </div>`;
      }).join('');

      const missCls = (tfName) => missingFlows.includes(tfName) ? 'is-missing-tf' : '';
      const missTip = (tfName) => missingFlows.includes(tfName)
        ? `⚠️ На ТФ ${t.timeframeLat} индикатор скрывает поток ${tfName} (взят кэш/график). Нажмите для ручной смены`
        : `Поток ${tfName} (нажмите для смены)`;

      const fuelInitHtml = defFuel30 === 'да'
        ? `<span style="color:#94a3b8;">Топливо:</span><b style="color:#e5c158;">Да</b>`
        : `<span style="opacity:0.6;">Топливо:</span><b style="opacity:0.6;font-weight:500;">нет</b>`;
      const againstInitHtml = defAgainst30 === 'да'
        ? `<span style="color:#94a3b8;">Против:</span><b style="color:#f87171;">Да</b>`
        : `<span style="opacity:0.6;">Против:</span><b style="opacity:0.6;font-weight:500;">нет</b>`;

      html += `
        <div style="background:#131621;border:1px solid ${hasFlowWarn ? 'rgba(245,158,11,0.42)' : '#242a3d'};border-top:none;border-radius:0 0 8px 8px;padding:4px 8px 5px;margin-top:-9px;margin-bottom:7px;display:flex;justify-content:space-between;align-items:center;gap:4px;white-space:nowrap;">
          <div style="display:flex;align-items:center;gap:3px;">
            ${flowBolt}
            <span style="font-size:8.5px;color:#64748b;font-weight:700;margin-right:1px;">ПОТОК:</span>
            <button type="button" class="bt-mflow bt-flow-pill ${missCls('Д')}" id="f-fl-d" data-tf="Д" data-dir="${state.flowD}" title="${missTip('Д')}"></button>
            <button type="button" class="bt-mflow bt-flow-pill ${missCls('4Ч')}" id="f-fl-4h" data-tf="4Ч" data-dir="${state.flow4H}" title="${missTip('4Ч')}"></button>
            <button type="button" class="bt-mflow bt-flow-pill ${missCls('1Ч')}" id="f-fl-1h" data-tf="1Ч" data-dir="${state.flow1H}" title="${missTip('1Ч')}"></button>
            <button type="button" class="bt-mflow bt-flow-pill ${missCls('15м')}" id="f-fl-15m" data-tf="15м" data-dir="${state.flow15M}" title="${missTip('15м')}"></button>
          </div>
          <div style="display:flex;align-items:center;gap:3px;">
            <button type="button" class="bt-mtog bt-switch" id="f-hit1r" data-val="${t.reached1R ? 'да' : 'нет'}" data-mode="gold" style="width:auto;height:20px;margin:0;padding:0 5px;" title="Дошла до 1:1 (нажмите для смены)"></button>
            <button type="button" class="bt-mtog bt-switch" id="f-hit3r" data-val="${t.reached3R ? 'да' : 'нет'}" data-mode="gold" style="width:auto;height:20px;margin:0;padding:0 5px;" title="Дошла до 1:3 (нажмите для смены)"></button>
            <button type="button" class="bt-mtog bt-switch" id="f-check30" data-val="да" data-mode="check" style="width:auto;height:20px;margin:0;padding:0 5px;" title="Чек-лист 7/7 (нажмите для смены)"></button>
          </div>
        </div>

        <div class="bt-section-box">
          <div style="display:grid;grid-template-columns:1.18fr 0.82fr;gap:6px;align-items:center;margin-bottom:3px;">
            <div style="display:flex;align-items:center;justify-content:space-between;gap:3px;min-width:0;overflow:hidden;">
              <div class="bt-lbl" style="margin:0;flex-shrink:0;">ЗОНА ${zoneStatusHtml}</div>
              ${zoneHeaderPills}
            </div>
            <div class="bt-lbl" style="margin:0;">СТРУКТУРА ${structStatusHtml}</div>
          </div>
          <div style="display:grid;grid-template-columns:1.14fr 0.86fr;gap:6px;margin-bottom:6px;">
            <div style="min-width:0;">
              ${makeSelect('f-zone30', zoneSelectOptions, defZone30, '0px', mz ? 'is-auto-field' : 'is-manual-fallback')}
            </div>
            <div style="min-width:0;">
              ${makeSelect('f-struct', ['пробой ↑', 'пробой ↑ · пробой ↑', 'пробой ↑ · пробой ↑ · слом ↓', 'пробой ↓', 'пробой ↓ · пробой ↓', 'пробой ↓ · пробой ↓ · слом ↑', 'слом ↑', 'слом ↓', 'ложный пробой ↑ · слом ↓', 'ложный пробой ↓ · слом ↑', 'внутр. слом ↑', 'внутр. слом ↓', 'смена ↑ на 4Ч', 'смена ↓ на 4Ч'], defStruct, '0px', ind.detectedStructure ? 'is-auto-field' : '')}
            </div>
          </div>

          <div style="display:grid;grid-template-columns:1.08fr 0.92fr 0.98fr 0.92fr;gap:4px;align-items:center;margin-bottom:7px;">
            <div id="bt-cc-wrap" style="position:relative;min-width:0;">
              <input type="hidden" id="f-candle-col" value="${defCandleCol}">
              <div id="bt-cc-box" class="${ind.candleMatched ? 'is-auto-field' : ''}" style="display:flex;height:26px;background:#0b0d14;border:1px solid #262c3f;border-radius:6px;overflow:hidden;box-sizing:border-box;transition:all 0.14s ease;">
                <button type="button" id="bt-cc-cycle-btn" title="Свеча входа (кликните для быстрой смены 3 → 2 → 1)" style="flex:1;background:transparent;border:none;color:#f4f4f5;padding:0 5px;font-size:10px;font-weight:600;cursor:pointer;display:flex;align-items:center;gap:3.5px;white-space:nowrap;min-width:0;">
                  <span id="bt-cc-dot" style="width:9px;height:9px;border-radius:2px;background:${cInfo.hex};flex-shrink:0;box-shadow:0 0 0 1px rgba(255,255,255,0.22);"></span>
                  <span id="bt-cc-label">${defCandleCol}</span>
                  <span id="bt-cc-bolt" class="bt-bolt" style="display:${ind.candleMatched ? 'inline' : 'none'};">⚡</span>
                </button>
                <button type="button" id="bt-cc-btn" title="Открыть список всех цветов свечей" style="background:rgba(255,255,255,0.03);border:none;border-left:1px solid #232838;color:#94a3b8;padding:0 3.5px;font-size:9px;cursor:pointer;flex-shrink:0;">▾</button>
              </div>
              <div id="bt-cc-menu" style="display:none;position:absolute;top:28px;left:0;width:185px;background:#181a24;border:1px solid #e5c158;border-radius:6px;box-shadow:0 12px 28px rgba(0,0,0,0.95);z-index:300;padding:3px;">${ccMenuItemsHtml}</div>
            </div>
            <button type="button" class="bt-prop-pill bt-switch" id="f-fuel30" data-prop="fuel" data-val="${defFuel30}" style="margin:0;" title="Наличие надписи «топливо» у зоны входа">${fuelInitHtml}</button>
            <div class="bt-group bt-seg-well" id="f-touch30" style="width:100%;height:26px;justify-content:space-between;" title="Касание зоны в момент входа">
              <span style="font-size:8.5px;color:#64748b;padding:0 2.5px;font-weight:700;">КАС:</span>
              <button type="button" class="bt-chip ${defTouch30 === '1' ? 'active' : ''}" style="flex:1;padding:0 2px!important;" data-val="1">1</button>
              <button type="button" class="bt-chip ${defTouch30 === '2' ? 'active' : ''}" style="flex:1;padding:0 2px!important;" data-val="2">2</button>
              <button type="button" class="bt-chip ${defTouch30 === '3 и больше' ? 'active' : ''}" style="flex:1;padding:0 2px!important;" data-val="3 и больше">3+</button>
            </div>
            <button type="button" class="bt-prop-pill bt-switch" id="f-against30" data-prop="against" data-val="${defAgainst30}" style="margin:0;" title="Наличие надписи «против потока» у зоны входа">${againstInitHtml}</button>
          </div>

          <div style="padding-top:5px;border-top:1px solid rgba(255,255,255,0.04);">
            <div style="display:grid;grid-template-columns:1.14fr 0.86fr;gap:6px;align-items:center;margin-bottom:3px;">
              <div style="display:flex;align-items:center;justify-content:space-between;gap:4px;min-width:0;overflow:hidden;">
                <div class="bt-lbl" style="margin:0;flex-shrink:0;">СНЯТИЕ ${sweepStatusHtml}</div>
                ${sweepHeaderPills}
              </div>
              <div class="bt-lbl" style="margin:0;">УРОВЕНЬ СНЯТИЯ ${msweep ? BOLT : ''}</div>
            </div>
            <div style="display:grid;grid-template-columns:1.14fr 0.86fr;gap:6px;align-items:center;">
              <div style="min-width:0;">
                ${makeSelect('f-swept', sweepSelectOptions, defSwept30, '0px', msweep ? 'is-auto-field' : 'is-manual-fallback')}
              </div>
              <div style="position:relative;display:flex;align-items:center;">
                <input type="text" id="f-swept-lvl" class="bt-price-inp ${msweep ? 'is-auto-field' : ''}" value="${defSweptPrice !== '' ? numStr(defSweptPrice) : ''}" placeholder="Без уровня..." style="width:100%;padding:0 22px 0 8px;font-size:11px;font-weight:600;outline:none;">
                <span id="bt-swept-price-bolt" style="position:absolute;right:7px;font-size:10px;color:#e5c158;cursor:pointer;display:${msweep ? 'inline' : 'none'};" title="Цена подставлена автоматически с графика">⚡</span>
              </div>
            </div>
          </div>
        </div>`;
    }

    if (has('liq') || has('conf') || has('react')) {
      const autoLiq = ind.matchedSweep ? 'да' : state.liq;
      html += `<div style="display:grid;grid-template-columns:1.1fr 1.9fr;gap:7px;">`;
      if (has('liq')) html += `<div><div class="bt-lbl">СНЯТИЕ ЛИКВ. ${ind.matchedSweep ? BOLT : ''}</div>${makeSwitch('f-liq', autoLiq, 'gold')}</div>`;
      if (has('conf')) {
        html += `<div><div class="bt-lbl">ПОДТВЕРЖДЕНИЕ</div>${makeSelect('f-conf', [
          { val: 'Слом структуры', label: 'Слом структуры' },
          { val: 'Реакция от зоны', label: 'Реакция от зоны' },
          { val: 'Поглощение', label: 'Поглощение' },
          { val: 'Снятие ликвидности', label: 'Снятие ликвидности' },
          { val: 'Без подтверждения', label: 'Без подтверждения' }
        ], state.conf)}</div>`;
      } else if (has('react')) {
        html += `<div><div class="bt-lbl">РЕАКЦИЯ ОТ ЗОНЫ</div>${makeSwitch('f-react', 'да', 'gold')}</div>`;
      }
      html += `</div>`;
    }

    if (has('res_30')) {
      const defRes30 = t.autoResult === 'Тейк' ? 'цель' : (t.autoResult === 'Стоп' ? 'стоп' : 'безубыток');
      const cs = t.chartSafe;
      const has2Safes = !!(cs && cs.secondSafe);
      const chartSafeSegBtn = cs
        ? `<button type="button" class="bt-safe-quick-trig" style="flex:1.25;" data-smode="chart" title="Взять линии сейфа с графика">${has2Safes ? '⚡ 2 сейфа' : `⚡ ${numStr(cs.price)}`}</button>`
        : '';
      const hasCommentOrScreen = has('comment') || has('comment_30') || has('screenshot');

      html += `
        <div class="bt-section-box" id="bt-outcome-card" style="padding:6px 8px;${hasCommentOrScreen ? 'border-radius:8px 8px 0 0!important;' : ''}">
          <div style="display:none;">` +
            makeChips('f-res30', [
              { val: 'цель', label: 'Цель' },
              { val: 'стоп', label: 'Стоп' },
              { val: 'безубыток', label: 'БУ' },
              { val: 'частичная фиксация', label: 'Сейф' },
              { val: 'закрыл руками', label: 'Руками' }
            ], defRes30, 5) +
          `</div>
          <div style="display:grid;grid-template-columns:1.14fr 0.86fr;gap:6px;align-items:center;margin-bottom:3px;">
            <span class="bt-lbl" id="bt-outcome-main-lbl" style="margin:0;">ИСХОД СДЕЛКИ ${BOLT}</span>
            <span class="bt-lbl" style="margin:0;color:#cbd5e1;">${SHIELD_SVG}<span>ЧАСТИЧНЫЙ СЕЙФ</span> ${cs ? BOLT : ''}</span>
          </div>
          <div style="display:grid;grid-template-columns:1.14fr 0.86fr;gap:6px;align-items:center;">
            <div id="bt-outcome-wrap" style="position:relative;min-width:0;">
              <button type="button" id="bt-outcome-cycle-btn" title="Определено автоматически по свечам на графике. Нажмите, чтобы изменить вручную" style="width:100%;height:27px;padding:0 8px;border-radius:6px;border:1px solid rgba(16,185,129,0.42);background:rgba(16,185,129,0.14);color:#34d399;font-size:11px;cursor:pointer;display:flex;align-items:center;justify-content:space-between;gap:6px;box-sizing:border-box;transition:0.12s;"></button>
              <div id="bt-outcome-menu" style="display:none;position:absolute;top:29px;left:0;width:100%;background:#181a24;border:1px solid #e5c158;border-radius:7px;box-shadow:0 12px 28px rgba(0,0,0,0.95);z-index:300;padding:3px;box-sizing:border-box;">
                <div class="bt-out-item" data-oval="цель" style="display:flex;justify-content:space-between;align-items:center;padding:5px 8px;border-radius:4px;cursor:pointer;font-size:11px;color:#34d399;font-weight:600;"><span>🎯 Цель (Тейк)</span></div>
                <div class="bt-out-item" data-oval="стоп" style="display:flex;justify-content:space-between;align-items:center;padding:5px 8px;border-radius:4px;cursor:pointer;font-size:11px;color:#f87171;font-weight:600;"><span>🛑 Стоп-лосс</span></div>
                <div class="bt-out-item" data-oval="безубыток" style="display:flex;justify-content:space-between;align-items:center;padding:5px 8px;border-radius:4px;cursor:pointer;font-size:11px;color:#e2e8f0;font-weight:600;"><span>⚖️ Безубыток (0R)</span></div>
                <div class="bt-out-item" data-oval="закрыл руками" style="display:flex;justify-content:space-between;align-items:center;padding:5px 8px;border-radius:4px;cursor:pointer;font-size:11px;color:#cbd5e1;font-weight:600;"><span>✋ Закрыл руками</span></div>
              </div>
            </div>
            <div class="bt-seg-well" style="width:100%;justify-content:space-between;">
              <button type="button" class="bt-safe-quick-trig" style="flex:1;" data-smode="off" title="Без сейфа">Нет</button>
              ${chartSafeSegBtn}
              <button type="button" class="bt-safe-quick-trig" style="flex:1;" data-smode="2r" title="Сейф на 1:2">1:2</button>
              <button type="button" class="bt-safe-quick-trig" style="flex:1;" data-smode="3r" title="Сейф на 1:3">1:3</button>
              ${!cs ? `<button type="button" class="bt-safe-quick-trig" style="flex:1;" data-smode="mid" title="Сейф на 50% пути к тейку">50%</button>` : ''}
            </div>
          </div>

          <div id="bt-safe-details" style="display:none;margin-top:6px;padding-top:6px;border-top:1px solid rgba(255,255,255,0.05);">
            <div style="display:flex;align-items:center;justify-content:space-between;gap:6px;white-space:nowrap;">
              <div style="display:flex;align-items:center;gap:4px;min-width:0;">
                <input type="text" id="bt-safe-price-inp" class="bt-price-inp" placeholder="Цена 1..." title="Цена 1-го сейфа" style="width:72px;height:25px!important;padding:0 5px;font-size:10.5px;font-weight:700;outline:none;text-align:center;">
                <div id="bt-safe-price2-wrap" style="display:none;align-items:center;gap:4px;">
                  <input type="text" id="bt-safe-price2-inp" class="bt-price-inp is-auto-field" placeholder="Цена 2..." title="Цена 2-го сейфа" style="width:72px;height:25px!important;padding:0 5px;font-size:10.5px;font-weight:700;outline:none;text-align:center;">
                </div>
                <div class="bt-seg-well" id="bt-safe-share-well" style="height:25px;flex-shrink:0;" title="Распределение объёма на сейфах">
                  <button type="button" class="bt-safe-btn bt-safe-share-btn" style="padding:0 6px!important;" data-sshare="0.5">1/2</button>
                  <button type="button" class="bt-safe-btn bt-safe-share-btn" style="padding:0 6px!important;" data-sshare="0.33">1/3</button>
                </div>
              </div>
              <span id="bt-safe-calc-hint" style="font-size:10.5px;color:#cbd5e1;white-space:nowrap;flex-shrink:0;"></span>
            </div>
          </div>

          <div id="bt-manual-r-wrap" style="display:none;justify-content:space-between;align-items:center;background:#0b0d14;padding:4px 8px;border-radius:5px;border:1px solid #3d3522;margin-top:5px;">
            <span style="font-size:10px;color:#f3d26b;">Фактический итог в R (при закрытии руками):</span>
            <input type="text" id="f-manual-r30" placeholder="0,0" style="width:60px;background:#07080c;color:#ffffff;border:1px solid #4a3f23;border-radius:4px;padding:2px 6px;font-size:11px;font-weight:700;text-align:center;outline:none;">
          </div>
        </div>`;
    } else if (has('res_off')) {
      const defOff = t.autoResult === 'Безубыток' ? 'Сейф/БУ' : t.autoResult;
      html += `<div class="bt-lbl">РЕЗУЛЬТАТ СДЕЛКИ ${BOLT}</div>` + makeChips('f-res', [
        { val: 'Тейк', label: 'Тейк', tone: 'green' },
        { val: 'Стоп', label: 'Стоп', tone: 'red' },
        { val: 'Сейф/БУ', label: 'Сейф/БУ' },
        { val: 'Бу', label: 'Бу' }
      ], defOff, 4);
    } else if (has('result') || has('r_out')) {
      html += `<div class="bt-lbl">РЕЗУЛЬТАТ СДЕЛКИ ${BOLT}</div>` + makeChips('f-res', [
        { val: 'Тейк', label: 'Тейк', tone: 'green' },
        { val: 'Стоп', label: 'Стоп', tone: 'red' },
        { val: 'Безубыток', label: 'БУ' },
        { val: 'Пропуск', label: 'Пропуск' }
      ], t.autoResult, 4);
    }

    if (has('error') || has('error_poi')) {
      const errList = has('error_poi')
        ? ['—', 'Без снятия', 'Без реакции', 'Один фактор', 'Область на M15', 'Вход по касанию', 'Зона отработана', 'R:R ниже нормы', 'Догонял цену', 'Другое']
        : ['—', 'Ранний вход', 'Нет подтверждения', 'Вход без снятия ликвидности', 'Стоп слишком узкий', 'Стоп слишком широкий', 'Ранний выход', 'Пересидел цель', 'Нарушен минимальный R:R', 'Вход вне сессии', 'Другое'];
      html += `<div style="display:flex;justify-content:space-between;align-items:center;"><div class="bt-lbl">ПРИЧИНА ОШИБКИ ${BOLT}</div><span id="bt-err-reason" style="font-size:10px;color:#f3d26b;font-weight:600;margin-bottom:4px;"></span></div>${makeSelect('f-err', errList, '—')}`;
    }

    dynFieldsEl.innerHTML = html;

    const extraInputsEl = document.getElementById('bt-extra-inputs');
    if (extraInputsEl) {
      extraInputsEl.style.display = (has('comment') || has('comment_30') || has('screenshot')) ? 'flex' : 'none';
    }

    if (tradeCardEl) {
      tradeCardEl.style.borderRadius = hasAnyFlowOr30 ? '8px 8px 0 0' : '8px';
    }

    const zone5Btn = dynFieldsEl.querySelector('#f-zone .bt-chip:nth-child(5)');
    if (zone5Btn) zone5Btn.style.gridColumn = 'span 2';
    const form5Btn = dynFieldsEl.querySelector('#f-form-m1 .bt-chip:nth-child(5)');
    if (form5Btn) form5Btn.style.gridColumn = 'span 2';

    const stBoltEl = document.getElementById('bt-struct-bolt');
    if (stBoltEl && ind.detectedStructure) {
      stBoltEl.onclick = () => {
        const stSel = document.getElementById('f-struct');
        if (stSel) {
          stSel.value = ind.detectedStructure;
          saveState({ struct30: ind.detectedStructure });
          syncPresetHighlights();
          updatePreviewBar();
        }
      };
    }

    const outcomeCycleBtn = document.getElementById('bt-outcome-cycle-btn');
    const outcomeMenu = document.getElementById('bt-outcome-menu');
    if (outcomeCycleBtn && outcomeMenu) {
      outcomeCycleBtn.onclick = e => {
        e.stopPropagation();
        outcomeMenu.style.display = outcomeMenu.style.display === 'block' ? 'none' : 'block';
      };
      outcomeMenu.querySelectorAll('.bt-out-item').forEach(item => {
        item.onclick = e => {
          e.stopPropagation();
          const chosen = item.dataset.oval;
          outcomeMenu.style.display = 'none';
          if (activeSafeState.enabled && chosen !== 'закрыл руками') {
            activeSafeState.remOutcome = chosen;
            const manInp = document.getElementById('f-manual-r30');
            if (manInp) manInp.value = '';
            syncSafeUI();
            updatePreviewBar();
          } else {
            activeSafeState.enabled = false;
            activeSafeState.mode = 'off';
            setChipVal('f-res30', chosen);
            syncSafeUI();
            updatePreviewBar();
          }
        };
      });
    }

    dynFieldsEl.querySelectorAll('.bt-zone-cand').forEach(btn => {
      btn.onclick = () => {
        const idx = parseInt(btn.dataset.zidx, 10);
        const zObj = currentTrade?.ind?.zoneOptions?.[idx];
        if (!zObj) return;
        const zSel = document.getElementById('f-zone30');
        if (zSel) zSel.value = zObj.zoneType;
        updateCompactPropPillUI(document.getElementById('f-fuel30'), zObj.fuel);
        setChipVal('f-touch30', zObj.touch);
        updateCompactPropPillUI(document.getElementById('f-against30'), zObj.against);
        syncPresetHighlights();
        updatePreviewBar();
      };
    });

    dynFieldsEl.querySelectorAll('.bt-sweep-cand').forEach(btn => {
      btn.onclick = () => {
        const swType = btn.dataset.swept;
        const swPrice = btn.dataset.sprice;
        const swSel = document.getElementById('f-swept');
        const swInp = document.getElementById('f-swept-lvl');
        if (swSel && swType) swSel.value = swType;
        if (swInp) swInp.value = (swType === 'ничего не снято' || !swPrice) ? '' : numStr(swPrice);
        syncPresetHighlights();
        updatePreviewBar();
      };
    });

    dynFieldsEl.querySelectorAll('.bt-safe-quick-trig').forEach(btn => {
      btn.onclick = () => applySafePresetMode(btn.dataset.smode);
    });
    dynFieldsEl.querySelectorAll('.bt-safe-share-btn').forEach(btn => {
      btn.onclick = () => {
        activeSafeState.share = parseFloat(btn.dataset.sshare) || 0.5;
        saveState({ safeShare: activeSafeState.share });
        const manInp = document.getElementById('f-manual-r30');
        if (manInp) manInp.value = '';
        syncSafeUI();
        updatePreviewBar();
      };
    });
    dynFieldsEl.querySelectorAll('.bt-safe-exp-btn').forEach(btn => {
      btn.onclick = () => {
        activeSafeState.exportMode = btn.dataset.sexp || '1row';
        saveState({ safeExportMode: activeSafeState.exportMode });
        syncSafeUI();
        updatePreviewBar();
      };
    });
    const safePriceInp = document.getElementById('bt-safe-price-inp');
    if (safePriceInp) {
      safePriceInp.addEventListener('input', () => {
        activeSafeState.mode = 'custom';
        activeSafeState.price = safePriceInp.value;
        const manInp = document.getElementById('f-manual-r30');
        if (manInp) manInp.value = '';
        syncSafeUI();
        updatePreviewBar();
      });
    }
    const safePrice2Inp = document.getElementById('bt-safe-price2-inp');
    if (safePrice2Inp) {
      safePrice2Inp.addEventListener('input', () => {
        activeSafeState.price2 = safePrice2Inp.value;
        const manInp = document.getElementById('f-manual-r30');
        if (manInp) manInp.value = '';
        syncSafeUI();
        updatePreviewBar();
      });
    }

    dynFieldsEl.querySelectorAll('.bt-flow-pill').forEach(pill => {
      pill.onclick = () => {
        const order = ['вверх', 'вниз', '—'];
        const curr = pill.dataset.dir || 'вверх';
        const next = order[(order.indexOf(curr) + 1) % order.length];
        pill.dataset.dir = next;
        const valSpan = pill.querySelector('.bt-flow-val');
        if (valSpan) valSpan.innerText = next === 'вверх' ? '↑ вверх' : (next === 'вниз' ? '↓ вниз' : '—');
        if (pill.id === 'f-fl-d') saveState({ flowD: next });
        if (pill.id === 'f-fl-4h') saveState({ flow4H: next });
        if (pill.id === 'f-fl-1h') saveState({ flow1H: next });
        if (pill.id === 'f-fl-15m') saveState({ flow15M: next });
        updateAutoStripSummaryUI();
        updatePreviewBar();
      };
    });

    dynFieldsEl.querySelectorAll('.bt-switch').forEach(sw => {
      sw.onclick = () => {
        const next = sw.dataset.val === 'да' ? 'нет' : 'да';
        if (sw.classList.contains('bt-prop-pill')) {
          updateCompactPropPillUI(sw, next);
        } else {
          updateSwitchUI(sw, next);
        }
        if (sw.id === 'f-liq') {
          saveState({ liq: next });
          if (next === 'нет' && getChipVal('f-res') === 'Пропуск') {
            const confEl = document.getElementById('f-conf');
            if (confEl) confEl.value = 'Без подтверждения';
          }
          autoSuggestError();
        }
        updateAutoStripSummaryUI();
        updatePreviewBar();
      };
    });

    const ccCycleBtn = document.getElementById('bt-cc-cycle-btn');
    const ccBtn = document.getElementById('bt-cc-btn');
    const ccMenu = document.getElementById('bt-cc-menu');
    if (ccCycleBtn) {
      ccCycleBtn.onclick = e => {
        e.stopPropagation();
        const cur = document.getElementById('f-candle-col')?.value || (isLong ? 'Лонг 3' : 'Шорт 3');
        const cycle = cur.startsWith('Шорт') ? ['Шорт 3', 'Шорт 2', 'Шорт 1'] : ['Лонг 3', 'Лонг 2', 'Лонг 1'];
        const next = cycle[(cycle.indexOf(cur) + 1) % cycle.length];
        setCandleColVal(next);
        updatePreviewBar();
      };
    }
    if (ccBtn && ccMenu) {
      ccBtn.onclick = e => { e.stopPropagation(); ccMenu.style.display = ccMenu.style.display === 'block' ? 'none' : 'block'; };
      ccMenu.querySelectorAll('.bt-cc-item').forEach(item => {
        item.onclick = () => { setCandleColVal(item.dataset.val); ccMenu.style.display = 'none'; updatePreviewBar(); };
      });
    }

    const swPriceBolt = document.getElementById('bt-swept-price-bolt');
    if (swPriceBolt) {
      swPriceBolt.onclick = () => {
        const inp = document.getElementById('f-swept-lvl');
        const swSelVal = document.getElementById('f-swept')?.value;
        const matchOpt = currentTrade?.ind?.sweepOptions?.find(o => o.sweptType === swSelVal);
        const fallbackPrice = matchOpt?.price ?? currentTrade?.ind?.matchedSweep?.price ?? currentTrade?.entryExtreme ?? '';
        if (inp && fallbackPrice !== '') {
          inp.value = numStr(fallbackPrice);
          syncPresetHighlights();
          updatePreviewBar();
        }
      };
    }

    dynFieldsEl.querySelectorAll('select, input').forEach(el => {
      if (el.id === 'bt-safe-price-inp' || el.id === 'bt-safe-price2-inp') return;
      el.addEventListener('change', () => {
        if (el.id === 'f-swept') {
          const lvlInp = document.getElementById('f-swept-lvl');
          if (lvlInp) {
            const matchOpt = currentTrade?.ind?.sweepOptions?.find(o => o.sweptType === el.value);
            const fallbackPrice = matchOpt?.price ?? currentTrade?.ind?.matchedSweep?.price ?? currentTrade?.entryExtreme ?? '';
            lvlInp.value = el.value === 'ничего не снято' ? '' : (fallbackPrice !== '' ? numStr(fallbackPrice) : '');
          }
          syncPresetHighlights();
        }
        if (el.id === 'f-zone30') {
          const matchZ = currentTrade?.ind?.zoneOptions?.find(o => o.zoneType === el.value);
          if (matchZ) {
            updateCompactPropPillUI(document.getElementById('f-fuel30'), matchZ.fuel);
            setChipVal('f-touch30', matchZ.touch);
            updateCompactPropPillUI(document.getElementById('f-against30'), matchZ.against);
          }
          syncPresetHighlights();
        }
        if (el.id === 'f-struct') { saveState({ struct30: el.value }); syncPresetHighlights(); }
        if (el.id === 'f-emotion') saveState({ emotion: el.value });
        if (el.id === 'f-conf') { saveState({ conf: el.value }); autoSuggestError(); }
        updatePreviewBar();
      });
      el.addEventListener('input', () => {
        if (el.id === 'f-swept-lvl') syncPresetHighlights();
        updatePreviewBar();
      });
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

    initTradeSafeDefault(currentTrade);
    autoSuggestError();
    restoreSavedTradeState();
    setCandleColVal(document.getElementById('f-candle-col')?.value || defCandleCol);
    syncSafeUI();
    syncPresetHighlights();
    updateAutoStripSummaryUI();
    updatePreviewBar();
  }

  function syncLiveChartChanges() {
    if (modalEl.style.display !== 'block') return;
    const prevId = currentTrade?.id;
    const prevSig = currentTrade?.signature;

    if (!scanChartTrades(true, false) || !currentTrade) return;

    if (currentTrade.id !== prevId) {
      manualSessionOverride = null;
      screenshotUrl = '';
      if (commentInpEl) commentInpEl.value = '';
      renderMainFields();
      updateHeaderUI();
      pulseTradeCard('#089981');
      return;
    }

    if (currentTrade.signature !== prevSig) {
      const savedComment = commentInpEl ? commentInpEl.value : '';
      const savedScreen = screenshotUrl;
      renderMainFields();
      if (commentInpEl && savedComment && !commentInpEl.value) commentInpEl.value = savedComment;
      if (savedScreen && !screenshotUrl) screenshotUrl = savedScreen;
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

  let toolbarRo = null;
  function initToolbarObservers() {
    mountNativeToolbarButton();
    if (typeof ResizeObserver !== 'function') return;

    let rafId = 0;
    const scheduleReposition = () => {
      if (rafId) return;
      rafId = requestAnimationFrame(() => {
        rafId = 0;
        mountNativeToolbarButton();
      });
    };

    window.addEventListener('resize', scheduleReposition, { passive: true });

    const tryAttach = () => {
      const toolbarEl = document.getElementById('drawing-toolbar');
      const favBtn = document.getElementById('drawing-toolbar-favorite-drawings');
      if (!toolbarEl || !favBtn) return false;

      if (!toolbarRo) {
        toolbarRo = new ResizeObserver(scheduleReposition);
        toolbarRo.observe(toolbarEl);
        const spacerEl = toolbarEl.querySelector('[class*="fill-"]');
        if (spacerEl) toolbarRo.observe(spacerEl);

        const mo = new MutationObserver(scheduleReposition);
        mo.observe(toolbarEl, { childList: true, subtree: true });
      }
      scheduleReposition();
      return true;
    };

    if (!tryAttach()) {
      const bootMo = new MutationObserver(() => {
        if (tryAttach()) bootMo.disconnect();
      });
      bootMo.observe(document.body, { childList: true, subtree: true });
    }
  }

  initToolbarObservers();

  window.addEventListener('pointerdown', () => {
    const model = getTvModel();
    if (!model) return;
    syncRadarOnBarClose(model);
    try {
      const hov = typeof model.hoveredSource === 'function' ? model.hoveredSource() : model._hoveredSource;
      if (hov && isSupportedTool(hov)) lastHoveredToolId = getToolId(hov);
    } catch (e) {}
  }, { capture: true, passive: true });

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

  function getTradeShortOutcomeLabel(item) {
    if (!item || typeof item !== 'object') return 'бу';
    if (item.safeInfo || item.res30Val === 'частичная фиксация' || item.activeResult === 'частичная фиксация') {
      const rem = item.safeInfo?.remOut || (item.rOut > 1.5 ? 'цель' : (item.rOut < 0.9 ? 'стоп' : 'бу'));
      if (rem === 'цель') return 'сейф/тейк';
      if (rem === 'стоп') return 'сейф/стоп';
      return 'сейф/бу';
    }
    const res = String(item.activeResult || item.res30Val || item.resStd || '').toLowerCase();
    if (res === 'тейк' || res === 'цель') return 'тейк';
    if (res === 'стоп') return 'стоп';
    if (res.includes('сейф')) return 'сейф/бу';
    if (res.includes('руками')) return 'руками';
    return 'бу';
  }

  function computeBatchStats() {
    const bucket = getTplBucket();
    let wins = 0, losses = 0, totalR = 0;
    const outcomesList = [];

    for (const item of bucket.batchQueue) {
      if (!item || typeof item !== 'object') continue;
      outcomesList.push(getTradeShortOutcomeLabel(item));
      const res = item.activeResult || item.result;
      const rrVal = Number(item.riskReward ?? item.rr ?? 0);
      if (typeof item.rOut === 'number' && !isNaN(item.rOut)) {
        totalR += item.rOut;
        if (item.rOut > 0) wins++;
        else if (item.rOut < 0) losses++;
      } else if (res === 'Тейк' || res === 'цель') { wins++; totalR += rrVal; }
      else if (res === 'Стоп' || res === 'стоп') { losses++; totalR -= 1; }
    }

    const count = bucket.batchQueue.length;
    const riskPct = Number(state.riskPctPer1R) || 0.3;
    const totalPct = Number((totalR * riskPct).toFixed(2));
    const closed = wins + losses;
    const wr = closed > 0 ? Math.round((wins / closed) * 100) : 0;
    const sign = totalR > 0 ? '+' : '';
    const bMode = state.badgeMode || 'pct';

    const valStr = bMode === 'pct'
      ? `${sign}${numStr(totalPct)}%`
      : `${sign}${numStr(Number(totalR.toFixed(2)))}R`;

    const toolbarBadgeText = bMode === 'cnt'
      ? String(count)
      : (bMode === 'pct' ? `${sign}${totalPct.toFixed(1)}%` : `${sign}${totalR.toFixed(1)}R`);

    const outcomesLine = outcomesList.length > 0
      ? outcomesList.map((o, i) => i === 0 ? o.charAt(0).toUpperCase() + o.slice(1) : o).join(', ')
      : '—';

    const reportText = `Сделок за день: ${count}\n${outcomesLine}\nРезультат дня: ${sign}${numStr(totalPct)}%`;

    return {
      count, totalR, totalPct, outcomesLine, reportText, toolbarBadgeText,
      shortText: count > 0 ? `${count} · ${valStr}` : '0',
      tooltip: `${reportText}\nВинрейт: ${wr}% (${wins}+ / ${losses}-) | ЛКМ — скопировать таблицу, ПКМ — список и отчёт`
    };
  }

  function renderBatchPopover() {
    const pop = document.getElementById('bt-batch-popover');
    if (!pop) return;
    const bStats = computeBatchStats();
    const bucket = getTplBucket();
    const riskPct = Number(state.riskPctPer1R) || 0.3;

    const repPre = document.getElementById('bt-daily-report-text');
    if (repPre) repPre.innerText = bStats.reportText;

    pop.querySelectorAll('.bt-bmode-btn').forEach(btn => {
      const isAct = (state.badgeMode || 'pct') === btn.dataset.bm;
      btn.style.background = isAct ? '#e5c158' : 'transparent';
      btn.style.color = isAct ? '#09090b' : '#9ca3af';
    });

    const listEl = document.getElementById('bt-batch-items-list');
    if (!listEl) return;
    if (!bucket.batchQueue.length) {
      listEl.innerHTML = `<div style="font-size:10.5px;color:#9ca3af;text-align:center;padding:8px 0;">В коробке пока нет сохранённых сделок</div>`;
      return;
    }

    listEl.innerHTML = bucket.batchQueue.map((item, idx) => {
      const rVal = typeof item.rOut === 'number' && !isNaN(item.rOut) ? item.rOut : 0;
      const pctVal = Number((rVal * riskPct).toFixed(2));
      const sign = rVal > 0 ? '+' : '';
      const col = rVal > 0 ? '#34d399' : (rVal < 0 ? '#f87171' : '#a1a1aa');
      const outLbl = getTradeShortOutcomeLabel(item);
      return `
        <div style="display:flex;justify-content:space-between;align-items:center;background:#111116;border:1px solid #252530;border-radius:5px;padding:4px 7px;font-size:10.5px;">
          <div style="display:flex;align-items:center;gap:5px;min-width:0;overflow:hidden;">
            <b style="color:#e5c158;">#${idx + 1}</b>
            <span style="color:#9ca3af;">${item.mskTime || ''}</span>
            <span style="color:${item.isLong ? '#34d399' : '#f87171'};font-weight:700;">${item.isLong ? 'L' : 'S'}</span>
            <span style="color:#e4e4e7;overflow:hidden;text-overflow:ellipsis;">${outLbl}</span>
          </div>
          <div style="display:flex;align-items:center;gap:6px;flex-shrink:0;">
            <b style="color:${col};">${sign}${numStr(pctVal)}% <span style="font-weight:400;color:#9ca3af;">(${sign}${numStr(Number(rVal.toFixed(2)))}R)</span></b>
            <button type="button" class="bt-del-batch-item" data-bidx="${idx}" title="Удалить эту сделку из списка" style="background:rgba(242,54,69,0.15);color:#f87171;border:none;border-radius:4px;padding:1px 5px;font-size:10px;font-weight:700;cursor:pointer;">✕</button>
          </div>
        </div>`;
    }).join('');

    listEl.querySelectorAll('.bt-del-batch-item').forEach(delBtn => {
      delBtn.onclick = e => {
        e.stopPropagation();
        const delIdx = parseInt(delBtn.dataset.bidx, 10);
        const curQueue = [...getTplBucket().batchQueue];
        if (delIdx >= 0 && delIdx < curQueue.length) {
          curQueue.splice(delIdx, 1);
          const totals = recalculateQueueTotals(curQueue);
          saveTplBucket({ batchQueue: curQueue, cumR: totals.cumR, maxCumR: totals.maxCumR });
          renderBatchPopover();
          updateHeaderUI();
          showToast(`Сделка #${delIdx + 1} удалена из списка`);
        }
      };
    });
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

    const bStats = computeBatchStats();
    const batchCntEl = document.getElementById('bt-batch-count');
    batchCntEl.innerText = bStats.shortText;
    batchCntEl.style.color = bStats.totalR > 0 ? '#34d399' : (bStats.totalR < 0 ? '#f87171' : '#f3d26b');
    document.getElementById('bt-copy-batch-btn').title = `${bStats.tooltip}\nНажмите, чтобы скопировать все сохранённые сделки сразу (для вставки в столбец ${pasteCol})`;

    mountNativeToolbarButton();
    updateScreenshotPillUI();
    updateVersionBannerUI();
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
      const rowMsg = rendered.isTwoRows ? '✓ Скопированы 2 строки (Сейф + Остаток)!' : '✓ Скопировано!';
      showToast(`${rowMsg} Вставьте (Ctrl+V) в столбец ${rendered.pasteCol} (${firstColShort})`);
    }).catch(() => {
      isInternalCopy = false;
      showToast('Не удалось скопировать в буфер обмена', 'error');
    });
  }

  function applyProfile(prof, copyImmediately = false) {
    saveState({ outputProfile: prof, decimalSep: prof === 'en_form' ? '.' : ',' });
    const sweptLvlEl = document.getElementById('f-swept-lvl');
    if (sweptLvlEl && sweptLvlEl.value) sweptLvlEl.value = numStr(sweptLvlEl.value);
    const safeP = document.getElementById('bt-safe-price-inp');
    if (safeP && safeP.value) safeP.value = numStr(safeP.value);
    updateHeaderUI();
    const labels = { gs_ru: '📊 Google Docs RU', ru_form: '📗 Excel RU', en_form: '🌐 EN' };
    if (copyImmediately && currentTrade) buildAndCopyRow(true);
    else showToast(`Выбран формат: ${labels[prof]}`);
  }

  
  
  if (quickTplBtn && quickTplMenu) {
    quickTplBtn.onclick = e => {
      e.stopPropagation();
      quickTplMenu.style.display = quickTplMenu.style.display === 'block' ? 'none' : 'block';
    };
    quickTplMenu.querySelectorAll('.bt-qtpl-item').forEach(item => {
      item.onclick = e => {
        e.stopPropagation();
        const val = item.dataset.tpl;
        quickTplMenu.style.display = 'none';
        if (!val) return;
        if (val === 'custom') {
          showColumnEditor = true;
          saveState({ presetKey: 'custom', customColumns: getActiveColumns() });
          renderMainFields();
          toggleSettingsView(true);
          return;
        }
        showColumnEditor = false;
        saveState({ presetKey: val, customColumns: null });
        manualSessionOverride = null;
        renderMainFields();
        if (isSettingsOpen) toggleSettingsView(false);
        updateHeaderUI();
        showToast(`Шаблон переключён: ${PRESET_SCHEMAS[val]?.name || val}`);
      };
    });
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
      showToast('Скопируйте ссылку на скриншот и нажмите Ctrl+V в этом окне');
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
    const baseKey = state.presetKey === 'custom' ? 'bt_30' : state.presetKey;
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

  const toggleBatchPop = e => {
    e.preventDefault();
    e.stopPropagation();
    const pop = document.getElementById('bt-batch-popover');
    if (!pop) return;
    const nextOpen = pop.style.display === 'none';
    pop.style.display = nextOpen ? 'block' : 'none';
    if (nextOpen) renderBatchPopover();
  };
  document.getElementById('bt-copy-batch-btn').addEventListener('contextmenu', toggleBatchPop);
  document.getElementById('bt-open-report-btn').onclick = toggleBatchPop;

  const closeBatchPopBtn = document.getElementById('bt-close-batch-pop');
  if (closeBatchPopBtn) {
    closeBatchPopBtn.onclick = e => {
      e.stopPropagation();
      const pop = document.getElementById('bt-batch-popover');
      if (pop) pop.style.display = 'none';
    };
  }

  document.getElementById('bt-copy-report-btn').onclick = e => {
    e.stopPropagation();
    const bStats = computeBatchStats();
    isInternalCopy = true;
    navigator.clipboard.writeText(bStats.reportText).then(() => {
      isInternalCopy = false;
      showToast('📋 Список сделок скопирован');
    }).catch(() => { isInternalCopy = false; });
  };

  document.querySelectorAll('.bt-bmode-btn').forEach(btn => {
    btn.onclick = e => {
      e.stopPropagation();
      saveState({ badgeMode: btn.dataset.bm });
      renderBatchPopover();
      updateHeaderUI();
    };
  });

  const riskPctInp = document.getElementById('bt-risk-pct-inp');
  if (riskPctInp) {
    riskPctInp.addEventListener('input', e => {
      const v = parseFloat(String(e.target.value).replace(',', '.'));
      if (!isNaN(v) && v > 0 && v <= 100) {
        saveState({ riskPctPer1R: v });
        renderBatchPopover();
        updateHeaderUI();
      }
    });
  }

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
    if (quickTplMenu && !e.target.closest('#bt-quick-tpl-menu') && !e.target.closest('#bt-quick-tpl-btn')) quickTplMenu.style.display = 'none';
    const ccMenu = document.getElementById('bt-cc-menu');
    if (ccMenu && !e.target.closest('#bt-cc-wrap')) ccMenu.style.display = 'none';
    const bPop = document.getElementById('bt-batch-popover'); if (bPop && !e.target.closest('#bt-batch-popover') && !e.target.closest('#bt-open-report-btn')) bPop.style.display = 'none';
    const outMenu = document.getElementById('bt-outcome-menu');
    if (outMenu && !e.target.closest('#bt-outcome-wrap')) outMenu.style.display = 'none';
  });

  mainCopyBtn.onclick = () => buildAndCopyRow(false);
  document.getElementById('bt-close-x').onclick = () => { modalEl.style.display = 'none'; };
  document.getElementById('bt-fab-btn').onclick = () => openModal();

  mountNativeToolbarButton();
  checkGithubScriptUpdate();

  (function initDrag() {
    const handle = document.getElementById('bt-header-drag');
    let dragging = false, sx = 0, sy = 0, it = 0, ir = 0;
    handle.addEventListener('mousedown', e => {
      if (e.target.closest('button') || e.target.closest('#bt-quick-tpl-menu')) return;
      dragging = true; sx = e.clientX; sy = e.clientY;
      const rect = modalEl.getBoundingClientRect();
      it = rect.top; ir = window.innerWidth - rect.right;
      e.preventDefault();
    });
    window.addEventListener('mousemove', e => {
      if (!dragging) return;
      const minTop = 48;
      const maxTop = Math.max(minTop, window.innerHeight - 120);
      const nextTop = Math.min(maxTop, Math.max(minTop, it + (e.clientY - sy)));
      const nextRight = Math.min(Math.max(10, window.innerWidth - 415), Math.max(10, ir - (e.clientX - sx)));
      modalEl.style.top = nextTop + 'px';
      modalEl.style.right = nextRight + 'px';
    });
    window.addEventListener('mouseup', () => {
      if (!dragging) return;
      dragging = false;
      saveState({ posTop: parseInt(modalEl.style.top, 10) || 56, posRight: parseInt(modalEl.style.right, 10) || 20 });
      clampModalPosition();
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
    if (quickTplMenu) quickTplMenu.style.display = 'none';
    toggleSettingsView(false);
    renderMainFields();
    updateHeaderUI();
    clampModalPosition();
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
      const outMenu = document.getElementById('bt-outcome-menu');
      const bPop = document.getElementById('bt-batch-popover');
      if (bPop && bPop.style.display === 'block') bPop.style.display = 'none';
      else if (outMenu && outMenu.style.display === 'block') outMenu.style.display = 'none';
      else if (commentMenuEl && commentMenuEl.style.display === 'block') commentMenuEl.style.display = 'none';
      else if (quickTplMenu && quickTplMenu.style.display === 'block') quickTplMenu.style.display = 'none';
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
  }, true);

})();
