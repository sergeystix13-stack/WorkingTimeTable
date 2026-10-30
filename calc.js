/*
 * Табель — логика расчётов (перенос calc.py, vacation.py и calendar_ru.py
 * из настольной версии). Без интерфейса: только данные на входе и результат.
 * Сверяется с Python-версией тестом tests/compare.mjs.
 *
 * Даты — строки «ГГГГ-ММ-ДД», длительности — целые минуты.
 */
(function (root) {
  "use strict";

  // ------------------------------------------------------------ константы
  const DEFAULT_WORK = 480, DEFAULT_LUNCH = 30;
  // Режимы нормы: по дням (полная ставка) и три режима для полуставки,
  // где переработка считается только по итогам месяца.
  const MODES = { DAY: "day", RATE: "rate", WEEK: "week", MONTH: "month" };
  const MODE_LABELS = { day: "По дням (полная ставка)", rate: "Доля ставки", week: "Часов в неделю", month: "Часов в месяц" };
  const DEFAULT_RATE = 0.5, DEFAULT_WEEK = 1200, DEFAULT_MONTH = 4800;
  const roundHalfUp = x => Math.floor(x + 0.5 + 1e-9);   // как ROUND в Excel и round_half_up в Python
  const MIN_WORK = 60, MAX_WORK = 16 * 60, MAX_LUNCH = 180;

  const WEEKDAYS = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота", "Воскресенье"];
  const WEEKDAYS_SHORT = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"];
  const MONTHS = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль",
    "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"];
  const MONTHS_GEN = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля",
    "августа", "сентября", "октября", "ноября", "декабря"];

  const MARK = { NONE: "", OFF: "off", VACATION: "vacation", SICK: "sick" };
  const MARK_LABELS = { "": "—", off: "Отгул за переработку", vacation: "Отпуск", sick: "Больничный" };
  const DAY = { WORK: "workday", WEEKEND: "weekend", HOLIDAY: "holiday" };

  class InputError extends Error {}

  // ------------------------------------------------------------ даты
  function iso(y, m, d) {
    return `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  }
  function parts(s) {
    const [y, m, d] = s.split("-").map(Number);
    return { y, m, d };
  }
  function toUTC(s) { const p = parts(s); return Date.UTC(p.y, p.m - 1, p.d); }
  function fromUTC(ms) { const t = new Date(ms); return iso(t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate()); }
  function addDays(s, n) { return fromUTC(toUTC(s) + n * 86400000); }
  function diffDays(a, b) { return Math.round((toUTC(a) - toUTC(b)) / 86400000); } // a − b
  function weekday(s) { return (new Date(toUTC(s)).getUTCDay() + 6) % 7; } // 0 = пн
  function daysInMonth(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); }
  function todayISO(now) {
    const t = now || new Date();
    return iso(t.getFullYear(), t.getMonth() + 1, t.getDate());
  }
  function shiftMonth(y, m, delta) {
    const idx = y * 12 + (m - 1) + delta;
    return [Math.floor(idx / 12), (idx % 12 + 12) % 12 + 1];
  }
  function monthDates(y, m) {
    const out = [];
    for (let d = 1; d <= daysInMonth(y, m); d++) out.push(iso(y, m, d));
    return out;
  }
  function fmtDate(s) { const p = parts(s); return `${String(p.d).padStart(2, "0")}.${String(p.m).padStart(2, "0")}.${p.y}`; }
  function parseRuDate(text) {
    const m = /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/.exec((text || "").trim());
    if (!m) return null;
    const d = +m[1], mo = +m[2], y = +m[3];
    if (mo < 1 || mo > 12 || d < 1 || d > daysInMonth(y, mo)) return null;
    return iso(y, mo, d);
  }

  // ------------------------------------------------------------ время
  function parseTime(text) {
    const raw = (text || "").trim();
    if (!raw) throw new InputError("Время не указано.");
    let h, mi, m = /^(\d{1,2})[:.,\-\s](\d{2})$/.exec(raw);
    if (m) { h = +m[1]; mi = +m[2]; }
    else {
      m = /^(\d{3,4})$/.exec(raw);
      if (!m) throw new InputError(`«${raw}» — неверный формат времени. Введите часы и минуты, например 09:00.`);
      const dg = m[1].padStart(4, "0");
      h = +dg.slice(0, 2); mi = +dg.slice(2);
    }
    if (h > 23) throw new InputError(`«${raw}»: часов может быть от 0 до 23.`);
    if (mi > 59) throw new InputError(`«${raw}»: минут может быть от 00 до 59.`);
    return h * 60 + mi;
  }
  function parseDuration(text) {
    const raw = (text || "").trim();
    const m = /^(\d{1,3})(?:[:.,](\d{2}))?$/.exec(raw);
    if (!m) throw new InputError(`«${raw}» — укажите длительность как Ч:ММ, например 8:00 или 0:30.`);
    const mi = +(m[2] || 0);
    if (mi > 59) throw new InputError(`«${raw}»: минут может быть от 00 до 59.`);
    return +m[1] * 60 + mi;
  }
  function normalizeTime(text) {
    if (!(text || "").trim()) return "";
    const v = parseTime(text);
    return `${String(Math.floor(v / 60)).padStart(2, "0")}:${String(v % 60).padStart(2, "0")}`;
  }
  function fmtMinutes(v, empty = "") {
    if (v === null || v === undefined) return empty;
    const sign = v < 0 ? "−" : "";
    v = Math.abs(v);
    return `${sign}${Math.floor(v / 60)}:${String(v % 60).padStart(2, "0")}`;
  }
  function plural(n, one, few, many) {
    n = Math.abs(n);
    if (n % 10 === 1 && n % 100 !== 11) return one;
    if (n % 10 >= 2 && n % 10 <= 4 && !(n % 100 >= 12 && n % 100 <= 14)) return few;
    return many;
  }
  function daysPhrase(n) { return `${n} ${plural(n, "день", "дня", "дней")}`; }
  function fmtHoursWords(v) {
    v = Math.max(0, v);
    const h = Math.floor(v / 60), mi = v % 60;
    return mi === 0 ? `${h} ${plural(h, "час", "часа", "часов")}` : `${h} ч ${mi} мин`;
  }

  // ------------------------------------------------------------ правила нормы
  function fmtRate(x) { return x.toFixed(2).replace(/0+$/, "").replace(/\.$/, "").replace(".", ","); }
  function makeRules(workMin = DEFAULT_WORK, lunchMin = DEFAULT_LUNCH, mode = MODES.DAY,
                     rate = DEFAULT_RATE, weekMin = DEFAULT_WEEK, monthMin = DEFAULT_MONTH) {
    const r = {
      workMin, lunchMin, mode, rate, weekMin, monthMin,
      presence: workMin + lunchMin,
      flexible: mode !== MODES.DAY,
    };
    // Норма одного рабочего дня в минутах (может быть дробной)
    r.perDay = workdays => {
      if (mode === MODES.RATE) return workMin * rate;
      if (mode === MODES.WEEK) return weekMin / 5;
      if (mode === MODES.MONTH) return workdays ? monthMin / workdays : 0;
      return workMin;
    };
    r.describe = () => {
      if (mode === MODES.RATE) return `${fmtRate(rate)} ставки`;
      if (mode === MODES.WEEK) return `${fmtMinutes(weekMin)} в неделю`;
      if (mode === MODES.MONTH) return `${fmtMinutes(monthMin)} в месяц`;
      return "полная ставка";
    };
    return r;
  }
  function validateRules(r) {
    if (!MODE_LABELS[r.mode]) throw new InputError("Неизвестный режим нормы.");
    if (!(r.workMin >= MIN_WORK && r.workMin <= MAX_WORK)) throw new InputError("Рабочий день может быть от 1:00 до 16:00.");
    if (!(r.lunchMin >= 0 && r.lunchMin <= MAX_LUNCH)) throw new InputError("Обед может быть от 0:00 до 3:00.");
    if (r.mode === MODES.RATE && !(r.rate >= 0.05 && r.rate <= 1.5)) throw new InputError("Доля ставки может быть от 0,05 до 1,5 (например 0,5).");
    if (r.mode === MODES.WEEK && !(r.weekMin >= 60 && r.weekMin <= 3600)) throw new InputError("Часов в неделю может быть от 1:00 до 60:00.");
    if (r.mode === MODES.MONTH && !(r.monthMin >= 0 && r.monthMin <= 18000)) throw new InputError("Часов в месяц может быть от 0:00 до 300:00.");
  }

  // ------------------------------------------------------------ день
  const EMPTY = Object.freeze({ presence: null, worked: null, overtime: null, undertime: null, complete: false });

  function validatePair(a, b) {
    const hasIn = !!(a || "").trim(), hasOut = !!(b || "").trim();
    if (hasIn) parseTime(a);
    if (hasOut) {
      const out = parseTime(b);
      if (!hasIn) throw new InputError("Сначала укажите время прихода, потом — ухода.");
      if (out <= parseTime(a)) {
        throw new InputError(`Время ухода (${normalizeTime(b)}) должно быть позже времени прихода (${normalizeTime(a)}).`);
      }
    }
  }
  function calcDay(a, b, rules) {
    rules = rules || makeRules();
    validatePair(a, b);
    if (!(a || "").trim() || !(b || "").trim()) return EMPTY;
    const p = parseTime(b) - parseTime(a);
    const worked = Math.max(0, p - rules.lunchMin);
    if (rules.flexible) {
      // У полуставки нет дневной нормы — переработка считается по месяцу
      return { presence: p, worked, overtime: null, undertime: null, complete: true };
    }
    return {
      presence: p,
      worked,
      overtime: Math.max(0, p - rules.presence),
      undertime: Math.max(0, rules.presence - p),
      complete: true,
    };
  }

  // ------------------------------------------------------------ календарь
  function dayType(s, special) {
    const k = special[s];
    if (k === "workday") return DAY.WORK;
    if (k === "holiday") return DAY.HOLIDAY;
    return weekday(s) >= 5 ? DAY.WEEKEND : DAY.WORK;
  }
  function countWorkdays(y, m, special) {
    return monthDates(y, m).filter(s => dayType(s, special) === DAY.WORK).length;
  }

  // ------------------------------------------------------------ месяц
  function rec(records, s) { return records[s] || { in: "", out: "", mark: "" }; }

  function buildRows(y, m, special, records, titles, rules) {
    titles = titles || {};
    return monthDates(y, m).map(s => {
      const type = dayType(s, special), r = rec(records, s);
      let result = EMPTY;
      if (type === DAY.WORK && !r.mark) {
        try { result = calcDay(r.in, r.out, rules); } catch (e) { result = EMPTY; }
      }
      return { day: s, type, title: titles[s] || "", record: r, result, editable: type === DAY.WORK };
    });
  }

  function summarizeMonth(y, m, special, records, titles, rules) {
    rules = rules || makeRules();
    const s = {
      rules, workdays: 0, vacationDays: 0, sickDays: 0, offDays: 0, norm: 0, worked: 0,
      remaining: 0, overtimeGross: 0, undertimeGross: 0, overtimeNet: 0, balance: 0,
      availableOffDays: 0, earnedLevel: 0, filled: 0,
      rows: buildRows(y, m, special, records, titles, rules),
    };
    for (const row of s.rows) {
      if (row.type !== DAY.WORK) continue;
      s.workdays++;
      const mk = row.record.mark;
      if (mk === MARK.VACATION) s.vacationDays++;
      else if (mk === MARK.SICK) s.sickDays++;
      else if (mk === MARK.OFF) s.offDays++;
      else if (row.result.complete) {
        s.worked += row.result.worked;
        s.overtimeGross += row.result.overtime || 0;
        s.undertimeGross += row.result.undertime || 0;
        s.filled++;
      }
    }
    s.normDays = s.workdays - s.vacationDays - s.sickDays;
    const perDay = rules.perDay(s.workdays);
    s.dayCost = roundHalfUp(perDay);
    s.norm = roundHalfUp(perDay * s.normDays);
    s.remaining = s.norm - s.worked;
    if (rules.flexible) {
      // Переработка по итогам месяца — сверх нормы за дни, которые нужно было отработать
      const required = s.norm - s.offDays * s.dayCost;
      s.overtimeGross = Math.max(0, s.worked - required);
      s.undertimeGross = Math.max(0, required - s.worked);
    }
    const cost = s.dayCost;
    s.overtimeNet = Math.max(0, s.overtimeGross - s.undertimeGross);
    s.balance = s.overtimeNet - s.offDays * cost;
    s.availableOffDays = cost > 0 ? Math.floor(Math.max(0, s.balance) / cost) : 0;
    s.earnedLevel = cost > 0 ? Math.floor(s.overtimeNet / cost) : 0;
    return s;
  }

  function checkMarkChange(y, m, special, records, day, mark, rules) {
    rules = rules || makeRules();
    const p = parts(day);
    if (p.y !== y || p.m !== m) throw new InputError("Отгул можно взять только в пределах того же месяца.");
    if (dayType(day, special) !== DAY.WORK && mark !== MARK.NONE && mark !== MARK.VACATION) {
      throw new InputError("На выходной или праздник можно поставить только отпуск.");
    }
    if (mark !== MARK.OFF) return;
    const old = rec(records, day);
    const trial = Object.assign({}, records, { [day]: { in: "", out: "", mark: MARK.OFF } });
    const after = summarizeMonth(y, m, special, trial, null, rules);
    if (after.balance < 0) {
      const before = summarizeMonth(y, m, special,
        Object.assign({}, records, { [day]: { in: old.in, out: old.out, mark: "" } }), null, rules);
      const have = Math.max(0, before.balance);
      const cost = after.dayCost;
      const need = cost - have;
      const hint = rules.flexible ? " У полуставки переработка считается по итогам месяца — сначала нужно отработать норму месяца." : "";
      throw new InputError(`Недостаточно переработки для отгула. Сейчас накоплено ${fmtMinutes(have)}, ` +
        `нужно ${fmtMinutes(cost)} — не хватает ${fmtMinutes(Math.max(need, 0))}.${hint}`);
    }
  }

  function bannerText(s) {
    if (s.balance < 0) return `Внимание: после отгула появилась недоработка — переработки не хватает на ${fmtMinutes(-s.balance)}.`;
    const n = s.availableOffDays;
    if (n <= 0) return "";
    return `Вы накопили ${fmtHoursWords(s.balance)} переработки — можно взять ${daysPhrase(n)} отгула`;
  }
  function popupText(s) {
    const earned = s.earnedLevel * s.dayCost;
    let t = `Вы накопили ${fmtHoursWords(earned)} переработки — можно взять ${daysPhrase(s.earnedLevel)} отгула.`;
    if (s.offDays) t += `\nУже взято в этом месяце: ${daysPhrase(s.offDays)}, доступно ещё: ${daysPhrase(s.availableOffDays)}.`;
    return t + "\n\nЧтобы взять отгул, откройте нужный рабочий день этого месяца и выберите «Отгул за переработку». " +
      "Переработки не переносятся на следующий месяц.";
  }

  // ------------------------------------------------------------ отпуск (ТК РФ)
  const HALF_MONTH = 15;
  const PUBLIC = new Set(["01-01", "01-02", "01-03", "01-04", "01-05", "01-06", "01-07", "01-08",
    "02-23", "03-08", "05-01", "05-09", "06-12", "11-04"]);

  function addMonths(s, n) {
    const p = parts(s);
    const [y, m] = shiftMonth(p.y, p.m, n);
    return iso(y, m, Math.min(p.d, daysInMonth(y, m)));
  }
  function serviceMonths(hire, onDate, excluded) {
    const start = addDays(hire, Math.max(0, excluded || 0));
    const end = addDays(onDate, 1);
    if (diffDays(end, start) <= 0) return 0;
    const ps = parts(start), pe = parts(end);
    let months = (pe.y - ps.y) * 12 + pe.m - ps.m;
    while (months > 0 && diffDays(addMonths(start, months), end) > 0) months--;
    const rest = diffDays(end, addMonths(start, months));
    if (rest >= HALF_MONTH) months++;
    return months;
  }
  function round2(x) { return Math.round((x + Number.EPSILON) * 100) / 100; }
  function isPublicHoliday(s, special) { return PUBLIC.has(s.slice(5)) && special[s] !== "workday"; }
  function countVacationDays(a, b, special) {
    let n = 0;
    for (let s = a; diffDays(b, s) >= 0; s = addDays(s, 1)) if (!isPublicHoliday(s, special)) n++;
    return n;
  }
  function vacationPeriods(dates, special) {
    const list = Array.from(new Set(dates)).sort();
    const res = [];
    for (const d of list) {
      if (res.length) {
        const last = res[res.length - 1];
        let bridged = true;
        for (let g = addDays(last[1], 1); diffDays(d, g) > 0; g = addDays(g, 1)) {
          if (dayType(g, special) === DAY.WORK) { bridged = false; break; }
        }
        if (bridged) { last[1] = d; continue; }
      }
      res.push([d, d]);
    }
    return res.map(([a, b]) => [a, b, countVacationDays(a, b, special)]);
  }
  function vacationStatus(v, dates, special, today) {
    if (!v.hireDate) return { configured: false };
    const periods = vacationPeriods(dates, special);
    let past = 0, planned = 0;
    for (const [a, b] of periods) {
      if (diffDays(today, b) >= 0) past += countVacationDays(a, b, special);
      else if (diffDays(a, today) > 0) planned += countVacationDays(a, b, special);
      else {
        past += countVacationDays(a, today, special);
        planned += countVacationDays(addDays(today, 1), b, special);
      }
    }
    const months = serviceMonths(v.hireDate, today, v.excluded || 0);
    const accrued = round2(v.perYear * months / 12);
    const available = round2(accrued - (v.usedBefore || 0) - past - planned);
    return { configured: true, months, accrued, usedBefore: v.usedBefore || 0, usedPast: past, planned, available, periods };
  }
  function fmtDays(x) {
    let t = x.toFixed(2).replace(/0+$/, "").replace(/\.$/, "");
    t = t.replace(".", ",").replace("-", "−");
    return (t === "−0" || t === "") ? "0" : t;
  }

  // ------------------------------------------------------------ календарь РФ
  const H = "holiday", W = "workday";
  const CAL = {
    2026: [[1, 1, H, "Новогодние каникулы"], [1, 2, H, "Новогодние каникулы"], [1, 3, H, "Новогодние каникулы"],
      [1, 4, H, "Новогодние каникулы"], [1, 5, H, "Новогодние каникулы"], [1, 6, H, "Новогодние каникулы"],
      [1, 7, H, "Рождество Христово"], [1, 8, H, "Новогодние каникулы"], [1, 9, H, "Перенос выходного с 3 января"],
      [2, 23, H, "День защитника Отечества"], [3, 8, H, "Международный женский день"],
      [3, 9, H, "Перенос праздника с 8 марта"], [5, 1, H, "Праздник Весны и Труда"], [5, 9, H, "День Победы"],
      [5, 11, H, "Перенос праздника с 9 мая"], [6, 12, H, "День России"], [11, 4, H, "День народного единства"],
      [12, 31, H, "Перенос выходного с 4 января"]],
    2027: [[1, 1, H, "Новогодние каникулы"], [1, 2, H, "Новогодние каникулы"], [1, 3, H, "Новогодние каникулы"],
      [1, 4, H, "Новогодние каникулы"], [1, 5, H, "Новогодние каникулы"], [1, 6, H, "Новогодние каникулы"],
      [1, 7, H, "Рождество Христово"], [1, 8, H, "Новогодние каникулы"],
      [2, 20, W, "Рабочая суббота (перенос на 22 февраля)"], [2, 22, H, "Перенос выходного с 20 февраля"],
      [2, 23, H, "День защитника Отечества"], [3, 8, H, "Международный женский день"],
      [5, 1, H, "Праздник Весны и Труда"], [5, 3, H, "Перенос праздника с 1 мая"], [5, 9, H, "День Победы"],
      [5, 10, H, "Перенос праздника с 9 мая"], [6, 12, H, "День России"], [6, 14, H, "Перенос праздника с 12 июня"],
      [11, 4, H, "День народного единства"], [11, 5, H, "Перенос выходного с 2 января"],
      [12, 31, H, "Перенос выходного с 3 января"]],
  };
  const FIXED = [[1, 1, "Новогодние каникулы"], [1, 2, "Новогодние каникулы"], [1, 3, "Новогодние каникулы"],
    [1, 4, "Новогодние каникулы"], [1, 5, "Новогодние каникулы"], [1, 6, "Новогодние каникулы"],
    [1, 7, "Рождество Христово"], [1, 8, "Новогодние каникулы"], [2, 23, "День защитника Отечества"],
    [3, 8, "Международный женский день"], [5, 1, "Праздник Весны и Труда"], [5, 9, "День Победы"],
    [6, 12, "День России"], [11, 4, "День народного единства"]];
  function hasOfficialCalendar(y) { return Object.prototype.hasOwnProperty.call(CAL, y); }
  function defaultSpecialDays(y) {
    if (hasOfficialCalendar(y)) return CAL[y].map(([m, d, k, t]) => [iso(y, m, d), k, t]);
    return FIXED.map(([m, d, t]) => [iso(y, m, d), H, t]);
  }

  const api = {
    DEFAULT_WORK, DEFAULT_LUNCH, MODES, MODE_LABELS, DEFAULT_RATE, DEFAULT_WEEK, DEFAULT_MONTH, roundHalfUp, fmtRate, WEEKDAYS, WEEKDAYS_SHORT, MONTHS, MONTHS_GEN, MARK, MARK_LABELS, DAY,
    InputError, iso, parts, addDays, diffDays, weekday, daysInMonth, todayISO, shiftMonth, monthDates,
    fmtDate, parseRuDate, parseTime, parseDuration, normalizeTime, fmtMinutes, plural, daysPhrase,
    fmtHoursWords, makeRules, validateRules, validatePair, calcDay, dayType, countWorkdays, buildRows,
    summarizeMonth, checkMarkChange, bannerText, popupText, addMonths, serviceMonths, isPublicHoliday,
    countVacationDays, vacationPeriods, vacationStatus, fmtDays, hasOfficialCalendar, defaultSpecialDays,
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Calc = api;
})(typeof window !== "undefined" ? window : globalThis);
