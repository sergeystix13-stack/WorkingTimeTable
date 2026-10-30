/*
 * Табель — планировщик отпуска (перенос planner.py программы для ПК; сверяется тестом).
 *
 * По дням: перебираются периоды отпуска на 12 месяцев вперёд, которые начинаются и
 * заканчиваются в рабочий день; дни отпуска — календарные без праздников (ст. 120 ТК РФ);
 * к периоду прибавляются примыкающие выходные — «дни отдыха».
 * По деньгам (примерно): отпускные = средний дневной (оклад ÷ 29,3) × дни отпуска;
 * зарплата месяца меньше на оклад ÷ рабочие дни месяца за каждый пропущенный рабочий день.
 */
(function (root) {
  "use strict";
  const C = (typeof Calc !== "undefined") ? Calc : require("./calc.js");

  const AVG_MONTH_DAYS = 29.3;
  const MONEY_MIN_DAYS = 7;
  const MAX_EDGE = 40;

  function findBlocks(start, end, special, minDays, maxDays, busy, salary, avgDaily) {
    busy = busy || new Set();
    if (salary && !avgDaily) avgDaily = salary / AVG_MONTH_DAYS;
    const isWork = d => C.dayType(d, special) === C.DAY.WORK;
    const wdCache = {};
    const monthWorkdays = key => {
      if (!(key in wdCache)) { const [y, m] = key.split("-").map(Number); wdCache[key] = C.countWorkdays(y, m, special); }
      return wdCache[key];
    };
    const blocks = [];
    for (let s = start; C.diffDays(end, s) >= 0; s = C.addDays(s, 1)) {
      if (!isWork(s) || busy.has(s)) continue;
      let charged = 0, workIn = 0;
      const missed = {};
      for (let e = s; C.diffDays(end, e) >= 0 && !busy.has(e); e = C.addDays(e, 1)) {
        if (!C.isPublicHoliday(e, special)) charged++;
        if (charged > maxDays) break;
        if (isWork(e)) {
          workIn++;
          const key = e.slice(0, 7);
          missed[key] = (missed[key] || 0) + 1;
          if (charged >= minDays) blocks.push(makeBlock(s, e, charged, workIn, Object.assign({}, missed), isWork, monthWorkdays, salary, avgDaily));
        }
      }
    }
    return blocks;
  }

  function makeBlock(s, e, charged, workIn, missed, isWork, monthWorkdays, salary, avgDaily) {
    let left = C.addDays(s, -1);
    for (let i = 0; i < MAX_EDGE && !isWork(left); i++) left = C.addDays(left, -1);
    let right = C.addDays(e, 1);
    for (let i = 0; i < MAX_EDGE && !isWork(right); i++) right = C.addDays(right, 1);
    const restStart = C.addDays(left, 1), restEnd = C.addDays(right, -1);
    const b = { start: s, end: e, charged, restStart, restEnd, rest: C.diffDays(restEnd, restStart) + 1,
      workdays: workIn, vacPay: null, lost: null, delta: null };
    b.ratio = b.rest / b.charged;
    if (salary) {
      let lost = 0;
      for (const key of Object.keys(missed).sort()) {
        const w = monthWorkdays(key);
        if (w) lost += salary * missed[key] / w;
      }
      b.vacPay = C.roundHalfUp(avgDaily * charged);
      b.lost = C.roundHalfUp(lost);
      b.delta = b.vacPay - b.lost;
    }
    b.perDay = b.delta === null ? null : b.delta / b.charged;
    return b;
  }

  const cmp = (a, b) => (a < b ? -1 : a > b ? 1 : 0);
  function rank(blocks, priority) {
    if (!blocks.length) return [];
    const hasMoney = blocks[0].delta !== null;
    const out = blocks.slice();
    if (priority === "money" && hasMoney) {
      out.sort((a, b) => cmp(b.perDay, a.perDay) || cmp(b.ratio, a.ratio) || cmp(a.start, b.start) || cmp(a.charged, b.charged));
    } else {
      out.sort((a, b) => cmp(b.ratio, a.ratio) || cmp(b.rest, a.rest) || cmp(a.start, b.start) || cmp(a.charged, b.charged));
    }
    return out;
  }

  const overlaps = (b, list) => list.some(c => !(b.restEnd < c.restStart || b.restStart > c.restEnd));

  function makePlan(blocks, budget, priority, minDays, maxDays, need14, mainMin) {
    mainMin = Math.max(minDays, mainMin || minDays);
    const ranked = rank(blocks, priority);
    const chosen = [];
    let remaining = Math.max(0, Math.floor(budget));
    if (need14 && remaining >= 14) {
      const long = ranked.find(b => b.charged === 14);
      if (long) { chosen.push(long); remaining -= 14; }
    }
    const passes = mainMin > minDays ? [mainMin, minDays] : [minDays];
    for (const lo of passes) {
      for (const b of ranked) {
        if (remaining < lo) break;
        if (b.charged >= lo && b.charged <= Math.min(maxDays, remaining) && !overlaps(b, chosen)) {
          chosen.push(b);
          remaining -= b.charged;
        }
      }
    }
    chosen.sort((a, b) => cmp(a.start, b.start));
    return { plan: chosen, left: remaining };
  }

  function alternatives(blocks, priority, plan, limit, minDays, maxDays) {
    const out = [];
    for (const b of rank(blocks, priority)) {
      if (out.length >= limit) break;
      if (b.charged >= minDays && b.charged <= maxDays && !overlaps(b, plan) && !overlaps(b, out)) out.push(b);
    }
    return out;
  }

  function horizon(today) { return [C.addDays(today, 1), C.addDays(today, 365)]; }

  function planVacation(today, special, budget, minDays, maxDays, priority, need14, busy, salary, avgDaily) {
    const [start, end] = horizon(today);
    minDays = Math.max(1, Math.floor(minDays));
    maxDays = Math.max(minDays, Math.floor(maxDays));
    const searchMax = Math.max(maxDays, need14 ? 14 : maxDays);
    const blocks = findBlocks(start, end, special, minDays, searchMax, busy, salary, avgDaily);
    const byMoney = priority === "money" && !!salary;
    const mainMin = byMoney ? Math.max(minDays, Math.min(MONEY_MIN_DAYS, maxDays)) : minDays;
    const { plan, left } = makePlan(blocks, budget, priority, minDays, maxDays, need14, mainMin);
    const alts = alternatives(blocks, priority, plan, 8, mainMin, maxDays);
    return { plan, left, alternatives: alts, start, end };
  }

  const api = { findBlocks, rank, makePlan, alternatives, horizon, planVacation, AVG_MONTH_DAYS };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Planner = api;
})(typeof window !== "undefined" ? window : globalThis);
