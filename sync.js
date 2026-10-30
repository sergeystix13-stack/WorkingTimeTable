/*
 * Табель — проверка и объединение файлов синхронизации (телефон ↔ ПК).
 *
 * Формат файла — тот же, что у резервной копии телефона:
 *   { app: "tabel", version: 2, exported, settingsT, days, special, monthNorms, seeded, notified, settings }
 * Каждый день хранит время последней правки t (мс). Очищенный день не удаляется,
 * а остаётся пустой записью с t — так удаление тоже переезжает на другое устройство.
 *
 * Правила объединения (одинаковые здесь и в sync.py программы для ПК):
 *   - день: остаётся запись с более поздним t (при равенстве — своя);
 *   - праздники и переносы: объединяются, при совпадении даты побеждает файл;
 *   - настройки нормы и отпуска, нормы месяцев: целиком с той стороны, где они менялись позже;
 *   - тема и шрифт не синхронизируются — у каждого устройства свои.
 */
(function (root) {
  "use strict";

  const MARKS = new Set(["", "off", "vacation", "sick"]);
  const MODES = new Set(["day", "rate", "week", "month"]);
  const SYNC_KEYS = ["workMin", "lunchMin", "normMode", "rate", "weekMin", "monthMin",
    "hireDate", "perYear", "usedBefore", "excluded"];
  const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
  const TIME_RE = /^([01]\d|2[0-3]):[0-5]\d$/;
  const MAX_DAYS = 20000;

  class SyncError extends Error {}

  const own = (o, k) => Object.prototype.hasOwnProperty.call(o, k);
  const isObj = v => v !== null && typeof v === "object" && !Array.isArray(v);
  const num = (v, lo, hi, dflt) => (typeof v === "number" && isFinite(v) && v >= lo && v <= hi) ? v : dflt;
  const int = (v, lo, hi, dflt) => Number.isInteger(v) && v >= lo && v <= hi ? v : dflt;
  const validDate = s => {
    if (typeof s !== "string" || !DATE_RE.test(s)) return false;
    const [y, m, d] = s.split("-").map(Number);
    if (y < 1990 || y > 2200 || m < 1 || m > 12 || d < 1) return false;
    return d <= new Date(Date.UTC(y, m, 0)).getUTCDate();
  };
  const cleanText = (s, max) => (typeof s === "string" ? s : "").replace(/[\u0000-\u001f\u007f]/g, " ").slice(0, max);

  const DEFAULT_SETTINGS = {
    workMin: 480, lunchMin: 30, normMode: "day", rate: 0.5, weekMin: 1200, monthMin: 4800,
    hireDate: "", perYear: 28, usedBefore: 0, excluded: 0, theme: "system", font: "Monomakh",
  };

  /**
   * Проверяет и очищает данные из файла или хранилища. Всё неизвестное отбрасывается,
   * неверные значения заменяются значениями по умолчанию. Объекты создаются без прототипа
   * (Object.create(null)), поэтому ключи вроде «__proto__» ничего не ломают.
   */
  function sanitize(raw) {
    if (!isObj(raw) || !isObj(raw.days)) throw new SyncError("Это не файл Табеля.");
    if (raw.app !== undefined && raw.app !== "tabel") throw new SyncError("Это не файл Табеля.");
    const out = {
      app: "tabel", version: 2,
      settingsT: int(raw.settingsT, 0, 1e15, 0),
      days: {}, special: {}, monthNorms: {}, seeded: {}, notified: {},
      settings: Object.assign({}, DEFAULT_SETTINGS),
    };
    const dayKeys = Object.keys(raw.days);
    if (dayKeys.length > MAX_DAYS) throw new SyncError("Файл слишком большой для Табеля.");
    for (const d of dayKeys) {
      const r = raw.days[d];
      if (!validDate(d) || !isObj(r)) continue;
      const rec = {
        in: typeof r.in === "string" && TIME_RE.test(r.in) ? r.in : "",
        out: typeof r.out === "string" && TIME_RE.test(r.out) ? r.out : "",
        mark: typeof r.mark === "string" && MARKS.has(r.mark) ? r.mark : "",
        t: int(r.t, 0, 1e15, 0),
      };
      if (rec.mark) { rec.in = ""; rec.out = ""; }
      if (rec.out && (!rec.in || rec.out <= rec.in)) rec.out = "";
      out.days[d] = rec;
    }
    if (isObj(raw.special)) {
      for (const d of Object.keys(raw.special)) {
        const v = raw.special[d];
        if (!validDate(d) || !isObj(v) || (v.kind !== "holiday" && v.kind !== "workday")) continue;
        out.special[d] = { kind: v.kind, title: cleanText(v.title, 120) };
      }
    }
    if (isObj(raw.monthNorms)) {
      for (const k of Object.keys(raw.monthNorms)) {
        if (/^\d{4}-\d{2}$/.test(k)) {
          const v = int(raw.monthNorms[k], 0, 18000, null);
          if (v !== null) out.monthNorms[k] = v;
        }
      }
    }
    if (isObj(raw.seeded)) for (const k of Object.keys(raw.seeded)) if (/^\d{4}$/.test(k) && raw.seeded[k] === true) out.seeded[k] = true;
    if (isObj(raw.notified)) {
      for (const k of Object.keys(raw.notified)) {
        if (/^\d{4}-\d{2}$/.test(k)) { const v = int(raw.notified[k], 0, 1000, null); if (v !== null) out.notified[k] = v; }
      }
    }
    const s = isObj(raw.settings) ? raw.settings : {};
    const st = out.settings;
    st.workMin = int(s.workMin, 60, 960, st.workMin);
    st.lunchMin = int(s.lunchMin, 0, 180, st.lunchMin);
    st.normMode = MODES.has(s.normMode) ? s.normMode : st.normMode;
    st.rate = num(s.rate, 0.05, 1.5, st.rate);
    st.weekMin = int(s.weekMin, 60, 3600, st.weekMin);
    st.monthMin = int(s.monthMin, 0, 18000, st.monthMin);
    st.hireDate = validDate(s.hireDate) ? s.hireDate : "";
    st.perYear = num(s.perYear, 1, 100, st.perYear);
    st.usedBefore = num(s.usedBefore, 0, 10000, st.usedBefore);
    st.excluded = int(s.excluded, 0, 10000, st.excluded);
    st.theme = ["light", "dark", "system"].includes(s.theme) ? s.theme : st.theme;
    st.font = ["Monomakh", "Shafarik", "Ruslan", "OldStandard", "system"].includes(s.font) ? s.font : st.font;
    return out;
  }

  /** Объединяет свои данные с данными из файла. Возвращает {data, stats}. */
  function merge(local, incoming) {
    local = sanitize(local);
    incoming = sanitize(incoming);
    const res = sanitize(local);              // копия своих данных
    const stats = { added: 0, updated: 0, removed: 0, kept: 0, settings: false };

    for (const d of Object.keys(incoming.days)) {
      const inc = incoming.days[d], mine = local.days[d];
      const incEmpty = !inc.in && !inc.out && !inc.mark;
      if (!mine) {
        res.days[d] = inc;
        if (!incEmpty) stats.added++;
        continue;
      }
      if (inc.t > mine.t) {
        const same = inc.in === mine.in && inc.out === mine.out && inc.mark === mine.mark;
        res.days[d] = inc;
        if (!same) {
          const mineEmpty = !mine.in && !mine.out && !mine.mark;
          if (incEmpty) stats.removed++;
          else if (mineEmpty) stats.added++;
          else stats.updated++;
        }
      } else if (inc.in !== mine.in || inc.out !== mine.out || inc.mark !== mine.mark) {
        stats.kept++;
      }
    }
    for (const d of Object.keys(incoming.special)) res.special[d] = incoming.special[d];
    for (const k of Object.keys(incoming.seeded)) res.seeded[k] = true;
    for (const k of Object.keys(incoming.notified)) res.notified[k] = Math.max(res.notified[k] || 0, incoming.notified[k]);
    if (incoming.settingsT > local.settingsT) {
      for (const k of SYNC_KEYS) res.settings[k] = incoming.settings[k];
      res.monthNorms = Object.assign({}, incoming.monthNorms);
      res.settingsT = incoming.settingsT;
      stats.settings = true;
    }
    return { data: res, stats };
  }

  function statsText(st) {
    const parts = [];
    if (st.added) parts.push(`добавлено дней: ${st.added}`);
    if (st.updated) parts.push(`обновлено: ${st.updated}`);
    if (st.removed) parts.push(`очищено: ${st.removed}`);
    if (st.kept) parts.push(`оставлено своих (они новее): ${st.kept}`);
    if (st.settings) parts.push("настройки взяты из файла");
    return parts.length ? parts.join(", ") : "изменений нет — данные уже совпадают";
  }

  const api = { sanitize, merge, statsText, SyncError, DEFAULT_SETTINGS, SYNC_KEYS };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Sync = api;
})(typeof window !== "undefined" ? window : globalThis);
