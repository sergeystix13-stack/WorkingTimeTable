/*
 * Табель — интерфейс мобильного приложения.
 * Расчёты — в calc.js, выгрузка в Excel — в xlsx.js.
 * Данные хранятся в одном JSON: на телефоне — в файле приложения (через AndroidBridge),
 * в обычном браузере (для проверки) — в localStorage.
 */
(function () {
  "use strict";
  const C = window.Calc;
  const bridge = window.AndroidBridge || null;
  const $ = id => document.getElementById(id);

  // =============================================================== данные
  // Формат данных и их проверка — в sync.js (общий с программой для ПК)
  const DEFAULTS = () => Sync.sanitize({ days: {} });
  let data = DEFAULTS();

  function load() {
    let raw = null;
    try { raw = bridge ? bridge.load() : localStorage.getItem("tabel"); } catch (e) { raw = null; }
    if (raw) {
      try { data = Sync.sanitize(JSON.parse(raw)); } catch (e) { data = DEFAULTS(); }
    }
  }
  const touchSettings = () => { data.settingsT = Date.now(); };
  function save() {
    const json = JSON.stringify(data);
    try { if (bridge) bridge.save(json); else localStorage.setItem("tabel", json); }
    catch (e) { toast("Не удалось сохранить: " + e.message); }
  }
  function ensureSeeded(year) {
    if (data.seeded[year] || !C.hasOfficialCalendar(year)) return;
    for (const [d, kind, title] of C.defaultSpecialDays(year)) data.special[d] = { kind, title };
    data.seeded[year] = true;
    save();
  }
  const specialKinds = () => Object.fromEntries(Object.entries(data.special).map(([d, v]) => [d, v.kind]));
  const specialTitles = () => Object.fromEntries(Object.entries(data.special).map(([d, v]) => [d, v.title]));
  const monthKey = (y, m) => `${y}-${String(m).padStart(2, "0")}`;
  // Правила нормы; для режима «Часов в месяц» — с учётом нормы, заданной на этот месяц
  function rules(y, m) {
    const s = data.settings;
    let monthMin = s.monthMin;
    if (y && m && s.normMode === "month" && data.monthNorms[monthKey(y, m)] !== undefined) {
      monthMin = data.monthNorms[monthKey(y, m)];
    }
    return C.makeRules(s.workMin, s.lunchMin, s.normMode, s.rate, s.weekMin, monthMin);
  }
  const monthRecords = (y, m) => {
    const pre = `${y}-${String(m).padStart(2, "0")}-`, out = {};
    for (const [d, r] of Object.entries(data.days)) if (d.startsWith(pre) && (r.in || r.out || r.mark)) out[d] = r;
    return out;
  };
  // Каждая правка дня запоминает время (t) — по нему объединяются данные с ПК.
  // Очищенный день остаётся пустой записью, чтобы очистка тоже переехала на ПК.
  function setDay(day, rec) {
    const empty = !rec.in && !rec.out && !rec.mark;
    if (empty && !data.days[day]) return;
    data.days[day] = { in: rec.in || "", out: rec.out || "", mark: rec.mark || "", t: Date.now() };
    save();
  }
  const vacationDates = () => Object.entries(data.days).filter(([, r]) => r.mark === "vacation").map(([d]) => d);
  function vacStatus() {
    const s = data.settings;
    return C.vacationStatus({ hireDate: s.hireDate, perYear: +s.perYear || 28, usedBefore: +s.usedBefore || 0,
      excluded: +s.excluded || 0 }, vacationDates(), specialKinds(), C.todayISO());
  }

  // =============================================================== вид
  const view = { y: 0, m: 0, page: "sheet", editing: null, summary: null };

  function nowHM() {
    const t = new Date();
    return `${String(t.getHours()).padStart(2, "0")}:${String(t.getMinutes()).padStart(2, "0")}`;
  }
  function longDate(s) {
    const p = C.parts(s);
    return `${C.WEEKDAYS[C.weekday(s)].toLowerCase()}, ${p.d} ${C.MONTHS_GEN[p.m - 1]}`;
  }
  function el(tag, cls, text) {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text !== undefined) e.textContent = text;
    return e;
  }

  // ---------------------------------------------------------------- табель
  function renderSheet(changed) {
    ensureSeeded(view.y);
    const today = C.todayISO();
    const s = C.summarizeMonth(view.y, view.m, specialKinds(), monthRecords(view.y, view.m), specialTitles(), rules(view.y, view.m));
    view.summary = s;
    $("month-title").textContent = `${C.MONTHS[view.m - 1]} ${view.y}`;

    // --- сегодня ---
    const tp = C.parts(today);
    ensureSeeded(tp.y);
    const isCurrent = tp.y === view.y && tp.m === view.m;
    $("to-today").hidden = isCurrent;
    $("today-label").textContent = "Сегодня, " + longDate(today);
    const tType = C.dayType(today, specialKinds());
    const tRec = data.days[today] || { in: "", out: "", mark: "" };
    let tText;
    if (tType !== C.DAY.WORK) tText = (specialTitles()[today] || "Выходной") + " — время не учитывается";
    else if (tRec.mark) tText = C.MARK_LABELS[tRec.mark];
    else tText = `Пришёл ${tRec.in || "—"}, ушёл ${tRec.out || "—"}`;
    $("today-times").textContent = tText;
    const canPunch = tType === C.DAY.WORK && !tRec.mark;
    $("punch-in").disabled = !canPunch;
    $("punch-out").disabled = !canPunch;

    // --- итоги ---
    const r = s.rules;
    $("s-remaining").textContent = C.fmtMinutes(Math.max(0, s.remaining));
    let normText;
    if (r.mode === "month") normText = `норма на месяц ${C.fmtMinutes(r.monthMin)} — изменить: нажмите «Норма»`;
    else {
      normText = `${s.normDays} ${C.plural(s.normDays, "рабочий день", "рабочих дня", "рабочих дней")} × ${C.fmtMinutes(s.dayCost)}`;
      if (r.flexible) normText += ` (${r.describe()})`;
    }
    if (r.flexible) normText += ". Переработка — по итогам месяца";
    $("s-remaining-sub").textContent = s.remaining <= 0
      ? (s.remaining < 0 ? `норма выполнена, сверх нормы ${C.fmtMinutes(-s.remaining)}` : "норма выполнена")
      : normText;
    $("s-norm").textContent = C.fmtMinutes(s.norm);
    $("s-worked").textContent = C.fmtMinutes(s.worked);
    $("s-over").textContent = C.fmtMinutes(Math.max(0, s.balance));
    $("s-off").textContent = String(s.availableOffDays);
    const vs = vacStatus();
    $("s-vac").textContent = vs.configured ? C.fmtDays(vs.available) : "—";

    const banner = C.bannerText(s);
    $("banner").hidden = !banner;
    $("banner").textContent = banner;
    $("banner").classList.toggle("warn", s.balance < 0);

    // --- дни ---
    const list = $("days");
    list.textContent = "";
    for (const row of s.rows) list.appendChild(dayItem(row, row.day === today));

    // --- уведомление о пороге (один раз на каждый новый «рабочий день» переработки) ---
    const key = `${view.y}-${String(view.m).padStart(2, "0")}`;
    const shown = data.notified[key] || 0;
    if (s.earnedLevel > shown) {
      data.notified[key] = s.earnedLevel;
      save();
      if (changed) {
        modal({ figure: C.fmtMinutes(s.earnedLevel * s.dayCost), title: "Можно взять отгул",
          text: C.popupText(s), buttons: [{ label: "Понятно", primary: true }] });
      }
    }
  }

  function dayItem(row, isToday) {
    const li = el("li");
    const b = el("button", "day");
    const rec = row.record, res = row.result;
    const off = row.type !== C.DAY.WORK;
    if (off) b.classList.add("off");
    if (row.type === C.DAY.HOLIDAY) b.classList.add("holiday");
    if (isToday) b.classList.add("today");

    const left = el("div");
    left.appendChild(el("div", "day-num", String(C.parts(row.day).d)));
    left.appendChild(el("div", "day-wd", C.WEEKDAYS_SHORT[C.weekday(row.day)]));
    const main = el("div", "day-main");
    const delta = el("div", "day-delta");

    if (off) {
      if (rec.mark === "vacation") main.appendChild(el("span", "chip", "Отпуск"));
      else main.appendChild(el("div", "day-times", row.title || "Выходной"));
    } else if (rec.mark) {
      main.appendChild(el("span", "chip" + (rec.mark === "off" ? " offday" : ""), C.MARK_LABELS[rec.mark]));
    } else {
      const t = el("div", "day-times");
      if (rec.in || rec.out) t.textContent = `${rec.in || "—"} – ${rec.out || "…"}`;
      else { t.textContent = "не заполнено"; t.classList.add("empty"); }
      main.appendChild(t);
      if (res.complete) main.appendChild(el("div", "day-info", "отработано " + C.fmtMinutes(res.worked)));
      else if (row.title) main.appendChild(el("div", "day-info", row.title));
      if (res.overtime) { delta.textContent = "+" + C.fmtMinutes(res.overtime); delta.classList.add("over"); }
      else if (res.undertime) { delta.textContent = "−" + C.fmtMinutes(res.undertime); delta.classList.add("under"); }
      else if (isToday && rec.in && !rec.out) { delta.textContent = "на работе"; delta.style.fontSize = "13px"; delta.style.color = "var(--accent)"; }
    }
    b.append(left, main, delta);
    b.addEventListener("click", () => openEditor(row.day));
    li.appendChild(b);
    return li;
  }

  function goMonth(delta) {
    [view.y, view.m] = C.shiftMonth(view.y, view.m, delta);
    renderSheet();
    window.scrollTo(0, 0);
  }
  function goToday(scroll) {
    const p = C.parts(C.todayISO());
    view.y = p.y; view.m = p.m;
    renderSheet();
    if (scroll) scrollToToday();
  }
  function scrollToToday() {
    const btn = document.querySelector(".day.today");
    if (btn) btn.scrollIntoView({ block: "center" });
  }

  // ---------------------------------------------------------------- Пришёл / Ушёл
  async function punch(which) {
    const today = C.todayISO();
    const p = C.parts(today);
    if (view.y !== p.y || view.m !== p.m) { view.y = p.y; view.m = p.m; }
    const rec = Object.assign({ in: "", out: "", mark: "" }, data.days[today]);
    const now = nowHM();
    if (which === "in") {
      if (rec.in && !(await confirmBox("Приход уже отмечен", `Сегодня приход записан: ${rec.in}. Заменить на ${now}?`, "Заменить"))) return;
      rec.in = now;
    } else {
      if (!rec.in) { await modal({ title: "Нет времени прихода", text: "Сначала нажмите «Пришёл» или введите время прихода.", buttons: [{ label: "Понятно", primary: true }] }); return; }
      if (rec.out && !(await confirmBox("Уход уже отмечен", `Сегодня уход записан: ${rec.out}. Заменить на ${now}?`, "Заменить"))) return;
      rec.out = now;
    }
    try { C.validatePair(rec.in, rec.out); }
    catch (e) { await modal({ title: "Неверное время", text: e.message, buttons: [{ label: "Понятно", primary: true }] }); return; }
    setDay(today, rec);
    renderSheet(true);
    toast((which === "in" ? "Приход" : "Уход") + " отмечен: " + now);
  }

  // =============================================================== редактор дня
  function openEditor(day) {
    view.editing = day;
    $("ed-err").textContent = "";
    $("ed-in").classList.remove("bad"); $("ed-out").classList.remove("bad");
    fillEditor();
    $("backdrop").hidden = false;
    $("sheet").hidden = false;
  }
  function closeEditor() {
    if (view.editing) commitTimes(true);
    view.editing = null;
    $("sheet").hidden = true;
    $("backdrop").hidden = true;
    renderSheet(true);
  }
  function fillEditor() {
    const day = view.editing;
    const type = C.dayType(day, specialKinds());
    const work = type === C.DAY.WORK;
    const rec = data.days[day] || { in: "", out: "", mark: "" };
    const t = longDate(day);
    $("ed-title").textContent = t.charAt(0).toUpperCase() + t.slice(1);
    const title = specialTitles()[day];
    $("ed-sub").textContent = work ? (title || (C.weekday(day) >= 5 ? "Рабочий день (перенос)" : ""))
      : (title || "Выходной") + " — можно отметить только отпуск";
    const timesOn = work && !rec.mark;
    $("ed-times").hidden = !timesOn;
    $("ed-now").hidden = !timesOn;
    $("ed-in").value = rec.in || "";
    $("ed-out").value = rec.out || "";
    renderResult();

    const chips = $("ed-marks");
    chips.textContent = "";
    const options = work ? ["", "off", "vacation", "sick"] : ["", "vacation"];
    for (const mk of options) {
      const b = el("button", "", mk ? C.MARK_LABELS[mk] : (work ? "Без отметки" : (title || "Выходной")));
      if ((rec.mark || "") === mk) b.classList.add("on");
      if (mk === "off") b.classList.add("offday");
      b.addEventListener("click", () => applyMark(mk));
      chips.appendChild(b);
    }
  }
  function renderResult() {
    const day = view.editing;
    const rec = data.days[day] || { in: "", out: "", mark: "" };
    const type = C.dayType(day, specialKinds());
    const box = $("ed-result");
    if (type !== C.DAY.WORK) { box.textContent = rec.mark === "vacation" ? "День отпуска" : "Нерабочий день"; return; }
    if (rec.mark) { box.textContent = `${C.MARK_LABELS[rec.mark]} — время не учитывается`; return; }
    let res;
    const p = C.parts(day);
    const r = rules(p.y, p.m);
    try { res = C.calcDay(rec.in, rec.out, r); } catch (e) { res = null; }
    if (!res || !res.complete) {
      box.textContent = rec.in ? "Укажите время ухода — и день посчитается" : "Укажите время прихода и ухода";
      return;
    }
    box.innerHTML = "";
    box.append("Отработано ", el("b", "", C.fmtMinutes(res.worked)));
    if (res.overtime) { box.append(", переработка "); const x = el("b", "", "+" + C.fmtMinutes(res.overtime)); x.style.color = "var(--over)"; box.append(x); }
    if (res.undertime) { box.append(", недоработка "); const x = el("b", "", "−" + C.fmtMinutes(res.undertime)); x.style.color = "var(--under)"; box.append(x); }
    if (r.flexible) box.append(el("div", "hint", "Переработка у полуставки считается по итогам месяца."));
  }
  function commitTimes(silent) {
    const day = view.editing;
    if (!day || C.dayType(day, specialKinds()) !== C.DAY.WORK) return true;
    const rec = Object.assign({ in: "", out: "", mark: "" }, data.days[day]);
    if (rec.mark) return true;
    const a = $("ed-in").value, b = $("ed-out").value;
    $("ed-in").classList.remove("bad"); $("ed-out").classList.remove("bad");
    try {
      C.validatePair(a, b);
    } catch (e) {
      $("ed-err").textContent = e.message;
      (a ? $("ed-out") : $("ed-in")).classList.add("bad");
      if (silent) toast("Время не сохранено: " + e.message);
      return false;
    }
    $("ed-err").textContent = "";
    if (a !== rec.in || b !== rec.out) {
      rec.in = a; rec.out = b;
      setDay(day, rec);
      if (!silent) { renderResult(); renderSheet(true); }
    }
    return true;
  }
  async function applyMark(mark) {
    const day = view.editing;
    const rec = Object.assign({ in: "", out: "", mark: "" }, data.days[day]);
    if ((rec.mark || "") === mark) return;
    const p = C.parts(day);
    try {
      C.checkMarkChange(p.y, p.m, specialKinds(), monthRecords(p.y, p.m), day, mark, rules(p.y, p.m));
    } catch (e) { $("ed-err").textContent = e.message; return; }
    if (mark && (rec.in || rec.out)) {
      const ok = await confirmBox("Время будет удалено",
        `За этот день введено время (${[rec.in, rec.out].filter(Boolean).join(" – ")}). С отметкой «${C.MARK_LABELS[mark]}» оно не учитывается и будет удалено.`, "Поставить отметку");
      if (!ok) return;
    }
    $("ed-err").textContent = "";
    setDay(day, { in: "", out: "", mark });
    fillEditor();
    renderSheet(true);
    if (mark === "off") toast(`Отгул: списано ${C.fmtMinutes(view.summary.dayCost)} переработки`);
    else if (mark === "vacation" || rec.mark === "vacation") toast("Остаток отпуска пересчитан");
  }
  async function clearDay() {
    const day = view.editing;
    if (!data.days[day]) return;
    if (!(await confirmBox("Очистить день?", "Время и отметка за этот день будут удалены.", "Очистить"))) return;
    setDay(day, { in: "", out: "", mark: "" });
    fillEditor();
    renderSheet(true);
  }

  // =============================================================== окна
  let modalResolve = null;
  function modal({ title = "", text = "", figure = "", buttons = [], input = null, validate = null }) {
    return new Promise(resolve => {
      if (modalResolve) modalResolve(null);
      modalResolve = resolve;
      $("modal-figure").textContent = figure;
      $("modal-title").textContent = title;
      $("modal-text").textContent = text;
      const inp = $("modal-input");
      inp.hidden = input === null;
      inp.classList.remove("bad");
      if (input !== null) inp.value = input;
      const box = $("modal-actions");
      box.textContent = "";
      for (const b of buttons) {
        const btn = el("button", "btn" + (b.primary ? " primary" : ""), b.label);
        btn.addEventListener("click", () => {
          if (b.value === "input") {
            // Кнопка «Сохранить» у окна с полем: проверить ввод и вернуть значение
            try {
              const v = validate ? validate(inp.value) : inp.value;
              closeModal(v === null ? "" : v);
            } catch (e) {
              inp.classList.add("bad");
              $("modal-text").textContent = e.message;
            }
            return;
          }
          closeModal(b.value === undefined ? true : b.value);
        });
        box.appendChild(btn);
      }
      $("modal").hidden = false;
      $("backdrop").hidden = false;
      $("backdrop").style.zIndex = 25;
    });
  }
  function closeModal(value) {
    $("modal").hidden = true;
    $("backdrop").style.zIndex = "";
    if ($("sheet").hidden) $("backdrop").hidden = true;
    const r = modalResolve; modalResolve = null;
    if (r) r(value);
  }
  const confirmBox = (title, text, yes) => modal({ title, text,
    buttons: [{ label: "Отмена", value: false }, { label: yes, primary: true, value: true }] });

  let toastTimer = null;
  function toast(text) {
    const t = $("toast");
    t.textContent = text; t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { t.hidden = true; }, 2600);
  }

  // =============================================================== настройки
  const holView = { y: 0 };

  function renderSettings() {
    const s = data.settings;
    $("set-mode").value = s.normMode;
    $("set-work").value = C.fmtMinutes(s.workMin);
    $("set-lunch").value = C.fmtMinutes(s.lunchMin);
    $("set-rate").value = C.fmtRate(s.rate);
    $("set-week").value = C.fmtMinutes(s.weekMin);
    $("set-month").value = C.fmtMinutes(s.monthMin);
    $("rules-msg").style.color = "";
    $("rules-msg").textContent = rulesSummary(rules());
    showModeFields();
    $("set-hire").value = s.hireDate || "";
    $("set-peryear").value = s.perYear;
    $("set-before").value = s.usedBefore;
    $("set-excl").value = s.excluded;
    renderVacReport();
    renderHolidays();
    for (const b of $("theme-seg").children) b.classList.toggle("on", b.dataset.v === s.theme);
    $("set-font").value = s.font;
  }
  function rulesSummary(r) {
    return r.flexible ? `Режим: ${r.describe()}. Переработка — по итогам месяца.`
      : `Норма присутствия в день: ${C.fmtMinutes(r.presence)}`;
  }
  function showModeFields() {
    const mode = $("set-mode").value;
    $("row-rate").hidden = mode !== "rate";
    $("row-week").hidden = mode !== "week";
    $("row-month").hidden = mode !== "month";
    $("mode-hint").textContent = {
      day: "",
      rate: "Норма дня = работа в день × доля ставки (8:00 × 0,5 = 4:00), норма месяца — по рабочим дням календаря.",
      week: "Норма дня = часы в неделю ÷ 5 (20:00 → 4:00), норма месяца — по рабочим дням календаря.",
      month: "Эта норма — на каждый месяц. Для отдельного месяца её можно поменять: в табеле нажмите «Норма».",
    }[mode];
  }
  function saveRules() {
    const s = data.settings;
    const mode = $("set-mode").value;
    let next;
    try {
      const w = C.parseDuration($("set-work").value || "0:00");
      const l = C.parseDuration($("set-lunch").value || "0:00");
      let rate = s.rate, week = s.weekMin, month = s.monthMin;
      if (mode === "rate") {
        rate = Number(($("set-rate").value || "").replace(",", "."));
        if (!isFinite(rate)) throw new C.InputError("Доля ставки должна быть числом, например 0,5.");
      }
      if (mode === "week") week = C.parseDuration($("set-week").value || "0:00");
      if (mode === "month") month = C.parseDuration($("set-month").value || "0:00");
      next = C.makeRules(w, l, mode, rate, week, month);
      C.validateRules(next);
    } catch (e) { $("rules-msg").textContent = e.message; $("rules-msg").style.color = "var(--under)"; return; }
    touchSettings();
    Object.assign(s, { workMin: next.workMin, lunchMin: next.lunchMin, normMode: next.mode,
      rate: next.rate, weekMin: next.weekMin, monthMin: next.monthMin });
    save();
    renderSettings();
    $("rules-msg").textContent = "Сохранено. " + rulesSummary(next);
  }
  // Норма на конкретный месяц (режим «Часов в месяц»)
  async function editMonthNorm() {
    const s = data.settings;
    if (s.normMode !== "month") {
      showPage("settings");
      setTimeout(() => $("workday-section").scrollIntoView(), 50);
      return;
    }
    const key = monthKey(view.y, view.m);
    const own = data.monthNorms[key];
    const value = await modal({
      title: `Норма на ${C.MONTHS[view.m - 1].toLowerCase()}`,
      text: `Сколько часов нужно отработать в этом месяце. Обычная норма — ${C.fmtMinutes(s.monthMin)}. ` +
        "Пустое поле — вернуть обычную.",
      input: own !== undefined ? C.fmtMinutes(own) : C.fmtMinutes(s.monthMin),
      validate: text => {
        if (!text.trim()) return null;
        const v = C.parseDuration(text);
        C.validateRules(C.makeRules(480, 0, "month", 0.5, 1200, v));
        return v;
      },
      buttons: [{ label: "Отмена", value: false }, { label: "Сохранить", primary: true, value: "input" }],
    });
    if (value === false || value === null || value === undefined) return;
    touchSettings();
    if (value === "") delete data.monthNorms[key];
    else data.monthNorms[key] = value;
    save();
    renderSheet(true);
    toast(value === "" ? "Норма месяца — обычная" : `Норма на месяц: ${C.fmtMinutes(value)}`);
  }
  function saveVacation() {
    const s = data.settings, err = $("vac-err");
    const num = (id, name, min, max) => {
      const raw = ($(id).value || "0").replace(",", ".");
      const v = Number(raw);
      if (!isFinite(v) || v < min || v > max) throw new C.InputError(`«${name}»: введите число от ${min} до ${max}.`);
      return v;
    };
    try {
      const hire = $("set-hire").value;
      if (hire && C.diffDays(hire, C.todayISO()) > 0) throw new C.InputError("Дата начала работы не может быть в будущем.");
      const perYear = num("set-peryear", "Дней отпуска в году", 1, 100);
      const before = num("set-before", "Использовано до начала учёта", 0, 10000);
      const excl = Math.round(num("set-excl", "Не входит в стаж", 0, 10000));
      Object.assign(s, { hireDate: hire, perYear, usedBefore: before, excluded: excl });
      touchSettings();
    } catch (e) { err.textContent = e.message; return; }
    err.textContent = "";
    save();
    renderVacReport();
  }
  function renderVacReport() {
    const st = vacStatus(), box = $("vac-report");
    box.textContent = "";
    if (!st.configured) { box.textContent = "Укажите дату начала работы — и программа посчитает, сколько дней отпуска вам положено."; return; }
    box.appendChild(el("div", "big", C.fmtDays(st.available) + " " + C.plural(Math.round(st.available), "день", "дня", "дней")));
    const lines = [
      `Отпускной стаж: ${st.months} ${C.plural(st.months, "месяц", "месяца", "месяцев")}`,
      `Начислено: ${C.fmtDays(st.accrued)}`,
    ];
    if (st.usedBefore) lines.push(`Использовано до программы: ${C.fmtDays(st.usedBefore)}`);
    lines.push(`Отгуляно по табелю: ${st.usedPast}`);
    if (st.planned) lines.push(`Запланировано: ${st.planned}`);
    for (const l of lines) box.appendChild(el("div", "", l));
    if (st.periods.length) {
      box.appendChild(el("div", "hint", "Отпуска в табеле:"));
      for (const [a, b, n] of st.periods.slice(-6)) {
        box.appendChild(el("div", "", `${C.fmtDate(a)}${a === b ? "" : " – " + C.fmtDate(b)}: ${C.daysPhrase(n)}`));
      }
    }
  }
  function renderHolidays() {
    ensureSeeded(holView.y);
    $("hol-year").textContent = holView.y;
    $("hol-hint").textContent = C.hasOfficialCalendar(holView.y)
      ? "Календарь РФ на этот год встроен (с переносами по постановлению Правительства)."
      : "Для этого года переносы не встроены: кнопка внизу добавит праздники из ТК РФ, переносы добавьте сами.";
    $("hol-reset").textContent = `Заполнить календарём РФ на ${holView.y}`;
    const list = $("hol-list");
    list.textContent = "";
    const items = Object.entries(data.special).filter(([d]) => d.startsWith(holView.y + "-")).sort();
    if (!items.length) list.appendChild(el("li", "hint", "Особых дней нет — считаются только субботы и воскресенья."));
    for (const [d, v] of items) {
      const li = el("li");
      const info = el("div");
      info.appendChild(el("div", "hol-date", `${C.fmtDate(d)}, ${C.WEEKDAYS_SHORT[C.weekday(d)].toLowerCase()}`));
      info.appendChild(el("div", "hol-title" + (v.kind === "workday" ? " work" : ""), v.kind === "workday" ? "Рабочий день: " + v.title : v.title));
      const del = el("button", "hol-del", "Удалить");
      del.addEventListener("click", () => { delete data.special[d]; save(); renderHolidays(); });
      li.append(info, del);
      list.appendChild(li);
    }
  }
  function addHoliday() {
    const d = $("hol-date").value, kind = $("hol-kind").value;
    const err = $("hol-err");
    if (!d) { err.textContent = "Выберите дату."; return; }
    if (kind === "workday" && C.weekday(d) < 5) { err.textContent = "Перенесённым рабочим днём бывает только суббота или воскресенье."; return; }
    data.special[d] = { kind, title: $("hol-title").value.trim() || (kind === "workday" ? "Рабочий день (перенос)" : "Праздник") };
    err.textContent = "";
    $("hol-date").value = ""; $("hol-title").value = "";
    holView.y = C.parts(d).y;
    save();
    renderHolidays();
  }
  async function resetHolidays() {
    const y = holView.y;
    const has = Object.keys(data.special).some(d => d.startsWith(y + "-"));
    if (has && !(await confirmBox("Заполнить календарь", `Праздники и переносы ${y} года будут заменены календарём РФ. Ваши изменения за этот год пропадут.`, "Заменить"))) return;
    for (const d of Object.keys(data.special)) if (d.startsWith(y + "-")) delete data.special[d];
    for (const [d, kind, title] of C.defaultSpecialDays(y)) data.special[d] = { kind, title };
    data.seeded[y] = true;
    save();
    renderHolidays();
  }

  // ---------------------------------------------------------------- тема и шрифт
  function systemNight() {
    try { if (bridge && bridge.isNight) return bridge.isNight(); } catch (e) { /* нет моста */ }
    return window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
  }
  function applyTheme() {
    const t = data.settings.theme;
    const dark = t === "dark" || (t === "system" && systemNight());
    document.documentElement.dataset.theme = dark ? "dark" : "light";
    try { if (bridge && bridge.setBars) bridge.setBars(dark ? "#14171D" : "#ECEEF2", !dark); } catch (e) { /* ок */ }
  }
  function applyFont() {
    const map = { Monomakh: '"Monomakh"', Shafarik: '"Shafarik"', Ruslan: '"Ruslan"', OldStandard: '"OldStandard"', system: "Roboto" };
    document.documentElement.style.setProperty("--display", `${map[data.settings.font] || map.Monomakh}, Georgia, serif`);
  }

  // ---------------------------------------------------------------- файлы
  function saveFile(name, mime, bytes) {
    if (bridge && bridge.saveFile) { bridge.saveFile(name, mime, Xlsx.toBase64(bytes)); return; }
    // iPhone и телефоны без моста: системное меню «Поделиться» (Сохранить в «Файлы», Telegram, почта…)
    try {
      const file = new File([bytes], name, { type: mime });
      if (navigator.canShare && navigator.canShare({ files: [file] })) {
        navigator.share({ files: [file], title: name }).then(
          () => toast("Готово"),
          e => { if (e && e.name !== "AbortError") download(name, mime, bytes); });
        return;
      }
    } catch (e) { /* нет поддержки — просто скачаем */ }
    download(name, mime, bytes);
  }
  function download(name, mime, bytes) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([bytes], { type: mime }));
    a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 60000);
  }
  window.onFileSaved = (ok, msg) => toast(ok ? "Файл сохранён" : (msg || "Файл не сохранён"));
  function exportMonth() {
    const bytes = Xlsx.monthXlsx(view.summary, view.y, view.m);
    saveFile(`Табель_${view.y}-${String(view.m).padStart(2, "0")}_${C.MONTHS[view.m - 1]}.xlsx`,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", bytes);
  }
  function backup() {
    const out = Object.assign({}, data, { exported: Date.now() });
    const bytes = new TextEncoder().encode(JSON.stringify(out, null, 1));
    saveFile(`Табель_${C.todayISO()}.json`, "application/json", bytes);
  }
  function restore() {
    if (bridge && bridge.openFile) { bridge.openFile("*/*"); return; }
    const inp = document.createElement("input");
    inp.type = "file"; inp.accept = ".json,application/json";
    inp.onchange = () => inp.files[0] && inp.files[0].text().then(window.onFileOpened);
    inp.click();
  }
  // Файл с ПК или резервная копия: объединить (по дням побеждает более поздняя правка) или заменить всё
  window.onFileOpened = async text => {
    let incoming;
    try {
      if (typeof text !== "string" || text.length > 20000000) throw new Sync.SyncError("Файл слишком большой.");
      incoming = Sync.sanitize(JSON.parse(text.replace(/^\uFEFF/, "")));
    } catch (e) { toast(e instanceof Sync.SyncError ? e.message : "Это не файл Табеля"); return; }
    const { data: merged, stats } = Sync.merge(data, incoming);
    const n = Object.values(incoming.days).filter(r => r.in || r.out || r.mark).length;
    const choice = await modal({
      title: "Загрузить данные?",
      text: `В файле ${n} ${C.plural(n, "заполненный день", "заполненных дня", "заполненных дней")}.\n\n` +
        `«Объединить» — добавить новое и обновить изменённое (по каждому дню остаётся более поздняя правка): ${Sync.statsText(stats)}.\n\n` +
        "«Заменить всё» — стереть данные на телефоне и взять всё из файла.",
      buttons: [{ label: "Отмена", value: false }, { label: "Заменить всё", value: "replace" },
        { label: "Объединить", primary: true, value: "merge" }],
    });
    if (choice !== "merge" && choice !== "replace") return;
    if (choice === "replace" && !(await confirmBox("Заменить всё?", "Все записи на телефоне будут заменены данными из файла.", "Заменить"))) return;
    const keep = { theme: data.settings.theme, font: data.settings.font };   // тема и шрифт — свои
    data = choice === "merge" ? merged : incoming;
    Object.assign(data.settings, keep);
    save(); applyTheme(); applyFont(); renderSettings(); renderSheet();
    toast(choice === "merge" ? "Объединено: " + Sync.statsText(stats) : "Данные заменены");
  };

  // =============================================================== планировщик отпуска
  // Настройки планировщика и оклад — только на этом устройстве (не в файле синхронизации).
  const PL_KEY = "tabel-planner";
  const plDefaults = () => ({ budget: null, min: 3, max: 14, priority: "rest", need14: true, salary: 0, avg: 0 });
  let pl = plDefaults();
  function plLoad() {
    try {
      const raw = JSON.parse(localStorage.getItem(PL_KEY) || "{}");
      const n = (v, lo, hi, d) => (typeof v === "number" && isFinite(v) && v >= lo && v <= hi ? v : d);
      pl = {
        budget: raw.budget === null ? null : n(raw.budget, 1, 200, null),
        min: n(raw.min, 1, 60, 3), max: n(raw.max, 1, 60, 14),
        priority: raw.priority === "money" ? "money" : "rest", need14: raw.need14 !== false,
        salary: n(raw.salary, 0, 1e8, 0), avg: n(raw.avg, 0, 1e7, 0),
      };
    } catch (e) { pl = plDefaults(); }
  }
  function plSave() { try { localStorage.setItem(PL_KEY, JSON.stringify(pl)); } catch (e) { /* ок */ } }

  const fmtRub = v => {
    const s = String(Math.abs(Math.round(v))).replace(/\B(?=(\d{3})+(?!\d))/g, "\u00a0");
    return (v < 0 ? "−" : "") + s + "\u00a0₽";
  };
  function fmtRange(a, b) {
    const pa = C.parts(a), pb = C.parts(b);
    if (a === b) return `${pa.d} ${C.MONTHS_GEN[pa.m - 1]} ${pa.y}`;
    if (pa.y === pb.y && pa.m === pb.m) return `${pa.d}–${pb.d} ${C.MONTHS_GEN[pa.m - 1]} ${pa.y}`;
    if (pa.y === pb.y) return `${pa.d} ${C.MONTHS_GEN[pa.m - 1]} – ${pb.d} ${C.MONTHS_GEN[pb.m - 1]} ${pa.y}`;
    return `${pa.d} ${C.MONTHS_GEN[pa.m - 1]} ${pa.y} – ${pb.d} ${C.MONTHS_GEN[pb.m - 1]} ${pb.y}`;
  }
  function fmtShort(a, b) {
    const pa = C.parts(a), pb = C.parts(b);
    if (pa.m === pb.m && pa.y === pb.y) return `${pa.d}–${pb.d} ${C.MONTHS_GEN[pa.m - 1]}`;
    return `${pa.d} ${C.MONTHS_GEN[pa.m - 1]} – ${pb.d} ${C.MONTHS_GEN[pb.m - 1]}`;
  }
  // Календарь для планирования: свои особые дни + праздники ТК для лет, которые ещё не заполнены
  function planSpecial(years) {
    for (const y of years) ensureSeeded(y);
    const out = {};
    for (const y of years) {
      if (!data.seeded[y]) for (const [d, k] of C.defaultSpecialDays(y)) out[d] = k;
    }
    return Object.assign(out, specialKinds());
  }
  function availableDays() {
    const st = vacStatus();
    return st.configured ? Math.max(0, Math.floor(st.available)) : 28;
  }

  function renderPlanner() {
    const today = C.todayISO();
    const y = C.parts(today).y;
    const special = planSpecial([y, y + 1]);
    const busy = new Set(Object.entries(data.days).filter(([, r]) => r.mark).map(([d]) => d));
    const budget = pl.budget === null ? availableDays() : pl.budget;
    $("pl-budget").value = budget;
    $("pl-min").value = pl.min; $("pl-max").value = Math.max(pl.min, pl.max);
    $("pl-salary").value = pl.salary || "";
    $("pl-avg").value = pl.avg || "";
    $("pl-14").checked = pl.need14;
    for (const b of $("pl-priority").children) b.classList.toggle("on", b.dataset.v === pl.priority);
    const salary = pl.salary > 0 ? pl.salary : null;
    const hints = [];
    if (pl.budget === null) hints.push(vacStatus().configured ? "Число дней — ваш остаток отпуска из счётчика." :
      "Укажите в «Настройках» дату начала работы — и здесь будет ваш остаток отпуска; пока взято 28.");
    if (pl.priority === "money") {
      hints.push(salary ? "По деньгам выгоднее брать отпуск целыми неделями в месяцы, где больше рабочих дней: план строится из периодов от 7 дней, остаток — короткими."
        : "Чтобы считать деньги, укажите оклад. Пока план — по отдыху.");
    }
    $("pl-hint").textContent = hints.join(" ");

    const r = Planner.planVacation(today, special, budget, pl.min, Math.max(pl.min, pl.max), pl.priority, pl.need14,
      busy, salary, pl.avg > 0 ? pl.avg : null);
    // --- итог ---
    const sum = $("pl-summary");
    sum.textContent = "";
    const charged = r.plan.reduce((a, b) => a + b.charged, 0);
    const rest = r.plan.reduce((a, b) => a + b.rest, 0);
    if (!r.plan.length) {
      sum.appendChild(el("div", "big", "Нет подходящих периодов"));
      sum.appendChild(el("p", "", budget < pl.min ? `Дней меньше, чем «за один раз от ${pl.min}». Уменьшите минимум.` :
        "Попробуйте изменить «за один раз» или число дней."));
    } else {
      sum.appendChild(el("div", "big", `${charged} ${C.plural(charged, "день", "дня", "дней")} отпуска → ${rest} ${C.plural(rest, "день", "дня", "дней")} отдыха`));
      sum.appendChild(el("p", "", `${r.plan.length} ${C.plural(r.plan.length, "часть", "части", "частей")}, в среднем ×${(rest / charged).toFixed(1).replace(".", ",")}` +
        (r.left ? `. Не распределено: ${r.left} ${C.plural(r.left, "день", "дня", "дней")} — добавьте к любой части или возьмите позже.` : ".")));
      if (salary) {
        const d = r.plan.reduce((a, b) => a + b.delta, 0);
        const p = el("p", "money-total", `Доход месяцев с отпуском изменится всего на ${d >= 0 ? "+" : ""}${fmtRub(d)}`);
        sum.appendChild(p);
      }
    }
    // --- списки ---
    const fill = (list, items, isPlan) => {
      list.textContent = "";
      items.forEach(b => list.appendChild(planItem(b, isPlan)));
    };
    fill($("pl-list"), r.plan, true);
    fill($("pl-alts"), r.alternatives, false);
    $("pl-alt-title").hidden = !r.alternatives.length;
  }

  function planItem(b, isPlan) {
    const li = el("li", "pl-item");
    const head = el("div", "pl-head");
    const dates = el("div", "pl-dates", fmtRange(b.start, b.end));
    if (b.charged >= 14 && isPlan) dates.appendChild(el("span", "pl-tag", "основная часть"));
    head.append(dates, el("span", "pl-ratio", "×" + b.ratio.toFixed(1).replace(".", ",")));
    const line = el("div", "pl-line");
    line.append(el("b", "", `${b.charged} ${C.plural(b.charged, "день", "дня", "дней")} отпуска`), " → ",
      el("b", "", `${b.rest} ${C.plural(b.rest, "день", "дня", "дней")} отдыха`),
      b.rest > b.charged ? ` (${fmtShort(b.restStart, b.restEnd)})` : "");
    li.append(head, line);
    if (b.delta !== null) {
      const m = el("div", "pl-money");
      m.append(`Отпускные ≈ ${fmtRub(b.vacPay)}, зарплата меньше на ${fmtRub(b.lost)} → `,
        el("span", b.delta >= 0 ? "plus" : "minus", `${b.delta >= 0 ? "+" : ""}${fmtRub(b.delta)}`), " к доходу");
      li.appendChild(m);
    }
    const btn = el("button", "btn" + (isPlan ? " primary" : ""), "Отметить в табеле");
    btn.addEventListener("click", () => markVacation(b));
    li.appendChild(btn);
    return li;
  }

  async function markVacation(b) {
    const days = [];
    const sp = planSpecial([C.parts(b.start).y, C.parts(b.end).y]);
    for (let d = b.start; C.diffDays(b.end, d) >= 0; d = C.addDays(d, 1)) {
      if (C.dayType(d, sp) === C.DAY.WORK) days.push(d);
    }
    const withTime = days.filter(d => data.days[d] && (data.days[d].in || data.days[d].out));
    const ok = await confirmBox("Отметить отпуск?",
      `${fmtRange(b.start, b.end)}: ${b.charged} ${C.plural(b.charged, "день", "дня", "дней")} отпуска. ` +
      `В табеле на рабочие дни встанет отметка «Отпуск», остаток отпуска пересчитается.` +
      (withTime.length ? ` На ${withTime.length} ${C.plural(withTime.length, "день", "дня", "дней")} уже введено время — оно будет удалено.` : ""),
      "Отметить");
    if (!ok) return;
    for (const d of days) setDay(d, { in: "", out: "", mark: "vacation" });
    if (pl.budget !== null) { pl.budget = Math.max(0, pl.budget - b.charged); plSave(); }
    renderPlanner();
    toast("Отпуск отмечен в табеле");
  }

  function plRead() {
    const num = (id, lo, hi) => { const v = Number(($(id).value || "").replace(",", ".")); return isFinite(v) ? Math.min(hi, Math.max(lo, v)) : null; };
    const budget = num("pl-budget", 0, 200);
    const avail = availableDays();
    pl.budget = budget === null || budget === avail ? null : Math.floor(budget);
    pl.min = Math.floor(num("pl-min", 1, 60) || 1);
    pl.max = Math.max(pl.min, Math.floor(num("pl-max", 1, 60) || pl.min));
    pl.salary = Math.floor(num("pl-salary", 0, 1e8) || 0);
    pl.avg = num("pl-avg", 0, 1e7) || 0;
    pl.need14 = $("pl-14").checked;
    plSave();
    renderPlanner();
  }

  // =============================================================== навигация
  function showPage(page) {
    view.page = page;
    $("page-sheet").hidden = page !== "sheet";
    $("page-settings").hidden = page !== "settings";
    $("page-plan").hidden = page !== "plan";
    for (const b of document.querySelectorAll(".tabbar button")) b.classList.toggle("active", b.dataset.page === page);
    if (page === "settings") { holView.y = view.y; renderSettings(); }
    else if (page === "plan") renderPlanner();
    else renderSheet();
    window.scrollTo(0, 0);
  }
  // Кнопка «Назад» на телефоне: закрыть окно → панель → вернуться в табель → выйти
  window.handleBack = () => {
    if (!$("modal").hidden) { closeModal(false); return true; }
    if (!$("sheet").hidden) { closeEditor(); return true; }
    if (view.page !== "sheet") { showPage("sheet"); return true; }
    return false;
  };
  window.onResumeApp = () => { applyTheme(); if (view.page === "sheet" && $("sheet").hidden) renderSheet(); };
  window.onSystemTheme = () => applyTheme();

  // =============================================================== запуск
  function init() {
    load();
    plLoad();
    // Веб-версия: попросить браузер не очищать данные приложения
    if (!bridge && navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
    applyTheme(); applyFont();
    const p = C.parts(C.todayISO());
    view.y = p.y; view.m = p.m; holView.y = p.y;
    ensureSeeded(p.y);

    $("prev").addEventListener("click", () => goMonth(-1));
    $("next").addEventListener("click", () => goMonth(1));
    $("month-title").addEventListener("click", () => goToday(true));
    $("to-today").addEventListener("click", () => goToday(true));
    $("punch-in").addEventListener("click", () => punch("in"));
    $("punch-out").addEventListener("click", () => punch("out"));
    $("export-btn").addEventListener("click", exportMonth);
    $("stat-vac").addEventListener("click", () => { showPage("settings"); setTimeout(() => $("vac-section").scrollIntoView(), 50); });
    for (const b of document.querySelectorAll(".tabbar button")) b.addEventListener("click", () => showPage(b.dataset.page));

    $("backdrop").addEventListener("click", () => { if (!$("modal").hidden) closeModal(false); else closeEditor(); });
    $("ed-in").addEventListener("change", () => commitTimes(false));
    $("ed-out").addEventListener("change", () => commitTimes(false));
    $("ed-in-now").addEventListener("click", () => { $("ed-in").value = nowHM(); commitTimes(false); });
    $("ed-out-now").addEventListener("click", () => { $("ed-out").value = nowHM(); commitTimes(false); });
    $("ed-done").addEventListener("click", closeEditor);
    $("ed-clear").addEventListener("click", clearDay);

    for (const id of ["set-work", "set-lunch", "set-rate", "set-week", "set-month"]) $(id).addEventListener("change", saveRules);
    $("set-mode").addEventListener("change", () => { showModeFields(); saveRules(); });
    $("stat-norm").addEventListener("click", editMonthNorm);
    for (const id of ["set-hire", "set-peryear", "set-before", "set-excl"]) $(id).addEventListener("change", saveVacation);
    $("hol-prev").addEventListener("click", () => { holView.y--; renderHolidays(); });
    $("hol-next").addEventListener("click", () => { holView.y++; renderHolidays(); });
    $("hol-add").addEventListener("click", addHoliday);
    $("hol-reset").addEventListener("click", resetHolidays);
    for (const b of $("theme-seg").children) {
      b.addEventListener("click", () => { data.settings.theme = b.dataset.v; save(); applyTheme(); renderSettings(); });
    }
    $("set-font").addEventListener("change", () => { data.settings.font = $("set-font").value; save(); applyFont(); });
    $("backup").addEventListener("click", backup);
    for (const id of ["pl-budget", "pl-min", "pl-max", "pl-salary", "pl-avg", "pl-14"]) $(id).addEventListener("change", plRead);
    for (const b of $("pl-priority").children) {
      b.addEventListener("click", () => { pl.priority = b.dataset.v; plSave(); renderPlanner(); });
    }
    $("restore").addEventListener("click", restore);
    if (window.matchMedia) window.matchMedia("(prefers-color-scheme: dark)").addEventListener("change", applyTheme);

    renderSheet();
    // Полночь: обновить «сегодня»
    let lastDay = C.todayISO();
    setInterval(() => { const d = C.todayISO(); if (d !== lastDay) { lastDay = d; if (view.page === "sheet") renderSheet(); } }, 60000);
  }
  init();
})();
