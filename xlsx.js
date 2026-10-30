/*
 * Маленький генератор .xlsx без сторонних библиотек (работает офлайн).
 * Собирает книгу из одного листа: табель месяца с итогами.
 * ZIP пишется без сжатия (метод «stored») — Excel такие файлы открывает.
 */
(function (root) {
  "use strict";
  const C = (typeof Calc !== "undefined") ? Calc : require("./calc.js");

  // ---------------------------------------------------------------- ZIP
  const CRC = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(bytes) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < bytes.length; i++) c = CRC[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }
  function zip(files) {
    const enc = new TextEncoder();
    const chunks = [], central = [];
    let offset = 0;
    for (const [name, text] of files) {
      const nameB = enc.encode(name), data = enc.encode(text), crc = crc32(data);
      const local = new DataView(new ArrayBuffer(30));
      local.setUint32(0, 0x04034b50, true); local.setUint16(4, 20, true); local.setUint16(6, 0x0800, true);
      local.setUint16(8, 0, true); local.setUint32(14, crc, true);
      local.setUint32(18, data.length, true); local.setUint32(22, data.length, true);
      local.setUint16(26, nameB.length, true);
      chunks.push(new Uint8Array(local.buffer), nameB, data);
      const cd = new DataView(new ArrayBuffer(46));
      cd.setUint32(0, 0x02014b50, true); cd.setUint16(4, 20, true); cd.setUint16(6, 20, true);
      cd.setUint16(8, 0x0800, true); cd.setUint32(16, crc, true);
      cd.setUint32(20, data.length, true); cd.setUint32(24, data.length, true);
      cd.setUint16(28, nameB.length, true); cd.setUint32(42, offset, true);
      central.push(new Uint8Array(cd.buffer), nameB);
      offset += 30 + nameB.length + data.length;
    }
    const cdSize = central.reduce((s, c) => s + c.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true);
    end.setUint16(10, files.length, true); end.setUint32(12, cdSize, true); end.setUint32(16, offset, true);
    const all = [...chunks, ...central, new Uint8Array(end.buffer)];
    const out = new Uint8Array(all.reduce((s, c) => s + c.length, 0));
    let p = 0;
    for (const c of all) { out.set(c, p); p += c.length; }
    return out;
  }

  // ---------------------------------------------------------------- листы
  const esc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const col = i => String.fromCharCode(65 + i);
  // Стили (индексы cellXfs ниже)
  const S = { title: 1, sub: 2, head: 3, date: 4, text: 5, time: 6, dur: 7, over: 8, under: 9,
    wDate: 10, wText: 11, total: 12, totalDur: 13, label: 14, value: 15, valueInt: 16, valueBold: 17 };

  const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="4"><numFmt numFmtId="164" formatCode="[h]:mm"/><numFmt numFmtId="165" formatCode="hh:mm"/><numFmt numFmtId="166" formatCode="dd.mm.yyyy"/><numFmt numFmtId="167" formatCode="[h]:mm;-[h]:mm;&quot;&quot;"/></numFmts>
<fonts count="7">
<font><sz val="10"/><color rgb="FF1F2A44"/><name val="Arial"/></font>
<font><b/><sz val="14"/><color rgb="FF1F2A44"/><name val="Arial"/></font>
<font><i/><sz val="9"/><color rgb="FF5B6475"/><name val="Arial"/></font>
<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font>
<font><sz val="10"/><color rgb="FF8A5A00"/><name val="Arial"/></font>
<font><sz val="10"/><color rgb="FFA3332F"/><name val="Arial"/></font>
<font><b/><sz val="10"/><color rgb="FF1F2A44"/><name val="Arial"/></font>
</fonts>
<fills count="6"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF2B3A67"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFECEEF1"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFDDE2EE"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF5F6F9"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color rgb="FFC9CED8"/></left><right style="thin"><color rgb="FFC9CED8"/></right><top style="thin"><color rgb="FFC9CED8"/></top><bottom style="thin"><color rgb="FFC9CED8"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="18">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="3" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" indent="1"/></xf>
<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="167" fontId="4" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="167" fontId="5" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="166" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" indent="1"/></xf>
<xf numFmtId="0" fontId="6" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="164" fontId="6" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>
<xf numFmtId="164" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="164" fontId="6" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

  const MIN = m => m / 1440;                       // минуты → доля суток для Excel
  const SERIAL = s => (Date.UTC(...(p => [p.y, p.m - 1, p.d])(C.parts(s))) - Date.UTC(1899, 11, 30)) / 86400000;

  function noteFor(row) {
    if (row.type !== C.DAY.WORK) {
      const base = row.type === C.DAY.WEEKEND ? "Выходной" : (row.title || "Праздник");
      return row.record.mark === C.MARK.VACATION ? `Отпуск (${base === "Выходной" ? "выходной" : base})` : base;
    }
    if (row.record.mark) return C.MARK_LABELS[row.record.mark];
    return row.title;
  }

  function monthXlsx(s, y, m) {
    const rows = [];      // [[номер строки, [[колонка, значение, стиль, тип]]]]
    const put = (r, cells) => rows.push([r, cells]);
    const r = s.rules;
    put(1, [[0, `Табель учёта рабочего времени — ${C.MONTHS[m - 1]} ${y}`, S.title, "s"]]);
    const now = new Date();
    const modeText = r.flexible ? `Режим: ${r.describe()}, норма в рабочий день ${C.fmtMinutes(s.dayCost)}, переработка по итогам месяца. `
      : `Работа в день ${C.fmtMinutes(r.workMin)}, обед ${C.fmtMinutes(r.lunchMin)}. `;
    put(2, [[0, modeText + (r.flexible ? `Обед ${C.fmtMinutes(r.lunchMin)}. ` : "") +
      `Выгружено ${String(now.getDate()).padStart(2, "0")}.${String(now.getMonth() + 1).padStart(2, "0")}.${now.getFullYear()} ` +
      `${String(now.getHours()).padStart(2, "0")}:${String(now.getMinutes()).padStart(2, "0")}`, S.sub, "s"]]);
    const heads = ["Дата", "День недели", "Приход", "Уход", "Отработано", "Переработка", "Недоработка", "Отметка"];
    put(4, heads.map((h, i) => [i, h, S.head, "s"]));
    let rn = 5;
    for (const row of s.rows) {
      const off = row.type !== C.DAY.WORK;
      const rec = row.record, res = row.result;
      const t = v => v ? MIN(C.parseTime(v)) : null;
      const cells = [
        [0, SERIAL(row.day), off ? S.wDate : S.date, "n"],
        [1, C.WEEKDAYS[C.weekday(row.day)], off ? S.wText : S.text, "s"],
        [2, !off && !rec.mark ? t(rec.in) : null, off ? S.wText : S.time, "n"],
        [3, !off && !rec.mark ? t(rec.out) : null, off ? S.wText : S.time, "n"],
        [4, res.complete ? MIN(res.worked) : null, off ? S.wText : S.dur, "n"],
        [5, res.complete && res.overtime !== null ? MIN(res.overtime) : null, off ? S.wText : S.over, "n"],
        [6, res.complete && res.undertime !== null ? MIN(res.undertime) : null, off ? S.wText : S.under, "n"],
        [7, noteFor(row), off ? S.wText : S.text, "s"],
      ];
      put(rn++, cells);
    }
    put(rn, [[0, "Итого за месяц", S.total, "s"], [1, "", S.total, "s"], [2, "", S.total, "s"], [3, "", S.total, "s"],
      [4, MIN(s.worked), S.totalDur, "n"], [5, MIN(s.overtimeGross), S.totalDur, "n"],
      [6, MIN(s.undertimeGross), S.totalDur, "n"], [7, "", S.total, "s"]]);
    rn += 2;
    put(rn++, [[0, "Итоги месяца", S.title, "s"]]);
    const summary = [
      ["Рабочих дней по календарю", s.workdays, S.valueInt],
      ["Дней отпуска (рабочих)", s.vacationDays, S.valueInt],
      ["Дней больничного", s.sickDays, S.valueInt],
      ["Отгулов за переработку", s.offDays, S.valueInt],
      ["Норма в рабочий день", MIN(s.dayCost), S.value],
      ["Норма месяца", MIN(s.norm), S.value],
      ["Отработано", MIN(s.worked), S.value],
      ["Осталось отработать", MIN(Math.max(0, s.remaining)), S.valueBold],
      ["Переработка за месяц", MIN(Math.max(0, s.balance)), S.valueBold],
      ["Доступно отгулов", s.availableOffDays, S.valueInt],
    ];
    for (const [label, v, st] of summary) {
      put(rn++, [[0, label, S.label, "s"], [1, "", S.label, "s"], [2, "", S.label, "s"], [3, v, st, "n"]]);
    }
    rn++;
    put(rn, [[0, (r.flexible ? "Переработка считается по итогам месяца — всё, что отработано сверх нормы. "
      : "Недоработка вычитается из переработок месяца. ") +
      "Отгул списывает норму одного рабочего дня. Переработки не переносятся на следующий месяц.", S.sub, "s"]]);

    const widths = [12, 15, 10, 10, 13, 14, 14, 36];
    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>`;
    widths.forEach((w, i) => { xml += `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`; });
    xml += "</cols><sheetData>";
    for (const [rnum, cells] of rows) {
      xml += `<row r="${rnum}">`;
      for (const [c, v, st, type] of cells) {
        const ref = `${col(c)}${rnum}`;
        if (v === null || v === undefined || v === "") xml += `<c r="${ref}" s="${st}"/>`;
        else if (type === "n") xml += `<c r="${ref}" s="${st}"><v>${v}</v></c>`;
        else xml += `<c r="${ref}" s="${st}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
      }
      xml += "</row>";
    }
    xml += `</sheetData><pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="portrait" fitToWidth="1" fitToHeight="1"/></worksheet>`;
    xml = xml.replace("<sheetViews>", "<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr><sheetViews>");

    const sheetName = esc(`${C.MONTHS[m - 1]} ${y}`);
    return zip([
      ["[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`],
      ["_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
      ["xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${sheetName}" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm.Print_Titles" localSheetId="0">'${sheetName}'!$4:$4</definedName></definedNames></workbook>`],
      ["xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`],
      ["xl/styles.xml", STYLES],
      ["xl/worksheets/sheet1.xml", xml],
    ]);
  }

  function toBase64(bytes) {
    let s = "";
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return (typeof btoa !== "undefined") ? btoa(s) : Buffer.from(bytes).toString("base64");
  }

  const api = { monthXlsx, toBase64, zip, crc32 };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Xlsx = api;
})(typeof window !== "undefined" ? window : globalThis);
