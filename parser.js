/*
 * Orari Negozio - parser
 * Converte il PDF (o l'Excel) degli orari settimanali in un modello dati
 * e lo codifica/decodifica in una stringa compatta da mettere nel link.
 *
 * Modello:
 * {
 *   v: 1,
 *   days:   [{ d: "2026-09-14", note: "" } x7],
 *   people: [{ name: "ILARIA",
 *              days: [{ shifts: [["10:00","13:00"],["14:00","19:00"]], h: "8,00", note: "" } x7],
 *              tot: "40,00", delta: "-0,30" }]
 * }
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.OrariParser = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var NUM_RE = /^-?\d+(?:[.,]\d+)?$/;
  var DATE_RE = /^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/;
  var TIME_RE = /^(\d{1,2})[.:,](\d{2})$/;

  function isNum(s) { return NUM_RE.test(String(s).trim()); }
  function toNum(s) { return parseFloat(String(s).trim().replace(',', '.')); }
  function pad(n) { return (n < 10 ? '0' : '') + n; }

  function parseDate(s) {
    var m = DATE_RE.exec(String(s).trim());
    if (!m) return null;
    var y = m[3].length === 2 ? '20' + m[3] : m[3];
    return y + '-' + pad(+m[2]) + '-' + pad(+m[1]);
  }

  // Data di Excel (numero seriale) -> ISO
  function excelSerialToISO(n) {
    var d = new Date(Math.round((n - 25569) * 86400 * 1000));
    return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate());
  }

  // 10.3 -> "10:30" (formato "ore,minuti" usato nel foglio); 0.4375 -> "10:30" (orario Excel)
  function numToTime(v) {
    if (v < 1) {
      var mins = Math.round(v * 24 * 60);
      return Math.floor(mins / 60) + ':' + pad(mins % 60);
    }
    var h = Math.floor(v);
    var m = Math.round((v - h) * 100);
    if (m >= 60) { // era in ore decimali (es. 10.75)
      m = Math.round((v - h) * 60);
    }
    return h + ':' + pad(m);
  }

  function strToTime(s) {
    var m = TIME_RE.exec(String(s).trim());
    if (!m) return null;
    return (+m[1]) + ':' + m[2];
  }

  function fmtHours(v) {
    if (v === null || v === undefined || v === '') return '';
    if (typeof v === 'number') return v.toFixed(2).replace('.', ',');
    return String(v).trim();
  }

  function pairShifts(times) {
    var shifts = [];
    for (var i = 0; i < times.length; i += 2) {
      shifts.push(i + 1 < times.length ? [times[i], times[i + 1]] : [times[i], '']);
    }
    return shifts;
  }

  function emptyDay() { return { shifts: [], h: '', note: '' }; }

  function addNote(obj, text) {
    text = String(text).trim();
    if (!text) return;
    obj.note = obj.note ? obj.note + ' · ' + text : text;
  }

  /* ------------------------------------------------------------------ */
  /* PDF                                                                 */
  /* ------------------------------------------------------------------ */

  // items: [{ s, x, y, w }] (testo, posizione, larghezza) di pdf.js
  function parsePdfItems(items) {
    var list = items.filter(function (i) { return String(i.s).trim() !== ''; })
      .map(function (i) { return { s: String(i.s).trim(), x: i.x, y: i.y, cx: i.x + (i.w || 0) / 2 }; })
      .sort(function (a, b) { return b.y - a.y || a.x - b.x; });

    // raggruppa in righe (tolleranza verticale)
    var rows = [];
    var cur = null;
    list.forEach(function (it) {
      if (!cur || Math.abs(cur.y - it.y) > 3) { cur = { y: it.y, items: [] }; rows.push(cur); }
      cur.items.push(it);
    });
    rows.forEach(function (r) { r.items.sort(function (a, b) { return a.cx - b.cx; }); });

    // riga delle date
    var dateRowIdx = -1, dates = [];
    for (var r = 0; r < rows.length; r++) {
      var found = rows[r].items.filter(function (i) { return parseDate(i.s); });
      if (found.length >= 5) { dateRowIdx = r; dates = found; break; }
    }
    if (dateRowIdx < 0) throw new Error('Non trovo la riga con le date (gg/mm/aaaa) nel PDF.');

    var days = dates.map(function (d) { return { d: parseDate(d.s), note: '' }; });
    var centers = dates.map(function (d) { return d.cx; });
    var n = centers.length;
    var left0 = centers[0] - (centers[1] - centers[0]) / 2;
    var rightN = centers[n - 1] + (centers[n - 1] - centers[n - 2]) / 2;
    var bounds = [];
    for (var i = 0; i < n - 1; i++) bounds.push((centers[i] + centers[i + 1]) / 2);

    function dayOf(cx) {
      if (cx < left0 || cx > rightN) return -1;
      for (var i = 0; i < bounds.length; i++) if (cx < bounds[i]) return i;
      return n - 1;
    }

    // riga intestazione "PERSONALE ... E U E U H"
    var headerIdx = -1;
    for (r = dateRowIdx + 1; r < rows.length; r++) {
      var labels = rows[r].items.filter(function (i) { return /^[EUH]$/.test(i.s); }).length;
      var pers = rows[r].items.some(function (i) { return /^PERSONALE/i.test(i.s); });
      if (pers || labels >= 5) { headerIdx = r; break; }
    }

    // righe dipendenti: prima cella testuale a sinistra delle date + almeno un numero nei giorni
    function isEmployeeRow(row) {
      var first = row.items[0];
      if (!first || first.cx >= left0 || isNum(first.s) || /^PERSONALE/i.test(first.s)) return false;
      return row.items.some(function (i) { return isNum(i.s) && dayOf(i.cx) >= 0; });
    }

    var startIdx = headerIdx >= 0 ? headerIdx + 1 : dateRowIdx + 1;
    var empRows = [];
    for (r = startIdx; r < rows.length; r++) if (isEmployeeRow(rows[r])) empRows.push(rows[r]);

    // note in testata (tra le date e l'intestazione)
    var noteEnd = headerIdx >= 0 ? headerIdx : (empRows.length ? rows.indexOf(empRows[0]) : rows.length);
    for (r = dateRowIdx + 1; r < noteEnd; r++) {
      rows[r].items.forEach(function (i) {
        var d = dayOf(i.cx);
        if (d >= 0 && !isNum(i.s)) addNote(days[d], i.s);
      });
    }
    // eventuali note sulla stessa riga delle date (testo non-data dentro un giorno)
    rows[dateRowIdx].items.forEach(function (i) {
      var d = dayOf(i.cx);
      if (d >= 0 && !parseDate(i.s) && !isNum(i.s)) addNote(days[d], i.s);
    });

    // colonna H (ore) di ogni giorno = valore numerico più a destra del giorno, su tutte le righe
    var hx = days.map(function () { return -Infinity; });
    empRows.forEach(function (row) {
      row.items.forEach(function (i) {
        var d = dayOf(i.cx);
        if (d >= 0 && isNum(i.s) && i.cx > hx[d]) hx[d] = i.cx;
      });
    });
    // confini raffinati: ogni giorno finisce subito dopo la sua colonna H
    var bounds2 = [];
    for (i = 0; i < n - 1; i++) bounds2.push(isFinite(hx[i]) ? hx[i] + 5 : bounds[i]);
    var rightN2 = isFinite(hx[n - 1]) ? hx[n - 1] + 5 : rightN;
    function dayOf2(cx) {
      if (cx < left0 || cx > rightN2) return -1;
      for (var i = 0; i < bounds2.length; i++) if (cx < bounds2[i]) return i;
      return n - 1;
    }

    var people = empRows.map(function (row) {
      var p = { name: row.items[0].s, days: days.map(emptyDay), tot: '', delta: '' };
      var timesByDay = days.map(function () { return []; });
      var afterWeek = [];
      row.items.slice(1).forEach(function (i) {
        var d = dayOf2(i.cx);
        if (d < 0) {
          if (i.cx > rightN2 && isNum(i.s)) afterWeek.push(i.s);
          return;
        }
        if (!isNum(i.s)) { addNote(p.days[d], i.s); return; }
        if (Math.abs(i.cx - hx[d]) < 6) { p.days[d].h = i.s; return; }
        timesByDay[d].push(numToTime(toNum(i.s)));
      });
      p.days.forEach(function (day, d) { day.shifts = pairShifts(timesByDay[d]); });
      p.tot = afterWeek[0] || '';
      p.delta = afterWeek[1] || '';
      return p;
    });

    if (!people.length) throw new Error('Non trovo righe con i nomi dei dipendenti nel PDF.');
    return { v: 1, days: days, people: people };
  }

  /* ------------------------------------------------------------------ */
  /* Excel (griglia di celle SheetJS: grid[r][c] = { v, t, w } | null)    */
  /* ------------------------------------------------------------------ */

  function gridFromWorksheet(ws, XLSX) {
    var range = XLSX.utils.decode_range(ws['!ref']);
    var grid = [];
    for (var r = range.s.r; r <= range.e.r; r++) {
      var row = [];
      for (var c = range.s.c; c <= range.e.c; c++) {
        var cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
        row.push(cell ? { v: cell.v, t: cell.t, w: cell.w } : null);
      }
      grid.push(row);
    }
    var merges = (ws['!merges'] || []).map(function (m) {
      return { r: m.s.r - range.s.r, c0: m.s.c - range.s.c, c1: m.e.c - range.s.c };
    });
    return { grid: grid, merges: merges };
  }

  function cellDate(cell) {
    if (!cell) return null;
    if (cell.t === 'd' && cell.v instanceof Date) {
      return cell.v.getFullYear() + '-' + pad(cell.v.getMonth() + 1) + '-' + pad(cell.v.getDate());
    }
    if (cell.w && parseDate(cell.w)) return parseDate(cell.w);
    if (cell.t === 's' && parseDate(cell.v)) return parseDate(cell.v);
    if (cell.t === 'n' && cell.v > 40000 && cell.v < 80000 && Number.isInteger(cell.v)) return excelSerialToISO(cell.v);
    return null;
  }

  function cellText(cell) {
    if (!cell || cell.v === null || cell.v === undefined) return '';
    return String(cell.w !== undefined ? cell.w : cell.v).trim();
  }

  function cellIsNum(cell) {
    return !!cell && (cell.t === 'n' || (cell.t === 's' && isNum(cell.v)));
  }

  function cellNum(cell) { return cell.t === 'n' ? cell.v : toNum(cell.v); }

  function parseGrid(data) {
    var grid = data.grid || data;
    var merges = data.merges || [];

    // riga delle date
    var dateRow = -1, dateCols = [], days = [];
    for (var r = 0; r < grid.length; r++) {
      var cols = [], ds = [];
      for (var c = 0; c < grid[r].length; c++) {
        var iso = cellDate(grid[r][c]);
        if (iso) { cols.push(c); ds.push(iso); }
      }
      if (cols.length >= 5) { dateRow = r; dateCols = cols; days = ds.map(function (d) { return { d: d, note: '' }; }); break; }
    }
    if (dateRow < 0) throw new Error('Non trovo la riga con le date nel foglio Excel.');
    var n = dateCols.length;

    // intervallo di colonne di ogni giorno (dalle celle unite se disponibili)
    var ranges = dateCols.map(function (c, i) {
      var m = merges.filter(function (m) { return m.r === dateRow && m.c0 === c; })[0];
      if (m) return [m.c0, m.c1];
      if (i < n - 1) return [c, dateCols[i + 1] - 1];
      var prev = i > 0 ? dateCols[i] - dateCols[i - 1] : 5;
      return [c, c + prev - 1];
    });
    function dayOfCol(c) {
      for (var i = 0; i < n; i++) if (c >= ranges[i][0] && c <= ranges[i][1]) return i;
      return -1;
    }

    // riga intestazione E U E U H / PERSONALE
    var headerRow = -1;
    for (r = dateRow + 1; r < grid.length; r++) {
      var labels = 0, pers = false;
      grid[r].forEach(function (cell) {
        var t = cellText(cell);
        if (/^[EUH]$/.test(t)) labels++;
        if (/^PERSONALE/i.test(t)) pers = true;
      });
      if (pers || labels >= 5) { headerRow = r; break; }
    }
    // colonna H di ogni giorno: etichetta "H" più a destra nel giorno, altrimenti ultima colonna
    var hCol = ranges.map(function (rg) { return rg[1]; });
    if (headerRow >= 0) {
      grid[headerRow].forEach(function (cell, c) {
        var d = dayOfCol(c);
        if (d >= 0 && cellText(cell) === 'H') hCol[d] = c;
      });
    }

    var nameColMax = dateCols[0];
    function nameOf(row) {
      for (var c = 0; c < nameColMax; c++) {
        var cell = row[c];
        if (cell && cell.t === 's' && cellText(cell) && !isNum(cell.v) && !/^PERSONALE/i.test(cell.v)) return cellText(cell);
      }
      return '';
    }
    function hasDayNumbers(row) {
      for (var c = dateCols[0]; c < row.length; c++) if (cellIsNum(row[c]) && dayOfCol(c) >= 0) return true;
      return false;
    }

    var start = headerRow >= 0 ? headerRow + 1 : dateRow + 1;
    var people = [];
    for (r = start; r < grid.length; r++) {
      var row = grid[r];
      var name = nameOf(row);
      if (!name || !hasDayNumbers(row)) continue;
      var p = { name: name, days: days.map(emptyDay), tot: '', delta: '' };
      var timesByDay = days.map(function () { return []; });
      var after = [];
      for (var c = dateCols[0]; c < row.length; c++) {
        var cell = row[c];
        if (!cell || cellText(cell) === '') continue;
        var d = dayOfCol(c);
        if (d < 0) {
          if (c > ranges[n - 1][1] && cellIsNum(cell)) after.push(fmtHours(cellNum(cell)));
          continue;
        }
        if (c === hCol[d]) { p.days[d].h = cellIsNum(cell) ? fmtHours(cellNum(cell)) : cellText(cell); continue; }
        if (cellIsNum(cell)) { timesByDay[d].push(numToTime(cellNum(cell))); continue; }
        var t = strToTime(cell.v);
        if (t) timesByDay[d].push(t); else addNote(p.days[d], cellText(cell));
      }
      p.days.forEach(function (day, d) { day.shifts = pairShifts(timesByDay[d]); });
      p.tot = after[0] || '';
      p.delta = after[1] || '';
      people.push(p);
    }

    // note in testata (righe fra le date e l'intestazione)
    var noteEnd = headerRow >= 0 ? headerRow : start;
    for (r = dateRow + 1; r < noteEnd; r++) {
      grid[r].forEach(function (cell, c) {
        var d = dayOfCol(c);
        if (d >= 0 && cell && cell.t === 's' && cellText(cell)) addNote(days[d], cellText(cell));
      });
    }

    if (!people.length) throw new Error('Non trovo righe con i nomi dei dipendenti nel foglio Excel.');
    return { v: 1, days: days, people: people };
  }

  /* ------------------------------------------------------------------ */
  /* Codifica nel link                                                   */
  /* ------------------------------------------------------------------ */

  function compact(model) {
    return {
      v: 1,
      d: model.days.map(function (d) { return [d.d, d.note || '']; }),
      p: model.people.map(function (p) {
        return [p.name, p.days.map(function (day) {
          var s = day.shifts.map(function (sh) { return sh[0] + '-' + sh[1]; }).join('|');
          return [s, day.h || '', day.note || ''];
        }), p.tot || '', p.delta || ''];
      })
    };
  }

  function expand(c) {
    return {
      v: c.v,
      days: c.d.map(function (d) { return { d: d[0], note: d[1] || '' }; }),
      people: c.p.map(function (p) {
        return {
          name: p[0],
          days: p[1].map(function (day) {
            return {
              shifts: day[0] ? day[0].split('|').map(function (s) { return s.split('-'); }) : [],
              h: day[1] || '', note: day[2] || ''
            };
          }),
          tot: p[2] || '', delta: p[3] || ''
        };
      })
    };
  }

  function b64urlEncode(bytes) {
    var s = '';
    for (var i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
    return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  }
  function b64urlDecode(str) {
    str = str.replace(/-/g, '+').replace(/_/g, '/');
    while (str.length % 4) str += '=';
    var s = atob(str);
    var bytes = new Uint8Array(s.length);
    for (var i = 0; i < s.length; i++) bytes[i] = s.charCodeAt(i);
    return bytes;
  }

  function encode(model, pako) {
    var json = JSON.stringify(compact(model));
    var bytes = pako.deflateRaw(new TextEncoder().encode(json), { level: 9 });
    return b64urlEncode(bytes);
  }

  function decode(str, pako) {
    var bytes = pako.inflateRaw(b64urlDecode(str));
    return expand(JSON.parse(new TextDecoder().decode(bytes)));
  }

  return {
    parsePdfItems: parsePdfItems,
    parseGrid: parseGrid,
    gridFromWorksheet: gridFromWorksheet,
    encode: encode,
    decode: decode,
    numToTime: numToTime,
    parseDate: parseDate
  };
});
