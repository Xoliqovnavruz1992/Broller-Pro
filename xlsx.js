/* Ross 308 — Excel (.xlsx) hisobot yaratuvchi. Tashqi kutubxonasiz, offline ishlaydi. */
(function (root) {
  'use strict';

  // ---------- ZIP (siqmasdan) ----------
  var CRC = (function () {
    var t = [], c, n, k;
    for (n = 0; n < 256; n++) { c = n; for (k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
    return t;
  })();
  function crc32(b) { var c = 0xFFFFFFFF, i; for (i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  function u16(v) { return [v & 255, (v >>> 8) & 255]; }
  function u32(v) { return [v & 255, (v >>> 8) & 255, (v >>> 16) & 255, (v >>> 24) & 255]; }
  function zip(files) {
    var enc = new TextEncoder(), parts = [], cd = [], off = 0, dt = ((2026 - 1980) << 9) | (10 << 5) | 16;
    files.forEach(function (f) {
      var name = enc.encode(f.name), data = enc.encode(f.text), c = crc32(data);
      var lh = [].concat(u32(0x04034b50), u16(20), u16(0), u16(0), u16(0), u16(dt), u32(c), u32(data.length), u32(data.length), u16(name.length), u16(0));
      parts.push(new Uint8Array(lh), name, data);
      var ch = [].concat(u32(0x02014b50), u16(20), u16(20), u16(0), u16(0), u16(0), u16(dt), u32(c), u32(data.length), u32(data.length), u16(name.length), u16(0), u16(0), u16(0), u16(0), u32(0), u32(off));
      cd.push(new Uint8Array(ch), name);
      off += lh.length + name.length + data.length;
    });
    var cdSize = 0; cd.forEach(function (x) { cdSize += x.length; });
    var end = new Uint8Array([].concat(u32(0x06054b50), u16(0), u16(0), u16(files.length), u16(files.length), u32(cdSize), u32(off), u16(0)));
    var all = parts.concat(cd, [end]), total = 0, p = 0;
    all.forEach(function (x) { total += x.length; });
    var out = new Uint8Array(total);
    all.forEach(function (x) { out.set(x, p); p += x.length; });
    return out;
  }

  // ---------- XML yordamchilar ----------
  function esc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
  function cn(ref, s, v) { return '<c r="' + ref + '" s="' + s + '"><v>' + v + '</v></c>'; }
  function cs(ref, s, v) { return '<c r="' + ref + '" s="' + s + '" t="inlineStr"><is><t xml:space="preserve">' + esc(v) + '</t></is></c>'; }
  function cb(ref, s) { return '<c r="' + ref + '" s="' + s + '"/>'; }
  function cf(ref, s, f, v) {
    if (typeof v === 'number' && isFinite(v)) return '<c r="' + ref + '" s="' + s + '"><f>' + esc(f) + '</f><v>' + v + '</v></c>';
    return '<c r="' + ref + '" s="' + s + '" t="str"><f>' + esc(f) + '</f><v>' + esc(v || '') + '</v></c>';
  }
  function row(r, cells) { return '<row r="' + r + '">' + cells.join('') + '</row>'; }
  var NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';

  var STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet ' + NS + '>' +
    '<numFmts count="4"><numFmt numFmtId="164" formatCode="dd\\.mm\\.yyyy"/><numFmt numFmtId="165" formatCode="0.0%"/><numFmt numFmtId="166" formatCode="#,##0.0"/><numFmt numFmtId="167" formatCode="0.00"/></numFmts>' +
    '<fonts count="4"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font>' +
    '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><b/><sz val="14"/><color rgb="FF0B6B5C"/><name val="Calibri"/></font></fonts>' +
    '<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FF0B6B5C"/><bgColor indexed="64"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF4CC"/><bgColor indexed="64"/></patternFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFE3F1EC"/><bgColor indexed="64"/></patternFill></fill></fills>' +
    '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
    '<border><left style="thin"><color rgb="FFD3DED8"/></left><right style="thin"><color rgb="FFD3DED8"/></right><top style="thin"><color rgb="FFD3DED8"/></top><bottom style="thin"><color rgb="FFD3DED8"/></bottom><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="12">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>' +
    '<xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>' +
    '<xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>' +
    '<xf numFmtId="167" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>' +
    '<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>' +
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>' +
    '<xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1"/>' +
    '<xf numFmtId="164" fontId="0" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>' +
    '<xf numFmtId="0" fontId="1" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>' +
    '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
    '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

  function serialOf(y, m, d) { return Date.UTC(y, m, d) / 864e5 + 25569; }

  // st = {start:'YYYY-MM-DD', birds, mort:{kun:bosh}, log:{kun:vazn}}; D = jadval; VET = veterinariya nomlari
  function makeReportXlsx(st, D, VET, opts) {
    st = st || {}; var N = D.length - 1, last = N + 2;
    var birds = +st.birds > 0 ? +st.birds : 0;
    var mort = st.mort || {}, log = st.log || {};
    var serial = null, m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(st.start || '');
    if (m) serial = serialOf(+m[1], +m[2] - 1, +m[3]);
    var now = (opts && opts.today) || new Date();
    var today = serialOf(now.getFullYear(), now.getMonth(), now.getDate());
    var cur = serial === null ? '' : Math.min(N, Math.max(0, today - serial));
    var SH = "'Kunlik hisobot'!";
    var rng = function (c) { return SH + c + '2:' + c + last; };

    // ----- Kunlik hisobot -----
    var heads = ['Kun', 'Sana', 'Standart vazn, g', 'Haqiqiy vazn, g', 'Farq, %', 'Ozuqa, g/bosh', 'Suv, ml/bosh', 'Harorat, °C', 'Veterinariya',
      'Ulim, bosh', 'Tirik bosh', 'Poda ozuqa, kg', 'Poda suv, litr', 'Poda vazn (std), kg', 'Kunlik o\'sish, g', 'Jami ozuqa (1 bosh), kg', 'FCR (std)'];
    var L = 'ABCDEFGHIJKLMNOPQ';
    var rows = [row(1, heads.map(function (h, i) { return cs(L[i] + '1', 1, h); }))];
    var cumM = 0, cumF = 0, sumFeed = 0, sumWater = 0, sumMortAll = 0, aliveAtCur = '', feedToCur = '', waterToCur = '';
    var lastW = '', lastWDay = '';
    for (var i = 0; i <= N; i++) {
      var r = i + 2, d = D[i], std = d[0], act = +log[i] > 0 ? +log[i] : null, mo = +mort[i] > 0 ? +mort[i] : null;
      cumM += mo || 0; cumF += d[1]; sumMortAll += mo || 0;
      var alive = birds ? birds - cumM : '';
      var lf = alive === '' ? '' : d[1] * alive / 1000, lw = alive === '' ? '' : d[2] * alive / 1000, ls = alive === '' ? '' : std * alive / 1000;
      if (act !== null) { lastW = act; lastWDay = i; }
      if (cur !== '' && i <= cur) { if (alive !== '') { aliveAtCur = alive; feedToCur = (feedToCur || 0) + lf; waterToCur = (waterToCur || 0) + lw; } }
      var c = [];
      c.push(cn('A' + r, 2, i));
      c.push(cf('B' + r, 7, 'IF(Xulosa!$B$2="","",Xulosa!$B$2+A' + r + ')', serial === null ? '' : serial + i));
      c.push(cn('C' + r, 3, std));
      c.push(act !== null ? cn('D' + r, 8, act) : cb('D' + r, 8));
      c.push(cf('E' + r, 6, 'IF(D' + r + '="","",(D' + r + '-C' + r + ')/C' + r + ')', act !== null ? (act - std) / std : ''));
      c.push(cn('F' + r, 3, d[1]));
      c.push(cn('G' + r, 3, d[2]));
      c.push(cn('H' + r, 2, d[3]));
      c.push(cs('I' + r, 2, VET[d[4]]));
      c.push(mo !== null ? cn('J' + r, 8, mo) : cb('J' + r, 8));
      c.push(cf('K' + r, 3, 'IF(Xulosa!$B$3="","",Xulosa!$B$3-SUM(J$2:J' + r + '))', alive));
      c.push(cf('L' + r, 4, 'IF(K' + r + '="","",F' + r + '*K' + r + '/1000)', lf));
      c.push(cf('M' + r, 4, 'IF(K' + r + '="","",G' + r + '*K' + r + '/1000)', lw));
      c.push(cf('N' + r, 4, 'IF(K' + r + '="","",C' + r + '*K' + r + '/1000)', ls));
      c.push(i === 0 ? cb('O' + r, 2) : cf('O' + r, 3, 'C' + r + '-C' + (r - 1), std - D[i - 1][0]));
      c.push(cf('P' + r, 5, 'SUM(F$2:F' + r + ')/1000', cumF / 1000));
      c.push(cf('Q' + r, 5, 'SUM(F$2:F' + r + ')/C' + r, cumF / std));
      rows.push(row(r, c));
    }
    var widths = [6, 12, 11, 11, 9, 11, 11, 10, 26, 10, 10, 12, 12, 14, 12, 14, 9];
    var cols1 = '<cols>' + widths.map(function (w, i) { return '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>'; }).join('') + '</cols>';
    var sheetDaily = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet ' + NS + '>' +
      '<sheetViews><sheetView workbookViewId="0"><pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="15"/>' + cols1 + '<sheetData>' + rows.join('') + '</sheetData></worksheet>';

    // ----- Xulosa -----
    var diff = '';
    if (lastW !== '') diff = (lastW - D[lastWDay][0]) / D[lastWDay][0];
    var lookW = 'IFERROR(LOOKUP(2,1/(' + rng('D') + '<>""),' + rng('D') + '),"")';
    var lookD = 'IFERROR(LOOKUP(2,1/(' + rng('D') + '<>""),' + rng('A') + '),"")';
    var sumTo = function (c) { return 'SUMIF(' + rng('A') + ',"<="&B4,' + rng(c) + ')'; };
    var needs = 'OR(B3="",B4="")';
    var aliveNow = (birds && cur !== '') ? birds - (function () { var s = 0; for (var k = 0; k <= cur; k++) s += +mort[k] > 0 ? +mort[k] : 0; return s; })() : '';
    var S = [];
    S.push(row(1, [cs('A1', 11, 'Ross 308 — hisobot')]));
    S.push(row(2, [cs('A2', 10, 'Joylashtirilgan sana'), serial === null ? cb('B2', 9) : cn('B2', 9, serial)]));
    S.push(row(3, [cs('A3', 10, 'Bosh soni'), birds ? cn('B3', 8, birds) : cb('B3', 8)]));
    S.push(row(4, [cs('A4', 10, 'Bugungi kun (0–' + N + ')'), cf('B4', 2, 'IF(B2="","",MIN(' + N + ',MAX(0,TODAY()-B2)))', cur)]));
    S.push(row(5, [cs('A5', 10, 'Jami ulim, bosh'), cf('B5', 3, 'SUM(' + rng('J') + ')', sumMortAll)]));
    S.push(row(6, [cs('A6', 10, 'Ulim, %'), cf('B6', 6, 'IF(B3="","",B5/B3)', birds ? sumMortAll / birds : '')]));
    S.push(row(7, [cs('A7', 10, 'Hozir tirik bosh'), cf('B7', 3, 'IF(' + needs + ',"",B3-' + sumTo('J') + ')', aliveNow)]));
    S.push(row(8, [cs('A8', 10, 'Podaga jami ozuqa (shu kungacha), kg'), cf('B8', 4, 'IF(' + needs + ',"",' + sumTo('L') + ')', (birds && cur !== '') ? feedToCur : '')]));
    S.push(row(9, [cs('A9', 10, 'Podaga jami suv (shu kungacha), litr'), cf('B9', 4, 'IF(' + needs + ',"",' + sumTo('M') + ')', (birds && cur !== '') ? waterToCur : '')]));
    S.push(row(10, [cs('A10', 10, 'Oxirgi kiritilgan vazn, g'), cf('B10', 4, lookW, lastW)]));
    S.push(row(11, [cs('A11', 10, 'Oxirgi vazn kiritilgan kun'), cf('B11', 2, lookD, lastWDay)]));
    S.push(row(12, [cs('A12', 10, 'Standartdan farq, %'), cf('B12', 6, 'IF(B11="","",(B10-INDEX(' + rng('C') + ',B11+1))/INDEX(' + rng('C') + ',B11+1))', diff)]));
    S.push(row(14, [cs('A14', 0, 'Sariq kataklarga (sana, bosh soni, haqiqiy vazn, ulim) ma\'lumot kiritsangiz, hisob avtomatik yangilanadi.')]));
    var sheetSum = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet ' + NS + '>' +
      '<sheetViews><sheetView workbookViewId="0" tabSelected="1"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>' +
      '<cols><col min="1" max="1" width="44" customWidth="1"/><col min="2" max="2" width="18" customWidth="1"/></cols>' +
      '<sheetData>' + S.join('') + '</sheetData></worksheet>';

    var files = [
      { name: '[Content_Types].xml', text: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>' },
      { name: '_rels/.rels', text: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' },
      { name: 'xl/workbook.xml', text: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook ' + NS + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="0"/></bookViews><sheets><sheet name="Xulosa" sheetId="1" r:id="rId1"/><sheet name="Kunlik hisobot" sheetId="2" r:id="rId2"/></sheets><calcPr fullCalcOnLoad="1"/></workbook>' },
      { name: 'xl/_rels/workbook.xml.rels', text: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>' },
      { name: 'xl/styles.xml', text: STYLES },
      { name: 'xl/worksheets/sheet1.xml', text: sheetSum },
      { name: 'xl/worksheets/sheet2.xml', text: sheetDaily }
    ];
    return zip(files);
  }

  root.makeReportXlsx = makeReportXlsx;
  if (typeof module !== 'undefined' && module.exports) module.exports = { makeReportXlsx: makeReportXlsx };
})(typeof self !== 'undefined' ? self : globalThis);
