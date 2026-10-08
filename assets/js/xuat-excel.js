/*
 * Xuất bảng dữ liệu ra file Excel (.xlsx) ngay trên trình duyệt, không cần thư viện ngoài.
 *
 * xuatExcel({
 *   file: 'ket-qua.xlsx', sheet: 'Kết quả',
 *   title: 'Tiêu đề dòng đầu (tuỳ chọn)',
 *   columns: [{ header: 'Tên', width: 30, money: false }, ...],
 *   rows: [[...], ...]          // ô số (money: true) truyền vào kiểu number
 * });
 */
(function () {
  var enc = new TextEncoder();

  var CRC = (function () {
    var t = new Uint32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(b) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  // Gói zip không nén (store) — đủ cho Excel đọc
  function zip(files) {
    var parts = [], central = [], offset = 0;
    files.forEach(function (f) {
      var name = enc.encode(f.name), data = enc.encode(f.data), crc = crc32(data);
      var h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true);
      h.setUint32(14, crc, true); h.setUint32(18, data.length, true); h.setUint32(22, data.length, true);
      h.setUint16(26, name.length, true);
      parts.push(h.buffer, name, data);
      var c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true);
      c.setUint32(16, crc, true); c.setUint32(20, data.length, true); c.setUint32(24, data.length, true);
      c.setUint16(28, name.length, true); c.setUint32(42, offset, true);
      central.push(c.buffer, name);
      offset += 30 + name.length + data.length;
    });
    var size = central.reduce(function (s, p) { return s + (p.byteLength || p.length); }, 0);
    var e = new DataView(new ArrayBuffer(22));
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true);
    e.setUint32(12, size, true); e.setUint32(16, offset, true);
    return new Blob(parts.concat(central, [e.buffer]), { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  function esc(v) {
    return String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function col(i) { var s = ''; i++; while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }

  window.xuatExcel = function (o) {
    var cols = o.columns, rows = [], r = 0, merge = '';
    function row(cells, style) {
      r++;
      rows.push('<row r="' + r + '">' + cells.map(function (v, i) {
        var ref = col(i) + r;
        if (v === null || v === undefined || v === '') return '<c r="' + ref + '" s="' + style(i, v) + '"/>';
        if (typeof v === 'number') return '<c r="' + ref + '" s="' + style(i, v) + '"><v>' + v + '</v></c>';
        return '<c r="' + ref + '" s="' + style(i, v) + '" t="inlineStr"><is><t xml:space="preserve">' + esc(v) + '</t></is></c>';
      }).join('') + '</row>');
    }
    if (o.title) {
      row([o.title], function () { return 3; });
      merge = '<mergeCells count="1"><mergeCell ref="A1:' + col(cols.length - 1) + '1"/></mergeCells>';
    }
    row(cols.map(function (c) { return c.header; }), function () { return 1; });
    var headRow = r;
    o.rows.forEach(function (cells) {
      row(cells, function (i, v) { return typeof v === 'number' && cols[i].money ? 2 : 4; });
    });

    var sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetViews><sheetView workbookViewId="0"><pane ySplit="' + headRow + '" topLeftCell="A' + (headRow + 1) + '" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
      '<cols>' + cols.map(function (c, i) { return '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + (c.width || 16) + '" customWidth="1"/>'; }).join('') + '</cols>' +
      '<sheetData>' + rows.join('') + '</sheetData>' +
      '<autoFilter ref="A' + headRow + ':' + col(cols.length - 1) + r + '"/>' + merge + '</worksheet>';

    var styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0"/></numFmts>' +
      '<fonts count="3"><font><sz val="11"/><name val="Times New Roman"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Times New Roman"/></font><font><b/><sz val="13"/><name val="Times New Roman"/></font></fonts>' +
      '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF0B4B8F"/></patternFill></fill></fills>' +
      '<borders count="2"><border/><border><left style="thin"><color rgb="FFB6C2D0"/></left><right style="thin"><color rgb="FFB6C2D0"/></right><top style="thin"><color rgb="FFB6C2D0"/></top><bottom style="thin"><color rgb="FFB6C2D0"/></bottom></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="5">' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
      '<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment vertical="top"/></xf>' +
      '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +
      '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

    var blob = zip([
      { name: '[Content_Types].xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>' },
      { name: '_rels/.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' },
      { name: 'xl/workbook.xml', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="' + esc((o.sheet || 'Sheet1').slice(0, 31)) + '" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">\'' + esc((o.sheet || 'Sheet1').slice(0, 31)) + '\'!$A$' + headRow + ':$' + col(cols.length - 1) + '$' + r + '</definedName></definedNames></workbook>' },
      { name: 'xl/_rels/workbook.xml.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>' },
      { name: 'xl/styles.xml', data: styles },
      { name: 'xl/worksheets/sheet1.xml', data: sheet }
    ]);

    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = o.file || 'du-lieu.xlsx';
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  };
})();
