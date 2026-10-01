var WORKER = "https://excel-tro-ly.daotaotuyendungvikoda.workers.dev"; /* đã điền sẵn */
var APP_VERSION = "1.2.0"; /* phải khớp VERSION trong worker.js; tăng mỗi lần sửa */
var SERVER_VERSION = "";
var DANGER_HEADER = /(lương|luong|salary|cccd|cmnd|stk|tài khoản|tai khoan|mst|mã số thuế|thưởng|thuong)/i;
var ERR_RE = /^#(REF!|N\/A|DIV\/0!|VALUE!|NAME\?|NUM!|NULL!)/;
var state = { code: "", lastQ: "", undo: [] };

function $(id) { return document.getElementById(id); }
function store(k, v) {
  try { if (v === undefined) return localStorage.getItem(k); localStorage.setItem(k, v); } catch (e) { return null; }
}
function fmt(n) { return Number(n).toLocaleString("vi-VN"); }
function colLetter(n) {
  var s = ""; n = n + 1;
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
function compat() {
  try { return Office.context.requirements.isSetSupported("ExcelApi", "1.12") ? "full" : "legacy"; }
  catch (e) { return "legacy"; }
}

/* ---------- Gọi Worker ---------- */
function api(path, method, body, cb) {
  var x = new XMLHttpRequest();
  x.open(method, WORKER + path);
  x.setRequestHeader("Content-Type", "application/json");
  x.setRequestHeader("X-User-Code", state.code);
  x.setRequestHeader("X-App-Version", APP_VERSION);
  x.onload = function () {
    var d;
    try { d = JSON.parse(x.responseText); } catch (e) { d = { error: "Phản hồi không hợp lệ" }; }
    if (x.status >= 200 && x.status < 300) cb(null, d); else cb(d.error || ("Lỗi " + x.status), d);
  };
  x.onerror = function () { cb("Không kết nối được máy chủ. Kiểm tra mạng.", null); };
  x.send(body ? JSON.stringify(body) : null);
}

/* ---------- Giao diện ---------- */
function busy(b) { $("sendBtn").disabled = b; $("scanBtn").disabled = b; }
function addMsg(cls, text) {
  var d = document.createElement("div");
  d.className = "msg " + cls;
  d.textContent = text;
  $("log").insertBefore(d, $("log").firstChild);
  return d;
}
function setBar(el, pct, invert) {
  /* pct: phần trăm đã dùng (invert=false) hoặc phần trăm còn lại (invert=true) */
  el.style.width = Math.max(0, Math.min(100, pct)) + "%";
  var bad = invert ? pct <= 10 : pct >= 90;
  var warn = invert ? pct <= 30 : pct >= 70;
  el.className = bad ? "bad" : (warn ? "warn" : "");
}
function updateQuota(q, meta) {
  if (!q) return;
  var pctUsed = q.limit_vnd > 0 ? Math.round(q.used_vnd * 100 / q.limit_vnd) : 100;
  $("quotaMe").textContent = "Hạn mức của chị: " + fmt(q.used_vnd) + " / " + fmt(q.limit_vnd) + " đ · hôm nay " + q.day_count + "/" + q.day_max + " lượt";
  setBar($("barMe"), pctUsed, false);
  $("quotaPool").textContent = "Quỹ chung còn " + q.pool_left_pct + "%";
  setBar($("barPool"), q.pool_left_pct, true);
  if (meta) $("badge").textContent = meta.label;
}
function showVersion() {
  var el = $("ver"); if (!el) return;
  var same = SERVER_VERSION === APP_VERSION;
  el.textContent = "Add-in v" + APP_VERSION + " · Máy chủ v" + (SERVER_VERSION || "?") + (SERVER_VERSION && !same ? " ⚠ chưa cập nhật đủ, nhờ admin dán lại code" : "");
  el.style.color = SERVER_VERSION && !same ? "#c00" : "#777";
}
function refreshMe() {
  api("/api/me", "GET", null, function (err, d) {
    if (err) {
      $("main").className = "hide"; $("login").className = "box";
      $("loginErr").textContent = state.code ? err : "";
      return;
    }
    $("login").className = "hide"; $("main").className = "";
    SERVER_VERSION = d.version || "cũ";
    showVersion();
    updateQuota(d.quota, null);
  });
}

/* ---------- Đọc ngữ cảnh bảng tính (gọn, có che dữ liệu nhạy cảm) ---------- */
function isDateFmt(nf) { return /[dy]/i.test(String(nf)) && !/[#0]/.test(String(nf)); }
function isTextNum(v) { return typeof v === "string" && /\d/.test(v) && /^\s*-?[\d.,]+\s*$/.test(v); }
function shortVal(v) {
  var s = String(v);
  if (/^\d{9,16}$/.test(s)) return "<ẨN>";
  return s.length > 30 ? s.substring(0, 30) + "…" : s;
}
function getContext(cb) {
  Excel.run(function (ctx) {
    var wb = ctx.workbook;
    var sel = wb.getSelectedRange(); sel.load("address");
    var sh = wb.worksheets.getActiveWorksheet(); sh.load("name");
    var sheets = wb.worksheets; sheets.load("items/name");
    var used = sh.getUsedRangeOrNullObject(); used.load("address,rowCount,columnCount,rowIndex,columnIndex");
    return ctx.sync().then(function () {
      var head = "Sheet đang mở: " + sh.name + "\nÔ đang chọn: " + sel.address + "\nCác sheet trong file: ";
      var names = [], i;
      for (i = 0; i < sheets.items.length; i++) names.push(sheets.items[i].name);
      head += names.join(", ") + "\n";
      if (used.isNullObject) return { text: head + "Sheet này đang trống." };
      var cols = Math.min(used.columnCount, 40);
      var statRows = Math.max(1, Math.min(used.rowCount, Math.floor(30000 / cols)));
      var rg = used.getCell(0, 0).getResizedRange(statRows - 1, cols - 1);
      rg.load("values,formulas,numberFormat");
      var others = [];
      for (i = 0; i < sheets.items.length && others.length < 6; i++) {
        if (sheets.items[i].name !== sh.name) {
          var u2 = sheets.items[i].getUsedRangeOrNullObject(); u2.load("address,columnIndex,rowIndex");
          others.push({ name: sheets.items[i].name, used: u2 });
        }
      }
      return ctx.sync().then(function () {
        var r0 = used.rowIndex, c0 = used.columnIndex;
        var v = rg.values, f = rg.formulas, nfm = rg.numberFormat;
        var out = [];
        out.push("Vùng dữ liệu: " + localAddr(used.address) + " (" + used.rowCount + " dòng x " + used.columnCount + " cột). Tiêu đề ở dòng " + (r0 + 1) + ", dữ liệu từ dòng " + (r0 + 2) + " đến dòng " + (r0 + used.rowCount) + "." + (statRows < used.rowCount ? " (thống kê chỉ trên " + statRows + " dòng đầu)" : ""));
        var c, r;
        for (c = 0; c < cols; c++) {
          var h = String(v[0][c]), L = colLetter(c0 + c);
          if (DANGER_HEADER.test(h)) { out.push(L + " (" + h + "): <ẨN> cột nhạy cảm"); continue; }
          var nonEmpty = 0, num = 0, txt = 0, err = 0, tn = 0, fcount = 0, date = 0, samples = [], fsample = "", dist = {}, nd = 0, tooMany = false;
          for (r = 1; r < v.length; r++) {
            var x = v[r][c];
            if (x === "" || x === null) continue;
            nonEmpty++;
            if (typeof f[r][c] === "string" && f[r][c].charAt(0) === "=") { fcount++; if (!fsample) fsample = "dòng " + (r0 + r + 1) + ": " + f[r][c]; }
            if (typeof x === "number") { num++; if (isDateFmt(nfm[r][c])) date++; }
            else if (typeof x === "string") {
              if (ERR_RE.test(x)) err++; else { txt++; if (isTextNum(x)) tn++; }
              if (!tooMany && !/^\d{9,16}$/.test(x)) { if (!(x in dist)) { dist[x] = 1; nd++; if (nd > 12) tooMany = true; } }
            }
            if (samples.length < 4) samples.push(shortVal(x));
          }
          var kind = num && !txt ? (date > num / 2 ? "ngày" : "số") : (txt && !num ? "chữ" : (nonEmpty ? "lẫn số và chữ" : "trống"));
          var line = L + " (" + h + "): " + kind + ", " + nonEmpty + " ô có dữ liệu; mẫu: " + samples.join(" | ");
          if (tn) line += "; " + tn + " ô là số lưu dạng chữ";
          if (err) line += "; " + err + " ô báo lỗi";
          if (fcount) line += "; có " + fcount + " công thức (" + fsample + ")";
          if (txt && !tooMany && nd > 0 && nd <= 12) line += "; giá trị khác nhau: " + Object.keys(dist).map(shortVal).join(", ");
          out.push(line);
        }
        var text = head + out.join("\n");
        if (others.length) {
          var hdrs = [];
          for (i = 0; i < others.length; i++) {
            if (others[i].used.isNullObject) continue;
            var hr = others[i].used.getCell(0, 0).getResizedRange(0, 11); hr.load("values");
            others[i].hr = hr; hdrs.push(others[i]);
          }
          return ctx.sync().then(function () {
            var extra = "";
            for (var k = 0; k < hdrs.length; k++) {
              var hv = hdrs[k].hr.values[0], hh = [];
              for (var m = 0; m < hv.length; m++) if (hv[m] !== "" && !DANGER_HEADER.test(String(hv[m]))) hh.push(colLetter(hdrs[k].used.columnIndex + m) + "=" + hv[m]);
              extra += "\nSheet '" + hdrs[k].name + "' (" + localAddr(hdrs[k].used.address) + "), tiêu đề: " + hh.join(", ");
            }
            return { text: text + extra };
          });
        }
        return { text: text };
      });
    });
  }).then(function (r) { cb(null, r); }, function (e) { cb(String(e && e.message || e), null); });
}

/* ---------- Gửi yêu cầu ---------- */
function parseReply(t) {
  var s = String(t || "").replace(/```json|```/g, "");
  var a = s.indexOf("{"), b = s.lastIndexOf("}");
  if (a >= 0 && b > a) {
    try {
      var j = JSON.parse(s.substring(a, b + 1));
      if (!j.actions || !j.actions.length) j.actions = [];
      return j;
    } catch (e) {}
  }
  return { message: String(t || ""), actions: [], warnings: null };
}
var OPS = { set_formula: 1, set_values: 1, format: 1, col_width: 1, freeze: 1, filter: 1, text_to_number: 1, trim_text: 1, add_sheet: 1, chart: 1 };
function describe(a) {
  var r = a.range ? " " + a.range : "";
  var sh = a.sheet ? " (sheet " + a.sheet + ")" : "";
  switch (a.op) {
    case "set_formula": return "Ghi công thức" + r + sh + ": " + a.formula;
    case "set_values": return "Điền nội dung" + r + sh;
    case "format":
      var bits = [];
      if (a.bold) bits.push("đậm"); if (a.fill) bits.push("nền " + a.fill); if (a.color) bits.push("chữ " + a.color);
      if (a.align) bits.push("căn " + a.align); if (a.wrap) bits.push("xuống dòng"); if (a.number_format) bits.push("định dạng số " + a.number_format);
      if (a.borders) bits.push("viền"); if (a.size) bits.push("cỡ chữ " + a.size);
      return "Định dạng" + r + sh + (bits.length ? ": " + bits.join(", ") : "");
    case "col_width": return "Chỉnh độ rộng cột" + r + sh + (a.width === "auto" ? " (tự vừa)" : "");
    case "freeze": return "Cố định " + (a.rows || 1) + " dòng đầu" + sh;
    case "filter": return "Bật bộ lọc" + r + sh;
    case "text_to_number": return "Đổi số lưu dạng chữ thành số" + r + sh;
    case "trim_text": return "Xóa khoảng trắng thừa" + r + sh;
    case "add_sheet": return "Thêm sheet " + a.name;
    case "chart": return "Vẽ biểu đồ " + (a.title || "") + " từ " + (a.range || "") + sh;
  }
  return String(a.op);
}

var RANGE_RE = /^\$?[A-Za-z]{1,3}(\$?\d{1,7})?(:\$?[A-Za-z]{1,3}(\$?\d{1,7})?)?$/;
function cleanRange(r) {
  r = String(r || "").replace(/\s/g, "");
  if (r.indexOf("!") >= 0) r = r.substring(r.lastIndexOf("!") + 1);
  r = r.replace(/\$/g, "");
  if (!RANGE_RE.test(r)) throw new Error("vùng '" + r + "' không hợp lệ");
  return r.toUpperCase();
}
function localAddr(a) { var i = a.lastIndexOf("!"); return i >= 0 ? a.substring(i + 1) : a; }
function sheetOf(ctx, a, base) { return ctx.workbook.worksheets.getItem(a.sheet ? String(a.sheet) : base); }

/* lấy vùng; nếu chỉ có cột (A:H) thì cắt theo vùng đã dùng */
function boundedRange(ctx, sheet, r) {
  if (/\d/.test(r)) return ctx.sync().then(function () { return sheet.getRange(r); });
  var used = sheet.getUsedRangeOrNullObject(); used.load("isNullObject");
  return ctx.sync().then(function () {
    if (used.isNullObject) return null;
    var inter = sheet.getRange(r).getIntersectionOrNullObject(used); inter.load("isNullObject");
    return ctx.sync().then(function () { return inter.isNullObject ? null : inter; });
  });
}

/* chụp trạng thái cũ để hoàn tác */
function snapRange(ctx, rng, wantFmt) {
  rng.load("address,rowCount,columnCount,formulas,numberFormat");
  rng.worksheet.load("name");
  return ctx.sync().then(function () {
    var s = { sheet: rng.worksheet.name, addr: localAddr(rng.address), formulas: rng.formulas, nf: rng.numberFormat, props: null, rows: rng.rowCount, cols: rng.columnCount, fmtSkipped: false };
    var n = rng.rowCount * rng.columnCount;
    if (!wantFmt) return s;
    if (n > 1500) { s.fmtSkipped = true; return s; }
    var cells = [], r, c;
    for (r = 0; r < rng.rowCount; r++) for (c = 0; c < rng.columnCount; c++) {
      var cell = rng.getCell(r, c);
      cell.format.fill.load("color"); cell.format.font.load("bold,italic,color,size"); cell.format.load("horizontalAlignment,wrapText");
      cells.push(cell);
    }
    return ctx.sync().then(function () {
      s.props = [];
      for (var i = 0; i < cells.length; i++) {
        var fm = cells[i].format;
        s.props.push({ fill: fm.fill.color, bold: fm.font.bold, italic: fm.font.italic, color: fm.font.color, size: fm.font.size, ha: fm.horizontalAlignment, wrap: fm.wrapText });
      }
      return s;
    });
  });
}
function restoreCells(ctx, s) {
  var rng = ctx.workbook.worksheets.getItem(s.sheet).getRange(s.addr);
  rng.numberFormat = s.nf;
  rng.formulas = s.formulas;
  if (s.props) {
    var i = 0;
    for (var r = 0; r < s.rows; r++) for (var c = 0; c < s.cols; c++) {
      var p = s.props[i++], cell = rng.getCell(r, c);
      if (p.fill === "" || p.fill === "#FFFFFF") cell.format.fill.clear(); else cell.format.fill.color = p.fill;
      cell.format.font.bold = p.bold; cell.format.font.italic = p.italic; cell.format.font.color = p.color; cell.format.font.size = p.size;
      cell.format.horizontalAlignment = p.ha; cell.format.wrapText = p.wrap;
    }
  }
}
var BORDER_IDS = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"];
function borderIds(rows, cols) {
  return BORDER_IDS.filter(function (id) { return !(id === "InsideHorizontal" && rows < 2) && !(id === "InsideVertical" && cols < 2); });
}
function parseNumberText(s) {
  var t = String(s).replace(/[\s ]/g, "");
  if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(t)) t = t.replace(/\./g, "").replace(",", ".");
  else if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(t)) t = t.replace(/,/g, "");
  else if (/^-?\d+,\d+$/.test(t)) t = t.replace(",", ".");
  if (!/^-?\d+(\.\d+)?$/.test(t)) return null;
  return Number(t);
}
var FORBID_F = /WEBSERVICE|CALL\(|REGISTER\.ID|EXEC\(|SHELL|DDE|IMPORTXML|\bFILTERXML\b/i;
var CHART_TYPES = { ColumnClustered: 1, BarClustered: 1, Line: 1, Pie: 1, Area: 1, ColumnStacked: 1 };

function doAction(ctx, a, base, undo) {
  var op = a.op;
  if (!OPS[op]) throw new Error("thao tác không được hỗ trợ");
  var sheet, rg;

  if (op === "add_sheet") {
    var nm = String(a.name || "").replace(/[:\\\/?*\[\]]/g, "").substring(0, 31);
    if (!nm) throw new Error("tên sheet không hợp lệ");
    var ex = ctx.workbook.worksheets.getItemOrNullObject(nm); ex.load("isNullObject");
    return ctx.sync().then(function () {
      if (!ex.isNullObject) return "Sheet " + nm + " đã có, dùng lại";
      ctx.workbook.worksheets.add(nm);
      return ctx.sync().then(function () { undo.push({ t: "addsheet", name: nm }); return "Đã thêm sheet " + nm; });
    });
  }

  sheet = sheetOf(ctx, a, base);

  if (op === "freeze") {
    var n = Math.max(1, Math.min(10, Number(a.rows) || 1));
    sheet.freezePanes.freezeRows(n);
    return ctx.sync().then(function () { undo.push({ t: "freeze", sheet: a.sheet || base }); return describe(a); });
  }

  if (op === "chart") {
    var type = CHART_TYPES[a.type] ? a.type : "ColumnClustered";
    var src = ctx.workbook.worksheets.getItem(String(a.source_sheet || a.sheet || base)).getRange(cleanRange(a.range));
    var chart = sheet.charts.add(type, src, "Auto");
    var cname = "AI_" + new Date().getTime();
    chart.name = cname;
    if (a.title) chart.title.text = String(a.title).substring(0, 100);
    if (a.at && /^[A-Za-z]{1,2}\d{1,4}$/.test(a.at)) chart.setPosition(a.at);
    return ctx.sync().then(function () { undo.push({ t: "chart", sheet: a.sheet || base, name: cname }); return describe(a); });
  }

  var r = cleanRange(a.range);

  if (op === "set_values") {
    if (!/\d/.test(r)) throw new Error("cần ghi rõ dòng");
    var vals = a.values;
    if (!vals || !vals.length || !vals[0] || !vals[0].length) throw new Error("thiếu nội dung");
    var cnt = 0, w = 0, k;
    for (k = 0; k < vals.length; k++) { if (vals[k].length > w) w = vals[k].length; cnt += vals[k].length; }
    if (cnt > 300) throw new Error("quá 300 ô");
    var grid = [];
    for (k = 0; k < vals.length; k++) { var row = []; for (var j = 0; j < w; j++) row.push(vals[k][j] === undefined || vals[k][j] === null ? "" : vals[k][j]); grid.push(row); }
    var tgt = sheet.getRange(r.split(":")[0]).getResizedRange(grid.length - 1, w - 1);
    return snapRange(ctx, tgt, false).then(function (s) {
      undo.push({ t: "cells", snap: s });
      tgt.values = grid;
      return ctx.sync().then(function () { return describe(a); });
    });
  }

  if (op === "set_formula") {
    var fm = String(a.formula || "");
    if (fm.charAt(0) !== "=" || fm.length > 1000 || FORBID_F.test(fm)) throw new Error("công thức không hợp lệ");
    if (!/\d/.test(r)) throw new Error("cần ghi rõ dòng");
    var tr = sheet.getRange(r);
    return snapRange(ctx, tr, false).then(function (s) {
      undo.push({ t: "cells", snap: s });
      var first = tr.getCell(0, 0);
      first.formulas = [[fm]];
      first.load("formulasR1C1");
      return ctx.sync().then(function () {
        var r1 = first.formulasR1C1[0][0];
        if (s.rows * s.cols > 1) {
          var arr = [];
          for (var i = 0; i < s.rows; i++) { var rw = []; for (var j2 = 0; j2 < s.cols; j2++) rw.push(r1); arr.push(rw); }
          tr.formulasR1C1 = arr;
        }
        tr.load("values");
        return ctx.sync().then(function () {
          var errs = 0, vv = tr.values;
          for (var x = 0; x < vv.length; x++) for (var y = 0; y < vv[x].length; y++) if (typeof vv[x][y] === "string" && ERR_RE.test(vv[x][y])) errs++;
          return describe(a) + (errs ? " ⚠ " + errs + " ô ra lỗi" : "");
        });
      });
    });
  }

  /* các thao tác còn lại làm trên vùng đã cắt theo vùng dữ liệu */
  return boundedRange(ctx, sheet, r).then(function (rng) {
    if (!rng) throw new Error("không có dữ liệu trong vùng này");

    if (op === "filter") {
      sheet.autoFilter.apply(rng);
      return ctx.sync().then(function () { undo.push({ t: "filter", sheet: a.sheet || base }); return describe(a); });
    }

    if (op === "col_width") {
      rng.load("columnCount");
      return ctx.sync().then(function () {
        var cs = [], i;
        for (i = 0; i < rng.columnCount; i++) { var cc = rng.getColumn(i); cc.format.load("columnWidth"); cs.push(cc); }
        return ctx.sync().then(function () {
          var ws = []; for (i = 0; i < cs.length; i++) ws.push(cs[i].format.columnWidth);
          rng.load("address");
          return ctx.sync().then(function () {
            undo.push({ t: "widths", sheet: a.sheet || base, addr: localAddr(rng.address), w: ws });
            if (a.width === "auto" || !a.width) rng.format.autofitColumns();
            else rng.format.columnWidth = Math.max(20, Math.min(400, Number(a.width) * 5.7));
            return ctx.sync().then(function () { return describe(a); });
          });
        });
      });
    }

    if (op === "format") {
      return snapRange(ctx, rng, true).then(function (s) {
        undo.push({ t: "cells", snap: s });
        var f = rng.format;
        if (a.bold !== undefined) f.font.bold = !!a.bold;
        if (a.italic !== undefined) f.font.italic = !!a.italic;
        if (typeof a.fill === "string" && /^#[0-9A-Fa-f]{6}$/.test(a.fill)) f.fill.color = a.fill;
        if (typeof a.color === "string" && /^#[0-9A-Fa-f]{6}$/.test(a.color)) f.font.color = a.color;
        if (a.size && Number(a.size) >= 6 && Number(a.size) <= 40) f.font.size = Number(a.size);
        var al = { left: "Left", center: "Center", right: "Right" }[String(a.align || "").toLowerCase()];
        if (al) f.horizontalAlignment = al;
        if (a.wrap !== undefined) f.wrapText = !!a.wrap;
        if (a.number_format) rng.numberFormat = String(a.number_format).substring(0, 60);
        if (a.borders) {
          var ids = borderIds(s.rows, s.cols);
          for (var i = 0; i < ids.length; i++) { var b = f.borders.getItem(ids[i]); b.style = "Continuous"; b.color = "#BFBFBF"; }
          undo.push({ t: "borders", sheet: s.sheet, addr: s.addr, rows: s.rows, cols: s.cols });
        }
        return ctx.sync().then(function () { return describe(a) + (s.fmtSkipped ? " (vùng lớn: màu/đậm không hoàn tác được)" : ""); });
      });
    }

    /* text_to_number, trim_text */
    rng.load("formulas,numberFormat");
    return ctx.sync().then(function () {
      var f2 = rng.formulas, nf2 = rng.numberFormat, changed = 0, i2, j3;
      var nw = [];
      for (i2 = 0; i2 < f2.length; i2++) {
        var rr = [];
        for (j3 = 0; j3 < f2[i2].length; j3++) {
          var cv = f2[i2][j3], nv = cv;
          if (typeof cv === "string" && cv.charAt(0) !== "=") {
            if (op === "trim_text") nv = cv.replace(/[\s ]+/g, " ").replace(/^\s+|\s+$/g, "");
            else if (isTextNum(cv)) { var pn = parseNumberText(cv); if (pn !== null) nv = pn; }
          }
          if (nv !== cv) changed++;
          rr.push(nv);
        }
        nw.push(rr);
      }
      if (!changed) return describe(a) + ": không có ô nào cần sửa";
      return snapRange(ctx, rng, false).then(function (s) {
        undo.push({ t: "cells", snap: s });
        if (op === "text_to_number") {
          for (i2 = 0; i2 < nw.length; i2++) for (j3 = 0; j3 < nw[i2].length; j3++)
            if (nw[i2][j3] !== f2[i2][j3] && nf2[i2][j3] === "@") rng.getCell(i2, j3).numberFormat = [["General"]];
        }
        rng.formulas = nw;
        return ctx.sync().then(function () { return describe(a) + " (" + changed + " ô)"; });
      });
    });
  });
}

/* chạy lần lượt từng thao tác, mỗi thao tác một Excel.run riêng để lỗi một chỗ không làm hỏng cả loạt */
function runActions(actions, done) {
  var undo = [], log = [];
  Excel.run(function (ctx) {
    var sh = ctx.workbook.worksheets.getActiveWorksheet(); sh.load("name");
    return ctx.sync().then(function () { return sh.name; });
  }).then(function (base) {
    var i = 0;
    function next() {
      if (i >= actions.length) { done(log, undo); return; }
      var a = actions[i++];
      Excel.run(function (ctx) { return doAction(ctx, a, base, undo); }).then(
        function (t) { log.push({ ok: true, t: t }); next(); },
        function (e) { log.push({ ok: false, t: describe(a) + " — không làm được: " + String(e && e.message || e) }); next(); }
      );
    }
    next();
  }, function (e) { log.push({ ok: false, t: "Không đọc được sheet: " + String(e && e.message || e) }); done(log, undo); });
}

function showResult(msgEl, log, undoBatch) {
  var box = document.createElement("div"); box.className = "small";
  var okN = 0;
  for (var i = 0; i < log.length; i++) {
    var l = document.createElement("div");
    l.textContent = (log[i].ok ? "✔ " : "✘ ") + log[i].t;
    if (!log[i].ok) l.style.color = "#c00"; else okN++;
    box.appendChild(l);
  }
  msgEl.appendChild(box);
  if (undoBatch.length) {
    state.undo.push(undoBatch);
    if (state.undo.length > 10) state.undo.shift();
    $("undoBtn").disabled = false;
    var t = document.createElement("div"); t.className = "small";
    t.textContent = "Không ưng thì bấm ↶ Hoàn tác (ở trên khung gõ).";
    msgEl.appendChild(t);
  }
}
function applyActions(actions, msgEl, btn) {
  if (btn) btn.disabled = true;
  busy(true);
  runActions(actions, function (log, undoBatch) {
    busy(false);
    showResult(msgEl, log, undoBatch);
  });
}

function showBot(j, meta, cut) {
  var d = document.createElement("div");
  d.className = "msg bot";
  var p = document.createElement("div");
  p.textContent = cut ? "Yêu cầu hơi lớn nên câu trả lời bị cắt giữa chừng. Chị chia nhỏ ra (vd: làm từng sheet, từng bước) rồi gửi lại nhé." : (j.message || "");
  d.appendChild(p);
  var acts = cut ? [] : j.actions.slice(0, 25);
  if (acts.length) {
    var ul = document.createElement("div"); ul.className = "small";
    for (var i = 0; i < acts.length; i++) {
      var li = document.createElement("div"); li.textContent = "• " + describe(acts[i]); ul.appendChild(li);
    }
    d.appendChild(ul);
  }
  if (j.warnings) { var w = document.createElement("div"); w.className = "small"; w.textContent = "Lưu ý: " + j.warnings; d.appendChild(w); }
  var m = document.createElement("div"); m.className = "meta";
  m.textContent = meta.label + " · " + fmt(meta.tokens_in) + " token vào / " + fmt(meta.tokens_out) + " ra · ~" + fmt(meta.cost_vnd) + "đ · vì: " + meta.reason;
  d.appendChild(m);
  $("log").insertBefore(d, $("log").firstChild);
  if (acts.length) {
    if ($("auto").checked) applyActions(acts, d, null);
    else {
      var ab = document.createElement("button"); ab.className = "pri"; ab.textContent = "Áp dụng " + acts.length + " thay đổi";
      ab.onclick = function () { ab.style.display = "none"; applyActions(acts, d, ab); };
      d.insertBefore(ab, m);
    }
  }
  var again = document.createElement("button"); again.textContent = "Làm kỹ hơn";
  again.onclick = function () { send(true); };
  d.insertBefore(again, m);
}
function send(deep) {
  var q = $("q").value.replace(/^\s+|\s+$/g, "");
  if (deep) q = state.lastQ;
  if (!q) return;
  state.lastQ = q;
  if (!deep) { addMsg("me", q); $("q").value = ""; }
  busy(true);
  getContext(function (err, c) {
    if (err) { busy(false); addMsg("bot", "Không đọc được bảng tính: " + err); return; }
    api("/api/chat", "POST", { question: q, context: c.text, compat: compat(), deep: !!deep }, function (e, d) {
      busy(false);
      if (e) { addMsg("bot", e); refreshMe(); return; }
      showBot(parseReply(d.reply), d.meta, d.cut);
      updateQuota(d.quota, d.meta);
    });
  });
}

/* ---------- Hoàn tác (lùi lại cả loạt thay đổi gần nhất) ---------- */
function undoEntry(e, cb) {
  Excel.run(function (ctx) {
    var wb = ctx.workbook, i;
    if (e.t === "cells") restoreCells(ctx, e.snap);
    else if (e.t === "widths") {
      var rng = wb.worksheets.getItem(e.sheet).getRange(e.addr);
      for (i = 0; i < e.w.length; i++) rng.getColumn(i).format.columnWidth = e.w[i];
    } else if (e.t === "borders") {
      var br = wb.worksheets.getItem(e.sheet).getRange(e.addr), ids = borderIds(e.rows, e.cols);
      for (i = 0; i < ids.length; i++) br.format.borders.getItem(ids[i]).style = "None";
    } else if (e.t === "freeze") wb.worksheets.getItem(e.sheet).freezePanes.unfreeze();
    else if (e.t === "filter") wb.worksheets.getItem(e.sheet).autoFilter.remove();
    else if (e.t === "addsheet") wb.worksheets.getItem(e.name).delete();
    else if (e.t === "chart") wb.worksheets.getItem(e.sheet).charts.getItem(e.name).delete();
    return ctx.sync();
  }).then(function () { cb(true); }, function () { cb(false); });
}
function undo() {
  var batch = state.undo.pop();
  if (!batch) return;
  $("undoBtn").disabled = true;
  var i = batch.length - 1, bad = 0;
  function next() {
    if (i < 0) {
      addMsg("bot", bad ? "Đã hoàn tác, nhưng có " + bad + " chỗ không lùi lại được (kiểm tra lại sheet)." : "Đã hoàn tác lần thay đổi gần nhất.");
      $("undoBtn").disabled = state.undo.length === 0;
      return;
    }
    undoEntry(batch[i--], function (ok) { if (!ok) bad++; next(); });
  }
  next();
}

/* ---------- Quét lỗi bằng code (miễn phí) ---------- */
function scan() {
  busy(true);
  $("scanOut").textContent = "Đang quét...";
  Excel.run(function (ctx) {
    var sh = ctx.workbook.worksheets.getActiveWorksheet(); sh.load("name");
    var used = sh.getUsedRange(); used.load("rowCount,columnCount,rowIndex,columnIndex");
    return ctx.sync().then(function () {
      var maxRows = Math.max(1, Math.min(used.rowCount, Math.floor(50000 / used.columnCount)));
      var rg = used.getCell(0, 0).getResizedRange(maxRows - 1, used.columnCount - 1);
      rg.load("values,formulasR1C1");
      return ctx.sync().then(function () {
        return { v: rg.values, f: rg.formulasR1C1, r0: used.rowIndex, c0: used.columnIndex, cut: maxRows < used.rowCount };
      });
    });
  }).then(function (d) {
    var items = [], v = d.v, f = d.f, i, c, k;
    for (c = 0; c < v[0].length; c++) {
      var counts = {}, nForm = 0, nNum = 0;
      for (i = 1; i < v.length; i++) {
        var fs = f[i][c];
        if (typeof fs === "string" && fs.charAt(0) === "=") { nForm++; counts[fs] = (counts[fs] || 0) + 1; }
        else if (typeof v[i][c] === "number") nNum++;
      }
      var top = null, topN = 0;
      for (k in counts) if (counts[k] > topN) { top = k; topN = counts[k]; }
      for (i = 1; i < v.length; i++) {
        var val = v[i][c], fs2 = f[i][c];
        var isF = typeof fs2 === "string" && fs2.charAt(0) === "=";
        var addr = colLetter(d.c0 + c) + (d.r0 + i + 1);
        if (typeof val === "string" && ERR_RE.test(val)) items.push({ a: addr, t: "Ô báo lỗi " + val });
        else if (isF && top && topN >= 3 && fs2 !== top && topN / nForm >= 0.7) items.push({ a: addr, t: "Công thức khác với phần còn lại của cột" });
        else if (!isF && typeof val === "number" && top && topN >= 3 && nForm / (nForm + nNum) >= 0.8) items.push({ a: addr, t: "Số gõ tay nằm giữa cột công thức" });
        else if (!isF && typeof val === "string" && /\d/.test(val) && /^\s*-?[\d.,]+\s*$/.test(val) && nNum > 0) items.push({ a: addr, t: "Số đang lưu dạng chữ" });
      }
    }
    var out = $("scanOut"); out.innerHTML = "";
    var head = document.createElement("div");
    head.textContent = items.length ? ("Tìm thấy " + items.length + " chỗ nghi ngờ" + (items.length > 100 ? " (hiện 100 đầu)" : "") + (d.cut ? ". Sheet lớn nên chỉ quét phần đầu." : ".")) : "Không thấy lỗi rõ ràng." + (d.cut ? " (chỉ quét phần đầu sheet lớn)" : "");
    out.appendChild(head);
    items.slice(0, 100).forEach(function (it) {
      var row = document.createElement("div"); row.className = "item";
      var b = document.createElement("button"); b.textContent = it.a;
      b.onclick = function () {
        Excel.run(function (ctx) { ctx.workbook.worksheets.getActiveWorksheet().getRange(it.a).select(); return ctx.sync(); });
      };
      row.appendChild(b);
      row.appendChild(document.createTextNode(" " + it.t));
      out.appendChild(row);
    });
    busy(false);
  }, function (e) { $("scanOut").textContent = "Không quét được: " + String(e && e.message || e); busy(false); });
}

/* ---------- Khởi động ---------- */
Office.onReady(function () {
  state.code = store("code") || "";
  var au = store("auto");
  $("auto").checked = au !== "0";
  $("auto").onchange = function () { store("auto", this.checked ? "1" : "0"); };
  $("loginBtn").onclick = function () {
    state.code = $("codeIn").value.replace(/\s/g, "");
    store("code", state.code);
    refreshMe();
  };
  $("sendBtn").onclick = function () { send(false); };
  $("scanBtn").onclick = scan;
  $("undoBtn").onclick = undo;
  var chips = document.querySelectorAll("button[data-q]");
  for (var i = 0; i < chips.length; i++) {
    chips[i].onclick = function () {
      $("q").value = this.getAttribute("data-q");
      send(false);
    };
  }
  if (state.code) refreshMe(); else $("login").className = "box";
});
