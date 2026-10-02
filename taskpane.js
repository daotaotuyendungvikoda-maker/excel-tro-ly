var WORKER = "https://excel-tro-ly.daotaotuyendungvikoda.workers.dev"; /* đã điền sẵn */
var APP_VERSION = "1.11.5"; /* phải khớp VERSION trong worker.js; tăng mỗi lần sửa */
var SERVER_VERSION = "";
var DANGER_HEADER = /(lương|luong|salary|cccd|cmnd|stk|tài khoản|tai khoan|mst|mã số thuế|thưởng|thuong)/i;
var ERR_RE = /^#(REF!|N\/A|DIV\/0!|VALUE!|NAME\?|NUM!|NULL!)/;
var state = { code: "", lastQ: "", undo: [], hist: [] };

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
function busy(b) {
  state.busy = b;
  $("sendBtn").className = b ? "hide" : "pri"; $("stopBtn").className = b ? "" : "hide";
  $("scanBtn").disabled = b;
  if (!b) state.cur = null;
}
function toBottom() { var c = $("chat"); c.scrollTop = c.scrollHeight; }
function addLog(d) {
  var w = $("welcome"); if (w) w.className = "hide";
  $("log").appendChild(d); toBottom();
}
function addMsg(cls, text) {
  var d = document.createElement("div");
  d.className = "msg " + cls;
  d.textContent = text;
  addLog(d);
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
  $("badge").textContent = "còn " + Math.max(0, 100 - pctUsed) + "%";
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
    var sel = wb.getSelectedRange(); sel.load("address,rowCount,columnCount,columnIndex,rowIndex");
    var sh = wb.worksheets.getActiveWorksheet(); sh.load("name");
    var sheets = wb.worksheets; sheets.load("items/name");
    return ctx.sync().then(function () {
      var used0 = sh.getUsedRangeOrNullObject(); used0.load("address,rowCount,columnCount,rowIndex,columnIndex");
      return ctx.sync().then(function () { return used0; }, function (e) { if (!LEGACY) throw e; return { isNullObject: true }; });
    }).then(function (used) {
      var head = "Sheet đang mở: " + sh.name + "\nÔ đang chọn: " + sel.address + "\nCác sheet trong file: ";
      var names = [], i;
      for (i = 0; i < sheets.items.length; i++) if (sheets.items[i].name.indexOf("~Lưu") !== 0) names.push(sheets.items[i].name);
      head += names.join(", ") + "\n";
      if (used.isNullObject) return { text: head + "Sheet này đang trống." };
      var cols = Math.min(used.columnCount, 40);
      var statRows = Math.max(1, Math.min(used.rowCount, Math.floor(30000 / cols)));
      var rg = used.getCell(0, 0).getResizedRange(statRows - 1, cols - 1);
      rg.load("values,formulas,numberFormat");
      var others = [];
      for (i = 0; i < sheets.items.length && others.length < 6 && !LEGACY; i++) {
        if (sheets.items[i].name !== sh.name) {
          var u2 = sheets.items[i].getUsedRangeOrNullObject(); u2.load("address,columnIndex,rowIndex");
          others.push({ name: sheets.items[i].name, used: u2 });
        }
      }
      return ctx.sync().then(function () {
        var r0 = used.rowIndex, c0 = used.columnIndex;
        var v = rg.values, f = rg.formulas, nfm = rg.numberFormat;
        var out = [], maskedAbs = {};
        out.push("Vùng dữ liệu: " + localAddr(used.address) + " (" + used.rowCount + " dòng x " + used.columnCount + " cột). Tiêu đề ở dòng " + (r0 + 1) + ", dữ liệu từ dòng " + (r0 + 2) + " đến dòng " + (r0 + used.rowCount) + "." + (statRows < used.rowCount ? " (thống kê chỉ trên " + statRows + " dòng đầu)" : ""));
        var c, r;
        for (c = 0; c < cols; c++) {
          var h = String(v[0][c]), L = colLetter(c0 + c);
          if (DANGER_HEADER.test(h)) { maskedAbs[c0 + c] = true; out.push(L + " (" + h + "): <ẨN> cột nhạy cảm"); continue; }
          var nonEmpty = 0, num = 0, txt = 0, err = 0, tn = 0, fcount = 0, date = 0, samples = [], fsample = "", dist = {}, nd = 0, tooMany = false, txtCells = [], errCells = [];
          for (r = 1; r < v.length; r++) {
            var x = v[r][c];
            if (x === "" || x === null) continue;
            nonEmpty++;
            if (typeof f[r][c] === "string" && f[r][c].charAt(0) === "=") { fcount++; if (!fsample) fsample = "dòng " + (r0 + r + 1) + ": " + f[r][c]; }
            if (typeof x === "number") { num++; if (isDateFmt(nfm[r][c])) date++; }
            else if (typeof x === "string") {
              var xa = L + (r0 + r + 1);
              if (ERR_RE.test(x)) { err++; if (errCells.length < 8) errCells.push(xa); } else { txt++; if (isTextNum(x)) tn++; if (txtCells.length < 8) txtCells.push(xa + "='" + shortVal(x) + "'"); }
              if (!tooMany && !/^\d{9,16}$/.test(x)) { if (!(x in dist)) { dist[x] = 1; nd++; if (nd > 12) tooMany = true; } }
            }
            if (samples.length < 4) samples.push(shortVal(x));
          }
          var oddF = [];
          if (fcount >= 3 && fcount / nonEmpty >= 0.6) {
            var cnt = {}, norm = [], rr, best = "", bestN = 0, k2;
            for (rr = 1; rr < v.length; rr++) {
              var fx = f[rr][c], nz = null;
              if (typeof fx === "string" && fx.charAt(0) === "=") { nz = fx.replace(new RegExp("([A-Z]+)" + (r0 + rr + 1) + "(?!\\d)", "g"), "$1#"); cnt[nz] = (cnt[nz] || 0) + 1; }
              norm.push(nz);
            }
            for (k2 in cnt) if (cnt[k2] > bestN) { best = k2; bestN = cnt[k2]; }
            for (rr = 1; rr < v.length && oddF.length < 6; rr++) {
              var fy = f[rr][c];
              if (fy === "" || fy === null) continue;
              if (norm[rr - 1] === null) oddF.push(L + (r0 + rr + 1) + " (số gõ tay " + shortVal(fy) + ")");
              else if (norm[rr - 1] !== best) oddF.push(L + (r0 + rr + 1) + " (" + fy + ")");
            }
          }
          var kind = num && !txt ? (date > num / 2 ? "ngày" : "số") : (txt && !num ? "chữ" : (nonEmpty ? "lẫn số và chữ" : "trống"));
          var line = L + " (" + h + "): " + kind + ", " + nonEmpty + " ô có dữ liệu; mẫu: " + samples.join(" | ");
          if (tn) line += "; " + tn + " ô là số lưu dạng chữ";
          if (err) line += "; " + err + " ô báo lỗi (" + errCells.join(", ") + ")";
          if (num > txt && txt > 0) line += "; Ô CHỮ LẪN TRONG CỘT SỐ: " + txtCells.join(", ");
          if (oddF.length) line += "; Ô LỆCH CÔNG THỨC SO VỚI CỘT: " + oddF.join(", ");
          if (fcount) line += "; có " + fcount + " công thức (" + fsample + ")";
          if (txt && !tooMany && nd > 0 && nd <= 12) line += "; giá trị khác nhau: " + Object.keys(dist).map(shortVal).join(", ");
          out.push(line);
        }
        function withSel(t) {
          var n = sel.rowCount * sel.columnCount;
          if (n > 60) return { text: t + "\nChị đang chọn vùng lớn " + localAddr(sel.address) + "." };
          sel.load("formulas");
          return ctx.sync().then(function () {
            var fs = sel.formulas, ln = [], a1, b1;
            for (a1 = 0; a1 < fs.length; a1++) for (b1 = 0; b1 < fs[a1].length; b1++) {
              var cv = maskedAbs[sel.columnIndex + b1] && sel.rowIndex !== r0 ? "<ẨN>" : shortVal(fs[a1][b1]);
              ln.push(colLetter(sel.columnIndex + b1) + (sel.rowIndex + a1 + 1) + "=" + cv);
            }
            return { text: t + "\nNội dung vùng chị đang chọn (" + localAddr(sel.address) + "): " + ln.join(" | ") };
          });
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
            return withSel(text + extra);
          });
        }
        return withSel(text);
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
var OPS = { query: 1, restore_formula: 1, mark_cells: 1, gridlines: 1, smart_format: 1, clean_data: 1, sort: 1, highlight: 1, dropdown: 1, set_formula: 1, set_values: 1, format: 1, col_width: 1, freeze: 1, filter: 1, text_to_number: 1, trim_text: 1, add_sheet: 1, chart: 1 };
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
    case "query": return "Truy vấn dữ liệu và xuất bảng kết quả" + sh;
    case "restore_formula": return "Khôi phục công thức cho " + (a.items || []).length + " ô" + sh;
    case "mark_cells": return "Tô nổi " + (a.cells || []).length + " ô cần chú ý" + sh;
    case "gridlines": return (a.show === false ? "Ẩn" : "Hiện") + " đường lưới" + sh;
    case "smart_format": return "Format cả bảng cho đẹp" + r + sh;
    case "clean_data": return "Dọn dữ liệu (khoảng trắng, số/ngày dạng chữ)" + r + sh;
    case "col_width": return "Chỉnh độ rộng cột" + r + sh + (a.width === "auto" ? " (tự vừa)" : "");
    case "freeze": return "Cố định " + (a.rows || 1) + " dòng đầu" + sh;
    case "filter": return "Bật bộ lọc" + r + sh;
    case "text_to_number": return "Đổi số lưu dạng chữ thành số" + r + sh;
    case "trim_text": return "Xóa khoảng trắng thừa" + r + sh;
    case "sort": return "Sắp xếp" + r + sh + " theo cột " + (a.by || "?") + (String(a.order).toLowerCase() === "desc" ? " (giảm dần)" : " (tăng dần)");
    case "highlight": return "Tô màu" + r + sh + " khi " + ({ duplicates: "trùng nhau", greater_than: "lớn hơn " + a.value, less_than: "nhỏ hơn " + a.value, equals: "bằng " + a.value, contains: "chứa " + a.value, formula: "thỏa điều kiện" }[String(a.rule)] || "thỏa điều kiện");
    case "dropdown": return "Tạo ô chọn sẵn" + r + sh + ": " + (a.items || []).slice(0, 6).join(", ");
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
  rng.load("address,rowCount,columnCount,columnIndex,formulas,numberFormat");
  rng.worksheet.load("name");
  return ctx.sync().then(function () {
    var s = { c0: rng.columnIndex, sheet: rng.worksheet.name, addr: localAddr(rng.address), formulas: rng.formulas, nf: rng.numberFormat, props: null, rows: rng.rowCount, cols: rng.columnCount, fmtSkipped: false };
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
function numFromText(s) {
  var m = /^\s*(-?\d[\d.,]*)\s*(ngày|ngay|công|đồng|vnđ|vnd|đ|₫)?\s*$/i.exec(String(s));
  return m ? parseNumberText(m[1]) : null;
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
    var wsAll = ctx.workbook.worksheets; wsAll.load("items/name");
    return ctx.sync().then(function () {
      var exists = false, q; for (q = 0; q < wsAll.items.length; q++) if (wsAll.items[q].name.toLowerCase() === nm.toLowerCase()) exists = true;
      if (exists) return "Sheet " + nm + " đã có, dùng lại";
      ctx.workbook.worksheets.add(nm);
      return ctx.sync().then(function () { undo.push({ t: "addsheet", name: nm }); return "Đã thêm sheet " + nm; });
    });
  }

  sheet = sheetOf(ctx, a, base);

  if (op === "restore_formula") {
    var its = (a.items || []).slice(0, 200).filter(function (x) { return x && /^[A-Za-z]{1,3}\d{1,7}$/.test(String(x.cell || "").replace(/\$/g, "")) && String(x.r1c1 || "").charAt(0) === "="; });
    if (!its.length) throw new Error("không có ô hợp lệ");
    return its.reduce(function (pr, it) {
      return pr.then(function () {
        var cc = sheet.getRange(String(it.cell).replace(/\$/g, ""));
        return snapRange(ctx, cc, false).then(function (sn) { undo.push({ t: "cells", snap: sn }); cc.formulasR1C1 = [[String(it.r1c1)]]; return ctx.sync(); });
      });
    }, Promise.resolve()).then(function () { return "Đã khôi phục công thức chuẩn cho " + its.length + " ô"; });
  }
  if (op === "mark_cells") {
    var cl = (a.cells || []).map(function (x) { return String(x).replace(/\$/g, ""); }).filter(function (x) { return /^[A-Za-z]{1,3}\d{1,7}$/.test(x); }).slice(0, 200);
    if (!cl.length) throw new Error("không có ô hợp lệ");
    var fillc = /^#[0-9A-Fa-f]{6}$/.test(String(a.fill || "")) ? a.fill : "#FFF2CC";
    var cobj = cl.map(function (x) { var c = sheet.getRange(x); c.format.fill.load("color"); return c; });
    return ctx.sync().then(function () {
      var old = cobj.map(function (c, i) { return { a: cl[i], color: c.format.fill.color }; });
      cobj.forEach(function (c) { c.format.fill.color = fillc; });
      return ctx.sync().then(function () { undo.push({ t: "marks", sheet: a.sheet || base, items: old }); return "Đã tô nổi " + cl.length + " ô (không đổi dữ liệu)"; });
    });
  }
  if (op === "gridlines") {
    sheet.load("showGridlines");
    return ctx.sync().then(function () {
      var was = sheet.showGridlines;
      sheet.showGridlines = a.show !== false;
      return ctx.sync().then(function () { undo.push({ t: "grid", sheet: a.sheet || base, was: was }); return describe(a); });
    });
  }
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

    if (op === "sort") {
      var by = String(a.by || "").replace(/[^A-Za-z]/g, "").toUpperCase();
      if (!by) throw new Error("thiếu cột để sắp xếp");
      var ci = 0, q2;
      for (q2 = 0; q2 < by.length; q2++) ci = ci * 26 + by.charCodeAt(q2) - 64;
      ci -= 1;
      return snapRange(ctx, rng, true).then(function (s) {
        var key = ci - s.c0;
        if (key < 0 || key >= s.cols) throw new Error("cột " + by + " nằm ngoài vùng");
        undo.push({ t: "cells", snap: s });
        rng.sort.apply([{ key: key, ascending: String(a.order).toLowerCase() !== "desc" }], false, true, "Rows");
        return ctx.sync().then(function () { return describe(a) + (s.fmtSkipped ? " (vùng lớn: màu/đậm không hoàn tác được)" : ""); });
      });
    }

    if (op === "highlight") {
      var rule = String(a.rule || "").toLowerCase();
      var fillC = /^#[0-9A-Fa-f]{6}$/.test(a.fill || "") ? a.fill : "#FFC7CE";
      var fontC = /^#[0-9A-Fa-f]{6}$/.test(a.color || "") ? a.color : "#9C0006";
      var v = String(a.value === undefined ? "" : a.value);
      var cf, fmtObj;
      if (rule === "duplicates") { cf = rng.conditionalFormats.add("PresetCriteria"); cf.preset.rule = { criterion: "DuplicateValues" }; fmtObj = cf.preset.format; }
      else if (rule === "greater_than" || rule === "less_than" || rule === "equals") {
        var f1 = /^-?\d+(\.\d+)?$/.test(v) ? "=" + v : (v.charAt(0) === "=" ? v : '="' + v.replace(/"/g, '""') + '"');
        cf = rng.conditionalFormats.add("CellValue");
        cf.cellValue.rule = { formula1: f1, operator: rule === "greater_than" ? "GreaterThan" : (rule === "less_than" ? "LessThan" : "EqualTo") };
        fmtObj = cf.cellValue.format;
      } else if (rule === "contains") {
        cf = rng.conditionalFormats.add("ContainsText");
        cf.textComparison.rule = { operator: "Contains", text: v };
        fmtObj = cf.textComparison.format;
      } else if (rule === "formula") {
        if (v.charAt(0) !== "=" || FORBID_F.test(v)) throw new Error("công thức điều kiện không hợp lệ");
        cf = rng.conditionalFormats.add("Custom");
        cf.custom.rule.formula = v;
        fmtObj = cf.custom.format;
      } else throw new Error("kiểu điều kiện chưa hỗ trợ");
      fmtObj.fill.color = fillC; fmtObj.font.color = fontC;
      rng.load("address"); rng.worksheet.load("name");
      return ctx.sync().then(function () {
        undo.push({ t: "cf", sheet: rng.worksheet.name, addr: localAddr(rng.address) });
        return describe(a);
      });
    }

    if (op === "dropdown") {
      var items = (a.items || []).slice(0, 30).map(function (x) { return String(x).replace(/,/g, " "); });
      if (!items.length) throw new Error("thiếu danh sách");
      rng.dataValidation.rule = { list: { inCellDropDown: true, source: items.join(",") } };
      rng.load("address"); rng.worksheet.load("name");
      return ctx.sync().then(function () {
        undo.push({ t: "validation", sheet: rng.worksheet.name, addr: localAddr(rng.address) });
        return describe(a);
      });
    }

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
            return ctx.sync().then(function () { return describe(a); }, function (e) {
              if (a.width === "auto" || !a.width) { rng.format.columnWidth = 90; return ctx.sync().then(function () { return describe(a) + " (đặt độ rộng 90 vì Excel không tự vừa được)"; }); }
              throw e;
            });
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
            else { var pn = numFromText(cv); if (pn !== null) nv = pn; }
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

/* ---------- Nhận câu trả lời chạy dần ---------- */
function canStream() { return typeof fetch === "function" && typeof ReadableStream !== "undefined" && typeof TextDecoder !== "undefined"; }
function streamChat(body, h) {
  var finished = false, aborted = false, ac = (typeof AbortController !== "undefined") ? new AbortController() : null;
  fetch(WORKER + "/api/chat", { signal: ac ? ac.signal : undefined, method: "POST", headers: { "Content-Type": "application/json", "X-User-Code": state.code, "X-App-Version": APP_VERSION }, body: JSON.stringify(body) }).then(function (r) {
    if (!r.ok) return r.json().then(function (d) { h.onError((d && d.error) || ("Lỗi " + r.status)); }, function () { h.onError("Lỗi " + r.status); });
    if (!r.body || !r.body.getReader) { h.onError("Trình duyệt không nhận được dữ liệu dần"); return; }
    var reader = r.body.getReader(), dec = new TextDecoder("utf-8"), buf = "", text = "";
    function pump() {
      return reader.read().then(function (x) {
        if (aborted) return;
        if (x.done) { if (!finished) h.onError("Mất kết nối giữa chừng. Thử lại."); return; }
        buf += dec.decode(x.value, { stream: true });
        var parts = buf.split("\n\n"); buf = parts.pop();
        for (var i = 0; i < parts.length; i++) {
          var ev; try { ev = JSON.parse(parts[i].replace(/^data:\s*/, "")); } catch (e) { continue; }
          if (ev.t) { text += ev.t; h.onText(text); }
          else if (ev.error) { finished = true; h.onError(ev.error); }
          else if (ev.done) { finished = true; h.onDone(ev, text); }
        }
        return pump();
      }, function () { if (!aborted && !finished) { finished = true; h.onError("Mất kết nối giữa chừng."); } });
    }
    return pump();
  }, function () { if (!aborted) h.onError("Không kết nối được máy chủ."); });
  return { abort: function () { aborted = true; finished = true; try { if (ac) ac.abort(); } catch (e) {} } };
}
/* lấy từng thao tác ra ngay khi AI viết xong nó */
function ActionParser() { this.i = 0; this.depth = 0; this.inStr = false; this.esc = false; this.start = -1; this.on = false; this.end = false; }
ActionParser.prototype.feed = function (buf) {
  var out = [];
  if (this.end) return out;
  if (!this.on) {
    var m = /"actions"\s*:\s*\[/.exec(buf);
    if (!m) return out;
    this.on = true; this.i = m.index + m[0].length; this.depth = 0;
  }
  for (; this.i < buf.length; this.i++) {
    var ch = buf.charAt(this.i);
    if (this.inStr) { if (this.esc) this.esc = false; else if (ch === "\\") this.esc = true; else if (ch === '"') this.inStr = false; continue; }
    if (ch === '"') { this.inStr = true; continue; }
    if (ch === "{") { if (this.depth === 0) this.start = this.i; this.depth++; }
    else if (ch === "}") {
      this.depth--;
      if (this.depth === 0 && this.start >= 0) { try { out.push(JSON.parse(buf.substring(this.start, this.i + 1))); } catch (e) {} this.start = -1; }
    } else if (ch === "]" && this.depth === 0) { this.end = true; this.i++; break; }
  }
  return out;
};

/* chế độ không ghi đè */
function editOK() { var e = $("edit"); return !e || e.checked; }
var NO_OVERWRITE_NOTE = "\n\nCHẾ ĐỘ KHÔNG GHI ĐÈ đang BẬT: tuyệt đối không sửa hay ghi vào ô đang có dữ liệu, không dùng clean_data, text_to_number, trim_text, sort. Chỉ được ghi vào ô trống, cột mới bên phải bảng hoặc sheet mới, và tô nổi ô bằng mark_cells. Nếu cần sửa dữ liệu có sẵn thì chỉ liệt kê trong issues/formulas và nhắc chị bật 'Cho phép sửa trực tiếp'.";
var GUARD_ALWAYS = { restore_formula: 1, clean_data: 1, text_to_number: 1, trim_text: 1, sort: 1 };
function guardCheck(a, base, cb) {
  if (editOK()) { cb(null); return; }
  if (GUARD_ALWAYS[a.op]) { cb("đang bật chế độ không ghi đè nên không được sửa dữ liệu có sẵn (chỉ ghi ô trống/cột mới/sheet mới)"); return; }
  if (a.op !== "set_formula" && a.op !== "set_values") { cb(null); return; }
  var r = cleanRange(a.range);
  if (!/\d/.test(r)) { cb(null); return; }
  var first = r.split(":")[0], rr = r;
  Excel.run(function (ctx) {
    var sh = ctx.workbook.worksheets.getItem(a.sheet ? String(a.sheet) : base), rg;
    if (a.op === "set_values") {
      var v = a.values || [], w = 0, k; for (k = 0; k < v.length; k++) if ((v[k] || []).length > w) w = v[k].length;
      rg = sh.getRange(first).getResizedRange(Math.max(0, v.length - 1), Math.max(0, w - 1));
    } else rg = sh.getRange(rr);
    var used = sh.getUsedRangeOrNullObject(); used.load("isNullObject");
    return ctx.sync().then(function () {
      if (used.isNullObject) return false;
      var inter = rg.getIntersectionOrNullObject(used); inter.load("isNullObject,address");
      return ctx.sync().then(function () {
        if (inter.isNullObject) return false;
        inter.load("values"); return ctx.sync().then(function () {
          for (var i = 0; i < inter.values.length; i++) for (var j = 0; j < inter.values[i].length; j++) if (inter.values[i][j] !== "") return localAddr(inter.address);
          return false;
        });
      });
    });
  }).then(function (hit) { cb(hit ? "đang bật chế độ không ghi đè nhưng vùng " + hit + " đã có dữ liệu (hãy ghi vào ô trống/cột mới/sheet mới)" : null); }, function () { cb(null); });
}

/* bản sao an toàn: chép sheet thành sheet ẩn trước khi sửa, giữ 3 bản gần nhất */
var NEED_BACKUP = { set_formula: 1, set_values: 1, format: 1, text_to_number: 1, trim_text: 1, sort: 1, clean_data: 1, smart_format: 1, highlight: 1, dropdown: 1, col_width: 1 };
function backupSheet(base, cb) {
  var d = new Date(), pad = function (n) { return (n < 10 ? "0" : "") + n; };
  var nm = "~Lưu " + pad(d.getHours()) + pad(d.getMinutes()) + pad(d.getSeconds());
  if (!supports("1.7")) { cb(null); return; }
  Excel.run(function (ctx) {
    var act = ctx.workbook.worksheets.getActiveWorksheet(); act.load("name");
    var all = ctx.workbook.worksheets, cp;
    return ctx.sync().then(function () {
      cp = all.getItem(base).copy("End"); cp.name = nm; cp.visibility = "Hidden";
      all.load("items/name");
      return ctx.sync();
    }).then(function () {
      /* giữ nguyên sheet đang xem: sao chép sheet có thể làm Excel nhảy sang sheet khác */
      try { all.getItem(act.name).activate(); } catch (e0) {}
      return ctx.sync();
    }).then(function () {
      var bk = [], i;
      for (i = 0; i < all.items.length; i++) if (all.items[i].name.indexOf("~Lưu ") === 0) bk.push(all.items[i]);
      bk.sort(function (x, y) { return x.name < y.name ? -1 : 1; });
      for (i = 0; i < bk.length - 3; i++) bk[i].delete();
      return ctx.sync();
    });
  }).then(function () { cb(nm); }, function () { cb(null); });
}

/* hàng đợi: làm lần lượt từng thao tác, hiện kết quả ngay */
function Runner(bot, onDone) {
  this.bot = bot; this.q = []; this.running = false; this.closed = false; this.finished = false;
  this.undo = []; this.log = []; this.base = null; this.onDone = onDone;
}
Runner.prototype.push = function (a) { this.q.push(a); this.total = (this.total || 0) + 1; this.pump(); };
Runner.prototype.close = function () { this.closed = true; this.pump(); };
Runner.prototype.pump = function () {
  var self = this;
  if (self.running) return;
  if (!self.q.length) {
    if (self.closed && !self.finished) {
      self.finished = true;
      var nbad = 0, nwarn = 0, z;
      for (z = 0; z < self.log.length; z++) { if (!self.log[z].ok) nbad++; else if (self.log[z].warn) nwarn++; }
      var sm = self.bot.sum; sm.className = "step"; sm.style.color = nbad ? "#c00" : (nwarn ? "#b36b00" : "#57606a");
      sm.textContent = (self.bot.body.className === "hide" ? "▸ " : "▾ ") + (nbad ? "⚠ " : "✔ ") + "Đã làm " + self.log.length + " bước" + (nbad ? ", " + nbad + " bước không làm được" : "") + (nwarn ? ", " + nwarn + " chưa sửa được gì" : "") + " (bấm để xem chi tiết)";
      if (nbad || nwarn) { self.bot.body.className = ""; sm.textContent = sm.textContent.replace(/^▸/, "▾"); }
      self.onDone(self.log, self.undo);
    }
    return;
  }
  self.running = true;
  var a = self.q.shift();
  /* sau add_sheet, thao tác ghi không nói rõ sheet thì hiểu là ghi vào sheet mới, tránh đè lên sheet dữ liệu */
  if (a.op === "add_sheet") self.newSheet = String(a.name || "").replace(/[:\\\/?*\[\]]/g, "").substring(0, 31);
  else if (self.newSheet && !a.sheet && /^(set_values|set_formula|format|col_width|freeze|gridlines|smart_format|chart)$/.test(a.op)) a.sheet = self.newSheet;
  var line = document.createElement("div"); line.className = "step"; line.textContent = "⏳ " + describe(a);
  self.bot.body.appendChild(line);
  self.doneN = (self.doneN || 0) + 1;
  self.bot.sum.className = "step"; self.bot.sum.style.color = "#1a56c4";
  self.bot.sum.textContent = (self.bot.body.className === "hide" ? "▸ " : "▾ ") + "⏳ " + describe(a) + " (" + self.doneN + "/" + Math.max(self.total || 1, self.doneN) + ")";
  function fail(e) {
    var msg = describe(a) + " — không làm được: " + String(e && e.message || e);
    self.log.push({ ok: false, t: msg }); line.textContent = "✘ " + msg; line.style.color = "#c00";
    self.running = false; self.pump();
  }
  function special(base) {
    var o = { sheet: a.sheet ? String(a.sheet) : base, range: a.range };
    function fin(err, res, u) {
      if (err) { fail(new Error(err)); return; }
      for (var k = 0; k < (u || []).length; k++) self.undo.push(u[k]);
      var t;
      if (a.op === "smart_format" || a.op === "query") t = res;
      else {
        var ps = []; if (res.ch.trim) ps.push(res.ch.trim + " ô khoảng trắng"); if (res.ch.num) ps.push(res.ch.num + " ô chữ→số"); if (res.ch.date) ps.push(res.ch.date + " ô chữ→ngày");
        t = ps.length ? "Đã dọn: " + ps.join(", ") : "Dọn dữ liệu: không có ô nào cần sửa";
        if (res.errs) t += " · còn " + res.errs + " ô ra lỗi";
      }
      var warn = /không có ô nào cần sửa/.test(t);
      self.log.push({ ok: true, warn: warn, t: t });
      line.textContent = (warn ? "⚠ " : "✔ ") + t; line.style.color = warn ? "#b36b00" : "";
      self.running = false; self.pump();
    }
    if (a.op === "query") runQueryOp(a, o.sheet, fin); else if (a.op === "smart_format") smartFormat(fin, o); else localClean(fin, o);
  }
  function step(base) {
    self.base = base;
    guardCheck(a, base, function (gm) { if (gm) { fail(new Error(gm)); return; } step2(base); });
  }
  function step2(base) {
    if (!self.backedUp && NEED_BACKUP[a.op]) {
      self.backedUp = true;
      backupSheet(a.sheet ? String(a.sheet) : base, function (nm) {
        if (nm) noteLine(self.bot, "Đã lưu bản sao an toàn (sheet ẩn \"" + nm + "\": chuột phải tên sheet → Hiện/Unhide để lấy lại nếu cần).");
        go(base);
      });
      return;
    }
    go(base);
  }
  function go(base) {
    if (a.op === "smart_format" || a.op === "clean_data" || a.op === "query") { special(base); return; }
    Excel.run(function (ctx) { return doAction(ctx, a, base, self.undo); }).then(function (t) {
      var warn = /không có ô nào cần sửa/.test(t);
      self.log.push({ ok: true, warn: warn, t: t });
      line.textContent = (warn ? "⚠ " : "✔ ") + t + (warn ? " (chưa sửa được gì)" : "");
      if (warn) line.style.color = "#b36b00"; else line.style.color = "";
      self.running = false; self.pump();
    }, fail);
  }
  if (self.base) step(self.base);
  else Excel.run(function (ctx) {
    var sh = ctx.workbook.worksheets.getActiveWorksheet(); sh.load("name");
    return ctx.sync().then(function () { return sh.name; });
  }).then(step, fail);
};

function makeBot() {
  var d = document.createElement("div"); d.className = "msg bot";
  var st = document.createElement("div"); st.className = "status"; d.appendChild(st);
  var msg = document.createElement("div"); d.appendChild(msg);
  var log = document.createElement("div"); d.appendChild(log);
  var sum = document.createElement("div"); sum.className = "step hide"; sum.style.cssText = "cursor:pointer;user-select:none"; log.appendChild(sum);
  var body = document.createElement("div"); body.className = "hide"; body.style.cssText = "margin:2px 0 2px 6px;border-left:2px solid #d0d7de;padding-left:8px"; log.appendChild(body);
  sum.onclick = function () { var open = body.className === "hide"; body.className = open ? "" : "hide"; sum.textContent = sum.textContent.replace(/^[▸▾]/, open ? "▾" : "▸"); };
  var extra = document.createElement("div"); d.appendChild(extra);
  var meta = document.createElement("div"); meta.className = "meta"; d.appendChild(meta);
  addLog(d);
  return { el: d, st: st, msg: msg, log: log, sum: sum, body: body, extra: extra, meta: meta };
}
function setStatus(bot, t) { bot.st.textContent = t || ""; }
function noteLine(bot, text, color) {
  var n = document.createElement("div"); n.className = "small"; n.textContent = text; if (color) n.style.color = color; bot.extra.appendChild(n);
}

function needVerify(acts) {
  for (var i = 0; i < (acts || []).length; i++) if (/^(set_formula|set_values|add_sheet|sort|clean_data|text_to_number|highlight)$/.test(acts[i].op)) return true;
  return false;
}
function runVerify(bot, acts, base, opts) {
  var reads = [], seen = {};
  acts.forEach(function (a) {
    var sh = a.sheet ? String(a.sheet) : base, r = a.op === "add_sheet" ? "" : String(a.range || "");
    if (a.op === "add_sheet") sh = String(a.name || "");
    var key = sh + "!" + r;
    if (seen[key] || reads.length >= 4 || !sh) return;
    if (a.op !== "add_sheet" && !/\d/.test(r)) return;
    seen[key] = 1; reads.push({ sheet: sh, range: r });
  });
  if (!reads.length) { busy(false); return; }
  setStatus(bot, "⏳ Đang kiểm tra lại kết quả…");
  readRanges(reads, function (err, txt) {
    if (err || !txt || opts.stopped) { setStatus(bot, ""); busy(false); return; }
    var q = "[KIỂM TRA LẠI SAU KHI LÀM]\nYêu cầu của chị: " + state.lastQ + "\nĐã làm: " + acts.slice(0, 12).map(describe).join("; ") + "\nKết quả thực tế đọc lại từ Excel:" + txt + "\nHãy soát như người kế toán kiểm tra file: ô lỗi, số âm/0/trống bất thường, tổng không khớp, công thức trỏ sai vùng, tiêu đề sai, trình bày chưa gọn. Có gì sai hoặc chưa tốt thì sửa bằng actions (chỉ phần cần sửa) và nói ngắn đã sửa gì; nếu ổn thì actions rỗng và message bắt đầu bằng 'Đã kiểm tra:' kèm 1 câu những gì đã soát.";
    runChat(q, { verify: true, bot: bot });
  });
}
function finishRun(bot, log, undoB, opts, acts, base) {
  setStatus(bot, "");
  if (undoB.length) {
    if (opts.retry && state.undo.length) state.undo[state.undo.length - 1] = state.undo[state.undo.length - 1].concat(undoB);
    else state.undo.push(undoB);
    if (state.undo.length > 10) state.undo.shift();
    $("undoBtn").disabled = false;
    noteLine(bot, "Không ưng thì bấm ↶ Hoàn tác (ở trên khung gõ).");
  }
  var bad = [];
  for (var i = 0; i < log.length; i++) if (!log[i].ok || /ô ra lỗi/.test(log[i].t)) bad.push(log[i].t);
  if (bad.length && !opts.retry && !opts.stopped) {
    noteLine(bot, "Tự kiểm tra thấy chỗ chưa ổn, đang nhờ trợ lý sửa lại…", "#b36b00");
    runChat("Lần làm trước có chỗ chưa ổn: " + bad.join(" ; ").substring(0, 700) + ". Yêu cầu gốc của chị: " + state.lastQ + ". Hãy sửa lại cho đúng (chỉ làm phần cần sửa).", { deep: true, retry: true });
    return;
  }
  if (!opts.verify && !opts.retry && !opts.stopped && $("verify").checked && needVerify(acts) && base) { runVerify(bot, acts, base, opts); return; }
  busy(false);
}

/* ---------- Khối hiển thị trực quan trong khung chat: thẻ chỉ số, thẻ công thức, danh sách vấn đề ---------- */
function mk(tag, css, text) { var e = document.createElement(tag); if (css) e.style.cssText = css; if (text !== undefined) e.textContent = text; return e; }
function gotoRef(ref) {
  var sh = null, addr = String(ref).replace(/\$/g, ""), k = addr.lastIndexOf("!");
  if (k >= 0) { sh = addr.substring(0, k).replace(/^'|'$/g, ""); addr = addr.substring(k + 1); }
  Excel.run(function (ctx) {
    var ws = sh ? ctx.workbook.worksheets.getItem(sh) : ctx.workbook.worksheets.getActiveWorksheet();
    if (sh) ws.activate();
    ws.getRange(addr).select(); return ctx.sync();
  }).then(null, function () {});
}
function copyText(t) {
  try { var ta = document.createElement("textarea"); ta.value = t; ta.style.cssText = "position:fixed;opacity:0;top:0;left:0"; document.body.appendChild(ta); ta.select(); var ok = document.execCommand("copy"); document.body.removeChild(ta); return ok; } catch (e) { return false; }
}
function renderKpis(bot, kpis) {
  var g = mk("div", "display:flex;flex-wrap:wrap;margin:8px -3px 2px");
  kpis.slice(0, 8).forEach(function (k) {
    var c = mk("div", "flex:1 1 42%;min-width:110px;margin:3px;padding:8px 10px;background:#fff;border:1px solid #d8dee4;border-radius:10px;" + (k.ref ? "cursor:pointer" : ""));
    c.appendChild(mk("div", "font-size:11px;color:#57606a", String(k.label || "")));
    c.appendChild(mk("div", "font-size:16px;font-weight:600;color:#1F4E78;margin:2px 0", String(k.value === undefined ? "" : k.value)));
    if (k.note) c.appendChild(mk("div", "font-size:11px;color:#8c959f", String(k.note)));
    if (k.ref) { c.title = "Bấm để tới " + k.ref; c.onclick = function () { gotoRef(k.ref); }; }
    g.appendChild(c);
  });
  bot.extra.appendChild(g);
}
function insertFormula(f, cell, bot) {
  Excel.run(function (ctx) {
    var tgt;
    if (cell && /^[A-Za-z]{1,3}\d{1,7}$/.test(String(cell).replace(/\$/g, ""))) tgt = ctx.workbook.worksheets.getActiveWorksheet().getRange(String(cell).replace(/\$/g, ""));
    else tgt = ctx.workbook.getSelectedRange().getCell(0, 0);
    tgt.load("formulas,address");
    return ctx.sync().then(function () {
      if (!editOK() && tgt.formulas[0][0] !== "") throw new Error("Ô " + localAddr(tgt.address) + " đang có dữ liệu mà chế độ không ghi đè đang bật. Chị chọn một ô trống rồi bấm lại.");
      return snapRange(ctx, tgt, false).then(function (sn) {
        tgt.formulas = [[f]];
        return ctx.sync().then(function () { pushUndo([{ t: "cells", snap: sn }]); return localAddr(tgt.address); });
      });
    });
  }).then(function (ad) { noteLine(bot, "✔ Đã chèn công thức vào " + ad + ". Không ưng thì bấm ↶ Hoàn tác."); toBottom(); }, function (e) { noteLine(bot, String(e && e.message || e), "#c00"); toBottom(); });
}
function renderFormulas(bot, list) {
  list.slice(0, 5).forEach(function (f) {
    var fm = String(f.formula || ""); if (fm.charAt(0) !== "=") return;
    var c = mk("div", "margin:8px 0 2px;padding:8px 10px;background:#fff;border:1px solid #d8dee4;border-radius:10px");
    c.appendChild(mk("div", "font-weight:600;font-size:12.5px", String(f.title || "Gợi ý công thức")));
    c.appendChild(mk("div", "font-family:Consolas,monospace;font-size:12px;background:#f6f8fa;border-radius:6px;padding:5px 7px;margin:5px 0;word-break:break-all", fm));
    if (f.explain) c.appendChild(mk("div", "font-size:12px;color:#57606a;margin-bottom:4px", String(f.explain)));
    var b1 = mk("button", "", "Chép"); b1.onclick = function () { b1.textContent = copyText(fm) ? "Đã chép ✔" : "Bấm Ctrl+C trên ô công thức"; };
    var cellOk = f.cell && /^[A-Za-z]{1,3}\d{1,7}$/.test(String(f.cell).replace(/\$/g, ""));
    var b2 = mk("button", "", cellOk ? "Chèn vào " + String(f.cell) : "Chèn vào ô đang chọn"); b2.onclick = function () { insertFormula(fm, f.cell, bot); };
    c.appendChild(b1); c.appendChild(b2);
    bot.extra.appendChild(c);
  });
}
function renderIssues(bot, issues, sheetName, opts) {
  opts = opts || {};
  var items = [];
  issues.slice(0, 200).forEach(function (it) {
    var ref = String(it.ref || it.a || "").replace(/\$/g, "");
    if (!ref) return;
    items.push({ a: ref, t: String(it.text || it.t || ""), type: String(it.type || ""), fix: String(it.fix || ""), restore: it.restore || null });
  });
  if (!items.length) return;
  var box = mk("div", "margin:8px 0 2px;padding:6px 8px;background:#fff;border:1px solid #d8dee4;border-radius:10px");
  box.appendChild(mk("div", "font-weight:600;font-size:12.5px;margin-bottom:4px", "Tìm thấy " + issues.length + " chỗ cần xem" + (issues.length > 200 ? " (hiện 200 chỗ đầu)" : "")));
  items.slice(0, 12).forEach(function (it) {
    var row = mk("div", "padding:3px 0;border-bottom:1px solid #f0f2f5;font-size:12px");
    var b = mk("button", "padding:0 6px;margin-right:6px", it.a); b.onclick = function () { gotoRef(it.a); };
    row.appendChild(b);
    if (it.type) row.appendChild(mk("b", "", "[" + it.type + "] "));
    row.appendChild(document.createTextNode(it.t + (it.fix ? " → " + it.fix : "")));
    box.appendChild(row);
  });
  if (items.length > 12) box.appendChild(mk("div", "font-size:11px;color:#8c959f;padding-top:3px", "… và " + (items.length - 12) + " chỗ nữa (xuất ra sheet để xem đủ)."));
  issueButtons(box, items, sheetName, opts);
  bot.extra.appendChild(box);
}
/* các cách xem trực quan mà KHÔNG sửa dữ liệu: tô nổi ô, hoặc xuất danh sách ra sheet mới có liên kết bấm tới từng ô */
function splitRef(it, sheetName) {
  var ad = it.a, sh = sheetName || "", k = ad.lastIndexOf("!");
  if (k >= 0) { sh = ad.substring(0, k).replace(/^'|'$/g, ""); ad = ad.substring(k + 1); }
  return { sh: sh, ad: ad };
}
function issueButtons(box, items, sheetName, opts) {
  opts = opts || {};
  var b1 = mk("button", "margin-top:6px", opts.fill === "#FFC7CE" ? "🖍 Bôi đỏ các ô này" : "🖍 Tô nổi các ô này");
  var b2 = mk("button", "margin-top:6px", "📄 Xuất danh sách ra sheet mới");
  function runOps(ops, label) {
    if (state.busy) return;
    busy(true);
    var bot = makeBot(); setStatus(bot, "⏳ " + label);
    var rn = new Runner(bot, function (log, undoB) {
      setStatus(bot, ""); pushUndo(undoB);
      var bad = log.filter(function (l) { return !l.ok; });
      bot.msg.textContent = bad.length ? "Chưa làm được: " + bad[0].t : "✔ Xong. Không ưng thì bấm ↶ Hoàn tác.";
      busy(false);
    });
    rn.base = null;
    ops.forEach(function (o) { rn.push(o); }); rn.close();
  }
  b1.onclick = function () {
    var by = {}, ops = [];
    items.forEach(function (it) {
      var r = splitRef(it, sheetName);
      if (!/^[A-Za-z]{1,3}\d{1,7}$/.test(r.ad)) return;
      (by[r.sh] = by[r.sh] || []).push(r.ad);
    });
    for (var sh2 in by) { var arr = by[sh2]; for (var k = 0; k < arr.length; k += 200) ops.push({ op: "mark_cells", sheet: sh2 || undefined, cells: arr.slice(k, k + 200), fill: opts.fill || "#FFF2CC" }); }
    if (!ops.length) { noteLine({ extra: box }, "Các mục này là vùng nhiều ô, hãy bấm từng địa chỉ để xem."); return; }
    runOps(ops, "Đang tô nổi…");
  };
  b2.onclick = function () {
    var nm = uniqueName("Kiểm lỗi");
    var rows = [["Vị trí", "Loại", "Vấn đề", "Gợi ý sửa", "Đi tới"]];
    items.slice(0, 220).forEach(function (it) {
      var r = splitRef(it, sheetName);
      var link = r.sh ? '=HYPERLINK("#\'' + r.sh.replace(/"/g, "") + '\'!' + r.ad.split(":")[0] + '","Đi tới ' + r.ad + '")' : "";
      rows.push([it.a, it.type, it.t, it.fix, link]);
    });
    var ops = [{ op: "add_sheet", name: nm }];
    for (var k = 0; k < rows.length; k += 55) ops.push({ op: "set_values", sheet: nm, range: "A" + (k + 1), values: rows.slice(k, k + 55) });
    ops.push({ op: "smart_format", sheet: nm });
    runOps(ops, "Đang xuất danh sách…");
  };
  box.appendChild(b1); box.appendChild(b2);
  var rs = items.filter(function (it) { return it.restore; });
  if (opts.restore && rs.length) {
    var b3 = mk("button", "margin-top:6px", "🔧 Khôi phục công thức cho " + rs.length + " ô số gõ đè");
    b3.onclick = function () {
      if (!editOK()) { noteLine({ extra: box }, "Đang tắt \"Cho phép sửa trực tiếp\". Chị bật lên (dưới ô gõ) rồi bấm lại.", "#b36b00"); return; }
      var by = {};
      rs.forEach(function (it) { var r = splitRef(it, sheetName); var k = r.sh + "|" + it.restore.c; (by[k] = by[k] || { sh: r.sh, r1c1: it.restore.r1c1, items: [] }).items.push({ cell: r.ad, r1c1: it.restore.r1c1 }); });
      var ops = Object.keys(by).map(function (k) { return { op: "restore_formula", sheet: by[k].sh || undefined, items: by[k].items }; });
      runOps(ops, "Đang khôi phục công thức…");
    };
    box.appendChild(b3);
  }
}

/* địa chỉ ô trong lời nhắn bấm được để nhảy tới ô đó */
var REF_RE = /((?:'[^']+'|[A-Za-z0-9_À-ỹ]+)!)?\$?[A-Z]{1,3}\$?\d{1,6}(?::\$?[A-Z]{1,3}\$?\d{1,6})?(?![A-Za-z0-9])/g;
function setMsg(el, text) {
  el.innerHTML = "";
  text = String(text || "");
  var last = 0, m;
  REF_RE.lastIndex = 0;
  while ((m = REF_RE.exec(text))) {
    if (m.index > 0 && /[A-Za-z0-9_]/.test(text.charAt(m.index - 1)) && !m[1]) continue;
    el.appendChild(document.createTextNode(text.substring(last, m.index)));
    var a = document.createElement("a"); a.textContent = m[0]; a.href = "#"; a.style.cssText = "color:#1a56c4;text-decoration:underline;cursor:pointer";
    (function (ref) {
      a.onclick = function (ev) {
        ev.preventDefault();
        var sh = null, addr = ref.replace(/\$/g, ""), k = addr.lastIndexOf("!");
        if (k >= 0) { sh = addr.substring(0, k).replace(/^'|'$/g, ""); addr = addr.substring(k + 1); }
        Excel.run(function (ctx) {
          var ws = sh ? ctx.workbook.worksheets.getItem(sh) : ctx.workbook.worksheets.getActiveWorksheet();
          if (sh) ws.activate();
          ws.getRange(addr).select(); return ctx.sync();
        }).then(null, function () {});
      };
    })(m[0]);
    el.appendChild(a);
    last = m.index + m[0].length;
  }
  el.appendChild(document.createTextNode(text.substring(last)));
}
/* AI xin đọc thêm vùng ngoài ngữ cảnh: đọc giúp rồi hỏi tiếp (tối đa 2 vòng) */
function readRanges(reads, cb) {
  var list = (reads || []).slice(0, 3);
  Excel.run(function (ctx) {
    var act = ctx.workbook.worksheets.getActiveWorksheet(); act.load("name");
    var jobs = [];
    return ctx.sync().then(function () {
      list.forEach(function (rd) {
        var nm = rd.sheet ? String(rd.sheet) : act.name, ws = ctx.workbook.worksheets.getItemOrNullObject(nm);
        var used = ws.getUsedRangeOrNullObject(); used.load("isNullObject,rowCount,columnCount,columnIndex");
        var hdr = used.getRow(0); hdr.load("values");
        var r = String(rd.range || "").replace(/\s/g, "");
        var sub = /\d/.test(r) && RANGE_RE.test(r) ? ws.getRange(r) : used;
        var inter = sub.getIntersectionOrNullObject(used);
        jobs.push({ nm: nm, ws: ws, used: used, hdr: hdr, inter: inter });
      });
      return ctx.sync();
    }).then(function () {
      jobs.forEach(function (j) {
        if (j.ws.isNullObject || j.used.isNullObject || j.inter.isNullObject) return;
        j.inter.load("address,rowCount,columnCount,columnIndex");
      });
      return ctx.sync();
    }).then(function () {
      jobs.forEach(function (j) {
        if (j.ws.isNullObject || j.used.isNullObject || j.inter.isNullObject) return;
        var rows = Math.max(1, Math.min(j.inter.rowCount, Math.floor(1500 / j.inter.columnCount)));
        j.take = j.inter.getCell(0, 0).getResizedRange(rows - 1, j.inter.columnCount - 1);
        j.take.load("formulas,values"); j.rows = rows;
      });
      return ctx.sync();
    }).then(function () {
      var out = "";
      jobs.forEach(function (j) {
        if (!j.take) { out += "\n[" + j.nm + "]: không có dữ liệu ở vùng này\n"; return; }
        var f = j.take.formulas, v = j.take.values, mask = [], c, r, line;
        for (c = 0; c < j.inter.columnCount; c++) mask.push(DANGER_HEADER.test(String(j.hdr.values[0][j.inter.columnIndex - j.used.columnIndex + c] || "")));
        out += "\n[" + j.nm + "!" + localAddr(j.inter.address) + (j.rows < j.inter.rowCount ? " (chỉ " + j.rows + " dòng đầu)" : "") + "]\n";
        for (r = 0; r < f.length; r++) {
          line = [];
          for (c = 0; c < f[r].length; c++) {
            var x = f[r][c];
            line.push(mask[c] && r > 0 ? "<ẨN>" : (typeof x === "string" && x.charAt(0) === "=" ? x + " → " + v[r][c] : String(x)));
          }
          out += line.join(" | ") + "\n";
        }
      });
      return out;
    });
  }).then(function (t) { cb(null, t); }, function (e) { cb(String(e && e.message || e)); });
}

/* để Excel tự tính giúp: ghi công thức vào sheet tạm, đọc kết quả, xóa sheet tạm => số liệu trả lời chính xác 100% */
function evalCalcs(calcs, cb) {
  var list = (calcs || []).slice(0, 12), nm = "_tinh_tam";
  Excel.run(function (ctx) {
    var allS = ctx.workbook.worksheets; allS.load("items/name");
    return ctx.sync().then(function () {
      var q; for (q = 0; q < allS.items.length; q++) if (allS.items[q].name === nm) allS.items[q].delete();
      var ws = ctx.workbook.worksheets.add(nm), cells = [];
      list.forEach(function (c, i) {
        var f = String(c.formula || "");
        if (f.charAt(0) !== "=" || FORBID_F.test(f)) { cells.push(null); return; }
        var cell = ws.getRange("A" + (i + 1)); cell.formulas = [[f]]; cells.push(cell);
      });
      return ctx.sync().then(function () {
        cells.forEach(function (c) { if (c) c.load("values"); });
        return ctx.sync().then(function () {
          var out = "";
          list.forEach(function (c, i) { out += "- " + (c.label || ("Phép tính " + (i + 1))) + " [" + c.formula + "] = " + (cells[i] ? JSON.stringify(cells[i].values[0][0]) : "(công thức không hợp lệ)") + "\n"; });
          ws.delete();
          return ctx.sync().then(function () { return out; });
        });
      });
    });
  }).then(function (t) { cb(null, t); }, function (e) { cb(String(e && e.message || e)); });
}

function runChat(q, opts) {
  opts = opts || {};
  busy(true);
  var bot = opts.bot || makeBot();
  var msgEl = bot.msg;
  if (opts.verify) { msgEl = document.createElement("div"); msgEl.className = "small"; msgEl.style.marginTop = "6px"; bot.extra.appendChild(msgEl); }
  setStatus(bot, "⏳ Đang đọc bảng tính…");
  getContext(function (err, c) {
    if (err) { setStatus(bot, ""); msgEl.textContent = "Không đọc được bảng tính: " + err; busy(false); return; }
    setStatus(bot, opts.verify ? "⏳ Đang soát lại kết quả…" : opts.round ? "⏳ Đã đọc thêm, đang làm tiếp…" : "⏳ Đang suy nghĩ…");
    var auto = $("auto").checked || !!opts.retry || !!opts.verify;
    var parser = new ActionParser(), got = [], msgShown = false, runner = null, finalText = "", ended = false, attempt = 0, stream = null;
    state.cur = { abort: function () {
      opts.stopped = true; ended = true;
      if (stream) stream.abort();
      noteLine(bot, "Đã dừng.", "#b36b00");
      if (runner) { runner.q.length = 0; runner.close(); } else { setStatus(bot, ""); busy(false); }
    } };
    function ensure() {
      if (!runner) runner = new Runner(bot, function (log, undoB) { finishRun(bot, log, undoB, opts, got, runner && runner.base); });
      return runner;
    }
    function showMsg(text) {
      if (msgShown) return;
      var mm = /"message"\s*:\s*"((?:[^"\\]|\\.)*)"/.exec(text);
      if (!mm) return;
      msgShown = true;
      var mt; try { mt = JSON.parse('"' + mm[1] + '"'); } catch (e) { mt = mm[1]; }
      setMsg(msgEl, mt);
      setStatus(bot, auto ? "⏳ Đang làm…" : "");
    }
    function onText(text) {
      finalText = text; showMsg(text);
      var objs = parser.feed(text);
      for (var i = 0; i < objs.length; i++) { got.push(objs[i]); if (auto) ensure().push(objs[i]); }
    }
    function onError(e) {
      if (ended) return;
      if (!finalText && attempt < 1 && /kết nối/.test(e)) {
        attempt++; noteLine(bot, "Mạng chậm, đang thử lại…", "#b36b00"); setStatus(bot, "⏳ Đang kết nối lại…"); start(); return;
      }
      ended = true;
      setStatus(bot, ""); noteLine(bot, e, "#c00"); busy(false); refreshMe();
    }
    function onDone(d) {
      if (ended) return; ended = true;
      var j = parseReply(finalText);
      if (!msgShown) setMsg(msgEl, j.message || "");
      if (j.calc && j.calc.length && !j.actions.length && (opts.round || 0) < 2) {
        setStatus(bot, "⏳ Excel đang tính số liệu…");
        evalCalcs(j.calc, function (err3, t3) {
          if (err3 || !t3) { noteLine(bot, "Excel không tính được: " + (err3 || "trống"), "#c00"); busy(false); return; }
          setStatus(bot, ""); if (opts.stopped) { busy(false); return; }
          runChat(q + "\n\n[KẾT QUẢ DO EXCEL TÍNH THEO YÊU CẦU CỦA BẠN — dùng đúng các số này để trả lời, đừng tự tính lại]\n" + t3, { deep: opts.deep, retry: opts.retry, verify: opts.verify, bot: bot, round: (opts.round || 0) + 1 });
        });
        return;
      }
      if (j.reads && j.reads.length && !j.actions.length && (opts.round || 0) < 2) {
        setStatus(bot, "⏳ Đang đọc thêm vùng cần thiết…");
        readRanges(j.reads, function (err2, txt) {
          if (err2 || !txt) { noteLine(bot, "Không đọc thêm được: " + (err2 || "trống"), "#c00"); busy(false); return; }
          setStatus(bot, ""); if (opts.stopped) { busy(false); return; }
          runChat(q + "\n\n[DỮ LIỆU ĐỌC THÊM THEO YÊU CẦU CỦA BẠN — hãy làm luôn, đừng xin đọc nữa]" + txt, { deep: opts.deep, retry: opts.retry, verify: opts.verify, bot: bot, round: (opts.round || 0) + 1 });
        });
        return;
      }
      var rest = j.actions.slice(got.length);
      for (var i = 0; i < rest.length; i++) { got.push(rest[i]); if (auto) ensure().push(rest[i]); }
      if (j.kpis && j.kpis.length) renderKpis(bot, j.kpis);
      if (j.formulas && j.formulas.length) renderFormulas(bot, j.formulas);
      if (j.issues && j.issues.length) renderIssues(bot, j.issues, "");
      var m = d.meta;
      bot.meta.textContent = m.label + " · " + fmt(m.tokens_in) + " token vào / " + fmt(m.tokens_out) + " ra · ~" + fmt(m.cost_vnd) + "đ · vì: " + m.reason;
      updateQuota(d.quota, m);
      if (d.version) { SERVER_VERSION = d.version; showVersion(); }
      if (j.warnings) noteLine(bot, "Lưu ý: " + j.warnings);
      if (d.cut) noteLine(bot, "Câu trả lời dài quá nên bị cắt; phần đã làm xong vẫn giữ. Chị gõ tiếp phần còn lại nhé.", "#b36b00");
      if (!opts.retry && !opts.round && !opts.verify) {
        var summary = (j.message || "") + (got.length ? " | " + got.slice(0, 6).map(describe).join("; ") : "");
        state.hist.push({ q: state.lastQ, a: summary.substring(0, 500) });
        if (state.hist.length > 4) state.hist.shift();
      }
      if (!opts.verify) {
        var again = document.createElement("button"); again.textContent = "Làm kỹ hơn";
        again.onclick = function () { send(true); };
        bot.extra.appendChild(again);
      }
      if (j.suggest && j.suggest.length && !opts.retry && !opts.verify) {
        var sg = document.createElement("div"); sg.className = "chips";
        j.suggest.slice(0, 3).forEach(function (t) {
          if (typeof t !== "string" || !t) return;
          var b2 = document.createElement("button"); b2.textContent = t.substring(0, 60);
          b2.onclick = function () { if (state.busy) return; $("q").value = t; send(false); };
          sg.appendChild(b2);
        });
        bot.extra.appendChild(sg);
      }
      if (auto) {
        if (runner || got.length) ensure().close(); else { setStatus(bot, ""); busy(false); }
      } else {
        setStatus(bot, ""); busy(false);
        if (got.length) {
          var ab = document.createElement("button"); ab.className = "pri"; ab.textContent = "Áp dụng " + got.length + " thay đổi";
          var shown = document.createElement("div"); shown.className = "small";
          for (var k = 0; k < got.length; k++) { var li = document.createElement("div"); li.textContent = "• " + describe(got[k]); shown.appendChild(li); }
          bot.log.appendChild(shown);
          ab.onclick = function () {
            ab.style.display = "none"; shown.style.display = "none"; busy(true);
            var rn = new Runner(bot, function (log, undoB) { finishRun(bot, log, undoB, { retry: false }, got, rn.base); });
            for (var z = 0; z < got.length; z++) rn.push(got[z]);
            rn.close();
          };
          bot.extra.insertBefore(ab, bot.extra.firstChild);
        }
      }
    }
    var body = { question: q, context: c.text + (editOK() ? "" : NO_OVERWRITE_NOTE), compat: compat(), deep: !!opts.deep, history: state.hist.slice(-3), stream: canStream() };
    function start() {
    if (canStream()) stream = streamChat(body, { onText: onText, onDone: onDone, onError: onError });
    else api("/api/chat", "POST", body, function (e, d) {
      if (e) { onError(e); return; }
      onText(d.reply); onDone({ meta: d.meta, quota: d.quota, cut: d.cut, version: d.version });
    });
    }
    start();
  });
}

/* ---------- Việc làm ngay trên máy (không gọi AI, không tốn hạn mức) ---------- */
function supports(v) { try { return Office.context.requirements.isSetSupported("ExcelApi", v); } catch (e) { return false; } }
function targetRange(ctx, sh, rangeStr) {
  if (rangeStr && /\d/.test(rangeStr) && RANGE_RE.test(String(rangeStr).replace(/\s/g, ""))) {
    var xr = sh.getRange(String(rangeStr).replace(/\s/g, "")); xr.load("rowCount,columnCount,rowIndex,columnIndex,address");
    return ctx.sync().then(function () { return xr; });
  }
  var sel = ctx.workbook.getSelectedRange(); sel.load("rowCount,columnCount");
  var used = sh.getUsedRangeOrNullObject(); used.load("rowCount,columnCount");
  return ctx.sync().then(function () {
    if (used.isNullObject) throw new Error("Sheet đang trống, chưa có gì để làm.");
    if (sel.rowCount * sel.columnCount > 1) {
      var inter = sel.getIntersectionOrNullObject(used); inter.load("rowCount");
      return ctx.sync().then(function () {
        var r = inter.isNullObject ? used : inter;
        r.load("rowCount,columnCount,rowIndex,columnIndex,address");
        return ctx.sync().then(function () { return r; });
      });
    }
    used.load("rowCount,columnCount,rowIndex,columnIndex,address");
    return ctx.sync().then(function () { return used; });
  });
}
var MONEY_HDR = /tiền|lương|giá|thành|doanh thu|doanh số|số dư|phụ cấp|thưởng|chi phí|nợ|thu|chi|vnd|vnđ|amount|total|thuế|bhxh|trợ cấp|hoa hồng|phí/i;
var CODE_HDR = /(^|\s)(id|mã|stt|số thứ tự|số hiệu|cccd|cmnd|stk|sđt|điện thoại|mst|phone)(\s|$)/i;
var DATE_HDR = /ngày|date|hạn|sinh|vào làm/i;
function planFormat(v, nfs, truncated) {
  var rows = v.length, cols = v[0].length, c, r;
  function filled(row) { var n = 0; for (var k = 0; k < cols; k++) if (v[row][k] !== "" && v[row][k] !== null) n++; return n; }
  var h = 0;
  if (rows > 2 && cols > 2 && filled(0) <= 1) {
    for (r = 1; r < Math.min(rows, 6); r++) if (filled(r) >= Math.max(2, cols / 2)) { h = r; break; }
  }
  var first = h + 1, last = rows - 1, tot = -1;
  if (!truncated && last > first) {
    for (c = 0; c < Math.min(cols, 3); c++) if (/^\s*(tổng|cộng|total|sum)/i.test(String(v[last][c]))) tot = last;
    if (tot >= 0) last = last - 1;
  }
  var plan = { h: h, first: first, last: last, tot: tot, cols: [] };
  for (c = 0; c < cols; c++) {
    var head = String(v[h][c]);
    var num = 0, intN = 0, frac = 0, dateN = 0, txt = 0, maxLen = 0, sumLen = 0, maxAbs = 0, in01 = 0, money = false, nonEmpty = 0;
    var limit = Math.min(last, first + 2999);
    for (r = first; r <= limit; r++) {
      var x = v[r][c];
      if (x === "" || x === null) continue;
      nonEmpty++;
      if (typeof x === "number") {
        num++; var ax = Math.abs(x); if (ax > maxAbs) maxAbs = ax;
        if (Math.floor(x) === x) intN++; else frac++;
        if (ax <= 1.5) in01++;
        if (isDateFmt(nfs[r][c])) dateN++;
        if (/#,##0|₫|\$/.test(String(nfs[r][c]))) money = true;
      } else { txt++; var L = String(x).length; sumLen += L; if (L > maxLen) maxLen = L; }
    }
    var col = { head: head, type: "text", align: "Left", nf: null, wrap: false };
    if (nonEmpty === 0) col.type = "empty";
    else if (num > 0 && num >= nonEmpty * 0.7) {
      if (dateN >= num * 0.6 || (DATE_HDR.test(head) && intN === num && maxAbs >= 20000 && maxAbs <= 80000)) { col.type = "date"; col.align = "Center"; col.nf = "dd/mm/yyyy"; }
      else if (CODE_HDR.test(head)) { col.type = "code"; col.align = "Center"; }
      else if (/%|tỷ lệ|tỉ lệ|phần trăm|percent/i.test(head) && in01 === num) { col.type = "percent"; col.align = "Right"; col.nf = "0.0%"; }
      else if (MONEY_HDR.test(head) || maxAbs >= 10000 || money) { col.type = "money"; col.align = "Right"; col.nf = (frac > 0 && maxAbs < 1000) ? "#,##0.00" : "#,##0"; }
      else if (frac > 0) { col.type = "decimal"; col.align = "Right"; col.nf = "#,##0.00"; }
      else { col.type = "int"; col.align = "Right"; col.nf = maxAbs >= 1000 ? "#,##0" : "0"; }
    } else if (txt > 0) {
      var avg = sumLen / txt;
      col.type = (maxLen <= 14 && CODE_HDR.test(head)) ? "code" : "text";
      col.align = col.type === "code" ? "Center" : "Left";
      col.wrap = avg > 40;
    }
    plan.cols.push(col);
  }
  return plan;
}
var PROP_SPEC = { format: { fill: { color: true }, font: { bold: true, italic: true, color: true, size: true, name: true }, horizontalAlignment: true, verticalAlignment: true, wrapText: true,
  borders: { top: { style: true, color: true, weight: true }, bottom: { style: true, color: true, weight: true }, left: { style: true, color: true, weight: true }, right: { style: true, color: true, weight: true } } } };
function stripEmpty(o) {
  if (o === null || o === undefined || o === "") return undefined;
  if (typeof o !== "object") return o;
  var r = {}, any = false;
  for (var k in o) { var v = stripEmpty(o[k]); if (v !== undefined) { r[k] = v; any = true; } }
  return any ? r : undefined;
}
function restoreSmart(ctx, s) {
  var sh = ctx.workbook.worksheets.getItem(s.sheet), rng = sh.getRange(s.addr), i, r, c;
  rng.numberFormat = s.nf;
  rng.format.fill.clear();
  if (s.props) {
    var data = [];
    for (r = 0; r < s.props.length; r++) {
      var row = [];
      for (c = 0; c < s.props[r].length; c++) {
        var p = s.props[r][c] || {}, f = (p.format) || {}, o = { format: {} };
        if (f.font) o.format.font = stripEmpty(f.font);
        if (f.horizontalAlignment) o.format.horizontalAlignment = f.horizontalAlignment;
        if (f.verticalAlignment) o.format.verticalAlignment = f.verticalAlignment;
        if (f.wrapText !== undefined) o.format.wrapText = f.wrapText;
        if (f.borders) o.format.borders = stripEmpty(f.borders);
        if (f.fill && f.fill.color && f.fill.color !== "#FFFFFF") o.format.fill = { color: f.fill.color };
        row.push(o);
      }
      data.push(row);
    }
    rng.setCellProperties(data);
  } else {
    var ids = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal"];
    for (i = 0; i < ids.length; i++) { try { rng.format.borders.getItem(ids[i]).style = "None"; } catch (e) {} }
    rng.format.font.bold = false; rng.format.font.color = "#000000"; rng.format.horizontalAlignment = "General"; rng.format.wrapText = false;
  }
  for (i = 0; i < s.widths.length; i++) rng.getColumn(i).format.columnWidth = s.widths[i];
  if (typeof s.rowHeight === "number") rng.format.rowHeight = s.rowHeight; else rng.format.autofitRows();
}

/* Format cho đẹp: chạy trên máy, vài giây, không tốn hạn mức */
function smartFormat(done, opts) {
  opts = opts || {};
  var info = {}, undoE = [], notes = [];
  function fail(e) { done(String(e && e.message || e), null, null); }
  Excel.run(function (ctx) {
    var sh = opts.sheet ? ctx.workbook.worksheets.getItem(opts.sheet) : ctx.workbook.worksheets.getActiveWorksheet(); sh.load("name");
    return ctx.sync().then(function () { return targetRange(ctx, sh, opts.range); }).then(function (rg) {
      if (rg.columnCount > 80) throw new Error("Bảng quá rộng (trên 80 cột), chị chọn vùng nhỏ hơn rồi thử lại.");
      var statN = Math.min(rg.rowCount, 3000);
      var stat = rg.getCell(0, 0).getResizedRange(statN - 1, rg.columnCount - 1); stat.load("values,numberFormat");
      rg.load("numberFormat");
      var cs = [], i;
      for (i = 0; i < rg.columnCount; i++) { var cr = rg.getColumn(i); cr.format.load("columnWidth"); cs.push(cr); }
      rg.format.load("rowHeight");
      return ctx.sync().then(function () {
        info.sheet = sh.name; info.addr = localAddr(rg.address); info.rows = rg.rowCount; info.cols = rg.columnCount;
        info.r0 = rg.rowIndex; info.nf = rg.numberFormat; info.widths = [];
        for (i = 0; i < cs.length; i++) info.widths.push(cs[i].format.columnWidth);
        info.rowHeight = rg.format.rowHeight; info.statN = statN;
        info.plan = planFormat(stat.values, stat.numberFormat, rg.rowCount > statN); info.vals = stat.values;
      });
    });
  }).then(step2, fail);

  /* chụp lại màu/viền cũ để còn hoàn tác (nếu Excel hỗ trợ) */
  function step2() {
    if (!(info.rows * info.cols <= 20000 && supports("1.9"))) { info.props = null; step3(); return; }
    Excel.run(function (ctx) {
      var pr = ctx.workbook.worksheets.getItem(info.sheet).getRange(info.addr).getCellProperties(PROP_SPEC);
      return ctx.sync().then(function () { return pr.value; });
    }).then(function (v) { info.props = v; step3(); }, function () { info.props = null; step3(); });
  }

  function step3() {
    Excel.run(function (ctx) {
      var sh = ctx.workbook.worksheets.getItem(info.sheet), rg = sh.getRange(info.addr), P = info.plan, cols = info.cols, rows = info.rows, c, i;
      var h = P.h, bodyRows = rows - h;
      var tbl = rg.getCell(h, 0).getResizedRange(bodyRows - 1, cols - 1);
      var lastData = (rows > info.statN) ? rows - 1 : (P.tot >= 0 ? P.tot : P.last);
      var dn = lastData - P.first + 1;
      if (h > 0) { var title = rg.getCell(0, 0); title.format.font.bold = true; title.format.font.size = 14; title.format.font.color = "#1F4E78"; }
      tbl.format.verticalAlignment = "Center";
      var edges = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"];
      for (i = 0; i < edges.length; i++) { var b = tbl.format.borders.getItem(edges[i]); b.style = "Continuous"; b.color = "#B7C0CB"; }
      if (bodyRows > 1) { var ih = tbl.format.borders.getItem("InsideHorizontal"); ih.style = "Continuous"; ih.color = "#D9DEE5"; }
      var hdr = tbl.getRow(0);
      hdr.format.fill.color = "#1F4E78"; hdr.format.font.color = "#FFFFFF"; hdr.format.font.bold = true;
      hdr.format.horizontalAlignment = "Center"; hdr.format.wrapText = false;
      if (dn > 0) {
        for (c = 0; c < cols; c++) {
          var col = P.cols[c], cr = tbl.getCell(1, c).getResizedRange(dn - 1, 0);
          cr.format.horizontalAlignment = col.align;
          if (col.nf) cr.numberFormat = col.nf;
          if (col.wrap) cr.format.wrapText = true;
        }
      }
      if (P.tot >= 0 && rows <= info.statN) {
        var tr = tbl.getRow(P.tot - h);
        tr.format.font.bold = true; tr.format.fill.color = "#EAF1F8";
        var tb = tr.format.borders.getItem("EdgeTop"); tb.style = "Continuous"; tb.color = "#1F4E78";
      }
      return ctx.sync().then(function () {
        /* autofit đôi khi báo "current selection is invalid" (đang chọn biểu đồ/đang sửa ô): khi đó tự ước lượng độ rộng */
        tbl.format.autofitColumns();
        return ctx.sync().then(function () { return true; }, function () { return false; });
      }).then(function (autoOK) {
        var cs = [];
        if (autoOK) for (c = 0; c < cols; c++) { var cc = rg.getColumn(c); cc.format.load("columnWidth"); cs.push(cc); }
        return ctx.sync().then(function () {
          var wrapAny = false;
          for (c = 0; c < cols; c++) {
            var w, col2 = P.cols[c];
            if (autoOK) w = cs[c].format.columnWidth;
            else {
              var ml = 4, rr;
              for (rr = 0; rr < (info.vals || []).length; rr++) { var tx = info.vals[rr][c]; var L = tx === null || tx === undefined ? 0 : String(tx).length; if (L > ml) ml = L; }
              w = Math.min(ml, 40) * 6.4 + 8;
            }
            var maxW = col2.type === "text" ? 260 : 170;
            var nw = Math.max(52, Math.min(maxW, w + 14));
            if (w + 14 > maxW && col2.type === "text" && dn > 0) { tbl.getCell(1, c).getResizedRange(dn - 1, 0).format.wrapText = true; wrapAny = true; }
            rg.getColumn(c).format.columnWidth = nw;
            if (P.cols[c].wrap) wrapAny = true;
          }
          hdr.format.rowHeight = 30; hdr.format.wrapText = true;
          if (wrapAny && dn > 0) tbl.getCell(1, 0).getResizedRange(dn - 1, cols - 1).format.autofitRows();
          return ctx.sync().catch(function () { return null; });
        });
      });
    }).then(function () {
      undoE.push({ t: "smartfmt", sheet: info.sheet, addr: info.addr, nf: info.nf, widths: info.widths, rowHeight: info.rowHeight, props: info.props });
      if (!info.props) notes.push("màu/viền cũ không khôi phục được khi hoàn tác (chỉ lùi số liệu, độ rộng cột)");
      step4();
    }, fail);
  }

  /* cố định dòng tiêu đề nếu bảng dài và chưa cố định */
  function step4() {
    var n = info.r0 + info.plan.h + 1;
    if (info.rows < 12 || n > 10 || !supports("1.7")) { step5(); return; }
    Excel.run(function (ctx) {
      var sh = ctx.workbook.worksheets.getItem(info.sheet), loc = sh.freezePanes.getLocationOrNullObject(); loc.load("isNullObject");
      return ctx.sync().then(function () {
        if (!loc.isNullObject) return false;
        sh.freezePanes.freezeRows(n);
        return ctx.sync().then(function () { return true; });
      });
    }).then(function (did) { if (did) undoE.push({ t: "freeze", sheet: info.sheet }); step5(); }, function () { step5(); });
  }

  /* bật bộ lọc ở dòng tiêu đề nếu chưa có */
  function step5() {
    if (info.rows < 4 || !supports("1.9")) { finish(); return; }
    Excel.run(function (ctx) {
      var sh = ctx.workbook.worksheets.getItem(info.sheet); sh.autoFilter.load("enabled");
      return ctx.sync().then(function () {
        if (sh.autoFilter.enabled) return false;
        var P = info.plan, last = (info.rows > info.statN) ? info.rows - 1 : (P.tot >= 0 ? P.tot - 1 : P.last);
        if (last <= P.h) return false;
        var rg = sh.getRange(info.addr).getCell(P.h, 0).getResizedRange(last - P.h, info.cols - 1);
        sh.autoFilter.apply(rg);
        return ctx.sync().then(function () { return true; });
      });
    }).then(function (did) { if (did) undoE.push({ t: "filter", sheet: info.sheet }); finish(); }, function () { finish(); });
  }

  function finish() {
    var P = info.plan, types = { money: 0, date: 0, text: 0, int: 0, percent: 0, code: 0, decimal: 0 }, k;
    for (k = 0; k < P.cols.length; k++) if (types[P.cols[k].type] !== undefined) types[P.cols[k].type]++;
    var parts = [];
    if (P.h > 0) parts.push("tiêu đề trang");
    parts.push("dòng tiêu đề xanh");
    if (types.money) parts.push(types.money + " cột tiền có dấu phân cách");
    if (types.date) parts.push(types.date + " cột ngày dd/mm/yyyy");
    if (types.percent) parts.push(types.percent + " cột %");
    if (types.code) parts.push(types.code + " cột mã/số điện thoại căn giữa");
    if (P.tot >= 0) parts.push("dòng tổng in đậm");
    parts.push("kẻ viền, chỉnh độ rộng cột");
    var sum = "Đã format " + info.addr + " (" + info.rows + " dòng × " + info.cols + " cột): " + parts.join(", ") + ".";
    if (notes.length) sum += " Lưu ý: " + notes.join("; ") + ".";
    done(null, sum, undoE);
  }
}

/* ---------- Dọn dữ liệu trên máy: cắt khoảng trắng, chữ→số, chữ→ngày ---------- */
function dmyToSerial(s) {
  var m = /^\s*(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})\s*$/.exec(String(s));
  if (!m) return null;
  var d = Number(m[1]), mo = Number(m[2]), y = Number(m[3]);
  if (mo < 1 || mo > 12 || d < 1 || d > 31 || y < 1900 || y > 2200) return null;
  var t = Date.UTC(y, mo - 1, d);
  var chk = new Date(t);
  if (chk.getUTCDate() !== d) return null;
  return Math.round((t - Date.UTC(1899, 11, 30)) / 86400000);
}
function cleanPlan(f, nf, hdrs) {
  var rows = f.length, cols = f[0].length, out = [], c, r, changed = { trim: 0, num: 0, date: 0 }, nfset = [];
  for (r = 0; r < rows; r++) { out.push(f[r].slice(0)); nfset.push(new Array(cols)); }
  for (c = 0; c < cols; c++) {
    var nNum = 0, nDate = 0, nTxt = 0;
    for (r = 0; r < rows; r++) {
      var x = f[r][c];
      if (typeof x === "number") { if (isDateFmt(nf[r][c])) nDate++; else nNum++; }
      else if (typeof x === "string" && x.charAt(0) !== "=" && x !== "") { if (dmyToSerial(x) !== null) nDate++; nTxt++; }
    }
    var head = String(hdrs[c] === undefined ? "" : hdrs[c]);
    var numCol = nNum > 0 && !CODE_HDR.test(head);
    var dateCol = nDate >= 2 && (nDate >= (nNum + nTxt) * 0.4 || DATE_HDR.test(head));
    for (r = 0; r < rows; r++) {
      var v = f[r][c];
      if (typeof v !== "string" || v.charAt(0) === "=" || v === "") continue;
      var t = v.replace(/[\s ]+/g, " ").replace(/^\s+|\s+$/g, "");
      if (dateCol) { var sd = dmyToSerial(t); if (sd !== null) { out[r][c] = sd; nfset[r][c] = "dd/mm/yyyy"; changed.date++; continue; } }
      if (numCol && !/^0\d/.test(t)) { var pn = numFromText(t); if (pn !== null) { out[r][c] = pn; if (nf[r][c] === "@") nfset[r][c] = "General"; changed.num++; continue; } }
      if (t !== v) { out[r][c] = t; changed.trim++; }
    }
  }
  return { out: out, nfset: nfset, changed: changed };
}
function localClean(done, opts) {
  opts = opts || {};
  var undoE = [];
  Excel.run(function (ctx) {
    var sh = opts.sheet ? ctx.workbook.worksheets.getItem(opts.sheet) : ctx.workbook.worksheets.getActiveWorksheet(); sh.load("name");
    return ctx.sync().then(function () { return targetRange(ctx, sh, opts.range); }).then(function (rg) {
      if (rg.rowCount * rg.columnCount > 60000) throw new Error("Vùng quá lớn, chị bôi đen phần nhỏ hơn rồi thử lại.");
      rg.load("formulas,numberFormat,values");
      return ctx.sync().then(function () {
        var f = rg.formulas, hdrs = f[0], body = f.slice(1), nfb = rg.numberFormat.slice(1);
        var errs = 0, i, j, vv = rg.values;
        for (i = 0; i < vv.length; i++) for (j = 0; j < vv[i].length; j++) if (typeof vv[i][j] === "string" && ERR_RE.test(vv[i][j])) errs++;
        if (!body.length) return { n: 0, errs: errs, ch: { trim: 0, num: 0, date: 0 } };
        var plan = cleanPlan(body, nfb, hdrs), ch = plan.changed, n = ch.trim + ch.num + ch.date;
        if (!n) return { n: 0, errs: errs, ch: ch };
        var body0 = rg.getCell(1, 0).getResizedRange(body.length - 1, rg.columnCount - 1);
        return snapRange(ctx, body0, false).then(function (s) {
          undoE.push({ t: "cells", snap: s });
          for (i = 0; i < plan.nfset.length; i++) for (j = 0; j < plan.nfset[i].length; j++) if (plan.nfset[i][j]) body0.getCell(i, j).numberFormat = [[plan.nfset[i][j]]];
          body0.formulas = plan.out;
          return ctx.sync().then(function () { return { n: n, errs: errs, ch: ch }; });
        });
      });
    });
  }).then(function (r) { done(null, r, undoE); }, function (e) { done(String(e && e.message || e), null, null); });
}

/* ---------- Điều hướng: việc đơn giản làm trên máy, việc cần suy luận gửi AI ---------- */
function detectLocal(q) {
  if (q.length > 80) return "";
  if (/(tổng|màu|tô |điều kiện|cột [a-z]\b|chỉ |riêng|biểu đồ|dashboard|sắp xếp|ngày)/i.test(q)) return "";
  if (/(format|định dạng|làm đẹp|cho đẹp|trình bày)/i.test(q)) return "format";
  if (/(quét|dọn|làm sạch|sửa)/i.test(q) && /(lỗi|dữ liệu|sạch)/i.test(q) && !/[A-Z]\d|cột/i.test(q)) return "fix";
  return "";
}
function localBot(q) {
  var bot = makeBot();
  setStatus(bot, "⏳ Đang làm ngay trên máy…");
  return bot;
}
function pushUndo(undoE) {
  if (!undoE || !undoE.length) return;
  state.undo.push(undoE);
  if (state.undo.length > 10) state.undo.shift();
  $("undoBtn").disabled = false;
}
function runLocalFormat() {
  busy(true);
  var bot = localBot();
  smartFormat(function (err, sum, undoE) {
    setStatus(bot, "");
    if (err) { bot.msg.textContent = "Chưa format được: " + err; busy(false); return; }
    bot.msg.textContent = "✔ " + sum;
    pushUndo(undoE);
    noteLine(bot, "Làm trên máy nên không tốn hạn mức. Không ưng thì bấm ↶ Hoàn tác. Muốn khác đi (vd: đổi màu, thêm dòng tổng) thì gõ yêu cầu.");
    state.hist.push({ q: state.lastQ, a: "Đã format bảng bằng chức năng có sẵn." }); if (state.hist.length > 4) state.hist.shift();
    busy(false);
  });
}
function runLocalFix(q) {
  busy(true);
  var bot = localBot();
  localClean(function (err, r, undoE) {
    setStatus(bot, "");
    if (err) { bot.msg.textContent = "Chưa dọn được: " + err; busy(false); return; }
    pushUndo(undoE);
    var lines = [];
    if (r.ch.trim) lines.push("cắt khoảng trắng thừa ở " + r.ch.trim + " ô");
    if (r.ch.num) lines.push("đổi " + r.ch.num + " ô số lưu dạng chữ thành số");
    if (r.ch.date) lines.push("đổi " + r.ch.date + " ô ngày dạng chữ thành ngày thật");
    bot.msg.textContent = lines.length ? "✔ Đã " + lines.join(", ") + "." : "Không thấy dữ liệu nào cần dọn (khoảng trắng, số/ngày dạng chữ).";
    if (lines.length) noteLine(bot, "Làm trên máy nên không tốn hạn mức. Không ưng thì bấm ↶ Hoàn tác.");
    if (r.errs > 0) {
      noteLine(bot, "Còn " + r.errs + " ô công thức đang báo lỗi, nhờ trợ lý xử lý tiếp…", "#b36b00");
      busy(false);
      state.lastQ = "Sửa các ô công thức đang báo lỗi (#REF!, #VALUE!, #N/A, #DIV/0!...) trong sheet này";
      runChat(state.lastQ, { retry: false });
      return;
    }
    busy(false);
  });
}
function send(deep) {
  var q = $("q").value.replace(/^\s+|\s+$/g, "");
  if (deep) q = state.lastQ;
  if (!q) return;
  state.lastQ = q;
  if (!deep) { addMsg("me", q); $("q").value = ""; $("q").style.height = "auto"; }
  if (state.busy) return;
  runChat(q, { deep: !!deep });
}

/* ====================================================================
   LOGIC THUẦN (không gọi Excel): gộp file, soát công thức, bất thường,
   chênh lệch, truy vấn. Chạy trên máy, dữ liệu tiền/người không gửi đi.
   ==================================================================== */
function normName(s) {
  s = String(s === undefined || s === null ? "" : s).toLowerCase();
  try { s = s.normalize("NFD").replace(/[̀-ͯ]/g, ""); } catch (e) {}
  s = s.replace(/đ/g, "d");
  return s.replace(/[^a-z0-9]+/g, " ").replace(/^\s+|\s+$/g, "");
}
function colNum(L) { var n = 0; for (var i = 0; i < L.length; i++) n = n * 26 + (L.charCodeAt(i) - 64); return n; }
function toNum(x) {
  if (typeof x === "number") return x;
  if (x === "" || x === null || x === undefined) return null;
  return numFromText(x);
}
function isBlank(x) { return x === "" || x === null || x === undefined; }
function periodInfo(v) {
  function mk(y, m) { y = Number(y); m = Number(m); if (m < 1 || m > 12) return null; return { k: y + "-" + (m < 10 ? "0" : "") + m, n: y * 100 + m }; }
  if (typeof v === "number" && v >= 20000 && v <= 80000) {
    var d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
    return mk(d.getUTCFullYear(), d.getUTCMonth() + 1);
  }
  var s = String(v === undefined || v === null ? "" : v).replace(/^\s+|\s+$/g, ""), m1, r;
  if ((m1 = /^(?:t(?:háng)?\s*)?(\d{1,2})\s*[\/\-.]\s*(\d{4})$/i.exec(s)) && (r = mk(m1[2], m1[1]))) return r;
  if ((m1 = /^(\d{4})\s*[\/\-.]\s*(\d{1,2})$/.exec(s)) && (r = mk(m1[1], m1[2]))) return r;
  if ((m1 = /^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/.exec(s)) && (r = mk(m1[3], m1[2]))) return r;
  return { k: s, n: null };
}
function median(a) {
  if (!a.length) return 0;
  var b = a.slice(0).sort(function (x, y) { return x - y; }), m = Math.floor(b.length / 2);
  return b.length % 2 ? b[m] : (b[m - 1] + b[m]) / 2;
}
function money(n) { return fmt(Math.round(n)); }

/* ---------- 1. Làm sạch bảng từ file lạ ---------- */
function detectHeaderRow(v) {
  var lim = Math.min(15, v.length), counts = [], strs = [], mx = 0, r, c;
  for (r = 0; r < lim; r++) {
    var n = 0, s = 0;
    for (c = 0; c < v[r].length; c++) { var x = v[r][c]; if (!isBlank(x)) { n++; if (typeof x === "string") s++; } }
    counts.push(n); strs.push(s); if (n > mx) mx = n;
  }
  for (r = 0; r < lim - 1; r++) {
    if (counts[r] >= Math.max(2, Math.ceil(mx * 0.6)) && strs[r] >= counts[r] * 0.7 && counts[r + 1] >= 1) return r;
  }
  return 0;
}
function applyMerges(grid, merges) {
  var fixed = 0;
  (merges || []).forEach(function (m) {
    var tl = grid[m.s.r] && grid[m.s.r][m.s.c];
    if (tl === undefined || isBlank(tl)) return;
    for (var r = m.s.r; r <= m.e.r; r++) for (var c = m.s.c; c <= m.e.c; c++) {
      if (r === m.s.r && c === m.s.c) continue;
      if (grid[r]) { if (isBlank(grid[r][c])) grid[r][c] = tl; }
    }
    fixed++;
  });
  return fixed;
}
/* trả về {headers, rows, rowNos, dropped:{blank,repeat,total}, titleRows} */
function tableFromGrid(grid) {
  var cols = 0, i, j;
  for (i = 0; i < grid.length; i++) if (grid[i].length > cols) cols = grid[i].length;
  var v = grid.map(function (r) { var o = r.slice(0); while (o.length < cols) o.push(""); return o; });
  var h = detectHeaderRow(v), hdr = [], seen = {};
  for (j = 0; j < cols; j++) {
    var t = String(isBlank(v[h][j]) ? "" : v[h][j]).replace(/[\r\n]+/g, " ").replace(/\s+/g, " ").replace(/^\s+|\s+$/g, "");
    hdr.push(t);
  }
  var body = [], nos = [], dropped = { blank: 0, repeat: 0, total: 0 };
  var hn = hdr.map(normName);
  for (i = h + 1; i < v.length; i++) {
    var row = v[i], nonEmpty = 0, same = 0;
    for (j = 0; j < cols; j++) if (!isBlank(row[j])) { nonEmpty++; if (hn[j] && normName(row[j]) === hn[j]) same++; }
    if (!nonEmpty) { dropped.blank++; continue; }
    if (same >= 2 && same >= nonEmpty * 0.6) { dropped.repeat++; continue; }
    var tot = false;
    for (j = 0; j < Math.min(cols, 3); j++) if (typeof row[j] === "string" && /^\s*(tổng|tổng cộng|cộng|total|sum|grand total)\b/i.test(row[j])) tot = true;
    if (tot) { dropped.total++; continue; }
    body.push(row); nos.push(i + 1);
  }
  /* bỏ cột không tên và không có dữ liệu; đặt tên cho cột không tên có dữ liệu; chống trùng tên */
  var keep = [];
  for (j = 0; j < cols; j++) {
    var any = false;
    for (i = 0; i < body.length; i++) if (!isBlank(body[i][j])) { any = true; break; }
    if (!hdr[j] && !any) continue;
    if (!hdr[j]) hdr[j] = "Cột " + (j + 1);
    keep.push(j);
  }
  var heads = [];
  keep.forEach(function (j) {
    var nm = hdr[j], base = nm, k = 2;
    while (seen[normName(nm)]) nm = base + " (" + (k++) + ")";
    seen[normName(nm)] = 1; heads.push(nm);
  });
  var rows = body.map(function (r) { return keep.map(function (j) { return r[j]; }); });
  return { headers: heads, rows: rows, rowNos: nos, dropped: dropped, titleRows: h, colIdx: keep.slice(0) };
}
/* gom các tên cột giống nhau (sau khi bỏ dấu/viết hoa) để gửi AI ghép nghĩa */
function collectHeaders(tables) {
  var map = {}, order = [];
  tables.forEach(function (t) {
    var seenHere = {};
    t.headers.forEach(function (h, j) {
      var k = normName(h); if (!k) return;
      if (!map[k]) { map[k] = { name: h, key: k, count: 0, sample: [] }; order.push(k); }
      if (!seenHere[k]) { map[k].count++; seenHere[k] = 1; }
      if (map[k].sample.length < 3) { for (var i = 0; i < t.rows.length && map[k].sample.length < 3; i++) { var x = t.rows[i][j]; if (!isBlank(x) && map[k].sample.indexOf(x) < 0) map[k].sample.push(x); } }
    });
  });
  return order.map(function (k) { return map[k]; });
}
/* mapping: {khóa chuẩn hóa của tên cột gốc: tên cột đích ("" = bỏ)} ; cols: danh sách cột đích theo thứ tự */
function buildMaster(tables, mapping, cols) {
  var out = [], idx = {}, stat = { rows: 0, noKey: 0 };
  cols.forEach(function (c, i) { idx[c] = i; });
  tables.forEach(function (t) {
    var dest = t.headers.map(function (h) { var m = mapping[normName(h)]; return (m && idx[m] !== undefined) ? idx[m] : -1; });
    t.rows.forEach(function (r) {
      var o = new Array(cols.length + 2), j;
      for (j = 0; j < o.length; j++) o[j] = "";
      for (j = 0; j < dest.length; j++) { if (dest[j] >= 0 && !isBlank(r[j]) && isBlank(o[dest[j]])) o[dest[j]] = r[j]; }
      o[cols.length] = t.file; o[cols.length + 1] = t.sheet;
      out.push(o);
    });
  });
  stat.rows = out.length;
  return { rows: out, stat: stat };
}
function lookupKey(x) {
  if (isBlank(x)) return "";
  if (typeof x === "number") return String(x);
  var s = String(x).replace(/[\s ]+/g, " ").replace(/^\s+|\s+$/g, "").toLowerCase();
  return /^\d+$/.test(s) ? String(Number(s)) : s;
}

/* ---------- 2. Soát công thức ---------- */
var CONST_OK = { 0: 1, 1: 1, 2: 1, 100: 1, 1000: 1, 1000000: 1, 365: 1, 360: 1, 12: 1 };
function stripFormula(f) {
  return String(f).replace(/"(?:[^"]|"")*"/g, "\"\"").replace(/(?:'[^']+'|[A-Za-z0-9_.À-ỹ]+)!\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?/g, "REF").replace(/\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}(?![A-Za-z0-9(])/g, "REF");
}
function formulaRefs(f) {
  var s = stripFormula(f), re = /\$?([A-Z]{1,3})\$?(\d{1,7})(?::\$?([A-Z]{1,3})\$?(\d{1,7}))?(?![A-Za-z0-9(])/gi, m, out = [];
  while ((m = re.exec(s))) {
    var c1 = colNum(m[1].toUpperCase()), r1 = Number(m[2]), c2 = m[3] ? colNum(m[3].toUpperCase()) : c1, r2 = m[4] ? Number(m[4]) : r1;
    out.push({ c1: Math.min(c1, c2), c2: Math.max(c1, c2), r1: Math.min(r1, r2), r2: Math.max(r1, r2) });
  }
  return out;
}
function formulaConsts(f) {
  var s = stripFormula(f).replace(/\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?(?![A-Za-z0-9(])/g, "REF");
  var toks = s.match(/[A-Za-z_][A-Za-z0-9_.]*|\d+(?:\.\d+)?%?|\.\d+/g) || [], out = [];
  toks.forEach(function (t) {
    if (!/^[\d.]/.test(t)) return;
    var pct = /%$/.test(t), val = parseFloat(t);
    if (isNaN(val) || CONST_OK[val]) return;
    if (pct || /\./.test(t) || val >= 100) out.push(t);
  });
  return out;
}
/* f1c1, fa1, vals: mảng 2 chiều cùng kích thước; r0,c0: chỉ số dòng/cột (0-based) của ô đầu trong sheet */
function auditSheet(name, f1c1, fa1, vals, r0, c0) {
  var issues = [], rows = f1c1.length, cols = rows ? f1c1[0].length : 0, r, c, hdr = detectHeaderRow(vals);
  function addr(r_, c_) { return colLetter(c0 + c_) + (r0 + r_ + 1); }
  function ref(r_, c_) { return name + "!" + addr(r_, c_); }
  function isF(x) { return typeof x === "string" && x.charAt(0) === "="; }
  var isTot = {};
  for (r = hdr + 1; r < rows; r++) for (c = 0; c < Math.min(cols, 2); c++) if (typeof vals[r][c] === "string" && /^\s*(tổng|cộng|total)/i.test(vals[r][c])) isTot[r] = true;
  for (c = 0; c < cols; c++) {
    var pats = {}, nF = 0, nConst = 0, firstF = -1, lastF = -1;
    for (r = hdr + 1; r < rows; r++) {
      if (isTot[r]) continue;
      if (isF(f1c1[r][c])) { nF++; pats[f1c1[r][c]] = (pats[f1c1[r][c]] || 0) + 1; if (firstF < 0) firstF = r; lastF = r; }
    }
    var top = null, topN = 0, k;
    for (k in pats) if (pats[k] > topN) { top = k; topN = pats[k]; }
    var constSeen = {};
    for (r = hdr + 1; r < rows; r++) {
      if (isTot[r]) continue;
      var f = f1c1[r][c], v = vals[r][c], a1 = fa1[r][c];
      if (isF(f)) {
        if (typeof v === "string" && ERR_RE.test(v)) issues.push({ ref: ref(r, c), type: "Lỗi công thức", text: "Ô báo lỗi " + v, fix: "" });
        var refs = formulaRefs(a1);
        for (var q = 0; q < refs.length; q++) {
          var R = refs[q], rr = r0 + r + 1, cc = c0 + c + 1;
          if (rr >= R.r1 && rr <= R.r2 && cc >= R.c1 && cc <= R.c2 && !/^\s*=\s*SUBTOTAL/i.test(a1)) { issues.push({ ref: ref(r, c), type: "Công thức vòng", text: "Công thức tự trỏ vào chính ô này (vòng lặp)", fix: "sửa vùng tham chiếu" }); break; }
        }
        if (top && nF >= 3 && f !== top && topN / nF >= 0.7) issues.push({ ref: ref(r, c), type: "Công thức lệch", text: "Công thức khác với phần còn lại của cột", fix: "xem lại ô cùng cột" });
        var cs = formulaConsts(a1);
        cs.forEach(function (t) { var kk = c + "|" + t; if (!constSeen[kk]) constSeen[kk] = { n: 0, r: r, t: t }; constSeen[kk].n++; });
      } else if (!isBlank(f) && top && nF >= 3 && nF / (nF + 1) >= 0.5 && r >= firstF && r <= lastF && topN / nF >= 0.6) {
        nConst++;
        issues.push({ ref: ref(r, c), type: "Số gõ đè", text: "Số/chữ gõ tay (" + (typeof v === "number" ? fmt(v) : String(v).substring(0, 20)) + ") nằm giữa cột công thức", fix: "khôi phục công thức chuẩn của cột", restore: { c: c, r1c1: top, r: r } });
      }
    }
    for (k in constSeen) { var cn = constSeen[k]; issues.push({ ref: ref(cn.r, c), type: "Số cứng trong công thức", text: "Công thức có số cứng " + cn.t + (cn.n > 1 ? " (" + cn.n + " ô trong cột)" : ""), fix: "nên đặt ở ô tham số riêng" }); }
  }
  /* dòng tổng cộng bỏ sót dòng dữ liệu: gộp mỗi dòng tổng thành một cảnh báo */
  Object.keys(isTot).forEach(function (rk) {
    r = Number(rk);
    var bad = [], minE = 1e9, maxS = 0;
    for (c = 0; c < cols; c++) {
      var m = /^=\s*SUM\(\s*\$?([A-Z]{1,3})\$?(\d+)\s*:\s*\$?([A-Z]{1,3})\$?(\d+)\s*\)\s*$/i.exec(isF(fa1[r][c]) ? fa1[r][c] : "");
      if (!m) continue;
      var s1 = Number(m[2]), e1 = Number(m[4]), rowNo = r0 + r + 1, hdrNo = r0 + hdr + 1;
      if (s1 > hdrNo + 1 || e1 < rowNo - 1) { bad.push(colLetter(c0 + c)); if (e1 < minE) minE = e1; if (s1 > maxS) maxS = s1; }
    }
    if (bad.length) issues.push({ ref: ref(r, 0), type: "Tổng thiếu dòng", text: "Dòng tổng cộng các cột " + bad.join(", ") + " chỉ cộng đến dòng " + minE + ", trong khi dữ liệu còn tới dòng " + (r0 + r) + " (có thể bỏ sót dòng)", fix: "mở rộng vùng cộng" });
  });
  return issues;
}

/* ---------- 3. Bất thường & đối chiếu ---------- */
/* rows: mảng dòng; rowNos: số dòng trong sheet; idx: {key, period, amount, base, group} (-1 nếu không có) */
function groupSeries(rows, rowNos, idx) {
  var keys = {}, periods = {}, pk = [];
  rows.forEach(function (r, i) {
    var key = lookupKey(r[idx.key]); if (!key) return;
    var p = idx.period >= 0 ? periodInfo(r[idx.period]) : { k: "-", n: 0 };
    if (!periods[p.k]) { periods[p.k] = p; pk.push(p.k); }
    if (!keys[key]) keys[key] = { key: key, label: String(r[idx.key]), group: idx.group >= 0 ? String(r[idx.group] || "") : "", by: {} };
    var e = keys[key].by[p.k]; if (!e) e = keys[key].by[p.k] = { amt: 0, base: 0, n: 0, rowNo: rowNos[i], hasAmt: false };
    var a = toNum(r[idx.amount]), b = idx.base >= 0 ? toNum(r[idx.base]) : null;
    if (a !== null) { e.amt += a; e.hasAmt = true; }
    if (b !== null) e.base += b;
    e.n++;
  });
  var allN = pk.every(function (k) { return periods[k].n !== null; });
  var order = pk.slice(0);
  if (allN) order.sort(function (x, y) { return periods[x].n - periods[y].n; });
  return { keys: keys, order: order };
}
function findAnomalies(rows, rowNos, idx, opts) {
  opts = opts || {};
  var up = (opts.upPct === undefined ? 100 : opts.upPct) / 100, down = (opts.downPct === undefined ? 60 : opts.downPct) / 100;
  var S = groupSeries(rows, rowNos, idx), flags = [], latest = opts.latest && S.order.indexOf(opts.latest) >= 0 ? opts.latest : S.order[S.order.length - 1];
  var hist = S.order.slice(0, S.order.indexOf(latest));
  var keyList = Object.keys(S.keys), cur = [], i;
  function flag(e, type, text, label) { flags.push({ rowNo: e.rowNo, col: idx.amount, key: label, type: type, text: label + ": " + text }); }
  keyList.forEach(function (k) {
    var K = S.keys[k], e = K.by[latest]; if (!e) return;
    if (e.n > 1) flag(e, "Trùng mã", "xuất hiện " + e.n + " dòng trong kỳ " + latest + " (số liệu được cộng gộp)", K.label);
    if (e.hasAmt && e.amt < 0) flag(e, "Số âm", "giá trị âm " + money(e.amt), K.label);
    if (opts.resigned && opts.resigned[k] && e.hasAmt && e.amt !== 0) flag(e, "Đã nghỉ vẫn phát sinh", "trong danh sách đã nghỉ việc nhưng kỳ " + latest + " vẫn có " + money(e.amt), K.label);
    var ha = [], hr = [];
    hist.forEach(function (p) { var h = K.by[p]; if (h && h.hasAmt) { ha.push(h.amt); if (h.base > 0) hr.push(h.amt / h.base); } });
    if (e.hasAmt && ha.length >= 2) {
      var avg = ha.reduce(function (s, x) { return s + x; }, 0) / ha.length;
      if (avg > 0 && e.amt > avg * (1 + up)) flag(e, "Tăng đột biến", "kỳ " + latest + " là " + money(e.amt) + ", cao hơn " + Math.round((e.amt / avg - 1) * 100) + "% so với trung bình " + ha.length + " kỳ trước (" + money(avg) + ")", K.label);
      else if (avg > 0 && e.amt < avg * (1 - down)) flag(e, "Giảm đột biến", "kỳ " + latest + " là " + money(e.amt) + ", thấp hơn " + Math.round((1 - e.amt / avg) * 100) + "% so với trung bình " + ha.length + " kỳ trước (" + money(avg) + ")", K.label);
    }
    if (idx.base >= 0 && e.hasAmt) {
      if (e.amt > 0 && e.base <= 0) flag(e, "Thiếu cơ sở", "có khoản " + money(e.amt) + " nhưng cơ sở tính (doanh thu/chỉ tiêu) bằng 0 hoặc trống", K.label);
      else if (e.base > 0 && hr.length >= 2) {
        var mr = median(hr), nr = e.amt / e.base;
        if (mr > 0 && nr > mr * (1 + up)) flag(e, "Tỷ lệ tăng", "tỷ lệ " + (nr * 100).toFixed(1) + "% so với mức thường " + (mr * 100).toFixed(1) + "% (tăng " + Math.round((nr / mr - 1) * 100) + "%)", K.label);
      }
    }
    cur.push({ k: k, e: e, group: K.group, label: K.label });
  });
  /* so với đồng nghiệp cùng kỳ (độ lệch tuyệt đối trung vị) */
  var groups = {};
  cur.forEach(function (x) { if (x.e.hasAmt && x.e.amt > 0) { var g = idx.group >= 0 && x.group ? x.group : "*"; (groups[g] = groups[g] || []).push(x); } });
  var useGroups = Object.keys(groups).filter(function (g) { return g !== "*" && groups[g].length >= 8; }).length > 0;
  var pools = useGroups ? groups : { "*": cur.filter(function (x) { return x.e.hasAmt && x.e.amt > 0; }) };
  Object.keys(pools).forEach(function (g) {
    var arr = pools[g]; if (arr.length < 8) return;
    var vals = arr.map(function (x) { return x.e.amt; }), med = median(vals), mad = median(vals.map(function (x) { return Math.abs(x - med); }));
    if (mad === 0) return;
    arr.forEach(function (x) {
      var z = 0.6745 * (x.e.amt - med) / mad;
      if (z > 3.5) flag(x.e, "Cao hơn hẳn người khác", "kỳ " + latest + " là " + money(x.e.amt) + ", gấp " + (x.e.amt / med).toFixed(1) + " lần mức giữa của " + (g === "*" ? "cả nhóm" : "nhóm " + g) + " (" + money(med) + ")", x.label);
    });
  });
  return { flags: flags, latest: latest, periods: S.order, people: keyList.length };
}

/* ---------- 4. Chênh lệch quỹ giữa hai kỳ ---------- */
function computeVariance(rows, rowNos, idx, A, B) {
  var S = groupSeries(rows, rowNos, idx), t = { A: A, B: B, sumA: 0, sumB: 0, nA: 0, nB: 0, joined: 0, left: 0, joinedN: 0, leftN: 0, stay: 0, stayN: 0, vol: 0, rate: 0, other: 0, hasBase: idx.base >= 0, groups: {}, people: [] };
  Object.keys(S.keys).forEach(function (k) {
    var K = S.keys[k], a = K.by[A], b = K.by[B], av = a && a.hasAmt ? a.amt : 0, bv = b && b.hasAmt ? b.amt : 0;
    if (a && a.hasAmt) { t.sumA += av; t.nA++; }
    if (b && b.hasAmt) { t.sumB += bv; t.nB++; }
    var inA = a && a.hasAmt && av !== 0, inB = b && b.hasAmt && bv !== 0, d = bv - av, g = K.group || "(không có nhóm)";
    if (!t.groups[g]) t.groups[g] = { a: 0, b: 0 };
    t.groups[g].a += av; t.groups[g].b += bv;
    if (!inA && inB) { t.joined += bv; t.joinedN++; }
    else if (inA && !inB) { t.left -= av; t.leftN++; }
    else if (inA && inB) {
      t.stayN++; t.stay += d;
      if (idx.base >= 0 && a.base > 0 && b.base >= 0) {
        var rateA = av / a.base, rateB = b.base > 0 ? bv / b.base : 0;
        var vol = (b.base - a.base) * rateA, rate = (rateB - rateA) * b.base;
        t.vol += vol; t.rate += rate; t.other += d - vol - rate;
      } else t.other += d;
    }
    if (d !== 0) t.people.push({ label: K.label, d: d, a: av, b: bv });
  });
  t.delta = t.sumB - t.sumA;
  t.pct = t.sumA !== 0 ? t.delta / t.sumA * 100 : null;
  t.people.sort(function (x, y) { return Math.abs(y.d) - Math.abs(x.d); });
  return t;
}

/* ---------- 5. Bộ truy vấn chạy trên máy ---------- */
function resolveCol(headers, name) {
  var n = normName(name), i;
  for (i = 0; i < headers.length; i++) if (normName(headers[i]) === n) return i;
  for (i = 0; i < headers.length; i++) if (n && (normName(headers[i]).indexOf(n) >= 0 || n.indexOf(normName(headers[i])) >= 0) && normName(headers[i])) return i;
  var toks = n.split(" "), best = -1, bs = 0;
  headers.forEach(function (h, j) { var ht = normName(h).split(" "), s = 0; toks.forEach(function (t) { if (t && ht.indexOf(t) >= 0) s++; }); if (s > bs) { bs = s; best = j; } });
  if (best >= 0 && bs >= Math.max(1, Math.ceil(toks.length / 2))) return best;
  throw new Error("không thấy cột \"" + name + "\" (các cột có: " + headers.join(", ").substring(0, 200) + ")");
}
function cmpVal(cell, op, val) {
  var cn = toNum(cell), vn = (typeof val === "number") ? val : toNum(val);
  if (typeof val === "string" && /^\d{4}-\d{2}$/.test(val)) {
    var pi = periodInfo(cell).k;
    if (op === "=" || op === "==") return pi === val; if (op === "!=") return pi !== val;
    if (op === ">=") return pi >= val; if (op === "<=") return pi <= val; if (op === ">") return pi > val; if (op === "<") return pi < val;
  }
  switch (op) {
    case "=": case "==": return (cn !== null && vn !== null) ? cn === vn : normName(cell) === normName(val);
    case "!=": return (cn !== null && vn !== null) ? cn !== vn : normName(cell) !== normName(val);
    case ">": return cn !== null && vn !== null && cn > vn;
    case ">=": return cn !== null && vn !== null && cn >= vn;
    case "<": return cn !== null && vn !== null && cn < vn;
    case "<=": return cn !== null && vn !== null && cn <= vn;
    case "contains": return normName(cell).indexOf(normName(val)) >= 0;
    case "not_contains": return normName(cell).indexOf(normName(val)) < 0;
    case "not_empty": return !isBlank(cell);
    case "empty": return isBlank(cell);
    case "in": return (val || []).some(function (x) { return normName(x) === normName(cell); });
  }
  throw new Error("phép so sánh không hỗ trợ: " + op);
}
function aggInit() { return { sum: 0, n: 0, min: null, max: null, set: {} }; }
function aggAdd(a, x) {
  var v = toNum(x);
  if (!isBlank(x)) { a.set[String(x)] = 1; }
  if (v !== null) { a.sum += v; a.n++; if (a.min === null || v < a.min) a.min = v; if (a.max === null || v > a.max) a.max = v; }
  else if (!isBlank(x)) a.n++;
}
function aggVal(a, fn) {
  switch (fn) {
    case "sum": return a.sum; case "count": return a.n; case "avg": return a.n ? a.sum / a.n : 0;
    case "min": return a.min === null ? "" : a.min; case "max": return a.max === null ? "" : a.max;
    case "countd": return Object.keys(a.set).length;
  }
  throw new Error("hàm tổng hợp không hỗ trợ: " + fn);
}
function runQuery(headers, rows, q) {
  var data = rows, i, j;
  (q.where || []).forEach(function (w) {
    var ci = resolveCol(headers, w.col), c2 = w.col2 ? resolveCol(headers, w.col2) : -1;
    data = data.filter(function (r) { return c2 >= 0 ? cmpVal(r[ci], w.op, r[c2]) : cmpVal(r[ci], w.op, w.value !== undefined ? w.value : w.values); });
  });
  var outH, outR;
  if (q.group_by && q.group_by.length) {
    var gi = q.group_by.map(function (g) { return resolveCol(headers, g); });
    var aggs = (q.aggregates && q.aggregates.length) ? q.aggregates : [{ fn: "count", col: q.group_by[0], as: "Số dòng" }];
    var pv = q.pivot ? resolveCol(headers, q.pivot.col) : -1, pvals = [];
    if (pv >= 0) {
      if (q.pivot.values && q.pivot.values.length) pvals = q.pivot.values.map(String);
      else { var sp = {}; data.forEach(function (r) { var k = periodInfo(r[pv]).k; if (!sp[k]) { sp[k] = 1; pvals.push(k); } }); pvals.sort(); }
    }
    var groups = {}, order = [];
    data.forEach(function (r) {
      var key = gi.map(function (c) { return lookupKey(r[c]); }).join("\u0001");
      var g = groups[key];
      if (!g) { g = groups[key] = { lab: gi.map(function (c) { return r[c]; }), cells: {} }; order.push(key); }
      var pk = pv >= 0 ? periodInfo(r[pv]).k : "*";
      aggs.forEach(function (a, ai) {
        var ci = a.col ? resolveCol(headers, a.col) : gi[0], cellKey = pk + "|" + ai;
        if (!g.cells[cellKey]) g.cells[cellKey] = aggInit();
        aggAdd(g.cells[cellKey], r[ci]);
      });
    });
    outH = q.group_by.map(function (g, k) { return headers[gi[k]]; });
    var useP = pv >= 0 ? pvals : ["*"];
    aggs.forEach(function (a, ai) { useP.forEach(function (p) { outH.push((a.as || (a.fn + " " + (a.col || ""))) + (pv >= 0 ? " (" + p + ")" : "")); }); });
    var diffAt = [];
    if (pv >= 0 && pvals.length === 2) {
      aggs.forEach(function (a, ai) { outH.push("Chênh lệch " + (a.as || a.fn)); outH.push("% chênh " + (a.as || a.fn)); });
    }
    outR = order.map(function (key) {
      var g = groups[key], row = g.lab.slice(0), vals = [];
      aggs.forEach(function (a, ai) { useP.forEach(function (p) { var c = g.cells[p + "|" + ai]; var v = c ? aggVal(c, a.fn) : 0; row.push(v); vals.push(v); }); });
      if (pv >= 0 && pvals.length === 2) {
        aggs.forEach(function (a, ai) {
          var v0 = Number(vals[ai * 2]) || 0, v1 = Number(vals[ai * 2 + 1]) || 0;
          row.push(v1 - v0); row.push(v0 !== 0 ? (v1 - v0) / v0 : "");
        });
      }
      return row;
    });
  } else {
    var cols = (q.columns && q.columns.length) ? q.columns.map(function (c) { return resolveCol(headers, c); }) : headers.map(function (h, k) { return k; });
    outH = cols.map(function (c) { return headers[c]; });
    outR = data.map(function (r) { return cols.map(function (c) { return r[c]; }); });
  }
  (q.having || []).forEach(function (h) {
    var hi = resolveCol(outH, h.as || h.col);
    outR = outR.filter(function (r) { return cmpVal(r[hi], h.op, h.value); });
  });
  if (q.sort && q.sort.by) {
    var si = resolveCol(outH, q.sort.by), desc = String(q.sort.order).toLowerCase() !== "asc";
    outR = outR.map(function (r, k) { return { r: r, k: k }; }).sort(function (x, y) {
      var a = x.r[si], b = y.r[si], an = toNum(a), bn = toNum(b), d;
      if (an !== null && bn !== null) d = an - bn; else d = String(a).localeCompare(String(b));
      return (desc ? -d : d) || x.k - y.k;
    }).map(function (o) { return o.r; });
  }
  var total = outR.length;
  if (q.top) outR = outR.slice(0, Math.max(1, Math.min(1000, Number(q.top) || 20)));
  return { headers: outH, rows: outR, total: total, matched: data.length };
}

/* ---------- 6. Sinh nhật & thâm niên ---------- */
function dateParts(v) {
  var d, m, y, mm;
  if (typeof v === "number") {
    if (v < 1 || v > 80000) return null;
    var dt = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
    return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() };
  }
  var s = String(v === undefined || v === null ? "" : v).replace(/^\s+|\s+$/g, "");
  if ((mm = /^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/.exec(s))) { d = Number(mm[1]); m = Number(mm[2]); y = Number(mm[3]); }
  else if ((mm = /^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/.exec(s))) { y = Number(mm[1]); m = Number(mm[2]); d = Number(mm[3]); }
  else return null;
  if (m < 1 || m > 12 || d < 1 || d > 31 || y < 1900 || y > 2200) return null;
  return { y: y, m: m, d: d };
}
function dayDiff(a, b) { return Math.round((Date.UTC(b.y, b.m - 1, b.d) - Date.UTC(a.y, a.m - 1, a.d)) / 86400000); }
function twoD(n) { return (n < 10 ? "0" : "") + n; }
/* rows: mảng dòng; rowNos: số dòng trong sheet; idx: {name, dept, birth, join, status}; today:{y,m,d} */
function hrEvents(rows, rowNos, idx, today, opts) {
  opts = opts || {};
  var miles = opts.milestones || [5, 10, 15, 20, 25, 30], soon = opts.soonDays || 7, re = opts.leftRe || /nghỉ|thôi việc|resign/i;
  var out = { bToday: [], bSoon: [], bMonth: [], aToday: [], aMonth: [], aOther: 0, noBirth: 0, noJoin: 0, skipped: 0, people: 0 };
  rows.forEach(function (r, i) {
    if (idx.status >= 0 && re.test(String(r[idx.status]))) { out.skipped++; return; }
    out.people++;
    var name = String(isBlank(r[idx.name]) ? "(không tên)" : r[idx.name]), dept = idx.dept >= 0 ? String(r[idx.dept] || "") : "";
    var b = idx.birth >= 0 ? dateParts(r[idx.birth]) : null, j = idx.join >= 0 ? dateParts(r[idx.join]) : null;
    if (idx.birth >= 0 && !b) out.noBirth++;
    if (idx.join >= 0 && !j) out.noJoin++;
    if (b) {
      var bd = b.d, bm = b.m;
      if (bm === 2 && bd === 29) bd = 28;
      var thisYear = { y: today.y, m: bm, d: bd }, du = dayDiff(today, thisYear);
      var it = { name: name, dept: dept, label: twoD(bd) + "/" + twoD(bm), age: today.y - b.y, days: du, rowNo: rowNos[i], col: idx.birth, nameCol: idx.name };
      if (du === 0) out.bToday.push(it);
      else if (du > 0 && du <= soon) out.bSoon.push(it);
      if (bm === today.m) out.bMonth.push(it);
    }
    if (j) {
      var yrs = today.y - j.y, jd = j.d;
      var it2 = { name: name, dept: dept, label: twoD(jd) + "/" + twoD(j.m) + "/" + j.y, years: yrs, days: dayDiff(today, { y: today.y, m: j.m, d: jd }), rowNo: rowNos[i], col: idx.join, nameCol: idx.name };
      if (yrs >= 1 && j.m === today.m) {
        if (miles.indexOf(yrs) >= 0) { out.aMonth.push(it2); if (it2.days === 0) out.aToday.push(it2); }
        else out.aOther++;
      }
    }
  });
  function bySoon(a, b) { return a.days - b.days; }
  out.bMonth.sort(bySoon); out.bSoon.sort(bySoon); out.aMonth.sort(function (a, b) { return a.days - b.days || b.years - a.years; });
  return out;
}

/* ---------- 7. Soát bảng lương một kỳ ---------- */
function payrollDetect(headers) {
  var hn = headers.map(normName);
  function find(re, not) {
    for (var i = 0; i < hn.length; i++) if (re.test(hn[i]) && !(not && not.test(hn[i]))) return i;
    return -1;
  }
  var det = {
    key: find(/^(ma nv|ma nhan vien|msnv|ma so nv|ma so|ma|id)\b/), name: find(/ho ten|ho va ten|ten nhan vien|^ten\b/), dept: find(/phong|bo phan|chi nhanh|don vi|khu vuc/),
    base: find(/luong co ban|\blcb\b|luong cb|muc luong/, /cong|tong/), allow: find(/phu cap/), days: find(/ngay cong|so cong|cong thuc te|ngay lam viec|so ngay lam/, /chuan|quy dinh/),
    std: find(/cong chuan|ngay cong chuan|cong quy dinh|ngay chuan/), kpi: find(/kpi|ty le dat|hieu suat|danh gia/, /thuong|tien/), bonus: find(/thuong|hoa hong/),
    gross: find(/tong thu nhap|tong luong|gross|tong cong thu nhap|luong gop/), net: find(/thuc linh|thuc nhan|\bnet\b|con lai/), pay: find(/luong theo cong|luong thuc te|luong thang|luong san pham/)
  };
  return det;
}
function numCol(rows, c) {
  var out = [], cnt = 0;
  for (var i = 0; i < rows.length; i++) { var x = rows[i][c]; var v = (typeof x === "number") ? x : null; if (v !== null) cnt++; out.push(v); }
  return { v: out, frac: rows.length ? cnt / rows.length : 0 };
}
function mineRelations(names, cols, n) {
  var m = names.length, res = [], nz = cols.map(function (c) { var k = 0; c.forEach(function (x) { if (x !== null && x !== 0) k++; }); return k / Math.max(1, n); });
  var meds = cols.map(function (c) { return median(c.filter(function (x) { return x !== null && x !== 0; }).map(Math.abs)); });
  var step = Math.max(1, Math.floor(n / 300)), sample = [], i;
  for (i = 0; i < n; i += step) sample.push(i);
  function tol(c) { return Math.max(1, Math.abs(c) * 0.003); }
  function score(t, calc, idxs) {
    var ok = 0, tot = 0;
    for (var q = 0; q < idxs.length; q++) { var r = idxs[q], c = cols[t][r]; if (c === null) continue; var v = calc(r); if (v === null) continue; tot++; if (Math.abs(v - c) <= tol(c)) ok++; }
    return { ok: ok, tot: tot };
  }
  var all = []; for (i = 0; i < n; i++) all.push(i);
  for (var t = 0; t < m; t++) {
    if (nz[t] < 0.6) continue;
    var L = []; for (var a = 0; a < m; a++) if (a !== t && nz[a] >= 0.3) L.push(a);
    if (L.length > 12) L = L.slice(0, 12);
    var found = {};
    function consider(kind, terms, calc, text) {
      if (kind === "sum" || kind === "diff") { for (var q0 = 0; q0 < terms.length; q0++) if (meds[terms[q0]] < 0.02 * meds[t]) return; }
      var s = score(t, calc, sample);
      if (s.tot < Math.min(10, n) || s.ok / s.tot < 0.7) return;
      var f = score(t, calc, all);
      if (f.tot < Math.max(10, n * 0.6) || f.ok / f.tot < 0.8 || f.ok === f.tot) return;
      var cand = { target: t, kind: kind, terms: terms, calc: calc, text: text, rate: f.ok / f.tot, ok: f.ok, tot: f.tot };
      var sk = terms.slice(0).sort(function (a_, b_) { return a_ - b_; }).join(","), old = found[sk];
      if (!old || cand.rate > old.rate + 0.01 || (Math.abs(cand.rate - old.rate) <= 0.01 && terms.length < old.terms.length)) found[sk] = cand;
    }
    function val(x, r) { return cols[x][r]; }
    function sumOf(ts, r) { var s = 0; for (var k = 0; k < ts.length; k++) { var v = val(ts[k], r); if (v === null) return null; s += v; } return s; }
    function combos(arr, k, cb, start, cur) {
      start = start || 0; cur = cur || [];
      if (cur.length === k) { cb(cur.slice(0)); return; }
      for (var x = start; x < arr.length; x++) { cur.push(arr[x]); combos(arr, k, cb, x + 1, cur); cur.pop(); }
    }
    var j;
    for (j = 2; j <= 4; j++) combos(L, j, function (ts) { consider("sum", ts, function (r) { return sumOf(ts, r); }, names[t] + " = " + ts.map(function (x) { return names[x]; }).join(" + ")); });
    L.forEach(function (A) {
      var rest = L.filter(function (x) { return x !== A; });
      for (j = 1; j <= 3; j++) combos(rest, j, function (ts) {
        consider("diff", [A].concat(ts), function (r) { var a1 = val(A, r), s = sumOf(ts, r); return a1 === null || s === null ? null : a1 - s; }, names[t] + " = " + names[A] + " − " + ts.map(function (x) { return names[x]; }).join(" − "));
      });
    });
    L.forEach(function (A) { L.forEach(function (B) {
      if (B <= A) return;
      consider("prod", [A, B], function (r) { var a1 = val(A, r), b1 = val(B, r); return a1 === null || b1 === null ? null : a1 * b1; }, names[t] + " = " + names[A] + " × " + names[B]);
    }); });
    L.forEach(function (A) { L.forEach(function (B) { if (A === B) return; L.forEach(function (C) {
      if (C === A || C === B) return;
      consider("ratio", [A, B, C], function (r) { var a1 = val(A, r), b1 = val(B, r), c1 = val(C, r); return a1 === null || b1 === null || c1 === null || b1 === 0 ? null : a1 / b1 * c1; }, names[t] + " = " + names[A] + " ÷ " + names[B] + " × " + names[C]);
    }); }); });
    Object.keys(found).forEach(function (k) { res.push(found[k]); });
  }
  return res;
}
function rankNorm(arr) {
  var idx = arr.map(function (v, i) { return { v: v, i: i }; }).filter(function (x) { return x.v !== null; }).sort(function (a, b) { return a.v - b.v; });
  var out = new Array(arr.length), n = idx.length, i = 0;
  while (i < n) { var j = i; while (j + 1 < n && idx[j + 1].v === idx[i].v) j++; var r = (i + j) / 2 / Math.max(1, n - 1); for (var k = i; k <= j; k++) out[idx[k].i] = r; i = j + 1; }
  return out;
}
function peerOutliers(vals, groups, minGroup) {
  /* trả về mảng {i, side, med, z} */
  var res = [], by = {}, gi;
  vals.forEach(function (v, i) { if (v === null) return; var g = groups ? groups[i] : "*"; (by[g] = by[g] || []).push(i); });
  var useG = groups && Object.keys(by).filter(function (g) { return by[g].length >= minGroup; }).length > 1;
  var pools = {};
  if (useG) { Object.keys(by).forEach(function (g) { if (by[g].length >= minGroup) pools[g] = by[g]; }); }
  else pools["*"] = [].concat.apply([], Object.keys(by).map(function (g) { return by[g]; }));
  Object.keys(pools).forEach(function (g) {
    var ids = pools[g]; if (ids.length < 8) return;
    var xs = ids.map(function (i) { return vals[i]; }), med = median(xs), mad = median(xs.map(function (x) { return Math.abs(x - med); }));
    ids.forEach(function (i) {
      var x = vals[i];
      if (mad === 0) {
        var same = xs.filter(function (y) { return y === med; }).length / xs.length;
        if (same >= 0.7 && med !== 0 && Math.abs(x - med) / Math.abs(med) > 1) res.push({ i: i, side: x > med ? "hi" : "lo", med: med, group: g });
        else if (same >= 0.7 && med === 0 && x !== 0) res.push({ i: i, side: "hi", med: med, group: g });
        return;
      }
      var z = 0.6745 * (x - med) / mad;
      if (Math.abs(z) > 3.5 && Math.abs(x - med) > Math.abs(med) * 0.25) res.push({ i: i, side: z > 0 ? "hi" : "lo", med: med, group: g, z: z });
    });
  });
  return res;
}
function payrollCheck(headers, rows, rowNos) {
  var det = payrollDetect(headers), flags = [], n = rows.length, i, c;
  function add(i_, col, type, text) { flags.push({ rowNo: rowNos[i_], col: col, type: type, text: text, who: who(i_) }); }
  function who(i_) { var r = rows[i_]; return String(det.name >= 0 && !isBlank(r[det.name]) ? r[det.name] : (det.key >= 0 ? r[det.key] : "dòng " + rowNos[i_])); }
  var numeric = {}, groups = det.dept >= 0 ? rows.map(function (r) { return String(isBlank(r[det.dept]) ? "" : r[det.dept]); }) : null;
  for (c = 0; c < headers.length; c++) { var nc = numCol(rows, c); if (nc.frac >= 0.7) numeric[c] = nc.v; }
  /* 1. mã, tên, phòng ban */
  if (det.key >= 0) {
    var seen = {};
    rows.forEach(function (r, i_) { var k = lookupKey(r[det.key]); if (!k) { add(i_, det.key, "Thiếu thông tin", "thiếu mã nhân viên"); return; } (seen[k] = seen[k] || []).push(i_); });
    Object.keys(seen).forEach(function (k) { if (seen[k].length > 1) seen[k].forEach(function (i_) { add(i_, det.key, "Trùng mã", "mã " + rows[i_][det.key] + " xuất hiện " + seen[k].length + " lần (dòng " + seen[k].map(function (x) { return rowNos[x]; }).join(", ") + ")"); }); });
  }
  if (det.dept >= 0) rows.forEach(function (r, i_) { if (isBlank(r[det.dept])) add(i_, det.dept, "Thiếu thông tin", "chưa có phòng ban/bộ phận"); });
  /* 2. ngày công */
  if (det.days >= 0 && numeric[det.days]) {
    var stdv = det.std >= 0 ? numeric[det.std] : null;
    numeric[det.days].forEach(function (d, i_) {
      if (d === null) return;
      var lim = stdv && stdv[i_] !== null ? stdv[i_] : 31;
      if (d < 0) add(i_, det.days, "Ngày công", "ngày công âm (" + d + ")");
      else if (d > lim + 0.01) add(i_, det.days, "Ngày công", "ngày công " + d + " vượt " + (stdv && stdv[i_] !== null ? "công chuẩn " + lim : "số ngày tối đa trong tháng"));
    });
  }
  /* 3. KPI */
  var kpiV = det.kpi >= 0 ? numeric[det.kpi] : null, kscale = 1;
  if (kpiV) {
    var kmax = Math.max.apply(null, kpiV.filter(function (x) { return x !== null; }).concat([0])), kmed = median(kpiV.filter(function (x) { return x !== null; }));
    kscale = kmed > 3 ? 100 : 1;
    kpiV.forEach(function (k, i_) { if (k === null) return; var p = k / kscale * 100; if (p < 0) add(i_, det.kpi, "KPI", "KPI âm (" + Math.round(p) + "%)"); else if (p > 150) add(i_, det.kpi, "KPI", "KPI " + Math.round(p) + "% vượt xa mức thường gặp (trên 150%)"); });
    if (det.bonus >= 0 && numeric[det.bonus]) {
      var bonusN = numeric[det.bonus].map(function (b, i_) { var bs = det.base >= 0 && numeric[det.base] ? numeric[det.base][i_] : null; return b === null ? null : (bs ? b / bs : b); });
      var rk = rankNorm(kpiV), rb = rankNorm(bonusN);
      kpiV.forEach(function (k, i_) {
        if (k === null || rk[i_] === undefined || rb[i_] === undefined) return;
        var p = k / kscale * 100;
        if (p >= 0 && p <= 150) {
          if (rb[i_] - rk[i_] > 0.6 && numeric[det.bonus][i_] > 0) add(i_, det.bonus, "Thưởng lệch KPI", "thưởng " + money(numeric[det.bonus][i_]) + " thuộc nhóm cao nhưng KPI chỉ " + Math.round(p) + "% (thấp)");
          else if (rk[i_] - rb[i_] > 0.6 && p >= 100) add(i_, det.bonus, "Thưởng lệch KPI", "KPI " + Math.round(p) + "% thuộc nhóm cao nhưng thưởng chỉ " + money(numeric[det.bonus][i_] || 0));
        }
      });
    }
  }
  /* 4. cao/thấp bất thường so với đồng nghiệp cùng phòng */
  var peerHits = [];
  var checkCols = [["base", "Lương cơ bản"], ["allow", "Phụ cấp"], ["bonus", "Thưởng"], ["pay", "Lương theo công"], ["gross", "Tổng thu nhập"], ["net", "Thực lĩnh"], ["days", "Ngày công"]];
  checkCols.forEach(function (cc) {
    var col = det[cc[0]]; if (col < 0 || !numeric[col]) return;
    peerOutliers(numeric[col], groups, 8).forEach(function (o) {
      var x = numeric[col][o.i];
      if (o.side === "lo" && x === 0 && (cc[0] === "bonus" || cc[0] === "allow")) return;
      peerHits.push({ i: o.i, col: col, side: o.side });
      add(o.i, col, o.side === "hi" ? "Cao bất thường" : "Thấp bất thường", headers[col] + " " + (cc[0] === "days" ? x : money(x)) + (o.side === "hi" ? " cao hơn hẳn" : " thấp hơn hẳn") + " mức thường gặp" + (o.group && o.group !== "*" ? " của " + o.group : " của cả bảng") + " (" + (cc[0] === "days" ? o.med : money(o.med)) + ")");
    });
  });
  /* 5. công thức tính khác với số đông (phát hiện quan hệ giữa các cột) */
  var names = [], cols = [], map = [];
  Object.keys(numeric).forEach(function (k) { if (Number(k) === det.kpi) return; names.push(headers[k]); cols.push(numeric[k]); map.push(Number(k)); });
  var rels = names.length >= 3 && n >= 10 ? mineRelations(names, cols, n) : [];
  (function () {
    var bySet = {};
    rels.forEach(function (rl) {
      var key = [rl.target].concat(rl.terms).sort(function (a, b) { return a - b; }).join(",");
      var cur = bySet[key];
      if (!cur || map[rl.target] > map[cur.target] || (map[rl.target] === map[cur.target] && rl.rate > cur.rate)) bySet[key] = rl;
    });
    rels = Object.keys(bySet).map(function (k) { return bySet[k]; });
  })();
  var relInfo = [];
  rels.forEach(function (rl) {
    relInfo.push(rl.text + " (đúng với " + rl.ok + "/" + rl.tot + " dòng)");
    for (var r = 0; r < n; r++) {
      var c1 = cols[rl.target][r], v = rl.calc(r);
      if (c1 === null || v === null) continue;
      if (Math.abs(v - c1) > Math.max(1, Math.abs(c1) * 0.003)) add(r, map[rl.target], "Khác công thức chung", "đa số dòng tính " + rl.text.split(" = ")[1] + ", dòng này " + money(c1) + " nhưng theo công thức phải là " + money(v) + " (lệch " + money(c1 - v) + ")");
    }
  });
  /* 6. thực lĩnh so với tổng thu nhập */
  if (det.gross >= 0 && det.net >= 0 && numeric[det.gross] && numeric[det.net]) {
    var ratio = numeric[det.gross].map(function (g, i_) { var nn = numeric[det.net][i_]; return g && nn !== null && g > 0 ? nn / g : null; });
    rows.forEach(function (r, i_) { var g = numeric[det.gross][i_], nn = numeric[det.net][i_]; if (g !== null && nn !== null && nn > g) add(i_, det.net, "Thực lĩnh lớn hơn thu nhập", "thực lĩnh " + money(nn) + " lớn hơn tổng thu nhập " + money(g)); if (nn !== null && nn < 0) add(i_, det.net, "Số âm", "thực lĩnh âm"); });
    peerOutliers(ratio, null, 8).forEach(function (o) { add(o.i, det.net, "Tỷ lệ thực lĩnh lạ", "thực lĩnh chỉ bằng " + Math.round(ratio[o.i] * 100) + "% tổng thu nhập (thường " + Math.round(o.med * 100) + "%)"); });
  }
  /* gộp trùng: cùng ô, cùng loại chỉ giữ một */
  var seenF = {}, uniq = [];
  flags.forEach(function (f) { var k = f.rowNo + "|" + f.col + "|" + f.type; if (!seenF[k]) { seenF[k] = 1; uniq.push(f); } });
  (function () {
    var groupsF = {};
    uniq.forEach(function (f, k) { if (f.type === "Cao bất thường" || f.type === "Thấp bất thường") { var key = f.rowNo + "|" + f.type; (groupsF[key] = groupsF[key] || []).push(k); } });
    var drop = {};
    Object.keys(groupsF).forEach(function (key) {
      var ks = groupsF[key]; if (ks.length < 2) return;
      ks.sort(function (a, b) { return uniq[a].col - uniq[b].col; });
      var main = uniq[ks[0]], others = ks.slice(1).map(function (k) { drop[k] = 1; return headers[uniq[k].col]; });
      main.text += " (kéo theo: " + others.join(", ") + ")";
    });
    uniq = uniq.filter(function (f, k) { return !drop[k]; });
  })();
  /* 7. tóm tắt */
  var S = { n: n, det: det, relations: relInfo, kpiScale: kscale, dept: [] };
  function sum(c) { return c >= 0 && numeric[c] ? numeric[c].reduce(function (s, x) { return s + (x || 0); }, 0) : null; }
  S.grossSum = sum(det.gross); S.netSum = sum(det.net); S.baseSum = sum(det.base);
  S.netAvg = S.netSum !== null && n ? S.netSum / n : null;
  S.daysAvg = det.days >= 0 && numeric[det.days] ? sum(det.days) / Math.max(1, numeric[det.days].filter(function (x) { return x !== null; }).length) : null;
  S.kpiAvg = kpiV ? kpiV.reduce(function (s, x) { return s + (x || 0); }, 0) / Math.max(1, kpiV.filter(function (x) { return x !== null; }).length) / kscale * 100 : null;
  if (det.dept >= 0) {
    var dg = {};
    rows.forEach(function (r, i_) { var g = groups[i_] || "(chưa có phòng ban)"; var o = dg[g] = dg[g] || { n: 0, gross: 0, net: 0, base: 0 }; o.n++; if (det.gross >= 0 && numeric[det.gross]) o.gross += numeric[det.gross][i_] || 0; if (det.net >= 0 && numeric[det.net]) o.net += numeric[det.net][i_] || 0; if (det.base >= 0 && numeric[det.base]) o.base += numeric[det.base][i_] || 0; });
    S.dept = Object.keys(dg).map(function (g) { return { name: g, n: dg[g].n, gross: dg[g].gross, net: dg[g].net, base: dg[g].base }; }).sort(function (a, b) { return (b.net || b.gross) - (a.net || a.gross); });
  }
  return { flags: uniq, det: det, summary: S };
}

/* ====================================================================
   CÔNG CỤ NÂNG CAO: gộp nhiều file, lấy dữ liệu từ file khác, soát công thức,
   bất thường, chênh lệch, truy vấn. Phần xử lý số liệu chạy trên máy.
   ==================================================================== */
function showTab(name) {
  $("chatView").className = name === "chat" ? "" : "hide";
  $("toolsView").className = name === "tools" ? "" : "hide";
  $("tabChat").className = "tab" + (name === "chat" ? " on" : "");
  $("tabTools").className = "tab" + (name === "tools" ? " on" : "");
  if (name === "chat") toBottom();
}
function uniqueName(base) {
  var d = new Date(), pad = function (n) { return (n < 10 ? "0" : "") + n; };
  return String(base).replace(/[:\\\/?*\[\]]/g, "").substring(0, 22) + " " + pad(d.getHours()) + pad(d.getMinutes()) + pad(d.getSeconds());
}
function escCell(x) { return (typeof x === "string" && x.charAt(0) === "=") ? "'" + x : x; }
/* ghi lưới dữ liệu lớn theo từng cụm 2000 dòng */
function writeGrid(ctx, ws, r0, c0, grid) {
  var w = 0, i; for (i = 0; i < grid.length; i++) if (grid[i].length > w) w = grid[i].length;
  var pos = 0;
  function step() {
    if (pos >= grid.length) return ctx.sync();
    var chunk = grid.slice(pos, pos + 2000).map(function (r) { var o = []; for (var j = 0; j < w; j++) o.push(j < r.length ? escCell(r[j]) : ""); return o; });
    ws.getRangeByIndexes(r0 + pos, c0, chunk.length, w).values = chunk;
    pos += chunk.length;
    return ctx.sync().then(step);
  }
  return step();
}
/* đọc một sheet thành bảng (tự nhận dòng tiêu đề, bỏ dòng trống/tổng) */
function readSheetTable(sheetName, rangeStr, cb) {
  Excel.run(function (ctx) {
    var ws = sheetName ? ctx.workbook.worksheets.getItem(sheetName) : ctx.workbook.worksheets.getActiveWorksheet(); ws.load("name");
    var used = ws.getUsedRangeOrNullObject(); used.load("rowCount,columnCount,rowIndex,columnIndex");
    return ctx.sync().then(function () {
      if (used.isNullObject) throw new Error("Sheet đang trống.");
      var base = used;
      if (rangeStr && /\d/.test(rangeStr) && RANGE_RE.test(String(rangeStr).replace(/\s/g, ""))) {
        var inter = ws.getRange(String(rangeStr).replace(/\s/g, "")).getIntersectionOrNullObject(used); inter.load("isNullObject,rowCount,columnCount,rowIndex,columnIndex");
        return ctx.sync().then(function () { if (inter.isNullObject) throw new Error("Vùng chọn nằm ngoài dữ liệu."); return fetchIt(inter); });
      }
      return fetchIt(base);
      function fetchIt(u) {
        var cols = u.columnCount, rows = Math.max(1, Math.min(u.rowCount, Math.floor(300000 / cols)));
        var rg = u.getCell(0, 0).getResizedRange(rows - 1, cols - 1); rg.load("values");
        return ctx.sync().then(function () { return { name: ws.name, r0: u.rowIndex, c0: u.columnIndex, cols: cols, cut: rows < u.rowCount, values: rg.values }; });
      }
    });
  }).then(function (d) {
    var t = tableFromGrid(d.values);
    cb(null, { name: d.name, r0: d.r0, c0: d.c0, cols: d.cols, cut: d.cut, t: t });
  }, function (e) { cb(String(e && e.message || e)); });
}

/* ---------- đọc nhiều file bằng thư viện SheetJS ---------- */
function readFileBuf(f, cb) {
  var fr = new FileReader();
  fr.onload = function () { cb(null, fr.result); };
  fr.onerror = function () { cb("Không đọc được file " + f.name); };
  fr.readAsArrayBuffer(f);
}
function loadFiles(files, opts, progress, cb) {
  if (typeof XLSX === "undefined") { cb("Chưa tải được thư viện đọc Excel (cần có mạng để tải). Đóng mở lại khung này rồi thử lại."); return; }
  var tables = [], notes = { files: 0, sheets: 0, skipped: [], merges: 0, blank: 0, repeat: 0, total: 0, errors: [] }, i = 0;
  function next() {
    if (i >= files.length) { cb(null, { tables: tables, notes: notes }); return; }
    var f = files[i++];
    progress("Đang đọc " + i + "/" + files.length + ": " + f.name);
    readFileBuf(f, function (err, buf) {
      if (err) { notes.errors.push(err); setTimeout(next, 0); return; }
      try {
        var wb = XLSX.read(buf, { type: "array" }); notes.files++;
        var names = wb.SheetNames.slice(0, opts.firstOnly ? 1 : wb.SheetNames.length);
        names.forEach(function (sn) {
          var ws = wb.Sheets[sn];
          if (!ws || !ws["!ref"]) { notes.skipped.push(f.name + " / " + sn); return; }
          var dr = XLSX.utils.decode_range(ws["!ref"]);
          var grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
          var merges = (ws["!merges"] || []).map(function (m) { return { s: { r: m.s.r - dr.s.r, c: m.s.c - dr.s.c }, e: { r: m.e.r - dr.s.r, c: m.e.c - dr.s.c } }; });
          notes.merges += applyMerges(grid, merges);
          var t = tableFromGrid(grid);
          if (t.headers.length < 2 || !t.rows.length) { notes.skipped.push(f.name + " / " + sn); return; }
          t.file = f.name; t.sheet = sn; tables.push(t); notes.sheets++;
          notes.blank += t.dropped.blank; notes.repeat += t.dropped.repeat; notes.total += t.dropped.total;
        });
      } catch (e) { notes.errors.push(f.name + ": " + String(e && e.message || e)); }
      setTimeout(next, 0);
    });
  }
  next();
}
function jsonOf(text) {
  var s = String(text || "").replace(/```json|```/g, ""), a = s.indexOf("{"), b = s.lastIndexOf("}");
  if (a < 0 || b <= a) return null;
  try { return JSON.parse(s.substring(a, b + 1)); } catch (e) { return null; }
}
function aiTask(task, question, context, cb) {
  api("/api/chat", "POST", { task: task, question: question, context: context, stream: false, compat: compat() }, function (e, d) {
    if (e) { cb(e); return; }
    updateQuota(d.quota, d.meta);
    cb(null, d.reply, d);
  });
}
function selectOf(options, value) {
  var s = document.createElement("select"); s.style.cssText = "width:100%;font-size:12px;padding:3px;margin:2px 0";
  options.forEach(function (o) { var op = document.createElement("option"); op.value = o.v; op.textContent = o.t; s.appendChild(op); });
  if (value !== undefined) s.value = value;
  return s;
}
function labelRow(text, ctrl) {
  var d = mk("div", "margin:4px 0"); d.appendChild(mk("div", "font-size:11.5px;color:#57606a", text)); d.appendChild(ctrl); return d;
}
function sheetNames(cb) {
  Excel.run(function (ctx) { var s = ctx.workbook.worksheets; s.load("items/name"); return ctx.sync().then(function () { return s.items.map(function (x) { return x.name; }).filter(function (n) { return n.indexOf("~Lưu") !== 0 && n !== "_tinh_tam"; }); }); })
    .then(function (n) { cb(null, n); }, function (e) { cb(String(e && e.message || e)); });
}
function guessCol(headers, re) { for (var i = 0; i < headers.length; i++) if (re.test(headers[i])) return i; return -1; }
function endBusy(bot) { setStatus(bot, ""); busy(false); }

/* ============ Ý 1: gộp nhiều file thành Master ============ */
function wizardMerge() {
  showTab("chat");
  if (state.busy) return;
  var bot = makeBot();
  bot.msg.textContent = "📥 Gộp nhiều file thành một bảng chuẩn (Master). Chọn các file báo cáo (xlsx, xls, csv). Trợ lý tự bỏ ô gộp, dòng trống, dòng tổng, tự ghép các tên cột khác nhau (vd LCB = Lương cơ bản) rồi gộp vào sheet mới. Chị duyệt bảng ghép trước khi gộp.";
  var fin = document.createElement("input"); fin.type = "file"; fin.multiple = true; fin.accept = ".xlsx,.xlsm,.xls,.csv"; fin.style.cssText = "margin-top:6px;font-size:12px";
  var tg = document.createElement("input"); tg.type = "text"; tg.placeholder = "Cột chuẩn muốn có (cách nhau dấu phẩy). Bỏ trống để trợ lý tự đề xuất"; tg.style.cssText = "margin-top:6px;font-size:12px";
  var l1 = mk("label", "font-size:12px;display:block;margin-top:4px"), c1 = document.createElement("input"); c1.type = "checkbox"; c1.style.width = "auto"; l1.appendChild(c1); l1.appendChild(document.createTextNode(" Chỉ lấy sheet đầu tiên của mỗi file"));
  var l2 = mk("label", "font-size:12px;display:block;margin-top:2px"), c2 = document.createElement("input"); c2.type = "checkbox"; c2.style.width = "auto"; l2.appendChild(c2); l2.appendChild(document.createTextNode(" Bỏ các dòng trùng hoàn toàn"));
  var go = mk("button", "margin-top:6px", "Đọc file và ghép cột"); go.className = "pri";
  bot.extra.appendChild(fin); bot.extra.appendChild(tg); bot.extra.appendChild(l1); bot.extra.appendChild(l2); bot.extra.appendChild(go);
  go.onclick = function () {
    if (state.busy) return;
    if (!fin.files || !fin.files.length) { noteLine(bot, "Chị chọn ít nhất một file trước.", "#c00"); return; }
    busy(true); go.disabled = true;
    loadFiles(Array.prototype.slice.call(fin.files), { firstOnly: c1.checked }, function (t) { setStatus(bot, "⏳ " + t); }, function (err, res) {
      if (err) { noteLine(bot, err, "#c00"); endBusy(bot); go.disabled = false; return; }
      var n = res.notes;
      if (!res.tables.length) { noteLine(bot, "Không thấy bảng dữ liệu nào trong các file đã chọn." + (n.errors.length ? " Lỗi: " + n.errors.join("; ") : ""), "#c00"); endBusy(bot); go.disabled = false; return; }
      noteLine(bot, "Đã đọc " + n.files + " file, " + n.sheets + " bảng. Tự xử lý: sửa " + n.merges + " vùng ô gộp, bỏ " + n.blank + " dòng trống, " + n.total + " dòng tổng cộng, " + n.repeat + " dòng tiêu đề lặp." + (n.skipped.length ? " Bỏ qua " + n.skipped.length + " sheet không có bảng." : "") + (n.errors.length ? " Không đọc được: " + n.errors.join("; ").substring(0, 200) : ""));
      var hs = collectHeaders(res.tables);
      var targets = tg.value.split(",").map(function (x) { return x.replace(/^\s+|\s+$/g, ""); }).filter(function (x) { return x; });
      setStatus(bot, "⏳ Đang ghép tên cột bằng trợ lý…");
      var payload = { targets: targets, headers: hs.slice(0, 150).map(function (h) { return { name: h.name, files: h.count, sample: DANGER_HEADER.test(h.name) ? [] : h.sample.map(function (x) { return String(x).substring(0, 30); }) }; }) };
      aiTask("map", "Ghép tên cột", JSON.stringify(payload), function (e2, reply) {
        var mj = e2 ? null : jsonOf(reply), canonical = [], amap = {};
        if (mj && mj.canonical) { canonical = mj.canonical.map(String); amap = mj.map || {}; }
        else {
          if (e2) noteLine(bot, "Trợ lý ghép tên cột chưa phản hồi (" + e2 + "), tạm ghép theo tên giống nhau.", "#b36b00");
          canonical = targets.length ? targets : hs.map(function (h) { return h.name; });
        }
        if (targets.length) canonical = targets;
        endBusy(bot);
        reviewMapping(bot, res.tables, hs, canonical, amap, c2.checked);
      });
    });
  };
}
function reviewMapping(bot, tables, hs, canonical, amap, dedupe) {
  var box = mk("div", "margin-top:8px;padding:6px 8px;background:#fff;border:1px solid #d8dee4;border-radius:10px");
  box.appendChild(mk("div", "font-weight:600;font-size:12.5px;margin-bottom:4px", "Bảng ghép cột (chị chỉnh nếu cần)"));
  var cn = {}; canonical.forEach(function (c) { cn[normName(c)] = c; });
  var opts = canonical.map(function (c) { return { v: "c:" + c, t: c }; });
  var sels = [];
  hs.slice(0, 150).forEach(function (h) {
    var row = mk("div", "padding:3px 0;border-bottom:1px solid #f0f2f5");
    row.appendChild(mk("div", "font-size:12px", h.name + "  (" + h.count + " file)"));
    var o = opts.concat([{ v: "self", t: "(giữ riêng cột này)" }, { v: "drop", t: "(bỏ cột này)" }]);
    var s = selectOf(o);
    var t = amap[h.name], nk = t ? normName(t) : normName(h.name);
    if (t && cn[normName(t)]) s.value = "c:" + cn[normName(t)]; else if (cn[normName(h.name)]) s.value = "c:" + cn[normName(h.name)]; else s.value = "self";
    row.appendChild(s); box.appendChild(row); sels.push({ h: h, s: s });
  });
  var go = mk("button", "margin-top:8px", "Gộp thành sheet Master"); go.className = "pri";
  box.appendChild(go); bot.extra.appendChild(box); toBottom();
  go.onclick = function () {
    if (state.busy) return;
    var mapping = {}, cols = [], used = {};
    sels.forEach(function (x) {
      var v = x.s.value, key = x.h.key;
      if (v === "drop") mapping[key] = "";
      else if (v === "self") { mapping[key] = x.h.name; if (!used[normName(x.h.name)]) { used[normName(x.h.name)] = 1; } }
      else mapping[key] = v.substring(2);
    });
    canonical.forEach(function (c) { var any = false; for (var k in mapping) if (mapping[k] === c) any = true; if (any && !used["c:" + c]) { used["c:" + c] = 1; cols.push(c); } });
    sels.forEach(function (x) { if (mapping[x.h.key] === x.h.name && x.s.value === "self" && cols.indexOf(x.h.name) < 0) cols.push(x.h.name); });
    if (!cols.length) { noteLine(bot, "Chưa chọn cột nào để gộp.", "#c00"); return; }
    go.disabled = true; busy(true); setStatus(bot, "⏳ Đang gộp dữ liệu…");
    var m = buildMaster(tables, mapping, cols), rows = m.rows, dups = 0;
    if (dedupe) {
      var seen = {}, keep = [];
      rows.forEach(function (r) { var k = JSON.stringify(r.slice(0, cols.length)); if (seen[k]) dups++; else { seen[k] = 1; keep.push(r); } });
      rows = keep;
    }
    var heads = cols.concat(["Nguồn file", "Sheet nguồn"]);
    var nf = rows.map(function (r) { return r.map(function () { return "General"; }); });
    var plan = rows.length ? cleanPlan(rows, nf, heads) : { out: rows, nfset: [], changed: { trim: 0, num: 0, date: 0 } };
    var dateCols = {};
    plan.nfset.forEach(function (r) { r.forEach(function (x, c) { if (x === "dd/mm/yyyy") dateCols[c] = 1; }); });
    var name = "Master";
    Excel.run(function (ctx) {
      var all = ctx.workbook.worksheets; all.load("items/name");
      return ctx.sync().then(function () {
        var exists = all.items.some(function (x) { return x.name === name; });
        if (exists) name = uniqueName("Master");
        var ws = ctx.workbook.worksheets.add(name);
        return ctx.sync().then(function () { return writeGrid(ctx, ws, 0, 0, [heads].concat(plan.out)); }).then(function () {
          for (var c in dateCols) ws.getRangeByIndexes(1, Number(c), Math.max(1, plan.out.length), 1).numberFormat = "dd/mm/yyyy";
          ws.activate();
          return ctx.sync();
        });
      });
    }).then(function () {
      smartFormat(function (err, sum, undoE) {
        pushUndo([{ t: "addsheet", name: name }]);
        endBusy(bot);
        var ch = plan.changed || { trim: 0, num: 0, date: 0 };
        bot.msg.textContent = "✔ Đã gộp " + fmt(plan.out.length) + " dòng từ " + tables.length + " bảng vào sheet \"" + name + "\" (" + cols.length + " cột + cột Nguồn file)." + (dups ? " Bỏ " + dups + " dòng trùng." : "") + (ch.num || ch.date || ch.trim ? " Đã chuẩn hóa " + (ch.num || 0) + " ô chữ→số, " + (ch.date || 0) + " ô chữ→ngày, " + (ch.trim || 0) + " ô khoảng trắng." : "");
        noteLine(bot, "Không ưng thì bấm ↶ Hoàn tác (xóa sheet Master). File gốc không bị đụng tới.");
        var sg = mk("div", "margin-top:6px"); sg.className = "chips";
        ["Soát lỗi và số liệu bất thường trong sheet này", "Tóm tắt nhanh sheet này", "Tổng hợp theo từng chi nhánh/phòng ban"].forEach(function (t) {
          var b = mk("button", "", t.substring(0, 40)); b.onclick = function () { if (state.busy) return; showTab("chat"); $("q").value = t; send(false); }; sg.appendChild(b);
        });
        bot.extra.appendChild(sg); toBottom();
      }, { sheet: name });
    }, function (e) { endBusy(bot); noteLine(bot, "Chưa gộp được: " + String(e && e.message || e), "#c00"); go.disabled = false; });
  };
}

/* ============ Ý 7: lấy dữ liệu từ file khác về bảng đang mở ============ */
function wizardImport() {
  showTab("chat");
  if (state.busy) return;
  var bot = makeBot();
  bot.msg.textContent = "🔗 Lấy dữ liệu từ một file khác. Chọn file, rồi chọn cột khóa chung (vd Mã NV) để điền thêm các cột từ file đó vào bên phải bảng đang mở. Không đè lên dữ liệu có sẵn. Hoặc nhập nguyên một sheet của file đó thành sheet mới.";
  var fin = document.createElement("input"); fin.type = "file"; fin.accept = ".xlsx,.xlsm,.xls,.csv"; fin.style.cssText = "margin-top:6px;font-size:12px";
  var go = mk("button", "margin-top:6px", "Đọc file"); go.className = "pri";
  bot.extra.appendChild(fin); bot.extra.appendChild(go);
  go.onclick = function () {
    if (state.busy) return;
    if (!fin.files || !fin.files.length) { noteLine(bot, "Chị chọn file trước.", "#c00"); return; }
    busy(true); go.disabled = true;
    loadFiles([fin.files[0]], { firstOnly: false }, function (t) { setStatus(bot, "⏳ " + t); }, function (err, res) {
      if (err || !res.tables.length) { noteLine(bot, err || "Không thấy bảng dữ liệu trong file này.", "#c00"); endBusy(bot); go.disabled = false; return; }
      readSheetTable(null, null, function (e2, cur) {
        endBusy(bot);
        if (e2) { noteLine(bot, "Không đọc được bảng đang mở: " + e2, "#c00"); go.disabled = false; return; }
        importForm(bot, res.tables, cur, fin.files[0].name);
      });
    });
  };
}
function importForm(bot, tables, cur, fname) {
  var box = mk("div", "margin-top:8px;padding:6px 8px;background:#fff;border:1px solid #d8dee4;border-radius:10px");
  var sheetSel = selectOf(tables.map(function (t, i) { return { v: String(i), t: t.sheet + " (" + t.rows.length + " dòng)" }; }));
  var curKey = selectOf(cur.t.headers.map(function (h, i) { return { v: String(i), t: h }; }));
  var fileKey = document.createElement("div");
  var listDiv = mk("div", "margin:4px 0;max-height:150px;overflow:auto");
  var fk = null, checks = [];
  var guess = guessCol(cur.t.headers, /^(mã|id|stt|số)|mã nv|mã nhân/i); if (guess >= 0) curKey.value = String(guess);
  function fill() {
    var t = tables[Number(sheetSel.value)];
    fileKey.innerHTML = ""; listDiv.innerHTML = ""; checks = [];
    fk = selectOf(t.headers.map(function (h, i) { return { v: String(i), t: h }; }));
    var g2 = -1, cname = normName(cur.t.headers[Number(curKey.value)]);
    t.headers.forEach(function (h, i) { if (g2 < 0 && normName(h) === cname) g2 = i; });
    if (g2 < 0) g2 = guessCol(t.headers, /^(mã|id)|mã nv|mã nhân/i);
    if (g2 >= 0) fk.value = String(g2);
    fileKey.appendChild(labelRow("Cột khóa trong file đó", fk));
    t.headers.forEach(function (h, i) {
      var lb = mk("label", "display:block;font-size:12px"), cb = document.createElement("input"); cb.type = "checkbox"; cb.style.width = "auto"; cb.checked = true;
      lb.appendChild(cb); lb.appendChild(document.createTextNode(" " + h)); listDiv.appendChild(lb); checks.push(cb);
    });
  }
  sheetSel.onchange = fill; curKey.onchange = fill; fill();
  box.appendChild(labelRow("Sheet trong file \"" + fname + "\"", sheetSel));
  box.appendChild(labelRow("Cột khóa trong bảng đang mở (" + cur.name + ")", curKey));
  box.appendChild(fileKey);
  box.appendChild(mk("div", "font-size:11.5px;color:#57606a", "Các cột muốn lấy về:"));
  box.appendChild(listDiv);
  var b1 = mk("button", "margin-top:6px", "Điền vào bên phải bảng đang mở"); b1.className = "pri";
  var b2 = mk("button", "margin-top:6px", "Nhập nguyên sheet này thành sheet mới");
  box.appendChild(b1); box.appendChild(b2); bot.extra.appendChild(box); toBottom();
  b2.onclick = function () {
    if (state.busy) return;
    var t = tables[Number(sheetSel.value)], name = uniqueName("Nhập " + t.sheet);
    busy(true); setStatus(bot, "⏳ Đang nhập…");
    Excel.run(function (ctx) {
      var ws = ctx.workbook.worksheets.add(name);
      return ctx.sync().then(function () { return writeGrid(ctx, ws, 0, 0, [t.headers].concat(t.rows)); }).then(function () { ws.activate(); return ctx.sync(); });
    }).then(function () {
      smartFormat(function () { pushUndo([{ t: "addsheet", name: name }]); endBusy(bot); noteLine(bot, "✔ Đã nhập " + t.rows.length + " dòng vào sheet \"" + name + "\". Không ưng thì bấm ↶ Hoàn tác."); }, { sheet: name });
    }, function (e) { endBusy(bot); noteLine(bot, "Chưa nhập được: " + String(e && e.message || e), "#c00"); });
  };
  b1.onclick = function () {
    if (state.busy) return;
    var t = tables[Number(sheetSel.value)], ck = Number(curKey.value), fki = Number(fk.value), pick = [];
    checks.forEach(function (c, i) { if (c.checked && i !== fki) pick.push(i); });
    if (!pick.length) { noteLine(bot, "Chưa chọn cột nào để lấy về.", "#c00"); return; }
    var dict = {}; t.rows.forEach(function (r) { var k = lookupKey(r[fki]); if (k && !dict.hasOwnProperty(k)) dict[k] = r; });
    var rows = cur.t.rows, nos = cur.t.rowNos;
    if (!rows.length) { noteLine(bot, "Bảng đang mở không có dòng dữ liệu.", "#c00"); return; }
    var first = nos[0], last = nos[nos.length - 1], arr = [], miss = [], hit = 0, i, j;
    for (i = 0; i < last - first + 1; i++) { var e = []; for (j = 0; j < pick.length; j++) e.push(""); arr.push(e); }
    rows.forEach(function (r, k) {
      var key = lookupKey(r[ck]); if (!key) return;
      var src = dict[key], at = nos[k] - first;
      if (!src) { miss.push({ ref: cur.name + "!" + colLetter(cur.c0 + cur.t.colIdx[ck]) + (cur.r0 + nos[k]), text: "Không tìm thấy \"" + String(r[ck]) + "\" trong file " + fname }); return; }
      hit++; for (j = 0; j < pick.length; j++) arr[at][j] = isBlank(src[pick[j]]) ? "" : src[pick[j]];
    });
    var hdrRow = cur.r0 + cur.t.titleRows, c0 = cur.c0 + cur.cols, r0 = cur.r0 + first;
    busy(true); b1.disabled = true; setStatus(bot, "⏳ Đang điền…");
    Excel.run(function (ctx) {
      var ws = ctx.workbook.worksheets.getItem(cur.name);
      var area = ws.getRangeByIndexes(hdrRow, c0, last - first + 1 + (r0 - hdrRow), pick.length);
      return snapRange(ctx, area, false).then(function (sn) {
        var hdrCells = ws.getRangeByIndexes(hdrRow, c0, 1, pick.length);
        hdrCells.values = [pick.map(function (p) { return t.headers[p]; })];
        var body = ws.getRangeByIndexes(r0, c0, arr.length, pick.length);
        var grid = arr.map(function (r) { return r.map(escCell); });
        body.values = grid;
        try { hdrCells.copyFrom(ws.getRangeByIndexes(hdrRow, c0 - 1, 1, 1), "Formats"); } catch (e1) {}
        return ctx.sync().then(function () {
          try { ws.getRangeByIndexes(hdrRow, c0, 1, pick.length).format.autofitColumns(); return ctx.sync(); } catch (e2) { return null; }
        }).then(function () { return sn; });
      });
    }).then(function (sn) {
      pushUndo([{ t: "cells", snap: sn }]);
      endBusy(bot);
      noteLine(bot, "✔ Đã điền " + pick.length + " cột mới vào bên phải bảng. Khớp " + hit + "/" + rows.length + " dòng." + (miss.length ? " " + miss.length + " dòng không tìm thấy trong file." : "") + " Không ưng thì bấm ↶ Hoàn tác.");
      if (miss.length) renderIssues(bot, miss, "");
    }, function (e) { endBusy(bot); b1.disabled = false; noteLine(bot, "Chưa điền được: " + String(e && e.message || e), "#c00"); });
  };
}

/* ============ Ý 2: soát công thức, số gõ đè, số cứng, vòng lặp ============ */
function wizardAudit() {
  showTab("chat");
  if (state.busy) return;
  busy(true);
  var bot = makeBot(); bot.msg.textContent = "🛡 Soát công thức toàn file";
  setStatus(bot, "⏳ Đang đọc các sheet…");
  Excel.run(function (ctx) {
    var sheets = ctx.workbook.worksheets; sheets.load("items/name,items/visibility");
    return ctx.sync().then(function () {
      var jobs = [];
      sheets.items.forEach(function (ws) {
        if (ws.name.indexOf("~Lưu") === 0 || ws.name === "_tinh_tam" || ws.visibility !== "Visible") return;
        var u = ws.getUsedRangeOrNullObject(); u.load("rowCount,columnCount,rowIndex,columnIndex");
        jobs.push({ name: ws.name, u: u });
      });
      return ctx.sync().then(function () {
        jobs.forEach(function (j) {
          if (j.u.isNullObject) return;
          var cols = j.u.columnCount, rows = Math.max(1, Math.min(j.u.rowCount, Math.floor(150000 / cols)));
          j.cut = rows < j.u.rowCount;
          j.rg = j.u.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
          j.rg.load("formulasR1C1,formulas,values");
        });
        return ctx.sync().then(function () { return jobs; });
      });
    });
  }).then(function (jobs) {
    var all = [], nSheets = 0, cut = [];
    jobs.forEach(function (j) {
      if (j.u.isNullObject) return; nSheets++;
      if (j.cut) cut.push(j.name);
      all = all.concat(auditSheet(j.name, j.rg.formulasR1C1, j.rg.formulas, j.rg.values, j.u.rowIndex, j.u.columnIndex));
    });
    endBusy(bot);
    var by = {}; all.forEach(function (i) { by[i.type] = (by[i.type] || 0) + 1; });
    var parts = Object.keys(by).map(function (k) { return by[k] + " " + k.toLowerCase(); });
    bot.msg.textContent = all.length ? "Đã soát " + nSheets + " sheet, thấy " + all.length + " chỗ cần xem: " + parts.join(", ") + "." : "Đã soát " + nSheets + " sheet: không thấy số gõ đè, công thức lệch, vòng lặp hay số cứng đáng ngờ. ✔";
    if (cut.length) noteLine(bot, "Sheet lớn nên chỉ soát phần đầu: " + cut.join(", "), "#b36b00");
    noteLine(bot, "Việc soát chạy trên máy, không tốn hạn mức và không đổi dữ liệu.");
    if (all.length) {
      var order = { "Số gõ đè": 0, "Công thức vòng": 1, "Tổng thiếu dòng": 2, "Lỗi công thức": 3, "Công thức lệch": 4, "Số cứng trong công thức": 5 };
      all.sort(function (a, b) { return (order[a.type] === undefined ? 9 : order[a.type]) - (order[b.type] === undefined ? 9 : order[b.type]); });
      renderIssues(bot, all, "", { fill: "#FFC7CE", restore: true });
    }
  }, function (e) { endBusy(bot); bot.msg.textContent = "Chưa soát được: " + String(e && e.message || e); });
}

/* ============ Ý 3 và 5: bất thường & chênh lệch ============ */
function wizardPeriod() {
  showTab("chat");
  if (state.busy) return;
  busy(true);
  var bot = makeBot(); bot.msg.textContent = "🚩 Đối chiếu và tìm bất thường / giải thích chênh lệch quỹ. Dữ liệu ở dạng mỗi dòng là một người trong một kỳ (tháng). Chị chọn cột tương ứng:";
  setStatus(bot, "⏳ Đang đọc bảng đang mở…");
  readSheetTable(null, null, function (err, cur) {
    if (err) { endBusy(bot); noteLine(bot, "Không đọc được: " + err, "#c00"); return; }
    sheetNames(function (e2, names) {
      endBusy(bot);
      periodForm(bot, cur, names || []);
    });
  });
}
function periodForm(bot, cur, names) {
  var H = cur.t.headers, none = { v: "-1", t: "(không có)" };
  function opts(opt) { var o = H.map(function (h, i) { return { v: String(i), t: h }; }); return opt ? [none].concat(o) : o; }
  var box = mk("div", "margin-top:6px;padding:6px 8px;background:#fff;border:1px solid #d8dee4;border-radius:10px");
  var sKey = selectOf(opts(false)), sPer = selectOf(opts(true)), sAmt = selectOf(opts(false)), sBase = selectOf(opts(true)), sGrp = selectOf(opts(true));
  function g(re, sel, optional) { var i = guessCol(H, re); if (i >= 0) sel.value = String(i); else if (optional) sel.value = "-1"; }
  g(/mã nv|mã nhân|mã|họ tên|tên|nhân viên|nhân sự/i, sKey); g(/tháng|kỳ|ngày|period|date/i, sPer, true);
  g(/hoa hồng|thưởng|thực lĩnh|tổng lương|lương|số tiền|amount/i, sAmt); g(/doanh (thu|số)|chỉ tiêu|cơ sở|base/i, sBase, true); g(/vùng|khu vực|phòng|chi nhánh|nhóm|region/i, sGrp, true);
  box.appendChild(labelRow("Cột người / mã nhân sự", sKey)); box.appendChild(labelRow("Cột kỳ (tháng) — bỏ trống nếu chỉ có một kỳ", sPer));
  box.appendChild(labelRow("Cột số tiền cần soát (hoa hồng, thưởng, lương…)", sAmt)); box.appendChild(labelRow("Cột cơ sở tính, vd doanh thu (không bắt buộc)", sBase)); box.appendChild(labelRow("Cột nhóm, vd vùng/phòng ban (không bắt buộc)", sGrp));
  var up = document.createElement("input"); up.type = "number"; up.value = "100"; up.style.cssText = "width:70px;font-size:12px"; var dn = document.createElement("input"); dn.type = "number"; dn.value = "60"; dn.style.cssText = "width:70px;font-size:12px";
  var th = mk("div", "font-size:12px;margin:6px 0"); th.appendChild(document.createTextNode("Báo khi tăng hơn ")); th.appendChild(up); th.appendChild(document.createTextNode(" % hoặc giảm hơn ")); th.appendChild(dn); th.appendChild(document.createTextNode(" % so với trung bình các kỳ trước")); box.appendChild(th);
  var rsSheet = selectOf([{ v: "", t: "(không có)" }].concat(names.map(function (n) { return { v: n, t: n }; }))), rsKey = selectOf([]), rsStat = selectOf([]), rsRe = document.createElement("input"); rsRe.type = "text"; rsRe.value = "nghỉ|thôi việc|resign|đã nghỉ"; rsRe.style.cssText = "font-size:12px";
  var rsBox = mk("div", ""); var rsT = null;
  rsSheet.onchange = function () {
    rsBox.innerHTML = ""; rsT = null;
    if (!rsSheet.value) return;
    readSheetTable(rsSheet.value, null, function (e, t) {
      if (e) { rsBox.appendChild(mk("div", "color:#c00;font-size:12px", e)); return; }
      rsT = t;
      rsKey = selectOf(t.t.headers.map(function (h, i) { return { v: String(i), t: h }; }));
      rsStat = selectOf([{ v: "-1", t: "(cả danh sách là người đã nghỉ)" }].concat(t.t.headers.map(function (h, i) { return { v: String(i), t: h }; })));
      var gi = guessCol(t.t.headers, /mã|tên/i); if (gi >= 0) rsKey.value = String(gi);
      var gs = guessCol(t.t.headers, /trạng thái|tình trạng|status/i); if (gs >= 0) rsStat.value = String(gs);
      rsBox.appendChild(labelRow("Cột mã/tên người trong danh sách", rsKey)); rsBox.appendChild(labelRow("Cột trạng thái", rsStat)); rsBox.appendChild(labelRow("Giá trị trạng thái coi là đã nghỉ (cách nhau dấu |)", rsRe));
    });
  };
  box.appendChild(labelRow("Danh sách người đã nghỉ việc (không bắt buộc): chọn sheet", rsSheet)); box.appendChild(rsBox);
  var b1 = mk("button", "margin-top:8px", "🚩 Tìm bất thường"); b1.className = "pri"; var b2 = mk("button", "margin-top:8px", "📈 Giải thích chênh lệch quỹ");
  box.appendChild(b1); box.appendChild(b2); bot.extra.appendChild(box); toBottom();
  function idx() { return { key: Number(sKey.value), period: Number(sPer.value), amount: Number(sAmt.value), base: Number(sBase.value), group: Number(sGrp.value) }; }
  function rowsAndNos() { return { rows: cur.t.rows, nos: cur.t.rowNos.map(function (n) { return cur.r0 + n; }) }; }
  function colRef(c) { return colLetter(cur.c0 + cur.t.colIdx[c]); }
  b1.onclick = function () {
    if (state.busy) return;
    var ix = idx(), rn = rowsAndNos(), resigned = null;
    if (rsSheet.value && rsT) {
      resigned = {}; var kc = Number(rsKey.value), sc = Number(rsStat.value), re = new RegExp(rsRe.value || "nghỉ", "i");
      rsT.t.rows.forEach(function (r) { if (sc < 0 || re.test(String(r[sc]))) { var k = lookupKey(r[kc]); if (k) resigned[k] = 1; } });
    }
    var res = findAnomalies(rn.rows, rn.nos, ix, { upPct: Number(up.value) || 100, downPct: Number(dn.value) || 60, resigned: resigned });
    var items = res.flags.map(function (f) { return { ref: cur.name + "!" + colRef(ix.amount) + f.rowNo, type: f.type, text: f.text, fix: "" }; });
    var msg = "Đã đối chiếu " + res.people + " người, kỳ mới nhất: " + res.latest + (res.periods.length > 1 ? " (so với " + (res.periods.length - 1) + " kỳ trước)" : " (chỉ có một kỳ nên chỉ so giữa các người với nhau)") + ". ";
    var m2 = mk("div", "margin-top:6px;font-size:12.5px", msg + (items.length ? "Có " + items.length + " dòng cần kiểm tra. Đây là cảnh báo để chị xem lại, chưa phải kết luận sai phạm." : "Không thấy bất thường theo các quy tắc đã chọn. ✔"));
    bot.extra.appendChild(m2);
    if (items.length) renderIssues(bot, items, "", { fill: "#FFC7CE" });
    toBottom();
  };
  b2.onclick = function () {
    if (state.busy) return;
    var ix = idx(); if (ix.period < 0) { noteLine(bot, "Để so sánh chênh lệch cần có cột kỳ (tháng).", "#c00"); return; }
    var rn = rowsAndNos(), S = groupSeries(rn.rows, rn.nos, ix);
    if (S.order.length < 2) { noteLine(bot, "Dữ liệu chỉ có một kỳ nên chưa so sánh được.", "#c00"); return; }
    var A = S.order[S.order.length - 2], B = S.order[S.order.length - 1];
    var sa = selectOf(S.order.map(function (p) { return { v: p, t: p }; }), A), sb = selectOf(S.order.map(function (p) { return { v: p, t: p }; }), B);
    var pb = mk("div", "margin-top:6px;padding:6px 8px;background:#f6f8fa;border-radius:8px");
    pb.appendChild(labelRow("So sánh kỳ", sa)); pb.appendChild(labelRow("với kỳ", sb));
    var go = mk("button", "", "Phân tích"); go.className = "pri"; pb.appendChild(go); bot.extra.appendChild(pb); toBottom();
    go.onclick = function () {
      if (state.busy) return;
      go.disabled = true; busy(true); setStatus(bot, "⏳ Đang tính…");
      var v = computeVariance(rn.rows, rn.nos, ix, sa.value, sb.value);
      writeVarianceSheet(v, ix, cur, function (err, name) {
        if (err) { endBusy(bot); noteLine(bot, "Chưa ghi được bảng phân tích: " + err, "#c00"); go.disabled = false; return; }
        pushUndo([{ t: "addsheet", name: name }]);
        var ctxText = varianceContext(v);
        setStatus(bot, "⏳ Trợ lý đang viết nhận xét…");
        aiTask("narrate", "Giải thích vì sao quỹ thay đổi giữa hai kỳ", ctxText, function (e3, reply) {
          endBusy(bot);
          var local = "Quỹ " + v.B + " " + (v.delta >= 0 ? "tăng " : "giảm ") + money(Math.abs(v.delta)) + (v.pct !== null ? " (" + Math.abs(v.pct).toFixed(1) + "%)" : "") + " so với " + v.A + ". Bảng chi tiết ở sheet \"" + name + "\".";
          var mm = mk("div", "margin-top:6px;font-size:12.5px;white-space:pre-wrap", e3 ? local + "\n(Trợ lý chưa viết được nhận xét: " + e3 + ")" : local + "\n\n" + String(reply).replace(/^\s+|\s+$/g, ""));
          bot.extra.appendChild(mm);
          noteLine(bot, "Phần bóc tách chạy trên máy, chỉ số tổng hợp theo nhóm được gửi cho trợ lý để viết nhận xét; tên người và số của từng người không rời khỏi máy chị. Nguyên nhân do chính sách hay quy định chỉ nêu được khi chị cho biết thêm.");
          toBottom();
        });
      });
    };
  };
}
function varianceContext(v) {
  var t = "Kỳ trước: " + v.A + ", kỳ sau: " + v.B + "\nTổng kỳ trước: " + money(v.sumA) + " (" + v.nA + " người)\nTổng kỳ sau: " + money(v.sumB) + " (" + v.nB + " người)\nChênh lệch: " + money(v.delta) + (v.pct !== null ? " (" + v.pct.toFixed(1) + "%)" : "") + "\n";
  t += "Người mới có khoản (kỳ trước không có): " + v.joinedN + " người, đóng góp " + money(v.joined) + "\nNgười không còn khoản: " + v.leftN + " người, đóng góp " + money(v.left) + "\nNgười ở lại cả hai kỳ: " + v.stayN + " người, thay đổi " + money(v.stay) + "\n";
  if (v.hasBase) t += "Trong phần người ở lại: do thay đổi cơ sở tính (doanh thu/chỉ tiêu) " + money(v.vol) + ", do thay đổi tỷ lệ " + money(v.rate) + ", khác " + money(v.other) + "\n";
  var gs = Object.keys(v.groups).map(function (k) { return { k: k, a: v.groups[k].a, b: v.groups[k].b, d: v.groups[k].b - v.groups[k].a }; }).sort(function (x, y) { return Math.abs(y.d) - Math.abs(x.d); }).slice(0, 8);
  t += "Theo nhóm (chênh lệch lớn nhất):\n" + gs.map(function (g) { return "- " + g.k + ": " + money(g.a) + " → " + money(g.b) + " (" + (g.d >= 0 ? "+" : "") + money(g.d) + ")"; }).join("\n");
  var top = v.people.slice(0, 5), topSum = top.reduce(function (s, x) { return s + x.d; }, 0);
  t += "\n5 người thay đổi nhiều nhất cộng lại: " + money(topSum) + " (" + (v.delta ? Math.round(topSum / v.delta * 100) : 0) + "% tổng chênh lệch)";
  return t;
}
function writeVarianceSheet(v, ix, cur, cb) {
  var name = uniqueName("Phân tích chênh lệch");
  var pct = function (x) { return v.delta ? x / v.delta : ""; };
  var blocks = [], rows;
  blocks.push({ title: "Tổng quan", h: ["Chỉ tiêu", "Giá trị"], r: [["Tổng kỳ " + v.A, v.sumA], ["Tổng kỳ " + v.B, v.sumB], ["Chênh lệch", v.delta], ["% chênh lệch", v.pct === null ? "" : v.pct / 100], ["Số người kỳ " + v.A, v.nA], ["Số người kỳ " + v.B, v.nB]], nf: { 1: "#,##0" }, special: { 3: "0.0%" } });
  rows = [["Người mới có khoản (" + v.joinedN + " người)", v.joined, pct(v.joined)], ["Người không còn khoản (" + v.leftN + " người)", v.left, pct(v.left)]];
  if (v.hasBase) { rows.push(["Người ở lại: do đổi cơ sở tính (doanh thu/chỉ tiêu)", v.vol, pct(v.vol)]); rows.push(["Người ở lại: do đổi tỷ lệ", v.rate, pct(v.rate)]); rows.push(["Người ở lại: khác", v.other, pct(v.other)]); }
  else rows.push(["Người ở lại: mức thay đổi", v.stay, pct(v.stay)]);
  rows.push(["Cộng", v.delta, v.delta ? 1 : ""]);
  blocks.push({ title: "Nguyên nhân chênh lệch", h: ["Yếu tố", "Số tiền", "% đóng góp"], r: rows, nf: { 1: "#,##0", 2: "0.0%" }, chart: true });
  var gs = Object.keys(v.groups).map(function (k) { return { k: k, a: v.groups[k].a, b: v.groups[k].b }; }).sort(function (x, y) { return Math.abs(y.b - y.a) - Math.abs(x.b - x.a); }).slice(0, 15);
  if (gs.length > 1 || (gs.length === 1 && gs[0].k !== "(không có nhóm)")) blocks.push({ title: "Theo nhóm", h: ["Nhóm", "Kỳ " + v.A, "Kỳ " + v.B, "Chênh lệch", "% đóng góp"], r: gs.map(function (g) { return [g.k, g.a, g.b, g.b - g.a, pct(g.b - g.a)]; }), nf: { 1: "#,##0", 2: "#,##0", 3: "#,##0", 4: "0.0%" } });
  blocks.push({ title: "10 người thay đổi nhiều nhất", h: ["Người", "Kỳ " + v.A, "Kỳ " + v.B, "Chênh lệch"], r: v.people.slice(0, 10).map(function (p) { return [p.label, p.a, p.b, p.d]; }), nf: { 1: "#,##0", 2: "#,##0", 3: "#,##0" } });
  Excel.run(function (ctx) {
    var ws = ctx.workbook.worksheets.add(name); ws.showGridlines = false;
    var r = 0, chartRange = null;
    ws.getRange("A1").values = [["Phân tích chênh lệch quỹ: kỳ " + v.B + " so với kỳ " + v.A]];
    ws.getRange("A1").format.font.bold = true; ws.getRange("A1").format.font.size = 15; ws.getRange("A1").format.font.color = "#1F4E78";
    r = 2;
    blocks.forEach(function (b) {
      var tc = ws.getRangeByIndexes(r, 0, 1, 1); tc.values = [[b.title]]; tc.format.font.bold = true; tc.format.font.color = "#1F4E78";
      var hr = ws.getRangeByIndexes(r + 1, 0, 1, b.h.length); hr.values = [b.h];
      hr.format.fill.color = "#1F4E78"; hr.format.font.color = "#FFFFFF"; hr.format.font.bold = true; hr.format.horizontalAlignment = "Center";
      if (b.r.length) {
        var body = ws.getRangeByIndexes(r + 2, 0, b.r.length, b.h.length); body.values = b.r.map(function (x) { return x.map(escCell); });
        body.format.borders.getItem("InsideHorizontal").style = "Continuous"; body.format.borders.getItem("InsideHorizontal").color = "#D9DEE5";
        body.format.borders.getItem("EdgeBottom").style = "Continuous"; body.format.borders.getItem("EdgeBottom").color = "#B7C0CB";
        for (var c in b.nf) ws.getRangeByIndexes(r + 2, Number(c), b.r.length, 1).numberFormat = b.nf[c];
        if (b.special) for (var sr in b.special) ws.getRangeByIndexes(r + 2 + Number(sr), 1, 1, 1).numberFormat = b.special[sr];
        if (b.title === "Nguyên nhân chênh lệch") { var lastRow = ws.getRangeByIndexes(r + 1 + b.r.length, 0, 1, b.h.length); lastRow.format.font.bold = true; lastRow.format.fill.color = "#EAF1F8"; }
        if (b.chart) chartRange = { r: r + 1, n: b.r.length - 1 };
      }
      r += 2 + b.r.length + 1;
    });
    ws.getRange("A:A").format.columnWidth = 300; ws.getRange("B:E").format.columnWidth = 110;
    if (chartRange) {
      var ch = ws.charts.add("BarClustered", ws.getRangeByIndexes(chartRange.r, 0, chartRange.n + 1, 2), "Columns");
      ch.title.text = "Đóng góp vào chênh lệch quỹ"; ch.legend.visible = false;
      ch.setPosition("G3", "O20");
    }
    ws.activate();
    return ctx.sync();
  }).then(function () { cb(null, name); }, function (e) { cb(String(e && e.message || e)); });
}

/* ============ Ý 4: truy vấn bằng ngôn ngữ thường (do trợ lý mô tả, máy thực hiện) ============ */
function runQueryOp(a, base, cb) {
  var src = a.sheet ? String(a.sheet) : base;
  readSheetTable(src, a.range, function (err, cur) {
    if (err) { cb(err); return; }
    var res;
    try { res = runQuery(cur.t.headers, cur.t.rows, a); } catch (e) { cb(String(e && e.message || e)); return; }
    if (!res.rows.length) { cb(null, "Truy vấn: không có dòng nào thỏa điều kiện (đã xét " + cur.t.rows.length + " dòng)", []); return; }
    var name = uniqueName((a.out && a.out.sheet) || "Kết quả");
    Excel.run(function (ctx) {
      var ws = ctx.workbook.worksheets.add(name);
      return ctx.sync().then(function () { return writeGrid(ctx, ws, 0, 0, [res.headers].concat(res.rows)); }).then(function () { ws.activate(); return ctx.sync(); });
    }).then(function () {
      smartFormat(function () {
        var chartType = a.out && a.out.chart;
        function done() { cb(null, "Truy vấn → sheet \"" + name + "\": " + res.rows.length + (res.total > res.rows.length ? "/" + res.total : "") + " dòng (từ " + res.matched + " dòng thỏa điều kiện)", [{ t: "addsheet", name: name }]); }
        if (!chartType || !CHART_TYPES[chartType] || res.rows.length < 2) { done(); return; }
        var gN = (a.group_by && a.group_by.length) || 1, vN = 1;
        if (a.pivot) { vN = res.headers.length - gN; if (vN > 3) vN = Math.min(3, (a.pivot.values && a.pivot.values.length) || 2); }
        var colsN = Math.min(res.headers.length, gN + vN), rowsN = Math.min(res.rows.length, 30) + 1;
        Excel.run(function (ctx) {
          var ws = ctx.workbook.worksheets.getItem(name);
          var ch = ws.charts.add(chartType, ws.getRangeByIndexes(0, gN - 1, rowsN, colsN - gN + 1), "Columns");
          ch.title.text = String((a.out && a.out.title) || "Kết quả truy vấn"); ch.setPosition(ws.getCell(1, colsN + 1), ws.getCell(18, colsN + 8));
          return ctx.sync();
        }).then(done, done);
      }, { sheet: name });
    }, function (e) { cb(String(e && e.message || e)); });
  });
}

/* ============ Ghi nhiều bảng nhỏ vào một sheet báo cáo ============ */
function writeBlocksSheet(base, title, blocks, cb) {
  var name = uniqueName(base);
  Excel.run(function (ctx) {
    var ws = ctx.workbook.worksheets.add(name); ws.showGridlines = false;
    ws.getRange("A1").values = [[title]];
    ws.getRange("A1").format.font.bold = true; ws.getRange("A1").format.font.size = 15; ws.getRange("A1").format.font.color = "#1F4E78";
    var r = 2;
    blocks.forEach(function (b) {
      var tc = ws.getRangeByIndexes(r, 0, 1, 1); tc.values = [[b.title]]; tc.format.font.bold = true; tc.format.font.color = "#1F4E78";
      var hr = ws.getRangeByIndexes(r + 1, 0, 1, b.h.length); hr.values = [b.h];
      hr.format.fill.color = "#1F4E78"; hr.format.font.color = "#FFFFFF"; hr.format.font.bold = true; hr.format.horizontalAlignment = "Center";
      if (b.r.length) {
        var body = ws.getRangeByIndexes(r + 2, 0, b.r.length, b.h.length); body.values = b.r.map(function (x) { return x.map(escCell); });
        body.format.borders.getItem("InsideHorizontal").style = "Continuous"; body.format.borders.getItem("InsideHorizontal").color = "#D9DEE5";
        body.format.borders.getItem("EdgeBottom").style = "Continuous"; body.format.borders.getItem("EdgeBottom").color = "#B7C0CB";
        for (var c in (b.nf || {})) ws.getRangeByIndexes(r + 2, Number(c), b.r.length, 1).numberFormat = b.nf[c];
      }
      r += 2 + b.r.length + 1;
    });
    ws.getRange("A:A").format.columnWidth = 260; ws.getRange("B:H").format.columnWidth = 120;
    ws.activate();
    return ctx.sync();
  }).then(function () { cb(null, name); }, function (e) { cb(String(e && e.message || e)); });
}

/* ============ Soát bảng lương một kỳ (HR trưởng phòng nhận file lương) ============ */
function wizardPayroll() {
  showTab("chat");
  if (state.busy) return;
  busy(true);
  var bot = makeBot(); bot.msg.textContent = "👥 Soát bảng lương";
  setStatus(bot, "⏳ Đang đọc bảng lương…");
  readSheetTable(null, null, function (err, cur) {
    if (err) { endBusy(bot); bot.msg.textContent = "Chưa đọc được: " + err; return; }
    var res;
    try { res = payrollCheck(cur.t.headers, cur.t.rows, cur.t.rowNos); } catch (e) { endBusy(bot); bot.msg.textContent = "Chưa soát được: " + String(e && e.message || e); return; }
    endBusy(bot);
    var S = res.summary, det = res.det, H = cur.t.headers;
    function nm(k) { return det[k] >= 0 ? H[det[k]] : null; }
    var found = [["Mã", "key"], ["Họ tên", "name"], ["Phòng ban", "dept"], ["Lương cơ bản", "base"], ["Phụ cấp", "allow"], ["Ngày công", "days"], ["KPI", "kpi"], ["Thưởng", "bonus"], ["Tổng thu nhập", "gross"], ["Thực lĩnh", "net"]];
    var okC = [], miss = [];
    found.forEach(function (f) { if (nm(f[1])) okC.push(f[0] + " = \"" + nm(f[1]) + "\""); else miss.push(f[0]); });
    var by = {}; res.flags.forEach(function (f) { by[f.type] = (by[f.type] || 0) + 1; });
    var parts = Object.keys(by).map(function (k) { return by[k] + " " + k.toLowerCase(); });
    bot.msg.textContent = "Đã soát " + S.n + " dòng lương (sheet " + cur.name + "). " + (res.flags.length ? "Có " + res.flags.length + " điểm cần kiểm tra: " + parts.join(", ") + "." : "Không thấy điểm bất thường theo các quy tắc. ✔");
    noteLine(bot, "Nhận diện cột: " + okC.join("; ") + (miss.length ? ". Không thấy cột: " + miss.join(", ") + " (bỏ qua phần kiểm tra liên quan)." : "."));
    var kp = [];
    kp.push({ label: "Số nhân sự", value: fmt(S.n) });
    if (S.grossSum !== null) kp.push({ label: "Tổng thu nhập", value: money(S.grossSum) + " đ" });
    if (S.netSum !== null) kp.push({ label: "Tổng thực lĩnh", value: money(S.netSum) + " đ" });
    if (S.netAvg !== null) kp.push({ label: "Thực lĩnh bình quân", value: money(S.netAvg) + " đ" });
    if (S.daysAvg !== null) kp.push({ label: "Ngày công bình quân", value: S.daysAvg.toFixed(1) });
    if (S.kpiAvg !== null) kp.push({ label: "KPI bình quân", value: S.kpiAvg.toFixed(0) + "%" });
    renderKpis(bot, kp);
    if (S.relations.length) noteLine(bot, "Công thức chung của bảng (suy ra từ số đông): " + S.relations.join("; ") + ". Dòng nào khác công thức này được báo là \"Khác công thức chung\".");
    if (S.dept.length > 1) {
      var tb = mk("div", "margin:8px 0 2px;padding:6px 8px;background:#fff;border:1px solid #d8dee4;border-radius:10px;font-size:12px");
      tb.appendChild(mk("div", "font-weight:600;margin-bottom:3px", "Theo phòng ban"));
      S.dept.slice(0, 8).forEach(function (d) { tb.appendChild(mk("div", "padding:2px 0;border-bottom:1px solid #f0f2f5", d.name + ": " + d.n + " người" + (S.netSum !== null ? ", thực lĩnh " + money(d.net) : (S.grossSum !== null ? ", thu nhập " + money(d.gross) : "")))); });
      bot.extra.appendChild(tb);
    }
    if (res.flags.length) {
      var order = { "Trùng mã": 0, "Khác công thức chung": 1, "Thực lĩnh lớn hơn thu nhập": 2, "Số âm": 2, "Ngày công": 3, "KPI": 4, "Thưởng lệch KPI": 5, "Cao bất thường": 6, "Thấp bất thường": 7, "Tỷ lệ thực lĩnh lạ": 8, "Thiếu thông tin": 9 };
      var items = res.flags.slice(0).sort(function (a, b) { return (order[a.type] === undefined ? 9 : order[a.type]) - (order[b.type] === undefined ? 9 : order[b.type]); }).map(function (f) {
        return { ref: cur.name + "!" + colLetter(cur.c0 + cur.t.colIdx[f.col]) + (cur.r0 + f.rowNo), type: f.type, text: f.who + ": " + f.text, fix: "" };
      });
      renderIssues(bot, items, "", { fill: "#FFC7CE" });
      noteLine(bot, "Đây là cảnh báo để chị kiểm tra, chưa phải kết luận sai phạm. Việc soát chạy hoàn toàn trên máy, không tốn hạn mức, không gửi số liệu đi đâu.");
    }
    var sb = mk("button", "margin-top:6px", "📋 Tạo sheet Tóm tắt lương");
    sb.onclick = function () {
      if (state.busy) return;
      busy(true); sb.disabled = true; setStatus(bot, "⏳ Đang tạo sheet…");
      var blocks = [], ov = [["Số nhân sự", S.n]];
      if (S.baseSum !== null) ov.push(["Tổng lương cơ bản", S.baseSum]);
      if (S.grossSum !== null) ov.push(["Tổng thu nhập", S.grossSum]);
      if (S.netSum !== null) ov.push(["Tổng thực lĩnh", S.netSum]);
      if (S.netAvg !== null) ov.push(["Thực lĩnh bình quân", S.netAvg]);
      if (S.daysAvg !== null) ov.push(["Ngày công bình quân", S.daysAvg]);
      if (S.kpiAvg !== null) ov.push(["KPI bình quân (%)", S.kpiAvg]);
      blocks.push({ title: "Tổng quan", h: ["Chỉ tiêu", "Giá trị"], r: ov, nf: { 1: "#,##0.0" } });
      if (S.dept.length > 1) blocks.push({ title: "Theo phòng ban", h: ["Phòng ban", "Số người", "Tổng thu nhập", "Tổng thực lĩnh", "Bình quân thực lĩnh"], r: S.dept.map(function (d) { return [d.name, d.n, d.gross, d.net, d.n ? d.net / d.n : 0]; }), nf: { 2: "#,##0", 3: "#,##0", 4: "#,##0" } });
      blocks.push({ title: "Cảnh báo theo loại", h: ["Loại", "Số điểm"], r: Object.keys(by).length ? Object.keys(by).map(function (k) { return [k, by[k]]; }) : [["Không có cảnh báo", 0]] });
      writeBlocksSheet("Tóm tắt lương", "Tóm tắt bảng lương (" + cur.name + ")", blocks, function (e2, name) {
        endBusy(bot);
        if (e2) { noteLine(bot, "Chưa tạo được: " + e2, "#c00"); sb.disabled = false; return; }
        pushUndo([{ t: "addsheet", name: name }]);
        noteLine(bot, "✔ Đã tạo sheet \"" + name + "\". Không ưng thì bấm ↶ Hoàn tác.");
      });
    };
    bot.extra.appendChild(sb); toBottom();
  });
}

/* ============ Sinh nhật và thâm niên ============ */
function wizardHR() {
  showTab("chat");
  if (state.busy) return;
  busy(true);
  var bot = makeBot(); bot.msg.textContent = "🎂 Sinh nhật và kỷ niệm ngày vào làm. Chọn sheet danh sách nhân sự (có cột ngày sinh, ngày vào làm):";
  setStatus(bot, "⏳ Đang đọc danh sách sheet…");
  sheetNames(function (e, names) {
    endBusy(bot);
    if (e) { noteLine(bot, e, "#c00"); return; }
    var box = mk("div", "margin-top:6px;padding:6px 8px;background:#fff;border:1px solid #d8dee4;border-radius:10px");
    var sh = selectOf(names.map(function (n) { return { v: n, t: n }; }));
    var ctrlDiv = mk("div", ""), cur = null, sels = {};
    var ms = document.createElement("input"); ms.type = "text"; ms.value = "5, 10, 15, 20, 25, 30"; ms.style.cssText = "font-size:12px";
    var sd = document.createElement("input"); sd.type = "number"; sd.value = "7"; sd.style.cssText = "width:70px;font-size:12px";
    var go = mk("button", "margin-top:8px", "Xem ngay"); go.className = "pri";
    function load() {
      ctrlDiv.innerHTML = ""; cur = null; go.disabled = true;
      readSheetTable(sh.value, null, function (e2, t) {
        if (e2) { ctrlDiv.appendChild(mk("div", "color:#c00;font-size:12px", e2)); return; }
        cur = t; go.disabled = false;
        var H = t.t.headers;
        function opt(none) { var o = H.map(function (h, i) { return { v: String(i), t: h }; }); return none ? [{ v: "-1", t: "(không có)" }].concat(o) : o; }
        function pick(re, none) { var s = selectOf(opt(none)); var i = guessCol(H.map(normName), re); if (i >= 0) s.value = String(i); else if (none) s.value = "-1"; return s; }
        sels = { name: pick(/ho ten|ho va ten|ten nhan vien|^ten\b/, false), dept: pick(/phong|bo phan|chi nhanh|don vi/, true), birth: pick(/ngay sinh|sinh nhat|\bdob\b|birth/, false), join: pick(/ngay vao|vao lam|bat dau|ngay nhan viec|join|ngay ky hd/, false), status: pick(/trang thai|tinh trang|status/, true) };
        ctrlDiv.appendChild(labelRow("Cột họ tên", sels.name)); ctrlDiv.appendChild(labelRow("Cột phòng ban (không bắt buộc)", sels.dept));
        ctrlDiv.appendChild(labelRow("Cột ngày sinh", sels.birth)); ctrlDiv.appendChild(labelRow("Cột ngày vào làm", sels.join)); ctrlDiv.appendChild(labelRow("Cột trạng thái (để bỏ người đã nghỉ, không bắt buộc)", sels.status));
      });
    }
    sh.onchange = load;
    var cur0 = null;
    box.appendChild(labelRow("Sheet danh sách nhân sự", sh)); box.appendChild(ctrlDiv);
    box.appendChild(labelRow("Báo kỷ niệm khi tròn bao nhiêu năm (cách nhau dấu phẩy)", ms));
    var sdRow = mk("div", "font-size:12px;margin:4px 0"); sdRow.appendChild(document.createTextNode("Sinh nhật sắp tới trong ")); sdRow.appendChild(sd); sdRow.appendChild(document.createTextNode(" ngày")); box.appendChild(sdRow);
    box.appendChild(go); bot.extra.appendChild(box); toBottom();
    try { Excel.run(function (ctx) { var a = ctx.workbook.worksheets.getActiveWorksheet(); a.load("name"); return ctx.sync().then(function () { return a.name; }); }).then(function (n) { if (names.indexOf(n) >= 0) sh.value = n; load(); }, load); } catch (e3) { load(); }
    go.onclick = function () {
      if (!cur || state.busy) return;
      var now = new Date(), today = { y: now.getFullYear(), m: now.getMonth() + 1, d: now.getDate() };
      var idx = { name: Number(sels.name.value), dept: Number(sels.dept.value), birth: Number(sels.birth.value), join: Number(sels.join.value), status: Number(sels.status.value) };
      var miles = ms.value.split(/[,\s;]+/).map(Number).filter(function (x) { return x > 0; });
      var ev = hrEvents(cur.t.rows, cur.t.rowNos, idx, today, { milestones: miles, soonDays: Number(sd.value) || 7 });
      var items = [], out = mk("div", "margin-top:8px");
      function ref(x, col) { return cur.name + "!" + colLetter(cur.c0 + cur.t.colIdx[col]) + (cur.r0 + x.rowNo); }
      function section(title, list, fmtLine, type) {
        var b = mk("div", "margin:6px 0 2px;padding:6px 8px;background:#fff;border:1px solid #d8dee4;border-radius:10px;font-size:12.5px");
        b.appendChild(mk("div", "font-weight:600;margin-bottom:3px", title + " (" + list.length + ")"));
        if (!list.length) b.appendChild(mk("div", "color:#8c959f", "Không có"));
        list.slice(0, 25).forEach(function (x) {
          var line = mk("div", "padding:2px 0;border-bottom:1px solid #f0f2f5"), bt = mk("button", "padding:0 6px;margin-right:6px", x.name);
          bt.onclick = function () { gotoRef(ref(x, x.nameCol)); };
          line.appendChild(bt); line.appendChild(document.createTextNode(fmtLine(x)));
          b.appendChild(line);
          items.push({ a: ref(x, x.nameCol), type: type, t: x.name + (x.dept ? " (" + x.dept + ")" : "") + " – " + fmtLine(x).replace(/^\s*[–-]\s*/, ""), fix: "" });
        });
        if (list.length > 25) b.appendChild(mk("div", "color:#8c959f;font-size:11px", "… và " + (list.length - 25) + " người nữa (xuất ra sheet để xem đủ)."));
        out.appendChild(b);
      }
      var dateLabel = twoD(today.d) + "/" + twoD(today.m) + "/" + today.y;
      var bm = ev.bMonth.filter(function (x) { return x.days !== 0; });
      section("🎂 Sinh nhật HÔM NAY (" + dateLabel + ")", ev.bToday, function (x) { return " – " + (x.dept ? x.dept + " – " : "") + "tròn " + x.age + " tuổi"; }, "Sinh nhật hôm nay");
      section("🎈 Sinh nhật " + (Number(sd.value) || 7) + " ngày tới", ev.bSoon, function (x) { return " – " + x.label + (x.dept ? " – " + x.dept : "") + " (" + x.age + " tuổi)"; }, "Sinh nhật sắp tới");
      section("📅 Tất cả sinh nhật trong tháng " + today.m, bm, function (x) { return " – " + x.label + (x.dept ? " – " + x.dept : "") + (x.days < 0 ? " (đã qua)" : ""); }, "Sinh nhật trong tháng");
      section("🏅 Kỷ niệm ngày vào làm trong tháng " + today.m + " (mốc " + miles.join(", ") + " năm)", ev.aMonth, function (x) { return " – " + x.years + " năm (vào làm " + x.label + ")" + (x.dept ? " – " + x.dept : "") + (x.days === 0 ? " – HÔM NAY" : ""); }, "Kỷ niệm vào làm");
      var warn = [];
      if (ev.noBirth) warn.push(ev.noBirth + " người thiếu hoặc sai định dạng ngày sinh");
      if (ev.noJoin) warn.push(ev.noJoin + " người thiếu hoặc sai ngày vào làm");
      if (ev.skipped) warn.push("đã bỏ " + ev.skipped + " người đã nghỉ");
      if (ev.aOther) warn.push(ev.aOther + " người có kỷ niệm tròn năm khác trong tháng (không thuộc mốc đã chọn)");
      if (warn.length) out.appendChild(mk("div", "font-size:11.5px;color:#b36b00;margin-top:4px", "Lưu ý: " + warn.join("; ") + "."));
      bot.extra.appendChild(out);
      if (items.length) { var bx = mk("div", "margin-top:4px"); issueButtons(bx, items, "", { fill: "#C6EFCE" }); bot.extra.appendChild(bx); }
      noteLine(bot, "Tính trên máy theo ngày hôm nay của máy chị, không tốn hạn mức.");
      toBottom();
    };
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
    } else if (e.t === "cf") wb.worksheets.getItem(e.sheet).getRange(e.addr).conditionalFormats.clearAll();
    else if (e.t === "validation") wb.worksheets.getItem(e.sheet).getRange(e.addr).dataValidation.clear();
    else if (e.t === "marks") {
      var mws = wb.worksheets.getItem(e.sheet);
      e.items.forEach(function (it) { var c = mws.getRange(it.a); if (!it.color || it.color === "#FFFFFF") c.format.fill.clear(); else c.format.fill.color = it.color; });
    } else if (e.t === "grid") wb.worksheets.getItem(e.sheet).showGridlines = e.was;
    else if (e.t === "smartfmt") restoreSmart(ctx, e);
    else if (e.t === "freeze") wb.worksheets.getItem(e.sheet).freezePanes.unfreeze();
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
      /* báo cho AI biết lần trước đã bị hủy, tránh nó tưởng vẫn còn */
      if (state.hist.length) { var lh = state.hist[state.hist.length - 1]; if (!/^\[ĐÃ HOÀN TÁC\]/.test(lh.a)) lh.a = "[ĐÃ HOÀN TÁC - không còn trong file] " + lh.a; }
      state.hist.push({ q: "(chị bấm Hoàn tác)", a: "Mọi thay đổi của lần làm liền trước, kể cả sheet đã tạo, đã bị XÓA khỏi file. Những gì đã nói là đã tạo thì hiện KHÔNG còn. Chỉ tin NGỮ CẢNH hiện tại; nếu chị yêu cầu lại thì phải làm lại từ đầu (add_sheet rồi ghi vào sheet mới)." });
      if (state.hist.length > 4) state.hist.shift();
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
        return { name: sh.name, v: rg.values, f: rg.formulasR1C1, r0: used.rowIndex, c0: used.columnIndex, cut: maxRows < used.rowCount };
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
    if (items.length) issueButtons(out, items.slice(0, 74), d.name);
    busy(false);
  }, function (e) { $("scanOut").textContent = "Không quét được: " + String(e && e.message || e); busy(false); });
}

/* ---------- Khởi động ---------- */
/* Excel 2016 cũ (ExcelApi dưới 1.4) không có các hàm ...OrNullObject: bù bằng hàm cũ tương đương */
var LEGACY = false;
function installCompatShim() {
  try {
    if (typeof Excel === "undefined") return;
    /* Excel 2016 rất cũ (ExcelApi 1.1-1.2) không có getResizedRange: ghép từ getOffsetRange + getBoundingRect (chỉ dùng cho ô đơn) */
    if (Excel.Range && typeof Excel.Range.prototype.getResizedRange !== "function")
      Excel.Range.prototype.getResizedRange = function (dr, dc) { return this.getBoundingRect(this.getOffsetRange(dr, dc)); };
    if (supports("1.4")) return;
    LEGACY = true;
    var wrap = function (o) {
      var orig = o.load; o.isNullObject = false;
      o.load = function (p) {
        if (typeof p === "string") p = p.split(",").map(function (x) { return x.replace(/\s/g, ""); }).filter(function (x) { return x && x !== "isNullObject"; }).join(",");
        return p ? orig.call(o, p) : o;
      };
      return o;
    };
    if (Excel.Worksheet && !Excel.Worksheet.prototype.getUsedRangeOrNullObject) Excel.Worksheet.prototype.getUsedRangeOrNullObject = function () { return wrap(this.getUsedRange(true)); };
    if (Excel.Range && !Excel.Range.prototype.getIntersectionOrNullObject) Excel.Range.prototype.getIntersectionOrNullObject = function (r) { return wrap(this.getIntersection(r)); };
    if (Excel.WorksheetCollection && !Excel.WorksheetCollection.prototype.getItemOrNullObject) Excel.WorksheetCollection.prototype.getItemOrNullObject = function (n) { return wrap(this.getItem(n)); };
  } catch (e) {}
}

Office.onReady(function () {
  installCompatShim();
  state.code = store("code") || "";
  var au = store("auto");
  $("auto").checked = au !== "0";
  var ed = store("edit"); $("edit").checked = ed !== "0";
  $("edit").onchange = function () { store("edit", this.checked ? "1" : "0"); };
  var vf = store("verify"); $("verify").checked = vf !== "0";
  $("verify").onchange = function () { store("verify", this.checked ? "1" : "0"); };
  $("auto").onchange = function () { store("auto", this.checked ? "1" : "0"); };
  $("loginBtn").onclick = function () {
    state.code = $("codeIn").value.replace(/\s/g, "");
    store("code", state.code);
    refreshMe();
  };
  $("sendBtn").onclick = function () { send(false); };
  $("tabChat").onclick = function () { showTab("chat"); };
  $("tabTools").onclick = function () { showTab("tools"); };
  var toolMap = { payroll: wizardPayroll, hr: wizardHR, merge: wizardMerge, importf: wizardImport, audit: wizardAudit, period: wizardPeriod };
  var tbs = document.querySelectorAll("button[data-tool]");
  for (var ti = 0; ti < tbs.length; ti++) tbs[ti].onclick = function () { if (state.busy) return; var f = toolMap[this.getAttribute("data-tool")]; if (f) f(); };
  var tqs = document.querySelectorAll("button[data-ask]");
  for (var tj = 0; tj < tqs.length; tj++) tqs[tj].onclick = function () { if (state.busy) return; showTab("chat"); $("q").value = this.getAttribute("data-ask"); send(false); };
  $("stopBtn").onclick = function () { if (state.cur) state.cur.abort(); };
  $("gearBtn").onclick = function () { var pn = $("panel"); pn.className = pn.className === "hide" ? "" : "hide"; };
  $("newBtn").onclick = function () { if (state.busy) return; $("log").innerHTML = ""; $("welcome").className = "welcome"; state.hist = []; state.lastQ = ""; };
  $("q").onkeydown = function (ev) {
    ev = ev || window.event;
    if ((ev.key === "Enter" || ev.keyCode === 13) && !ev.shiftKey) { if (ev.preventDefault) ev.preventDefault(); if (!state.busy) send(false); }
  };
  $("q").oninput = function () { this.style.height = "auto"; this.style.height = Math.min(120, this.scrollHeight) + "px"; };
  $("scanBtn").onclick = scan;
  $("undoBtn").onclick = undo;
  var chips = document.querySelectorAll("button[data-q]");
  for (var i = 0; i < chips.length; i++) {
    chips[i].onclick = function () {
      if (state.busy) return;
      $("q").value = this.getAttribute("data-q");
      send(false);
    };
  }
  if (state.code) refreshMe(); else $("login").className = "box";
});
