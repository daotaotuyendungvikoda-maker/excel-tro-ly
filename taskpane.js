var WORKER = "https://excel-tro-ly.daotaotuyendungvikoda.workers.dev"; /* <-- SỬA */
var DANGER_HEADER = /(lương|luong|salary|cccd|cmnd|stk|tài khoản|tai khoan|mst|mã số thuế|thưởng|thuong)/i;
var ERR_RE = /^#(REF!|N\/A|DIV\/0!|VALUE!|NAME\?|NUM!|NULL!)/;
var state = { code: "", task: "formula", lastQ: "", undo: [] };

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
function refreshMe() {
  api("/api/me", "GET", null, function (err, d) {
    if (err) {
      $("main").className = "hide"; $("login").className = "box";
      $("loginErr").textContent = state.code ? err : "";
      return;
    }
    $("login").className = "hide"; $("main").className = "";
    updateQuota(d.quota, null);
  });
}

/* ---------- Đọc ngữ cảnh bảng tính (có che dữ liệu nhạy cảm) ---------- */
function getContext(cb) {
  Excel.run(function (ctx) {
    var sel = ctx.workbook.getSelectedRange(); sel.load("address,rowCount,columnCount");
    var sh = ctx.workbook.worksheets.getActiveWorksheet(); sh.load("name");
    var used = sh.getUsedRange(); used.load("rowCount,columnCount");
    return ctx.sync().then(function () {
      var rows = Math.min(used.rowCount, 16), cols = Math.min(used.columnCount, 26);
      var sample = used.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
      sample.load("address,formulas");
      return ctx.sync().then(function () {
        var f = sample.formulas, head = f[0], masked = [];
        for (var c = 0; c < head.length; c++) masked[c] = DANGER_HEADER.test(String(head[c]));
        var lines = [];
        for (var r = 0; r < f.length; r++) {
          var cells = [];
          for (var c2 = 0; c2 < f[r].length; c2++) {
            var v = f[r][c2];
            if (r > 0 && masked[c2]) v = "<ẨN>";
            else if (typeof v === "string" && /^\d{9,16}$/.test(v)) v = "<ẨN>";
            cells.push(colLetter(c2) + "=" + v);
          }
          lines.push("Dòng " + (r + 1) + ": " + cells.join(" | "));
        }
        var text = "Sheet: " + sh.name + "\nVùng đích đang chọn: " + sel.address + " (" + sel.rowCount + " dòng x " + sel.columnCount + " cột)\n" +
          "Mẫu dữ liệu (" + sample.address + ", dòng 1 là tiêu đề; cột lương/CCCD/STK đã bị ẩn):\n" + lines.join("\n");
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
    try { return JSON.parse(s.substring(a, b + 1)); } catch (e) {}
  }
  return { message: String(t || ""), formula: null, warnings: null };
}
function showBot(j, meta) {
  var d = document.createElement("div");
  d.className = "msg bot";
  var p = document.createElement("div"); p.textContent = j.message || ""; d.appendChild(p);
  if (j.formula && String(j.formula).charAt(0) === "=") {
    var fb = document.createElement("div"); fb.className = "formula"; fb.textContent = j.formula; d.appendChild(fb);
    var ab = document.createElement("button"); ab.className = "pri"; ab.textContent = "Áp vào vùng đang chọn";
    ab.onclick = function () { applyFormula(j.formula, d); };
    d.appendChild(ab);
  }
  if (j.warnings) { var w = document.createElement("div"); w.className = "small"; w.textContent = "Lưu ý: " + j.warnings; d.appendChild(w); }
  var again = document.createElement("button"); again.textContent = "Làm kỹ hơn";
  again.onclick = function () { send(true); };
  d.appendChild(again);
  var m = document.createElement("div"); m.className = "meta";
  m.textContent = meta.label + " · " + fmt(meta.tokens_in) + " token vào / " + fmt(meta.tokens_out) + " ra · ~" + fmt(meta.cost_vnd) + "đ · vì: " + meta.reason;
  d.appendChild(m);
  $("log").insertBefore(d, $("log").firstChild);
}
function send(deep) {
  var q = $("q").value.replace(/^\s+|\s+$/g, "");
  if (deep) q = state.lastQ;
  if (!q) return;
  state.lastQ = q;
  if (!deep) addMsg("me", q);
  busy(true);
  getContext(function (err, c) {
    if (err) { busy(false); addMsg("bot", "Không đọc được bảng tính: " + err); return; }
    api("/api/chat", "POST", { task: state.task, question: q, context: c.text, compat: compat(), deep: !!deep }, function (e, d) {
      busy(false);
      if (e) { addMsg("bot", e); refreshMe(); return; }
      showBot(parseReply(d.reply), d.meta);
      updateQuota(d.quota, d.meta);
    });
  });
}

/* ---------- Áp công thức + Hoàn tác ---------- */
function localAddr(a) { var i = a.lastIndexOf("!"); return i >= 0 ? a.substring(i + 1) : a; }
function applyFormula(formula, msgEl) {
  Excel.run(function (ctx) {
    var rng = ctx.workbook.getSelectedRange();
    rng.load("address,rowCount,columnCount,formulas");
    var sh = rng.worksheet; sh.load("name");
    return ctx.sync().then(function () {
      var snap = { sheet: sh.name, addr: localAddr(rng.address), formulas: rng.formulas };
      var first = rng.getCell(0, 0);
      first.formulas = [[formula]];
      first.load("formulasR1C1");
      return ctx.sync().then(function () {
        var r1 = first.formulasR1C1[0][0], arr = [];
        for (var i = 0; i < rng.rowCount; i++) {
          var row = [];
          for (var j = 0; j < rng.columnCount; j++) row.push(r1);
          arr.push(row);
        }
        rng.formulasR1C1 = arr;
        rng.load("values");
        return ctx.sync().then(function () {
          var errs = 0, v = rng.values;
          for (var a = 0; a < v.length; a++) for (var b = 0; b < v[a].length; b++)
            if (typeof v[a][b] === "string" && ERR_RE.test(v[a][b])) errs++;
          state.undo.push(snap);
          if (state.undo.length > 10) state.undo.shift();
          $("undoBtn").disabled = false;
          return { cells: rng.rowCount * rng.columnCount, errs: errs };
        });
      });
    });
  }).then(function (r) {
    var t = "Đã ghi công thức vào " + r.cells + " ô." + (r.errs ? " ⚠ Có " + r.errs + " ô ra lỗi, bấm Hoàn tác nếu không đúng ý." : " Không thấy ô nào báo lỗi.");
    var n = document.createElement("div"); n.className = "small"; n.textContent = t; msgEl.appendChild(n);
  }, function (e) { addMsg("bot", "Không ghi được: " + String(e && e.message || e)); });
}
function undo() {
  var s = state.undo.pop();
  if (!s) return;
  Excel.run(function (ctx) {
    ctx.workbook.worksheets.getItem(s.sheet).getRange(s.addr).formulas = s.formulas;
    return ctx.sync();
  }).then(function () {
    addMsg("bot", "Đã hoàn tác vùng " + s.sheet + "!" + s.addr + ".");
    $("undoBtn").disabled = state.undo.length === 0;
  }, function (e) { addMsg("bot", "Không hoàn tác được: " + String(e && e.message || e)); });
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
  $("loginBtn").onclick = function () {
    state.code = $("codeIn").value.replace(/\s/g, "");
    store("code", state.code);
    refreshMe();
  };
  $("sendBtn").onclick = function () { send(false); };
  $("scanBtn").onclick = scan;
  $("undoBtn").onclick = undo;
  var chips = document.querySelectorAll("button[data-task]");
  for (var i = 0; i < chips.length; i++) {
    chips[i].onclick = function () {
      state.task = this.getAttribute("data-task");
      var h = this.getAttribute("data-hint");
      if (!$("q").value) $("q").value = h;
      $("q").focus();
    };
  }
  if (state.code) refreshMe(); else $("login").className = "box";
});
