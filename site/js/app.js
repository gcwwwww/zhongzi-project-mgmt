/* ============================================================
 * 科研创新小组 · 项目与人员费用管理系统
 * 纯前端单页应用，localStorage 持久化，SheetJS 处理 Excel
 * ============================================================ */
(function () {
  "use strict";

  /* ---------- 常量 ---------- */
  var LS_KEY = "rd_pm_store_v1";
  var WEEK = ["日", "一", "二", "三", "四", "五", "六"];

  /* ---------- 工具 ---------- */
  function el(id) { return document.getElementById(id); }
  function h(tag, attrs, children) {
    var n = document.createElement(tag);
    if (attrs) for (var k in attrs) {
      if (k === "class") n.className = attrs[k];
      else if (k === "html") n.innerHTML = attrs[k];
      else if (k.startsWith("on") && typeof attrs[k] === "function") n.addEventListener(k.slice(2), attrs[k]);
      else if (k === "dataset") for (var d in attrs[k]) n.dataset[d] = attrs[k][d];
      else n.setAttribute(k, attrs[k]);
    }
    if (children != null) {
      (Array.isArray(children) ? children : [children]).forEach(function (c) {
        if (c == null) return;
        n.appendChild(typeof c === "string" ? document.createTextNode(c) : c);
      });
    }
    return n;
  }
  function fmt(n, d) { if (n == null || isNaN(n)) return "—"; return Number(n).toFixed(d == null ? 2 : d); }
  function fmtMoney(n) { if (n == null || isNaN(n)) return "—"; return (Number(n) / 10000).toFixed(2); } // 元 -> 万元
  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  function uid() { return "id" + Math.random().toString(36).slice(2, 9); }
  function monthKey(d) { if (!d) return ""; var p = d.split("-"); return p[0] + "-" + ("0" + p[1]).slice(-2); }
  function daysInMonth(m) { var p = m.split("-"); return new Date(p[0], p[1], 0).getDate(); }
  function weekdayOf(m, d) { var p = m.split("-"); return new Date(p[0], p[1] - 1, d).getDay(); }
  /* 考勤周期：上报月份 M → 考勤月份 = M-1，周期 = (M-2)月26日 ~ (M-1)月25日 */
  function attDates(reportMonth) {
    var p = reportMonth.split("-"); var y = Number(p[0]), m = Number(p[1]);
    var end = new Date(y, m - 2, 25);
    var start = new Date(y, m - 3, 26);
    var arr = []; var cur = new Date(start);
    while (cur.getTime() <= end.getTime()) {
      arr.push({ day: cur.getDate(), month: cur.getMonth() + 1, weekday: cur.getDay() });
      cur.setDate(cur.getDate() + 1);
    }
    return arr;
  }
  function attLabel(reportMonth) {
    var p = reportMonth.split("-"); var y = Number(p[0]), m = Number(p[1]);
    var d = new Date(y, m - 1, 1); d.setMonth(d.getMonth() - 1);
    return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0");
  }
  function attCycleStr(reportMonth) {
    var ds = attDates(reportMonth); if (!ds.length) return "";
    var fmtD = function (o) { return o.month + "/" + o.day; };
    return fmtD(ds[0]) + " ~ " + fmtD(ds[ds.length - 1]);
  }
  /* 2026 年法定节假日（国务院发布，调休补班日另列） */
  var HOLIDAYS_2026 = [
    "2026-01-01",                                     // 元旦
    "2026-02-15", "2026-02-16", "2026-02-17", "2026-02-18", "2026-02-19", "2026-02-20", "2026-02-21", // 春节
    "2026-04-04", "2026-04-05", "2026-04-06",         // 清明
    "2026-05-01", "2026-05-02", "2026-05-03", "2026-05-04", "2026-05-05", // 劳动节
    "2026-06-19", "2026-06-20", "2026-06-21",         // 端午
    "2026-09-25", "2026-09-26", "2026-09-27",         // 中秋
    "2026-10-01", "2026-10-02", "2026-10-03", "2026-10-04", "2026-10-05", "2026-10-06", "2026-10-07"  // 国庆
  ];
  /* 调休补班日（周末但需上班） */
  var WORKDAY_WEEKENDS_2026 = [
    "2026-02-14", "2026-02-28",  // 春节调休
    "2026-04-26",                // 清明/劳动节调休
    "2026-09-28",                // 中秋/国庆调休
    "2026-10-10"                 // 国庆调休
  ];
  function dateKey(y, m, d) { return y + "-" + ("0" + m).slice(-2) + "-" + ("0" + d).slice(-2); }
  function isRestDay(o) {
    // o = { day, month, weekday }，需结合年份判断
    var y = Number(Store.data.settings.currentMonth.split("-")[0]);
    var key = dateKey(y, o.month, o.day);
    if (WORKDAY_WEEKENDS_2026.indexOf(key) >= 0) return false; // 调休补班
    if (HOLIDAYS_2026.indexOf(key) >= 0) return true;          // 法定假日
    return o.weekday === 0 || o.weekday === 6;                  // 周末
  }

  var toastTimer;
  function toast(msg, type) {
    var t = el("toast");
    t.textContent = msg;
    t.className = "toast" + (type ? " " + type : "");
    t.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { t.hidden = true; }, 2200);
  }

  /* ---------- 模态框 ---------- */
  function openModal(title, bodyNode, opts) {
    opts = opts || {};
    el("modalTitle").textContent = title;
    var b = el("modalBody"); b.innerHTML = ""; b.appendChild(bodyNode);
    var foot = el("modalFoot"); foot.innerHTML = "";
    if (opts.foot) opts.foot.forEach(function (btn) {
      var bEl = h("button", { class: "btn " + (btn.kind || ""), onclick: btn.onclick }, btn.text);
      foot.appendChild(bEl);
    });
    var m = el("modal");
    if (opts.wide) m.classList.add("wide"); else m.classList.remove("wide");
    el("modalMask").hidden = false;
  }
  function closeModal() { el("modalMask").hidden = true; }

  /* ---------- 数据层 ---------- */
  var Store = {
    data: null,
    load: function () {
      // 先用 localStorage 即时加载（避免空白等待）
      try {
        var raw = localStorage.getItem(LS_KEY);
        if (raw) { this.data = JSON.parse(raw); }
        else { this.data = this.fromSeed(); localStorage.setItem(LS_KEY, JSON.stringify(this.data)); }
      } catch (e) { this.data = this.fromSeed(); }
    },
    syncFromServer: function (cb) {
      // 从服务器拉取最新共享数据，成功后更新本地缓存并刷新视图
      var self = this;
      fetch("/api/data").then(function (r) { return r.json(); }).then(function (data) {
        if (data && data.projects) {
          // 合并缺失字段
          if (!data.achievements) data.achievements = [];
          if (!data.fundBudgets) data.fundBudgets = {};
          if (!data.fundRecords) data.fundRecords = {};
          if (!data.contacts || !data.contacts.cols) data.contacts = defaultContacts();
          if (!data.settings.signatures) data.settings.signatures = {};
          self.data = data;
          localStorage.setItem(LS_KEY, JSON.stringify(data));
          if (cb) cb(true);
        } else if (cb) cb(false);
      }).catch(function () { if (cb) cb(false); });
    },
    save: function () {
      // 本地立即保存
      try { localStorage.setItem(LS_KEY, JSON.stringify(this.data)); } catch (e) {
        console.warn("本地保存失败：", e);
        toast("保存失败：本地存储已满（可能因签名图片过大）。建议导出 JSON 备份后清理", "err");
      }
      // 异步同步到服务器（多人共享）
      try {
        fetch("/api/data", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(this.data) }).catch(function (e) { console.warn("服务器同步失败：", e); });
      } catch (e) { console.warn("服务器同步异常：", e); }
    },
    fromSeed: function () {
      var s = window.SEED || { projects: [], people: [], companyName: "xxx分公司", currentMonth: "2026-08", dailyRate: 300 };
      var projects = s.projects.map(function (p) {
        return {
        id: uid(), name: p.name, periodStart: p.periodStart, periodEnd: p.periodEnd,
        leader: p.leader, totalBudget: p.totalBudget || "", monthlyReports: {}
      };
      });
      function projId(name) {
        for (var i = 0; i < projects.length; i++) if (projects[i].name === name) return projects[i].id;
        return projects.length ? projects[0].id : null;
      }
      var personnel = s.people.map(function (p) {
        return {
          id: uid(), name: p.name, responsibility: p.responsibility,
          projectId: projId(p.project),
          attendance: {}
        };
      });
      // 8月考勤
      var m = s.currentMonth;
      personnel.forEach(function (pe, idx) {
        var src = s.people[idx];
        pe.attendance[m] = { days: src.days || {} };
      });
      return {
        settings: { companyName: s.companyName, currentMonth: m, dailyRate: s.dailyRate || 300 },
        projects: projects,
        personnel: personnel,
        fundBudgets: {},
        fundRecords: {},
        contacts: defaultContacts(),
        achievements: (s.achievements || []).map(function (a) {
          return { id: uid(), type: a.type || "研究报告", title: a.title || "", authors: a.authors || "", projectId: a.projectId || null, date: a.date || "", description: a.description || "" };
        })
      };
    },
    reset: function () { localStorage.removeItem(LS_KEY); this.data = this.fromSeed(); this.save(); }
  };
  function defaultContacts() {
    return {
      cols: ["序号","创新小组","人员工时表对接人","月度汇报对接人"],
      rows: [
        ["1","公路","周晶","周晶"],
        ["2","材料","田佳磊","田佳磊"],
        ["3","桥梁","汪辛","汪辛"],
        ["4","隧道","高兴","高兴"],
        ["5","交通","苟泉涛","苟泉涛"],
        ["6","桥隧养护","张若男","张若男"]
      ]
    };
  }

  /* ---------- 业务查询 ---------- */
  function projectById(id) { for (var i = 0; i < Store.data.projects.length; i++) if (Store.data.projects[i].id === id) return Store.data.projects[i]; return null; }
  function personById(id) { for (var i = 0; i < Store.data.personnel.length; i++) if (Store.data.personnel[i].id === id) return Store.data.personnel[i]; return null; }
  function peopleInProject(pid) { return Store.data.personnel.filter(function (p) { return p.projectId === pid; }); }

  function ensureMonth(p, m) {
    if (!p.monthlyReports) p.monthlyReports = {};
    if (!p.monthlyReports[m]) p.monthlyReports[m] = { completedWork: "", problems: "", nextPlan: "", monthlyFunds: "", cumulativeFunds: "" };
    return p.monthlyReports[m];
  }
  function ensureAtt(pe, m) {
    if (!pe.attendance) pe.attendance = {};
    if (!pe.attendance[m]) pe.attendance[m] = { days: {} };
    return pe.attendance[m];
  }
  function dayPair(att, d) {
    var o = att.days[String(d)];
    if (!o) { o = { rd: 0, nonRd: 0 }; att.days[String(d)] = o; }
    return o;
  }
  function sumRd(pe, m) {
    var att = ensureAtt(pe, m), s = 0;
    for (var d in att.days) s += Number(att.days[d].rd) || 0;
    return s;
  }
  function sumTotal(pe, m) {
    var att = ensureAtt(pe, m), s = 0;
    for (var d in att.days) s += (Number(att.days[d].rd) || 0) + (Number(att.days[d].nonRd) || 0);
    return s;
  }
  function projectRdSum(pid, m) {
    return peopleInProject(pid).reduce(function (s, pe) { return s + sumRd(pe, m); }, 0);
  }
  function projectPersonCost(pid, m) {
    return projectRdSum(pid, m) * (Store.data.settings.dailyRate || 0);
  }
  function cumulativeFundsUpto(pid, m) {
    var p = projectById(pid); if (!p) return 0;
    var target = m; var sum = 0;
    Object.keys(p.monthlyReports || {}).sort().forEach(function (mk) {
      if (mk <= target) { var v = p.monthlyReports[mk].cumulativeFunds; if (v) sum = Number(v); }
    });
    // 若未填写累计值，则按本月值累加
    if (!sum) {
      Object.keys(p.monthlyReports || {}).sort().forEach(function (mk) {
        if (mk <= target) sum += Number(p.monthlyReports[mk].monthlyFunds) || 0;
      });
    }
    return sum;
  }

  /* ---------- 路由 ---------- */
  var _hist = [], _hi = -1;
  var Router = {
    view: "dashboard",
    params: {},
    go: function (view, params, noPush) {
      this.view = view; this.params = params || {};
      // 记录导航历史（用于前进后退）
      if (!noPush) { _hist = _hist.slice(0, _hi + 1); _hist.push({ view: view, params: params }); _hi++; }
      Array.prototype.forEach.call(el("nav").querySelectorAll(".nav-item"), function (a) {
        a.classList.toggle("active", a.dataset.view === view || (view === "projectDetail" && a.dataset.view === "projects"));
      });
      var titles = {
        dashboard: "概览仪表盘", projects: "项目管理", projectDetail: "项目详情",
        attendance: "人员考勤", monthly: "月度执行情况",
        personnel: "人员台账", achievements: "成果管理", io: "数据导入导出",
        funds: "经费管理", contacts: "创新小组联系人"
      };
      el("viewTitle").textContent = titles[view] || view;
      // 更新前进后退按钮状态
      var btnBack = el("navBack"); var btnFwd = el("navFwd");
      if (btnBack) btnBack.disabled = !this.canGoBack();
      if (btnFwd) btnFwd.disabled = !this.canGoForward();
      el("content").innerHTML = "";
      var node = VIEWS[view] ? VIEWS[view](this.params) : h("div", { class: "empty" }, "未知视图");
      el("content").appendChild(node);
      var sb = el("sidebar"); if (sb) sb.classList.remove("open");
    },
    back: function () { if (_hi > 0) { _hi--; var s = _hist[_hi]; Router.go(s.view, s.params, true); } },
    fwd: function () { if (_hi < _hist.length - 1) { _hi++; var s = _hist[_hi]; Router.go(s.view, s.params, true); } },
    canGoBack: function () { return _hi > 0; },
    canGoForward: function () { return _hi < _hist.length - 1; }
  };

  /* ============================================================
   * 视图
   * ============================================================ */
  var VIEWS = {};

  /* ---- 仪表盘 ---- */
  VIEWS.dashboard = function () {
    var s = Store.data.settings;
    var m = s.currentMonth;
    var nProj = Store.data.projects.length;
    var nPeople = Store.data.personnel.length;
    var rdTotal = Store.data.projects.reduce(function (sum, p) { return sum + projectRdSum(p.id, m); }, 0);
    var costTotal = Store.data.projects.reduce(function (sum, p) { return sum + projectPersonCost(p.id, m); }, 0);
    var fundsTotal = Store.data.projects.reduce(function (sum, p) {
      var r = ensureMonth(p, m); return sum + (Number(r.monthlyFunds) || 0);
    }, 0);

    var root = h("div");
    root.appendChild(h("div", { class: "stat-grid" }, [
      statCard("项目总数", nProj + " 个", "在研科研项目", "📁"),
      statCard("研发人员", nPeople + " 人", "本月参与人数", "👥"),
      statCard("本月研发工时", fmt(rdTotal, 1) + " 天", "所有项目合计", "⏱️"),
      statCard("本月归集经费", fmt(fundsTotal, 2) + " 元", "已填报月度经费", "📈"),
      statCard("知识产权成果", (Store.data.achievements ? Store.data.achievements.length : 0) + " 项", "论文/专利/报告等", "📜")
    ]));

    // 项目进度列表
    var list = h("div", { class: "card" }, [
      h("div", { class: "card-head" }, [h("h2", {}, "项目进度一览（" + m + "）")]),
      h("div", { class: "card-body" })
    ]);
    var body = list.querySelector(".card-body");
    if (!nProj) { body.appendChild(h("div", { class: "empty" }, "暂无项目，请到「项目管理」新增。")); }
    else {
      var wrap = h("div", { class: "table-wrap" });
      var tbl = h("table", { class: "tbl" });
      tbl.innerHTML = "<thead><tr><th>项目名称</th><th>起止周期</th><th>负责人</th><th class='num'>总预算/元</th><th class='num'>本月研发工时</th><th class='num'>本月经费/元</th><th>进度</th><th>操作</th></tr></thead>";
      var tb = h("tbody");
      Store.data.projects.forEach(function (p) {
        var r = ensureMonth(p, m);
        var rd = projectRdSum(p.id, m);
        var pct = progressPct(p, m);
        tb.appendChild(h("tr", {}, [
          h("td", { class: "wrap" }, h("span", { class: "clickable", onclick: function () { Router.go("projectDetail", { id: p.id }); } }, p.name)),
          h("td", {}, (p.periodStart || "") + " ~ " + (p.periodEnd || "")),
          h("td", {}, p.leader || "—"),
          h("td", { class: "num" }, p.totalBudget ? fmt(Number(p.totalBudget) * 10000, 2) : "—"),
          h("td", { class: "num" }, fmt(rd, 1)),
          h("td", { class: "num" }, fmt(r.monthlyFunds || 0, 2)),
          h("td", { style: "min-width:120px" }, h("div", { class: "bar" }, h("span", { style: "width:" + pct + "%" }))),
          h("td", {}, h("button", { class: "btn btn-sm", onclick: function () { Router.go("projectDetail", { id: p.id }); } }, "详情"))
        ]));
      });
      tbl.appendChild(tb); wrap.appendChild(tbl); body.appendChild(wrap);
    }
    root.appendChild(list);

    // 经费趋势（按已填报月份）
    var months = collectMonths();
    if (months.length) {
      var card = h("div", { class: "card" }, [h("div", { class: "card-head" }, [h("h2", {}, "经费归集趋势")]), h("div", { class: "card-body" })]);
      var cw = h("div", { class: "table-wrap" });
      var ct = h("table", { class: "tbl" });
      ct.innerHTML = "<thead><tr><th>月份</th>" + Store.data.projects.map(function (p) { return "<th>" + esc(p.name.slice(0, 8)) + "…</th>"; }).join("") + "<th class='num'>合计/万</th></tr></thead>";
      var ctb = h("tbody");
      months.forEach(function (mk) {
        var cells = Store.data.projects.map(function (p) {
          var r = (p.monthlyReports || {})[mk];
          return r && r.monthlyFunds ? fmt(r.monthlyFunds, 2) : "—";
        });
        var total = Store.data.projects.reduce(function (s, p) {
          var r = (p.monthlyReports || {})[mk]; return s + (r && r.monthlyFunds ? Number(r.monthlyFunds) : 0);
        }, 0);
        ctb.appendChild(h("tr", {}, [h("td", {}, mk)].concat(
          cells.map(function (c) { return h("td", { class: "num" }, c); })
        ).concat([h("td", { class: "num" }, fmt(total, 2))])));
      });
      ct.appendChild(ctb); cw.appendChild(ct); card.querySelector(".card-body").appendChild(cw);
      root.appendChild(h("div", { style: "margin-top:16px" }, card));
    }
    return root;
  };

  function statCard(label, val, sub, icon) {
    return h("div", { class: "stat" }, [
      h("div", { class: "stat-label" }, label),
      h("div", { class: "stat-val" }, val),
      h("div", { class: "stat-sub" }, sub),
      h("div", { class: "stat-icon" }, icon)
    ]);
  }
  function progressPct(p, m) {
    if (!p.periodStart || !p.periodEnd) return 0;
    var s = new Date(p.periodStart + "-01"), e = new Date(p.periodEnd + "-01"), n = new Date(m + "-01");
    if (n <= s) return 0; if (n >= e) return 100;
    return Math.round((n - s) / (e - s) * 100);
  }
  function collectMonths() {
    var set = {};
    Store.data.projects.forEach(function (p) { Object.keys(p.monthlyReports || {}).forEach(function (mk) { set[mk] = 1; }); });
    if (Store.data.settings.currentMonth) set[Store.data.settings.currentMonth] = 1;
    return Object.keys(set).sort();
  }

  /* ---- 项目管理 ---- */
  VIEWS.projects = function () {
    var root = h("div");
    root.appendChild(h("div", { class: "toolbar" }, [
      h("div", { class: "left" }, h("div", { class: "section-title" }, "共 " + Store.data.projects.length + " 个在研项目")),
      h("div", { class: "right" }, h("button", { class: "btn btn-primary", onclick: function () { editProject(null); } }, "+ 新增项目"))
    ]));
    var grid = h("div", { class: "proj-list" });
    if (!Store.data.projects.length) {
      root.appendChild(h("div", { class: "empty card" }, "暂无项目，点击「新增项目」创建。"));
      return root;
    }
    Store.data.projects.forEach(function (p) {
      var m = Store.data.settings.currentMonth;
      var r = ensureMonth(p, m);
      var rd = projectRdSum(p.id, m);
      var pct = progressPct(p, m);
      var nP = peopleInProject(p.id).length;
      var card = h("div", { class: "proj-card", onclick: function () { Router.go("projectDetail", { id: p.id }); } }, [
        h("div", { class: "proj-name" }, p.name),
        h("div", { class: "proj-meta" }, [
          tag("负责人"), h("span", {}, p.leader || "—"),
          tag("周期"), h("span", {}, (p.periodStart || "") + " ~ " + (p.periodEnd || "")),
          tag("总预算"), h("span", {}, p.totalBudget ? fmt(Number(p.totalBudget) * 10000, 2) + " 元" : "—"),
          tag("人员"), h("span", {}, nP + " 人")
        ]),
        h("div", { class: "proj-meta" }, [
          tag("本月工时"), h("span", {}, fmt(rd, 1) + " 天"),
          tag("本月经费"), h("span", {}, fmt(r.monthlyFunds || 0, 2) + " 元")
        ]),
        h("div", { class: "proj-bar" }, [h("div", { class: "stat-label", style: "font-size:11px" }, "项目周期进度 " + pct + "%"), h("div", { class: "bar" }, h("span", { style: "width:" + pct + "%" }))])
      ]);
      grid.appendChild(card);
    });
    root.appendChild(grid);
    return root;
  };
  function tag(t) { return h("span", { class: "tag" }, t); }

  function editProject(pid) {
    var p = pid ? projectById(pid) : { name: "", periodStart: "", periodEnd: "", leader: "", totalBudget: "" };
    var form = h("div", { class: "form-grid" });
    form.appendChild(field("项目名称", h("input", { id: "f_name", value: p.name || "", style: "width:100%" })));
    form.appendChild(field("负责人", h("input", { id: "f_leader", value: p.leader || "", style: "width:100%" })));
    form.appendChild(field("项目总预算（元）", h("input", { type: "number", step: "0.01", min: "0", id: "f_budget", value: p.totalBudget ? (Number(p.totalBudget) * 10000).toString() : "", style: "width:100%", placeholder: "如 3000000" })));
    form.appendChild(field("起始月", h("input", { type: "month", id: "f_ps", value: p.periodStart || "", style: "width:100%" })));
    form.appendChild(field("结束月", h("input", { type: "month", id: "f_pe", value: p.periodEnd || "", style: "width:100%" })));
    openModal(pid ? "编辑项目" : "新增项目", form, { foot: [
      { text: "取消", onclick: closeModal },
      { text: "保存", kind: "btn-primary", onclick: function () {
        var name = el("f_name").value.trim();
        if (!name) { toast("请填写项目名称", "err"); return; }
        var budgetRaw = el("f_budget").value;
        // 用户输入元，存储为万元
        var budget = budgetRaw ? (Number(budgetRaw) / 10000).toString() : "";
        if (pid) { var x = projectById(pid); x.name = name; x.leader = el("f_leader").value.trim(); x.periodStart = el("f_ps").value; x.periodEnd = el("f_pe").value; x.totalBudget = budget; }
        else { Store.data.projects.push({ id: uid(), name: name, leader: el("f_leader").value.trim(), periodStart: el("f_ps").value, periodEnd: el("f_pe").value, totalBudget: budget, monthlyReports: {} }); }
        Store.save(); closeModal(); Router.go(Router.view, Router.params); toast("已保存", "ok");
      } }
    ] });
  }

  /* ---- 项目详情 ---- */
  VIEWS.projectDetail = function (params) {
    var p = projectById(params.id);
    if (!p) return h("div", { class: "empty" }, "项目不存在");
    var m = Store.data.settings.currentMonth;
    var root = h("div");
    root.appendChild(h("div", { class: "card" }, [h("div", { class: "card-body" }, [
      h("div", { class: "row", style: "justify-content:space-between" }, [
        h("div", {}, [h("h2", { style: "margin-bottom:6px" }, p.name), h("div", { class: "text-muted" }, (p.periodStart || "") + " ~ " + (p.periodEnd || "") + " · 负责人：" + (p.leader || "—") + (p.totalBudget ? " · 总预算：" + fmt(Number(p.totalBudget) * 10000, 2) + " 元" : ""))]),
        h("div", {}, [
          h("button", { class: "btn", onclick: function () { editProject(p.id); } }, "编辑项目"),
          h("button", { class: "btn", onclick: function () { editMonthly(p.id, m); } }, "编辑本月执行情况"),
          h("button", { class: "btn btn-danger", onclick: function () { delProject(p.id); } }, "删除项目")
        ])
      ])
    ])]));

    // 本月概要
    var r = ensureMonth(p, m);
    var rd = projectRdSum(p.id, m);
    var cost = projectPersonCost(p.id, m);
    root.appendChild(h("div", { class: "stat-grid" }, [
      statCard("本月研发工时", fmt(rd, 1) + " 天", "项目人员合计"),
      statCard("本月归集经费", fmt(r.monthlyFunds || 0, 2) + " 元", "已填报"),
      statCard("累计归集经费", fmt(cumulativeFundsUpto(p.id, m) || 0, 2) + " 元", "截至本月")
    ]));

    // 本月汇报内容
    var rep = h("div", { class: "card" }, [h("div", { class: "card-head" }, [h("h2", {}, m + " 月度执行情况"), h("button", { class: "btn btn-sm", onclick: function () { editMonthly(p.id, m); } }, "编辑")]), h("div", { class: "card-body" })]);
    var rb = rep.querySelector(".card-body");
    function block(title, val) { return h("div", { style: "margin-bottom:12px" }, [h("div", { class: "section-title" }, title), h("div", { class: "muted-block" }, val || "（未填写）")]); }
    rb.appendChild(block("本月完成工作（含成果产出）", r.completedWork));
    rb.appendChild(block("存在问题", r.problems));
    rb.appendChild(block("下步计划", r.nextPlan));
    var achCount = (Store.data.achievements || []).filter(function(a) { return a.projectId === p.id; }).length;
    rb.appendChild(block("指标完成情况（知识产权）", achCount + " 项成果"));
    // 自动检测问题说明
    var issues = [];
    if (!r.completedWork || !r.completedWork.trim()) issues.push("完成工作未填写");
    if (!r.monthlyFunds || Number(r.monthlyFunds) === 0) issues.push("本月经费未归集");
    if (achCount === 0) issues.push("知识产权无进展");
    rb.appendChild(block("问题说明", issues.length ? issues.join("；") : "正常"));
    root.appendChild(h("div", { style: "margin-top:16px" }, rep));

    // 项目相关成果
    var achs = (Store.data.achievements || []).filter(function (a) { return a.projectId === p.id; });
    var ac = h("div", { class: "card" }, [h("div", { class: "card-head" }, [h("h2", {}, "项目成果（" + achs.length + " 项）"), h("button", { class: "btn btn-sm", onclick: function () { Router.go("achievements"); } }, "去成果管理")]), h("div", { class: "card-body" })]);
    var acb = ac.querySelector(".card-body");
    if (!achs.length) acb.appendChild(h("div", { class: "empty" }, "该项目暂无成果记录"));
    else {
      var aw = h("div", { class: "table-wrap" });
      var at = h("table", { class: "tbl" });
      at.innerHTML = "<thead><tr><th>类型</th><th>名称/标题</th><th>作者</th><th>日期</th></tr></thead>";
      var atb = h("tbody");
      achs.forEach(function (a) {
        atb.appendChild(h("tr", {}, [
          h("td", {}, h("span", { class: "tag " + typeTagClass(a.type) }, a.type)),
          h("td", { class: "wrap" }, a.title),
          h("td", {}, a.authors || "—"),
          h("td", {}, a.date || "—")
        ]));
      });
      at.appendChild(atb); aw.appendChild(at); acb.appendChild(aw);
    }
    root.appendChild(h("div", { style: "margin-top:16px" }, ac));

    // 项目人员
    var people = peopleInProject(p.id);
    var pc = h("div", { class: "card" }, [h("div", { class: "card-head" }, [h("h2", {}, "项目研发人员（" + people.length + "）"), h("button", { class: "btn btn-sm", onclick: function () { editPerson(null, p.id); } }, "+ 添加人员")]), h("div", { class: "card-body" })]);
    var pcb = pc.querySelector(".card-body");
    if (!people.length) pcb.appendChild(h("div", { class: "empty" }, "该项目暂无人员"));
    else {
      var w = h("div", { class: "table-wrap" });
      var t = h("table", { class: "tbl" });
      t.innerHTML = "<thead><tr><th>姓名</th><th>职责分工</th><th class='num'>本月研发工时</th><th class='num'>本月总工时</th><th>操作</th></tr></thead>";
      var tb = h("tbody");
      people.forEach(function (pe) {
        tb.appendChild(h("tr", {}, [
          h("td", {}, pe.name),
          h("td", { class: "wrap" }, pe.responsibility || "—"),
          h("td", { class: "num" }, fmt(sumRd(pe, m), 1)),
          h("td", { class: "num" }, fmt(sumTotal(pe, m), 1)),
          h("td", {}, [
            h("button", { class: "btn btn-sm", onclick: function () { editPerson(pe.id, p.id); } }, "编辑"),
            h("button", { class: "btn btn-sm btn-danger", onclick: function () { delPerson(pe.id); } }, "删除")
          ])
        ]));
      });
      t.appendChild(tb); w.appendChild(t); pcb.appendChild(w);
    }
    root.appendChild(h("div", { style: "margin-top:16px" }, pc));

    // 历史月度执行情况
    var months = Object.keys(p.monthlyReports || {}).sort();
    if (months.length) {
      var hc = h("div", { class: "card" }, [h("div", { class: "card-head" }, [h("h2", {}, "历史月度执行情况")]), h("div", { class: "card-body" })]);
      var hcb = hc.querySelector(".card-body");
      var hw = h("div", { class: "table-wrap" });
      var ht = h("table", { class: "tbl" });
      ht.innerHTML = "<thead><tr><th>月份</th><th>完成工作</th><th>存在问题</th><th>下步计划</th><th>指标完成情况</th><th>问题说明</th><th class='num'>本月经费/元</th><th class='num'>累计经费/元</th><th>操作</th></tr></thead>";
      var htb = h("tbody");
      months.forEach(function (mk) {
        var rr = p.monthlyReports[mk];
        var histAchCount = (Store.data.achievements || []).filter(function(a) { return a.projectId === p.id; }).length;
        var histIssues = [];
        if (!rr.completedWork || !rr.completedWork.trim()) histIssues.push("完成工作未填写");
        if (!rr.monthlyFunds || Number(rr.monthlyFunds) === 0) histIssues.push("本月经费未归集");
        if (histAchCount === 0) histIssues.push("知识产权无进展");
        var histIssueStr = histIssues.length ? histIssues.join("；") : "正常";
        htb.appendChild(h("tr", {}, [
          h("td", {}, mk),
          h("td", { class: "wrap" }, esc(rr.completedWork || "").slice(0, 40) || "—"),
          h("td", { class: "wrap" }, esc(rr.problems || "").slice(0, 40) || "—"),
          h("td", { class: "wrap" }, esc(rr.nextPlan || "").slice(0, 40) || "—"),
          h("td", {}, histAchCount + " 项"),
          h("td", { class: "wrap" }, histIssueStr),
          h("td", { class: "num" }, fmt(rr.monthlyFunds || 0, 2)),
          h("td", { class: "num" }, fmt(rr.cumulativeFunds || 0, 2)),
          h("td", {}, h("button", { class: "btn btn-sm", onclick: function () { editMonthly(p.id, mk); } }, "编辑"))
        ]));
      });
      ht.appendChild(htb); hw.appendChild(ht); hcb.appendChild(hw);
      root.appendChild(h("div", { style: "margin-top:16px" }, hc));
    }
    return root;
  };

  function delProject(pid) {
    confirmModal("删除项目", "确认删除该项目及其月度汇报？项目下人员将转为未分配。", function () {
      Store.data.projects = Store.data.projects.filter(function (p) { return p.id !== pid; });
      Store.data.personnel.forEach(function (pe) { if (pe.projectId === pid) pe.projectId = null; });
      Store.save(); toast("已删除", "ok"); Router.go("projects");
    });
  }

  function editReport(pid, m) { editMonthly(pid, m); }
  function field(label, input) { return h("div", { class: "field" }, [h("label", {}, label), input]); }
  function fieldFull(label, input) { return h("div", { class: "field field-full" }, [h("label", {}, label), input]); }

  /* ---- 人员考勤 ---- */
  VIEWS.attendance = function () {
    var m = Store.data.settings.currentMonth;
    var dates = attDates(m);
    var n = dates.length;
    var pid = Router.params.pid || "all";
    var root = h("div");
    root.appendChild(h("div", { class: "toolbar" }, [
      h("div", { class: "left" }, [
        h("div", { class: "filter" }, [h("span", {}, "项目筛选"), projectSelect(pid, function (v) { Router.params.pid = v; Router.go("attendance", Router.params); })]),
        h("div", { class: "filter" }, [h("span", {}, "上报月份"), h("input", { type: "month", value: m, onchange: function (e) { Store.data.settings.currentMonth = e.target.value; Store.save(); Router.go("attendance"); } })])
      ]),
      h("div", { class: "right" }, [
        h("button", { class: "btn", onclick: function () { batchFillAttendance(m, pid); } }, "批量填充工时"),
        h("button", { class: "btn", onclick: function () { exportAttendance(m, pid); } }, "导出考勤表"),
        h("button", { class: "btn btn-primary", onclick: function () { editPerson(null, null); } }, "+ 添加人员")
      ])
    ]));

    var list = Store.data.personnel.filter(function (pe) { return pid === "all" || pe.projectId === pid; });
    if (!list.length) { root.appendChild(h("div", { class: "empty card" }, "暂无人员，请添加。")); return root; }

    var card = h("div", { class: "card" }, [h("div", { class: "card-head" }, [h("h2", {}, attLabel(m) + " 考勤明细表（周期 " + attCycleStr(m) + "，单位：天）")]), h("div", { class: "card-body" })]);
    var body = card.querySelector(".card-body");
    var wrap = h("div", { class: "table-wrap", style: "max-height:70vh" });
    var tbl = h("table", { class: "tbl att" });
    // 表头三行
    var thead = h("thead");
    var r0 = h("tr", {}, [h("th", { rowspan: 3, class: "left" }, "操作"), h("th", { rowspan: 3, class: "left" }, "序号"), h("th", { rowspan: 3, class: "left" }, "姓名"), h("th", { rowspan: 3, class: "left" }, "职责分工"), h("th", { rowspan: 3, class: "left" }, "研发课题名称"), h("th", { rowspan: 3 }, "研发人员研发工时"), h("th", { rowspan: 3 }, "研发人员总工时")]);
    for (var d = 0; d < n; d++) {
      var rest = isRestDay(dates[d]);
      r0.appendChild(h("th", { colspan: 2, class: rest ? "rest-day" : "" }, dates[d].month + "/" + dates[d].day));
    }
    thead.appendChild(r0);
    var r1 = h("tr", {});
    for (var d1 = 0; d1 < n; d1++) {
      var rest1 = isRestDay(dates[d1]);
      r1.appendChild(h("th", { colspan: 2, class: rest1 ? "rest-day" : "" }, (rest1 ? "休 " : "") + "周" + WEEK[dates[d1].weekday]));
    }
    thead.appendChild(r1);
    var r2 = h("tr", {});
    for (var d2 = 1; d2 <= n; d2++) { r2.appendChild(h("th", { class: "rd" }, "研发")); r2.appendChild(h("th", { class: "nrd" }, "非研发")); }
    thead.appendChild(r2);
    tbl.appendChild(thead);

    var tb = h("tbody");
    list.forEach(function (pe, idx) {
      var att = ensureAtt(pe, m);
      var tr = h("tr", {}, [
        h("td", { class: "left" }, h("button", { class: "btn btn-sm btn-danger", title: "删除该人员", onclick: function () { delPersonFromAtt(pe.id, idx); } }, "删除")),
        h("td", {}, String(idx + 1)),
        h("td", { class: "left" }, pe.name),
        h("td", { class: "left" }, pe.responsibility || ""),
        h("td", { class: "left" }, pe.projectId ? projectById(pe.projectId).name : "—")
      ]);
      tr.appendChild(h("td", { class: "sum-cell" }, h("span", { id: "rd_" + pe.id }, fmt(sumRd(pe, m), 1))));
      tr.appendChild(h("td", { class: "sum-cell" }, h("span", { id: "tt_" + pe.id }, fmt(sumTotal(pe, m), 1))));
      for (var d = 1; d <= n; d++) {
        var pair = dayPair(att, d);
        var rest2 = isRestDay(dates[d - 1]);
        tr.appendChild(attCell("rd", pe.id, d, pair.rd, rest2));
        tr.appendChild(attCell("nrd", pe.id, d, pair.nonRd, rest2));
      }
      tb.appendChild(tr);
    });
    // 合计行
    var tfoot = h("tfoot");
    var ftr = h("tr", {}, [h("td", { colspan: 5, class: "left sum-cell" }, "合计"), h("td", { class: "sum-cell", id: "foot_rd" }, fmt(list.reduce(function (s, pe) { return s + sumRd(pe, m); }, 0), 1)), h("td", { class: "sum-cell", id: "foot_tt" }, fmt(list.reduce(function (s, pe) { return s + sumTotal(pe, m); }, 0), 1))]);
    for (var d = 1; d <= n; d++) {
      var sR = 0, sN = 0;
      list.forEach(function (pe) { var pr = dayPair(ensureAtt(pe, m), d); sR += Number(pr.rd) || 0; sN += Number(pr.nonRd) || 0; });
      ftr.appendChild(h("td", { class: "sum-cell rd", id: "foot_rd_" + d }, fmt(sR, 1)));
      ftr.appendChild(h("td", { class: "sum-cell nrd", id: "foot_nrd_" + d }, fmt(sN, 1)));
    }
    tfoot.appendChild(ftr);
    tbl.appendChild(tb); tbl.appendChild(tfoot);
    wrap.appendChild(tbl); body.appendChild(wrap);
    root.appendChild(card);
    root.appendChild(signaturesBlock(m, pid));
    root.appendChild(h("div", { class: "text-muted", style: "margin-top:8px;font-size:12px" }, "说明：灰色列=周末/法定节假日（默认工时0，可手动修改）；研发人员研发工时/总工时=本周期累计（含25日）；「批量填充」可一键填充工作日；表下方为签名位。各月独立保存，切换月份不丢数据。"));
    return root;
  };
  function signaturesBlock(m, pid) {
    var s = Store.data.settings;
    var key = m + (pid && pid !== "all" ? "_" + pid : "");
    if (!s.signatures) s.signatures = {};
    if (!s.signatures[key]) s.signatures[key] = { leaderName: "", leaderSig: "", makerName: "", makerSig: "" };
    var sig = s.signatures[key];
    var block = h("div", { class: "card", style: "margin-top:16px" }, [h("div", { class: "card-body" }, [
      h("div", { class: "section-title" }, "签字确认（" + attLabel(m) + " 考勤）"),
      h("div", { class: "sig-row" }, [
        sigColumn("项目负责人", sig, "leader"),
        sigColumn("制表人", sig, "maker")
      ])
    ])]);
    return block;
  }
  function sigColumn(title, sig, kind) {
    var col = h("div", { class: "sig-col" }, [
      h("div", { class: "sig-label" }, title),
      h("div", { class: "sig-name-row" }, [
        h("span", {}, "姓名："),
        h("input", { type: "text", class: "sig-name-input", value: kind === "leader" ? sig.leaderName : sig.makerName, onchange: function (e) { if (kind === "leader") sig.leaderName = e.target.value; else sig.makerName = e.target.value; Store.save(); } })
      ]),
      h("div", { class: "sig-img-area" }, sigImgArea(sig, kind))
    ]);
    return col;
  }
  function sigImgArea(sig, kind) {
    var url = kind === "leader" ? sig.leaderSig : sig.makerSig;
    var frag = h("div", {});
    if (url) {
      frag.appendChild(h("img", { src: url, class: "sig-preview", title: "点击查看/替换", onclick: function () { viewSig(url, function () { uploadSig(function (data) { if (kind === "leader") sig.leaderSig = data; else sig.makerSig = data; Store.save(); Router.go("attendance", Router.params); }); }); } }));
      frag.appendChild(h("div", { class: "sig-btns" }, [
        h("button", { class: "btn btn-sm", onclick: function () { uploadSig(function (data) { if (kind === "leader") sig.leaderSig = data; else sig.makerSig = data; Store.save(); Router.go("attendance", Router.params); }); } }, "替换"),
        h("button", { class: "btn btn-sm btn-danger", onclick: function () { if (kind === "leader") sig.leaderSig = ""; else sig.makerSig = ""; Store.save(); Router.go("attendance", Router.params); } }, "清除")
      ]));
    } else {
      frag.appendChild(h("button", { class: "btn btn-sm", onclick: function () { uploadSig(function (data) { if (kind === "leader") sig.leaderSig = data; else sig.makerSig = data; Store.save(); Router.go("attendance", Router.params); }); } }, "上传签名"));
    }
    return frag;
  }
  function uploadSig(cb) {
    var inp = h("input", { type: "file", accept: "image/*", style: "display:none" });
    inp.addEventListener("change", function () {
      var f = inp.files[0]; if (!f) return;
      if (f.size > 1024 * 1024 * 2) { toast("图片过大（>2MB），请压缩后上传", "err"); return; }
      var reader = new FileReader();
      reader.onload = function (e) { compressImage(e.target.result, 400, 0.7, cb); };
      reader.readAsDataURL(f);
    });
    document.body.appendChild(inp); inp.click();
    setTimeout(function () { if (inp.parentNode) inp.parentNode.removeChild(inp); }, 5000);
  }
  function viewSig(url, onReplace) {
    var body = h("div", { style: "text-align:center" }, [
      h("img", { src: url, style: "max-width:100%;max-height:50vh;border:1px solid var(--border);border-radius:8px" })
    ]);
    openModal("签名预览", body, { wide: true, foot: [
      { text: "关闭", onclick: closeModal },
      { text: "替换", onclick: function () { closeModal(); onReplace(); } }
    ] });
  }
  function batchFillAttendance(m, pid) {
    var dates = attDates(m);
    var list = Store.data.personnel.filter(function (pe) { return pid === "all" || pe.projectId === pid; });
    if (!list.length) { toast("当前筛选无人员", "err"); return; }
    var form = h("div", { class: "form-grid" });
    var peSel = h("select", { id: "bf_person", style: "width:100%" });
    peSel.appendChild(h("option", { value: "all" }, "全部人员（" + list.length + " 人）"));
    list.forEach(function (pe) { peSel.appendChild(h("option", { value: pe.id }, pe.name)); });
    form.appendChild(field("人员", peSel));
    var scopeSel = h("select", { id: "bf_scope", style: "width:100%" });
    scopeSel.appendChild(h("option", { value: "workdays" }, "仅工作日（周末/节假日填0）"));
    scopeSel.appendChild(h("option", { value: "all" }, "全部日期"));
    scopeSel.appendChild(h("option", { value: "rest" }, "仅周末/节假日（填0）"));
    form.appendChild(field("适用范围", scopeSel));
    var modeSel = h("select", { id: "bf_mode", style: "width:100%", onchange: function () {
      var isRandom = el("bf_mode").value === "random";
      el("bf_rd").disabled = isRandom;
      el("bf_nrd").disabled = isRandom;
      el("bf_rd").style.background = isRandom ? "#f5f5f5" : "";
      el("bf_nrd").style.background = isRandom ? "#f5f5f5" : "";
      el("bf_hint").textContent = isRandom
        ? "随机模式：每天研发与非研发工时在0~1间随机生成，保留1位小数，且研发+非研发=1。"
        : "固定模式：工作日填充上方值，周末/法定节假日自动填0。调休补班日视为工作日。";
    } });
    modeSel.appendChild(h("option", { value: "fixed" }, "固定值模式"));
    modeSel.appendChild(h("option", { value: "random" }, "随机值模式（和=1）"));
    form.appendChild(field("填充模式", modeSel));
    form.appendChild(h("div", { class: "field" }, [h("label", {}, "研发工时（每天）"), h("input", { type: "number", id: "bf_rd", step: "0.1", min: "0", max: "1", value: "1", style: "width:100%" })]));
    form.appendChild(h("div", { class: "field" }, [h("label", {}, "非研发工时（每天）"), h("input", { type: "number", id: "bf_nrd", step: "0.1", min: "0", max: "1", value: "0", style: "width:100%" })]));
    form.appendChild(h("div", { class: "field field-full" }, [h("label", { id: "bf_hint" }, "固定模式：工作日填充上方值，周末/法定节假日自动填0。调休补班日视为工作日。")]));
    openModal("批量填充工时（" + attLabel(m) + "）", form, { wide: true, foot: [
      { text: "取消", onclick: closeModal },
      { text: "仅预览", onclick: function () {
        var scope = el("bf_scope").value;
        var mode = el("bf_mode").value;
        var work = 0, rest = 0;
        dates.forEach(function (dt) { if (isRestDay(dt)) rest++; else work++; });
        var msg = "周期共 " + dates.length + " 天：工作日 " + work + " 天，休息日 " + rest + " 天。";
        if (mode === "random") {
          msg += "\n随机模式：工作日随机生成（研发+非研发=1，1位小数），休息日填0。";
        } else {
          var rd = Number(el("bf_rd").value) || 0, nrd = Number(el("bf_nrd").value) || 0;
          if (scope === "workdays") msg += "\n将填充工作日：研发" + rd + " 非研发" + nrd + "，休息日填0。";
          else if (scope === "rest") msg += "\n将休息日全部填0。";
          else msg += "\n将全部日期填：研发" + rd + " 非研发" + nrd + "。";
        }
        alert(msg);
      } },
      { text: "执行填充", kind: "btn-primary", onclick: function () {
        var peVal = el("bf_person").value;
        var scope = el("bf_scope").value;
        var mode = el("bf_mode").value;
        var rd = Number(el("bf_rd").value) || 0, nrd = Number(el("bf_nrd").value) || 0;
        var targets = peVal === "all" ? list : list.filter(function (pe) { return pe.id === peVal; });
        var filled = 0;
        targets.forEach(function (pe) {
          var att = ensureAtt(pe, m);
          for (var d = 1; d <= dates.length; d++) {
            var rest = isRestDay(dates[d - 1]);
            var pair = dayPair(att, d);
            if (mode === "random") {
              if (scope === "workdays") {
                if (rest) { pair.rd = 0; pair.nonRd = 0; }
                else { var r = randomPairSum1(); pair.rd = r.rd; pair.nonRd = r.nrd; }
              } else if (scope === "rest") {
                if (rest) { pair.rd = 0; pair.nonRd = 0; }
              } else {
                var r2 = randomPairSum1(); pair.rd = r2.rd; pair.nonRd = r2.nrd;
              }
            } else {
              if (scope === "workdays") { if (rest) { pair.rd = 0; pair.nonRd = 0; } else { pair.rd = rd; pair.nonRd = nrd; } }
              else if (scope === "rest") { if (rest) { pair.rd = 0; pair.nonRd = 0; } }
              else { pair.rd = rd; pair.nonRd = nrd; }
            }
          }
          filled++;
        });
        Store.save(); closeModal(); Router.go("attendance", Router.params);
        toast("已批量填充 " + filled + " 人 × " + dates.length + " 天" + (mode === "random" ? "（随机模式）" : ""), "ok");
      } }
    ] });
  }
  function randomPairSum1() {
    var rd = Math.round(Math.random() * 10) / 10;
    if (rd > 1) rd = 1;
    var nrd = Math.round((1 - rd) * 10) / 10;
    return { rd: rd, nrd: nrd };
  }
  function delPersonFromAtt(id, idx) {
    var pe = personById(id);
    confirmModal("删除人员", "确认从考勤表删除「" + pe.name + "」？该人员所有月度考勤将一并删除。", function () {
      Store.data.personnel = Store.data.personnel.filter(function (p) { return p.id !== id; });
      Store.save(); toast("已删除", "ok"); Router.go("attendance", Router.params);
    });
  }
  function compressImage(dataUrl, maxW, quality, cb) {
    var img = new Image();
    img.onload = function () {
      var w = img.width, h = img.height;
      if (w > maxW) { h = h * (maxW / w); w = maxW; }
      var canvas = document.createElement("canvas");
      canvas.width = w; canvas.height = h;
      var ctx = canvas.getContext("2d");
      ctx.drawImage(img, 0, 0, w, h);
      try { cb(canvas.toDataURL("image/jpeg", quality)); }
      catch (e) { cb(dataUrl); }
    };
    img.onerror = function () { cb(dataUrl); };
    img.src = dataUrl;
  }
  function attCell(type, peId, d, val, rest) {
    var m = Store.data.settings.currentMonth;
    var input = h("input", { type: "number", step: "0.1", min: "0", max: "1", value: val, onchange: function (e) {
      var pe = personById(peId); var att = ensureAtt(pe, m);
      var v = Number(e.target.value); if (isNaN(v)) v = 0;
      v = Math.min(1, Math.max(0, v)); v = Math.round(v * 10) / 10;
      var pair = dayPair(att, d);
      if (type === "rd") { pair.rd = v; pair.nonRd = Math.round((1 - v) * 10) / 10; }
      else { pair.nonRd = v; pair.rd = Math.round((1 - v) * 10) / 10; }
      Store.save();
      var rdTd = document.getElementById("att_rd_" + peId + "_" + d);
      var nrdTd = document.getElementById("att_nrd_" + peId + "_" + d);
      if (rdTd) rdTd.querySelector("input").value = pair.rd;
      if (nrdTd) nrdTd.querySelector("input").value = pair.nonRd;
      el("rd_" + peId).textContent = fmt(sumRd(pe, m), 1);
      el("tt_" + peId).textContent = fmt(sumTotal(pe, m), 1);
      updateAttFoot(m);
      checkAttRowOverflow(peId, d);
    } });
    if (rest) { input.placeholder = "0"; input.style.background = "#f5f5f5"; }
    var td = h("td", { class: type + (rest ? " rest-cell" : "") }, input);
    td.id = "att_" + type + "_" + peId + "_" + d;
    setTimeout(function () { checkAttRowOverflow(peId, d, td); }, 0);
    return td;
  }
  function updateAttFoot(m) {
    var pid = Router.params.pid || "all";
    var list = Store.data.personnel.filter(function (pe) { return pid === "all" || pe.projectId === pid; });
    var footRd = el("foot_rd"), footTt = el("foot_tt");
    if (!footRd || !footTt) return;
    var totR = 0, totT = 0;
    list.forEach(function (pe) { totR += sumRd(pe, m); totT += sumTotal(pe, m); });
    footRd.textContent = fmt(totR, 1);
    footTt.textContent = fmt(totT, 1);
    var n = attDates(m).length;
    for (var d = 1; d <= n; d++) {
      var sR = 0, sN = 0;
      list.forEach(function (pe) { var pr = dayPair(ensureAtt(pe, m), d); sR += Number(pr.rd) || 0; sN += Number(pr.nonRd) || 0; });
      var cR = el("foot_rd_" + d), cN = el("foot_nrd_" + d);
      if (cR) cR.textContent = fmt(sR, 1);
      if (cN) cN.textContent = fmt(sN, 1);
    }
  }
  function checkAttRowOverflow(peId, d, forceTd) {
    var m = Store.data.settings.currentMonth;
    var pe = personById(peId); if (!pe) return;
    var pair = dayPair(ensureAtt(pe, m), d);
    var total = (Number(pair.rd) || 0) + (Number(pair.nonRd) || 0);
    var rdTd = forceTd && forceTd.classList.contains("rd") ? forceTd : document.getElementById("att_rd_" + peId + "_" + d);
    var nrdTd = forceTd && forceTd.classList.contains("nrd") ? forceTd : document.getElementById("att_nrd_" + peId + "_" + d);
    var markRed = function (td, on) { if (!td) return; if (on) { td.classList.add("overflow"); td.title = "当日研发+非研发=" + total.toFixed(1) + "，超过1，请修改"; } else { td.classList.remove("overflow"); td.title = ""; } };
    markRed(rdTd, total > 1.0001);
    markRed(nrdTd, total > 1.0001);
    if (total > 1.0001) toast((pe.name || "行" + peId) + " 第" + d + "天工时=" + total.toFixed(1) + " 超过1，请修改", "err");
  }
  function projectSelect(val, onchange) {
    var s = h("select", { onchange: function (e) { onchange(e.target.value); } });
    s.appendChild(h("option", { value: "all" }, "全部项目"));
    Store.data.projects.forEach(function (p) { var o = h("option", { value: p.id }, p.name); if (p.id === val) o.selected = true; s.appendChild(o); });
    return s;
  }

  /* ---- 月度汇报 ---- */
  /* ---- 月度执行情况（合并月度汇报+经费管理） ---- */
  VIEWS.monthly = function () {
    var m = Store.data.settings.currentMonth;
    var root = h("div");
    root.appendChild(h("div", { class: "toolbar" }, [
      h("div", { class: "left" }, [
        h("div", { class: "filter" }, [h("span", {}, "月份"), h("input", { type: "month", value: m, onchange: function (e) { Store.data.settings.currentMonth = e.target.value; Store.save(); Router.go("monthly"); } })]),
        h("div", { class: "filter" }, [h("span", {}, "日单价"), h("input", { type: "number", value: Store.data.settings.dailyRate, onchange: function (e) { Store.data.settings.dailyRate = Number(e.target.value) || 0; Store.save(); Router.go("monthly"); }, style: "width:80px" }), h("span", {}, "元/天")])
      ]),
      h("div", { class: "right" }, [h("button", { class: "btn", onclick: function () { exportMonthly(m); } }, "导出月度执行情况"), h("button", { class: "btn btn-primary", onclick: function () { batchEditMonthly(m); } }, "批量编辑本月")])
    ]));
    var totBudget = 0, totFunds = 0, totCum = 0;
    var wrap = h("div", { class: "table-wrap" });
    var tbl = h("table", { class: "tbl" });
    tbl.innerHTML = "<thead><tr><th>项目名称</th><th>起止周期</th><th>负责人</th><th class='num'>总预算/元</th><th class='num'>本月归集经费/元</th><th class='num'>累计归集经费/元</th><th>指标完成情况（知识产权）</th><th>本月完成工作</th><th>存在问题</th><th>下步计划</th><th>问题说明</th><th>操作</th></tr></thead>";
    var tb = h("tbody");
    Store.data.projects.forEach(function (p) {
      var r = ensureMonth(p, m);
      var cost = projectPersonCost(p.id, m);
      var cum = r.cumulativeFunds ? Number(r.cumulativeFunds) : cumulativeFundsUpto(p.id, m);
      totBudget += Number(p.totalBudget) || 0; totFunds += Number(r.monthlyFunds) || 0; totCum += cum;
      // 计算指标完成情况（知识产权）：p.targets 要求数量 vs 成果管理实际数量
      var tkeys = p.targets ? Object.keys(p.targets).filter(function (t) { return Number(p.targets[t]) > 0; }) : [];
      var achCell;
      if (tkeys.length) {
        achCell = h("td", { class: "wrap", style: "max-width:170px;font-size:12px" });
        tkeys.forEach(function (t, i) {
          var req = Number(p.targets[t]) || 0;
          var act = countAchievements(p.id, t);
          var span = h("span", {}, t + " " + act + "/" + req);
          if (act >= req) { span.style.color = "#1e8e3e"; span.style.fontWeight = "600"; }
          achCell.appendChild(span);
          if (i < tkeys.length - 1) achCell.appendChild(document.createTextNode("，"));
        });
      } else achCell = h("td", {}, "—");

      // 自动检测问题说明
      var issues = [];
      var cwEmpty = !r.completedWork || !String(r.completedWork).trim();
      var mf = Number(r.monthlyFunds) || 0;
      if (cwEmpty) issues.push("本月完成工作未填报");
      if (!r.monthlyFunds || mf === 0) issues.push("本月归集经费未填报");
      if (p.totalBudget && cum > Number(p.totalBudget) * 10000) issues.push("累计经费超预算");
      if (mf > 0 && cwEmpty) issues.push("有经费无工作进度");
      var ppl = peopleInProject(p.id);
      if (!ppl.length) issues.push("未关联人员");
      var crossNames = {};
      ppl.forEach(function (pe) {
        if (crossNames[pe.name]) return;
        var cross = Store.data.personnel.some(function (o) { return o.name === pe.name && o.projectId !== p.id; });
        if (cross) { crossNames[pe.name] = true; issues.push(pe.name + "跨多个项目"); }
      });
      if (tkeys.length) {
        var projAch = (Store.data.achievements || []).filter(function (a) { return a.projectId === p.id; });
        if (!projAch.length) issues.push("知识产权无成果记录");
        else if (!projAch.some(function (a) { return a.date && String(a.date).slice(0, 7) === m; })) issues.push("本月无新增成果");
      }
      var issueStr = issues.length ? issues.join("；") : "正常";

      tb.appendChild(h("tr", {}, [
        h("td", { class: "wrap" }, p.name),
        h("td", { style: "font-size:12px;white-space:nowrap" }, (p.periodStart || "—") + " ~ " + (p.periodEnd || "—")),
        h("td", {}, p.leader || "—"),
        h("td", { class: "num" }, p.totalBudget ? fmt(Number(p.totalBudget) * 10000, 2) : "—"),
        h("td", { class: "num" }, fmt(r.monthlyFunds || 0, 2)),
        h("td", { class: "num" }, fmt(cum, 2)),
        achCell,
        h("td", { class: "wrap", style: "max-width:200px" }, esc(r.completedWork || "").slice(0, 50) || "—"),
        h("td", { class: "wrap", style: "max-width:200px" }, esc(r.problems || "").slice(0, 50) || "—"),
        h("td", { class: "wrap", style: "max-width:200px" }, esc(r.nextPlan || "").slice(0, 50) || "—"),
        h("td", { class: "wrap", style: "max-width:160px;font-size:12px" }, issueStr),
        h("td", {}, [h("button", { class: "btn btn-sm", onclick: function () { editMonthly(p.id, m); } }, "填报"), h("button", { class: "btn btn-sm", onclick: function () {
          r.monthlyFunds = fmt(cost, 2); Store.save(); Router.go("monthly"); toast("已用估算值填充本月经费", "ok");
        } }, "用估算填充")])
      ]));
    });
    tb.appendChild(h("tr", { style: "font-weight:600;background:#eef3fa" }, [
      h("td", {}, "合计"), h("td", {}, "—"), h("td", {}, "—"),
      h("td", { class: "num" }, fmt(totBudget * 10000, 2)),
      h("td", { class: "num" }, fmt(totFunds, 2)),
      h("td", { class: "num" }, fmt(totCum, 2)),
      h("td", { colspan: 6 }, "")
    ]));
    tbl.appendChild(tb); wrap.appendChild(tbl); root.appendChild(wrap);
    root.appendChild(h("div", { class: "text-muted", style: "margin-top:8px;font-size:12px" }, "说明：总预算=万元×10000转为元；本月/累计归集经费单位均为元；「指标完成情况」=成果管理实际数/考核指标要求数（绿色加粗为已达标）；「问题说明」为系统自动检测生成。"));
    return root;
  };
  function editMonthly(pid, m) {
    var p = projectById(pid); if (!p) return;
    var r = ensureMonth(p, m);
    var rd = projectRdSum(p.id, m);
    var cost = projectPersonCost(p.id, m);
    var form = h("div", { class: "form-grid" });
    form.appendChild(fieldFull("项目名称", h("div", { class: "text-muted" }, p.name)));
    form.appendChild(h("div", { class: "field" }, [h("label", {}, "月份"), h("div", { class: "text-muted" }, m)]));
    var grid = h("div", { class: "form-grid" });
    grid.appendChild(field("研发工时（天）", h("div", { class: "text-muted" }, fmt(rd, 1))));
    form.appendChild(grid);
    form.appendChild(h("div", { class: "field field-full" }, [h("label", {}, "本月完成工作（含成果产出）"), h("textarea", { id: "f_cw", style: "width:100%;min-height:80px" }, r.completedWork || "")]));
    form.appendChild(h("div", { class: "field field-full" }, [h("label", {}, "存在问题"), h("textarea", { id: "f_pb", style: "width:100%;min-height:60px" }, r.problems || "")]));
    form.appendChild(h("div", { class: "field field-full" }, [h("label", {}, "下步计划"), h("textarea", { id: "f_np", style: "width:100%;min-height:60px" }, r.nextPlan || "")]));
    // 指标完成情况
    var achCount = (Store.data.achievements || []).filter(function(a) { return a.projectId === p.id; }).length;
    form.appendChild(h("div", { class: "field field-full" }, [h("label", {}, "指标完成情况（知识产权）"), h("div", { class: "text-muted" }, achCount + " 项成果（论文/专利/报告等），详见成果管理模块")]));
    var grid2 = h("div", { class: "form-grid" });
    grid2.appendChild(field("本月归集经费（元）", h("input", { id: "f_mf", value: r.monthlyFunds || "", type: "number", step: "0.01", style: "width:100%" })));
    grid2.appendChild(field("累计归集经费（元）", h("input", { id: "f_cf", value: r.cumulativeFunds || "", type: "number", step: "0.01", style: "width:100%" })));
    form.appendChild(grid2);
    form.appendChild(h("div", { class: "field field-full" }, [h("label", {}, "提示：研发工时 " + fmt(rd, 1) + " 天 × 日单价 " + (Store.data.settings.dailyRate || 0) + " 元/天 = " + fmt(cost, 2) + " 元"), h("button", { class: "btn btn-sm", onclick: function () {
      el("f_mf").value = fmt(cost, 2);
      toast("已用估算值填充本月经费", "ok");
    } }, "用估算值填充")]));
    openModal("月度执行情况填报", form, { wide: true, foot: [
      { text: "取消", onclick: closeModal },
      { text: "保存", kind: "btn-primary", onclick: function () {
        r.completedWork = el("f_cw").value; r.problems = el("f_pb").value; r.nextPlan = el("f_np").value;
        r.monthlyFunds = el("f_mf").value; r.cumulativeFunds = el("f_cf").value;
        Store.save(); closeModal(); Router.go(Router.view, Router.params); toast("已保存", "ok");
      } }
    ] });
  }
  function batchEditMonthly(m) {
    var form = h("div", {});
    Store.data.projects.forEach(function (p) {
      var r = ensureMonth(p, m);
      var cost = projectPersonCost(p.id, m);
      var fieldset = h("div", { class: "card", style: "margin-bottom:12px" }, [h("div", { class: "card-body" })]);
      var fb = fieldset.querySelector(".card-body");
      fb.appendChild(h("div", { class: "section-title" }, p.name));
      fb.appendChild(h("div", { class: "form-grid" }, [
        h("div", { class: "field" }, [h("label", {}, "本月归集经费/元"), h("input", { id: "mf_" + p.id, value: r.monthlyFunds || "", type: "number", step: "0.01", style: "width:100%" })])
      ]));
      fb.appendChild(h("div", { class: "form-grid" }, [
        h("div", { class: "field" }, [h("label", {}, "累计归集经费/元"), h("input", { id: "cf_" + p.id, value: r.cumulativeFunds || "", type: "number", step: "0.01", style: "width:100%" })]
        )]));
      fb.appendChild(field("本月完成工作", h("textarea", { id: "cw_" + p.id, style: "width:100%;min-height:50px" }, r.completedWork || "")));
      fb.appendChild(field("存在问题", h("textarea", { id: "pb_" + p.id, style: "width:100%;min-height:40px" }, r.problems || "")));
      fb.appendChild(field("下步计划", h("textarea", { id: "np_" + p.id, style: "width:100%;min-height:40px" }, r.nextPlan || "")));
      var achCount = (Store.data.achievements || []).filter(function(a) { return a.projectId === p.id; }).length;
      fb.appendChild(field("指标完成情况（知识产权）", h("div", { class: "text-muted" }, achCount + " 项成果")));
      form.appendChild(fieldset);
    });
    openModal("批量编辑 " + m + " 月度执行情况", form, { wide: true, foot: [
      { text: "取消", onclick: closeModal },
      { text: "全部保存", kind: "btn-primary", onclick: function () {
        Store.data.projects.forEach(function (p) {
          var r = ensureMonth(p, m);
          r.monthlyFunds = el("mf_" + p.id).value; r.cumulativeFunds = el("cf_" + p.id).value;
          r.completedWork = el("cw_" + p.id).value; r.problems = el("pb_" + p.id).value; r.nextPlan = el("np_" + p.id).value;
        });
        Store.save(); closeModal(); Router.go("monthly"); toast("已保存", "ok");
      } }
    ] });
  }
  function exportMonthly(m) {
    var rows = [["中咨养护检测 - 月度执行情况表"], ["月份：" + m + "    单位：元"], ["项目名称", "起止周期", "负责人", "总预算/元", "本月归集经费/元", "累计归集经费/元", "指标完成情况（知识产权）", "本月完成工作", "存在问题", "下步计划"]];
    Store.data.projects.forEach(function (p) {
      var r = ensureMonth(p, m);
      var cum = r.cumulativeFunds ? Number(r.cumulativeFunds) : cumulativeFundsUpto(p.id, m);
      var tkeys = p.targets ? Object.keys(p.targets).filter(function (t) { return Number(p.targets[t]) > 0; }) : [];
      var achStr = tkeys.length ? tkeys.map(function (t) { return t + " " + countAchievements(p.id, t) + "/" + (Number(p.targets[t]) || 0); }).join("，") : "—";
      rows.push([p.name, (p.periodStart || "") + "~" + (p.periodEnd || ""), p.leader || "", p.totalBudget ? (Number(p.totalBudget) * 10000).toFixed(2) : "", r.monthlyFunds || "", cum.toFixed(2), achStr, r.completedWork || "", r.problems || "", r.nextPlan || ""]);
    });
    // 合计行
    var totB = Store.data.projects.reduce(function (s, p) { return s + (Number(p.totalBudget) || 0); }, 0) * 10000;
    var totF = Store.data.projects.reduce(function (s, p) { return s + (Number(ensureMonth(p, m).monthlyFunds) || 0); }, 0);
    var totCum = Store.data.projects.reduce(function (s, p) { var r = ensureMonth(p, m); return s + (r.cumulativeFunds ? Number(r.cumulativeFunds) : cumulativeFundsUpto(p.id, m)); }, 0);
    rows.push(["合计", "", "", totB.toFixed(2), totF.toFixed(2), totCum.toFixed(2), "", "", "", ""]);
    var ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 9 } }, { s: { r: 1, c: 0 }, e: { r: 1, c: 9 } }];
    ws["!cols"] = [{ wch: 24 }, { wch: 18 }, { wch: 10 }, { wch: 14 }, { wch: 16 }, { wch: 16 }, { wch: 24 }, { wch: 30 }, { wch: 30 }, { wch: 30 }];
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, m + "月度执行");
    download("月度执行情况-" + m + ".xlsx", wb);
    toast("已导出月度执行情况", "ok");
  }

  /* ---- 人员台账 ---- */
  VIEWS.personnel = function () {
    var root = h("div");
    var selectedIds = {};
    var tb = h("tbody");
    var batchBar = h("div", { class: "toolbar", style: "display:none" });
    root.appendChild(h("div", { class: "toolbar" }, [
      h("div", { class: "left" }, h("div", { class: "section-title" }, "研发人员台账（" + Store.data.personnel.length + " 人）")),
      h("div", { class: "right" }, [
        h("button", { class: "btn", onclick: downloadPersonnelTemplate }, "下载导入模板"),
        h("button", { class: "btn", onclick: function () { triggerUpload(importPersonnel, ".xls,.xlsx"); } }, "上传Excel导入"),
        h("button", { class: "btn", onclick: function () { toggleBatchSelect(root, selectedIds, batchBar, tb); } }, "批量选取"),
        h("button", { class: "btn btn-primary", onclick: function () { editPerson(null, null); } }, "+ 添加人员")
      ])
    ]));
    // 批量操作工具栏
    batchBar.appendChild(h("div", { class: "left" }, [
      h("span", { class: "section-title", id: "batch-count" }, "已选 0 人")
    ]));
    batchBar.appendChild(h("div", { class: "right" }, [
      h("button", { class: "btn btn-sm", onclick: function () { selectAllRows(selectedIds, tb); } }, "全选"),
      h("button", { class: "btn btn-sm", onclick: function () { clearSelection(selectedIds, tb, batchBar); } }, "取消选择"),
      h("button", { class: "btn btn-sm btn-primary", onclick: function () { batchAssignProject(selectedIds); } }, "批量分配项目"),
      h("button", { class: "btn btn-sm btn-danger", onclick: function () { batchDeletePersons(selectedIds, batchBar, tb); } }, "批量删除"),
      h("button", { class: "btn btn-sm", onclick: function () { exitBatchMode(root, selectedIds, batchBar, tb); } }, "退出批量模式")
    ]));
    root.appendChild(batchBar);
    // 检查一人关联多项目的情况
    var conflicts = findPersonProjectConflicts();
    if (conflicts.length) {
      var warn = h("div", { class: "card warn-card" }, [h("div", { class: "card-body" }, [
        h("div", { class: "section-title", style: "color:#c0392b" }, "项目关联冲突（同一姓名关联了多个项目，请修改）"),
        h("div", { class: "text-muted", style: "font-size:12px" }, "规则：每个人员只能关联一个项目。以下人员存在多处关联：")
      ])]);
      var wbody = warn.querySelector(".card-body");
      conflicts.forEach(function (c) {
        wbody.appendChild(h("div", { style: "margin-top:4px" }, [
          h("span", { style: "font-weight:600" }, c.name + "："),
          h("span", { style: "color:#c0392b" }, c.projects.map(function (p) { return p ? p : "未分配"; }).join("、"))
        ]));
      });
      root.appendChild(warn);
    }
    var wrap = h("div", { class: "table-wrap" });
    var tbl = h("table", { class: "tbl" });
    tbl.innerHTML = "<thead><tr><th class='batch-col' style='display:none'>选</th><th>序号</th><th>姓名</th><th>职责分工</th><th>所属项目</th><th class='num'>本月研发工时</th><th class='num'>本月总工时</th><th>操作</th></tr></thead>";
    var m = Store.data.settings.currentMonth;
    Store.data.personnel.forEach(function (pe, i) {
      var tr = h("tr", {}, [
        h("td", { class: "batch-col", style: "display:none" }, h("input", { type: "checkbox", "data-id": pe.id, onchange: function (e) { onRowSelect(e, selectedIds, batchBar); } })),
        h("td", {}, String(i + 1)),
        h("td", {}, pe.name),
        h("td", { class: "wrap" }, pe.responsibility || "—"),
        h("td", {}, pe.projectId ? projectById(pe.projectId).name : h("span", { class: "text-muted" }, "未分配")),
        h("td", { class: "num" }, fmt(sumRd(pe, m), 1)),
        h("td", { class: "num" }, fmt(sumTotal(pe, m), 1)),
        h("td", {}, [
          h("button", { class: "btn btn-sm", onclick: function () { editPerson(pe.id, null); } }, "编辑"),
          h("button", { class: "btn btn-sm", onclick: function () { Router.params.pid = pe.projectId; Router.go("attendance", Router.params); } }, "考勤"),
          h("button", { class: "btn btn-sm btn-danger", onclick: function () { delPerson(pe.id); } }, "删除")
        ])
      ]);
      tb.appendChild(tr);
    });
    tbl.appendChild(tb); wrap.appendChild(tbl); root.appendChild(wrap);
    return root;
  };
  // 批量选取相关辅助函数
  function toggleBatchSelect(root, selectedIds, batchBar, tb) {
    batchBar.style.display = "flex";
    var cols = root.querySelectorAll(".batch-col");
    cols.forEach(function (c) { c.style.display = ""; });
    toast("已进入批量选取模式，勾选人员后可批量操作", "ok");
  }
  function exitBatchMode(root, selectedIds, batchBar, tb) {
    batchBar.style.display = "none";
    var cols = root.querySelectorAll(".batch-col");
    cols.forEach(function (c) { c.style.display = "none"; });
    Object.keys(selectedIds).forEach(function (k) { delete selectedIds[k]; });
    var boxes = tb.querySelectorAll("input[type=checkbox]");
    boxes.forEach(function (b) { b.checked = false; });
    Router.go(Router.view, Router.params);
  }
  function onRowSelect(e, selectedIds, batchBar) {
    var id = e.target.getAttribute("data-id");
    if (e.target.checked) selectedIds[id] = true;
    else delete selectedIds[id];
    var cnt = Object.keys(selectedIds).length;
    var lbl = el("batch-count");
    if (lbl) lbl.textContent = "已选 " + cnt + " 人";
  }
  function selectAllRows(selectedIds, tb) {
    var boxes = tb.querySelectorAll("input[type=checkbox]");
    boxes.forEach(function (b) { b.checked = true; b.dispatchEvent(new Event("change")); });
  }
  function clearSelection(selectedIds, tb, batchBar) {
    Object.keys(selectedIds).forEach(function (k) { delete selectedIds[k]; });
    var boxes = tb.querySelectorAll("input[type=checkbox]");
    boxes.forEach(function (b) { b.checked = false; });
    var lbl = el("batch-count");
    if (lbl) lbl.textContent = "已选 0 人";
  }
  function batchDeletePersons(selectedIds, batchBar, tb) {
    var ids = Object.keys(selectedIds);
    if (!ids.length) { toast("请先勾选要删除的人员", "err"); return; }
    var names = ids.map(function (id) { var p = personById(id); return p ? p.name : ""; }).filter(Boolean);
    confirmModal("批量删除", "确认删除以下 " + ids.length + " 名人员？删除后考勤记录一并清除：\n\n" + names.join("、"), function () {
      var idSet = {}; ids.forEach(function (id) { idSet[id] = true; });
      Store.data.personnel = Store.data.personnel.filter(function (p) { return !idSet[p.id]; });
      Store.save();
      toast("已删除 " + ids.length + " 名人员", "ok");
      exitBatchMode(document, selectedIds, batchBar, tb);
      Router.go(Router.view, Router.params);
    });
  }
  function batchAssignProject(selectedIds) {
    var ids = Object.keys(selectedIds);
    if (!ids.length) { toast("请先勾选要分配项目的人员", "err"); return; }
    var form = h("div", { class: "form-grid" });
    var cnt = h("div", { class: "text-muted", style: "margin-bottom:8px" }, "为 " + ids.length + " 名人员批量分配项目：");
    var sel = h("select", { id: "batch_proj", style: "width:100%" });
    sel.appendChild(h("option", { value: "" }, "未分配"));
    Store.data.projects.forEach(function (p) { sel.appendChild(h("option", { value: p.id }, p.name)); });
    form.appendChild(cnt);
    form.appendChild(field("所属项目", sel));
    openModal("批量分配项目", form, { foot: [
      { text: "取消", onclick: closeModal },
      { text: "确认分配", kind: "btn-primary", onclick: function () {
        var newPid = el("batch_proj").value || null;
        var cnt2 = 0;
        ids.forEach(function (id) { var p = personById(id); if (p) { p.projectId = newPid; cnt2++; } });
        Store.save();
        closeModal();
        Router.go(Router.view, Router.params);
        toast("已为 " + cnt2 + " 人分配" + (newPid ? "到「" + (projectById(newPid) ? projectById(newPid).name : "") + "」" : "为未分配"), "ok");
      } }
    ] });
  }
  function editPerson(id, fixedPid) {
    var pe = id ? personById(id) : { name: "", responsibility: "", projectId: fixedPid || null };
    var form = h("div", { class: "form-grid" });
    form.appendChild(field("姓名", h("input", { id: "p_name", value: pe.name || "", style: "width:100%" })));
    var sel = h("select", { id: "p_proj", style: "width:100%" });
    sel.appendChild(h("option", { value: "" }, "未分配"));
    Store.data.projects.forEach(function (p) { var o = h("option", { value: p.id }, p.name); if (p.id === pe.projectId) o.selected = true; sel.appendChild(o); });
    form.appendChild(field("所属项目", sel));
    form.appendChild(h("div", { class: "field field-full" }, [h("label", {}, "职责分工"), h("input", { id: "p_resp", value: pe.responsibility || "", style: "width:100%" })]));
    openModal(id ? "编辑人员" : "添加人员", form, { foot: [
      { text: "取消", onclick: closeModal },
      { text: "保存", kind: "btn-primary", onclick: function () {
        var name = el("p_name").value.trim();
        if (!name) { toast("请填写姓名", "err"); return; }
        var newPid = el("p_proj").value || null;
        // 校验：同一姓名只能关联一个项目
        var sameNameOthers = Store.data.personnel.filter(function (p) { return p.name === name && p.id !== id; });
        if (sameNameOthers.length && newPid) {
          var otherProj = sameNameOthers[0].projectId;
          if (otherProj && otherProj !== newPid) {
            var otherP = projectById(otherProj);
            var newP = projectById(newPid);
            confirmModal("项目关联冲突", "「" + name + "」已关联项目「" + (otherP ? otherP.name : "—") + "」。\n每个人只能关联一个项目。\n\n确认改关联为「" + (newP ? newP.name : "—") + "」吗？\n（原项目关联将被替换）", function () {
              if (id) { var x = personById(id); x.name = name; x.responsibility = el("p_resp").value.trim(); x.projectId = newPid; }
              else { Store.data.personnel.push({ id: uid(), name: name, responsibility: el("p_resp").value.trim(), projectId: newPid, attendance: {} }); }
              Store.save(); closeModal(); Router.go(Router.view, Router.params); toast("已保存（项目已替换）", "ok");
            });
            return;
          }
        }
        if (id) { var x = personById(id); x.name = name; x.responsibility = el("p_resp").value.trim(); x.projectId = newPid; }
        else { Store.data.personnel.push({ id: uid(), name: name, responsibility: el("p_resp").value.trim(), projectId: newPid, attendance: {} }); }
        Store.save(); closeModal(); Router.go(Router.view, Router.params); toast("已保存", "ok");
      } }
    ] });
  }
  function delPerson(id) {
    confirmModal("删除人员", "确认删除该人员？其考勤记录将一并删除。", function () {
      Store.data.personnel = Store.data.personnel.filter(function (p) { return p.id !== id; });
      Store.save(); toast("已删除", "ok"); Router.go(Router.view, Router.params);
    });
  }
  function findPersonProjectConflicts() {
    // 按姓名分组，若同一姓名关联了不同的项目，则为冲突
    var byName = {};
    Store.data.personnel.forEach(function (pe) {
      if (!byName[pe.name]) byName[pe.name] = [];
      var pn = pe.projectId ? (projectById(pe.projectId) ? projectById(pe.projectId).name : "未知项目") : "未分配";
      if (byName[pe.name].indexOf(pn) < 0) byName[pe.name].push(pn);
    });
    var conflicts = [];
    Object.keys(byName).forEach(function (name) {
      if (byName[name].length > 1) conflicts.push({ name: name, projects: byName[name] });
    });
    return conflicts;
  }

  /* ---- 成果管理 ---- */
  var ACH_TYPES = ["研究报告", "论文", "专利", "软件著作权", "标准", "其他"];
  function achievementById(id) { for (var i = 0; i < Store.data.achievements.length; i++) if (Store.data.achievements[i].id === id) return Store.data.achievements[i]; return null; }
  VIEWS.achievements = function () {
    var root = h("div");
    root.appendChild(h("div", { class: "toolbar" }, [
      h("div", { class: "left" }, [
        h("div", { class: "section-title" }, "知识产权与成果（" + Store.data.achievements.length + " 项）"),
        typeFilter(function (v) { Router.params.type = v; Router.go("achievements", Router.params); }, Router.params.type || "all")
      ]),
      h("div", { class: "right" }, [
        h("button", { class: "btn", onclick: exportAchievements }, "导出成果清单"),
        h("button", { class: "btn btn-primary", onclick: function () { editAchievement(null); } }, "+ 新增成果")
      ])
    ]));

    // 考核指标对照表
    root.appendChild(targetsCard());

    var list = Store.data.achievements;
    var ft = Router.params.type || "all";
    if (ft !== "all") list = list.filter(function (a) { return a.type === ft; });
    if (!list.length) { root.appendChild(h("div", { class: "empty card" }, "暂无成果记录，点击「新增成果」添加。")); return root; }
    var wrap = h("div", { class: "table-wrap" });
    var tbl = h("table", { class: "tbl" });
    tbl.innerHTML = "<thead><tr><th>序号</th><th>类型</th><th>名称/标题</th><th>作者</th><th>所属项目</th><th>日期</th><th>说明</th><th>操作</th></tr></thead>";
    var tb = h("tbody");
    list.forEach(function (a, i) {
      tb.appendChild(h("tr", {}, [
        h("td", {}, String(i + 1)),
        h("td", {}, h("span", { class: "tag " + typeTagClass(a.type) }, a.type)),
        h("td", { class: "wrap" }, a.title || "—"),
        h("td", { class: "wrap" }, a.authors || "—"),
        h("td", {}, a.projectId ? projectById(a.projectId).name : h("span", { class: "text-muted" }, "—")),
        h("td", {}, a.date || "—"),
        h("td", { class: "wrap" }, a.description || "—"),
        h("td", {}, [
          h("button", { class: "btn btn-sm", onclick: function () { editAchievement(a.id); } }, "编辑"),
          h("button", { class: "btn btn-sm btn-danger", onclick: function () { delAchievement(a.id); } }, "删除")
        ])
      ]));
    });
    tbl.appendChild(tb); wrap.appendChild(tbl); root.appendChild(wrap);
    return root;
  };
  function typeFilter(onchange, val) {
    var s = h("select", { onchange: function (e) { onchange(e.target.value); }, class: "filter" });
    [{ v: "all", t: "全部类型" }].concat(ACH_TYPES.map(function (t) { return { v: t, t: t }; })).forEach(function (o) {
      var op = h("option", { value: o.v }, o.t); if (o.v === val) op.selected = true; s.appendChild(op);
    });
    return h("div", { class: "filter" }, [h("span", {}, "类型"), s]);
  }
  function typeTagClass(t) {
    if (t === "专利") return "tag-ok";
    if (t === "论文") return "tag-warn";
    return "";
  }
  function editAchievement(id) {
    var a = id ? achievementById(id) : { type: "研究报告", title: "", authors: "", projectId: null, date: "", description: "" };
    var form = h("div", { class: "form-grid" });
    var typeSel = h("select", { id: "a_type", style: "width:100%" });
    ACH_TYPES.forEach(function (t) { var o = h("option", { value: t }, t); if (t === a.type) o.selected = true; typeSel.appendChild(o); });
    form.appendChild(field("类型", typeSel));
    form.appendChild(field("日期", h("input", { type: "date", id: "a_date", value: a.date || "", style: "width:100%" })));
    form.appendChild(h("div", { class: "field field-full" }, [h("label", {}, "名称/标题"), h("input", { id: "a_title", value: a.title || "", style: "width:100%" })]));
    form.appendChild(h("div", { class: "field" }, [h("label", {}, "作者/发明人"), h("input", { id: "a_authors", value: a.authors || "", style: "width:100%" })]));
    var projSel = h("select", { id: "a_proj", style: "width:100%" });
    projSel.appendChild(h("option", { value: "" }, "未关联项目"));
    Store.data.projects.forEach(function (p) { var o = h("option", { value: p.id }, p.name); if (p.id === a.projectId) o.selected = true; projSel.appendChild(o); });
    form.appendChild(field("所属项目", projSel));
    form.appendChild(h("div", { class: "field field-full" }, [h("label", {}, "说明"), h("textarea", { id: "a_desc", style: "width:100%;min-height:60px" }, a.description || "")]));
    openModal(id ? "编辑成果" : "新增成果", form, { wide: true, foot: [
      { text: "取消", onclick: closeModal },
      { text: "保存", kind: "btn-primary", onclick: function () {
        var title = el("a_title").value.trim();
        if (!title) { toast("请填写名称/标题", "err"); return; }
        var data = { type: el("a_type").value, title: title, authors: el("a_authors").value.trim(), projectId: el("a_proj").value || null, date: el("a_date").value, description: el("a_desc").value };
        if (id) { var x = achievementById(id); Object.keys(data).forEach(function (k) { x[k] = data[k]; }); }
        else { data.id = uid(); Store.data.achievements.push(data); }
        Store.save(); closeModal(); Router.go("achievements", Router.params); toast("已保存", "ok");
      } }
    ] });
  }
  function delAchievement(id) {
    confirmModal("删除成果", "确认删除该成果记录？", function () {
      Store.data.achievements = Store.data.achievements.filter(function (a) { return a.id !== id; });
      Store.save(); toast("已删除", "ok"); Router.go("achievements", Router.params);
    });
  }
  function ensureTargets(p) {
    if (!p.targets) p.targets = {};
    ACH_TYPES.forEach(function (t) { if (p.targets[t] == null) p.targets[t] = 0; });
    return p.targets;
  }
  function countAchievements(pid, type) {
    return Store.data.achievements.filter(function (a) { return a.projectId === pid && a.type === type; }).length;
  }
  function targetsCard() {
    var card = h("div", { class: "card", style: "margin-bottom:16px" }, [
      h("div", { class: "card-head" }, [h("h2", {}, "立项考核指标对照（要求 vs 实际完成）"), h("button", { class: "btn btn-sm", onclick: function () { editTargets(null); } }, "批量编辑指标")]),
      h("div", { class: "card-body" })
    ]);
    var body = card.querySelector(".card-body");
    if (!Store.data.projects.length) { body.appendChild(h("div", { class: "empty" }, "暂无项目")); return card; }
    var wrap = h("div", { class: "table-wrap" });
    var tbl = h("table", { class: "tbl" });
    var head = "<thead><tr><th>项目名称</th>";
    ACH_TYPES.forEach(function (t) { head += "<th class='num'>" + t + "<br><span style='font-weight:400;color:var(--text-2);font-size:11px'>(要求/实际)</span></th>"; });
    head += "<th>操作</th></tr></thead>";
    tbl.innerHTML = head;
    var tb = h("tbody");
    Store.data.projects.forEach(function (p) {
      ensureTargets(p);
      var cells = [h("td", { class: "wrap" }, p.name)];
      ACH_TYPES.forEach(function (t) {
        var req = p.targets[t] || 0;
        var actual = countAchievements(p.id, t);
        var cls = actual >= req && req > 0 ? "tag-ok" : (req > 0 && actual < req ? "tag-warn" : "");
        cells.push(h("td", { class: "num" }, h("span", { class: "tag " + cls }, req + " / " + actual)));
      });
      cells.push(h("td", {}, h("button", { class: "btn btn-sm", onclick: function () { editTargets(p.id); } }, "编辑指标")));
      tb.appendChild(h("tr", {}, cells));
    });
    tbl.appendChild(tb); wrap.appendChild(tbl); body.appendChild(wrap);
    return card;
  }
  function editTargets(pid) {
    var form = h("div", {});
    var targets = pid ? ensureTargets(projectById(pid)) : null;
    Store.data.projects.forEach(function (p) {
      if (pid && p.id !== pid) return;
      ensureTargets(p);
      var fs = h("div", { class: "card", style: "margin-bottom:12px" }, [h("div", { class: "card-body" })]);
      var fb = fs.querySelector(".card-body");
      fb.appendChild(h("div", { class: "section-title" }, p.name));
      var grid = h("div", { class: "form-grid" });
      ACH_TYPES.forEach(function (t) {
        var actual = countAchievements(p.id, t);
        grid.appendChild(h("div", { class: "field" }, [
          h("label", {}, t + " 要求数量（实际：" + actual + " 项）"),
          h("input", { type: "number", min: "0", id: "tgt_" + p.id + "_" + t, value: p.targets[t] || 0, style: "width:100%" })
        ]));
      });
      fb.appendChild(grid);
      form.appendChild(fs);
    });
    openModal(pid ? "编辑考核指标" : "批量编辑考核指标", form, { wide: true, foot: [
      { text: "取消", onclick: closeModal },
      { text: "保存", kind: "btn-primary", onclick: function () {
        Store.data.projects.forEach(function (p) {
          if (pid && p.id !== pid) return;
          ACH_TYPES.forEach(function (t) {
            var inp = el("tgt_" + p.id + "_" + t);
            if (inp) p.targets[t] = Number(inp.value) || 0;
          });
        });
        Store.save(); closeModal(); Router.go("achievements", Router.params); toast("指标已保存", "ok");
      } }
    ] });
  }

  /* ---- 经费管理 ---- */
  // 叶子科目（可编辑）；equipment 父行=3子项之和；_direct/_total 为自动计算行
  var FUND_LEAVES = ["research", "equipmentPurchase", "equipmentTrial", "equipmentLease", "material", "testing", "fuel", "travel", "publishing", "labor", "expert", "other", "indirect"];
  var FUND_DIRECT = ["research", "equipment", "material", "testing", "fuel", "travel", "publishing", "labor", "expert", "other"];
  var FUND_ROWS = [
    { key: "research", name: "研究人员费" },
    { key: "equipment", name: "设备费", auto: "equip" },
    { key: "equipmentPurchase", name: "其中：设备购置费", sub: true },
    { key: "equipmentTrial", name: "设备试制费", sub: true },
    { key: "equipmentLease", name: "设备改造与租赁费", sub: true },
    { key: "material", name: "材料费" },
    { key: "testing", name: "测试化验加工费" },
    { key: "fuel", name: "燃料动力费" },
    { key: "travel", name: "差旅/会议/国际合作交流费" },
    { key: "publishing", name: "出版/文献/信息传播/知识产权事务费" },
    { key: "labor", name: "劳务费" },
    { key: "expert", name: "专家咨询费" },
    { key: "other", name: "其他支出" },
    { key: "_direct", name: "直接费用小计", auto: "direct" },
    { key: "indirect", name: "间接费用" },
    { key: "_total", name: "合计", auto: "total" }
  ];
  // 预算初始值（单位：万元，按项目名称关键词匹配，仅在该 projectId 无记录时初始化）
  var FUND_SEEDS = [
    { kw: "桥梁集群智慧感知", b: { research: 365, material: 150, testing: 50, fuel: 5, travel: 50, publishing: 30, labor: 0, expert: 30, other: 140, indirect: 145 } },
    { kw: "岩土结构", b: {} },
    { kw: "薄层罩面", b: { research: 360, equipmentPurchase: 30, material: 280, testing: 30, fuel: 20, travel: 30, publishing: 5, labor: 30, expert: 30, other: 5, indirect: 80 } },
    { kw: "跨尺度", b: { research: 320, equipmentPurchase: 100, equipmentTrial: 10, equipmentLease: 20, material: 30, testing: 55, fuel: 10, travel: 10, publishing: 50, labor: 50, expert: 20, other: 100, indirect: 135 } },
    { kw: "公路路桥多维", b: { research: 663, equipmentLease: 214, other: 73 } },
    { kw: "极端环境", b: { research: 500, travel: 230, publishing: 20, labor: 30, expert: 20, other: 100 } }
  ];
  function ensureFundBudget(pid) {
    if (!Store.data.fundBudgets) Store.data.fundBudgets = {};
    if (!Store.data.fundBudgets[pid]) {
      var p = projectById(pid);
      var b = {};
      FUND_LEAVES.forEach(function (k) { b[k] = 0; });
      if (p) for (var i = 0; i < FUND_SEEDS.length; i++) {
        if (p.name.indexOf(FUND_SEEDS[i].kw) >= 0) {
          for (var k in FUND_SEEDS[i].b) b[k] = FUND_SEEDS[i].b[k];
          break;
        }
      }
      b.equipment = (Number(b.equipmentPurchase) || 0) + (Number(b.equipmentTrial) || 0) + (Number(b.equipmentLease) || 0);
      Store.data.fundBudgets[pid] = b;
    }
    return Store.data.fundBudgets[pid];
  }
  function fundMonthVal(pid, year, mon, key) {
    if (!Store.data.fundRecords) return 0;
    var rec = Store.data.fundRecords[pid + "|" + year + "-" + ("0" + mon).slice(-2)];
    if (!rec) return 0;
    if (key === "equipment") return (Number(rec.equipmentPurchase) || 0) + (Number(rec.equipmentTrial) || 0) + (Number(rec.equipmentLease) || 0);
    return Number(rec[key]) || 0;
  }
  function fundRowMonthVal(pid, year, mon, row) {
    if (row.auto === "direct") { var s = 0; FUND_DIRECT.forEach(function (k) { s += fundMonthVal(pid, year, mon, k); }); return s; }
    if (row.auto === "total") return fundRowMonthVal(pid, year, mon, { auto: "direct" }) + fundMonthVal(pid, year, mon, "indirect");
    return fundMonthVal(pid, year, mon, row.key);
  }
  function fundRowBudget(bs, row) {
    if (row.auto === "direct") { var s = 0; FUND_DIRECT.forEach(function (k) { s += Number(bs[k]) || 0; }); return s; }
    if (row.auto === "total") return fundRowBudget(bs, { auto: "direct" }) + (Number(bs.indirect) || 0);
    if (row.key === "equipment") return (Number(bs.equipmentPurchase) || 0) + (Number(bs.equipmentTrial) || 0) + (Number(bs.equipmentLease) || 0);
    return Number(bs[row.key]) || 0;
  }

  VIEWS.funds = function (params) {
    params = params || {};
    var projects = Store.data.projects;
    var root = h("div");
    if (!projects.length) { root.appendChild(h("div", { class: "empty" }, "暂无项目，请先在「项目管理」中添加")); return root; }
    var pid = (params.pid && projectById(params.pid)) ? params.pid : projects[0].id;
    var year = Number(params.year) || Number(Store.data.settings.currentMonth.split("-")[0]);
    var p = projectById(pid);
    var bs = ensureFundBudget(pid);

    root.appendChild(h("div", { class: "toolbar" }, [
      h("div", { class: "left" }, [
        h("div", { class: "filter" }, [h("span", {}, "项目"), (function () {
          var s = h("select", { onchange: function (e) { Router.go("funds", { pid: e.target.value, year: year }); } });
          projects.forEach(function (pr) { var o = h("option", { value: pr.id }, pr.name); if (pr.id === pid) o.selected = true; s.appendChild(o); });
          return s;
        })()]),
        h("div", { class: "filter" }, [h("span", {}, "年度"), h("input", { type: "number", value: year, style: "width:90px", onchange: function (e) { Router.go("funds", { pid: pid, year: Number(e.target.value) || year }); } })]),
        h("span", { style: "color:#c0392b;font-size:13px;font-weight:600;margin-left:10px" }, "修改后请点击保存")
      ]),
      h("div", { class: "right" }, [
        h("button", { class: "btn", onclick: function () { editFundBudget(pid, year); } }, "预算设置"),
        h("button", { class: "btn btn-primary", onclick: function () { Store.save(); toast("已保存", "ok"); } }, "保存")
      ])
    ]));

    var wrap = h("div", { class: "table-wrap" });
    var tbl = h("table", { class: "tbl" });
    var head = "<thead><tr><th>预算科目</th><th class='num'>预算金额(万元)</th>";
    for (var mi = 1; mi <= 12; mi++) head += "<th class='num'>" + mi + "月(元)</th>";
    head += "<th class='num'>年度合计(元)</th><th class='num'>占预算比例%</th></tr></thead>";
    tbl.innerHTML = head;
    var tb = h("tbody");
    FUND_ROWS.forEach(function (row) {
      var isAuto = !!row.auto;
      var budget = fundRowBudget(bs, row);
      var cells = [h("td", row.sub ? { style: "padding-left:22px" } : {}, row.name)];
      cells.push(h("td", { class: "num" }, fmt(budget, 2)));
      var yearSum = 0;
      for (var mo = 1; mo <= 12; mo++) {
        (function (mon) {
          var v = fundRowMonthVal(pid, year, mon, row);
          yearSum += v;
          if (isAuto) { cells.push(h("td", { class: "num" }, v ? fmt(v, 2) : "")); return; }
          var ym = year + "-" + ("0" + mon).slice(-2);
          cells.push(h("td", { class: "num" }, h("input", {
            type: "number", step: "0.01", value: v || "", style: "width:78px;text-align:right",
            onchange: function (e) {
              if (!Store.data.fundRecords) Store.data.fundRecords = {};
              var k = pid + "|" + ym;
              if (!Store.data.fundRecords[k]) Store.data.fundRecords[k] = {};
              Store.data.fundRecords[k][row.key] = Number(e.target.value) || 0;
              Router.go("funds", { pid: pid, year: year }, true);
            }
          })));
        })(mo);
      }
      cells.push(h("td", { class: "num" }, fmt(yearSum, 2)));
      cells.push(h("td", { class: "num" }, budget > 0 ? (yearSum / (budget * 10000) * 100).toFixed(1) + "%" : "—"));
      tb.appendChild(h("tr", isAuto ? { style: "font-weight:600;background:#eef3fa" } : {}, cells));
    });
    tbl.appendChild(tb); wrap.appendChild(tbl); root.appendChild(wrap);

    root.appendChild(h("div", { style: "margin-top:10px" }, [
      h("button", { class: "btn", onclick: function () { exportFunds(p, bs, year); } }, "导出经费表")
    ]));
    root.appendChild(h("div", { class: "text-muted", style: "margin-top:8px;font-size:12px" }, "说明：预算金额单位为万元，各月归集金额单位为元；设备费=购置+试制+改造租赁之和；直接费用小计、合计为自动计算；占预算比例=年度合计÷(科目预算×10000)。月份单元格修改后请点击右上角「保存」按钮生效。"));
    return root;
  };

  function editFundBudget(pid, year) {
    var p = projectById(pid); if (!p) return;
    var bs = ensureFundBudget(pid);
    var form = h("div", {});
    form.appendChild(h("div", { class: "section-title" }, p.name + "（单位：万元）"));
    var grid = h("div", { class: "form-grid" });
    FUND_ROWS.forEach(function (row) {
      if (row.auto === "direct" || row.auto === "total") return;
      if (row.key === "equipment") {
        grid.appendChild(field(row.name, h("div", { class: "text-muted" }, "自动 = 3子项之和：" + fmt(fundRowBudget(bs, row), 2) + " 万元")));
        return;
      }
      grid.appendChild(field(row.name, h("input", { type: "number", step: "0.01", min: "0", id: "fb_" + row.key, value: Number(bs[row.key]) || 0, style: "width:100%" })));
    });
    form.appendChild(grid);
    openModal("预算设置", form, { wide: true, foot: [
      { text: "取消", onclick: closeModal },
      { text: "保存", kind: "btn-primary", onclick: function () {
        FUND_LEAVES.forEach(function (k) { var inp = el("fb_" + k); if (inp) bs[k] = Number(inp.value) || 0; });
        bs.equipment = (Number(bs.equipmentPurchase) || 0) + (Number(bs.equipmentTrial) || 0) + (Number(bs.equipmentLease) || 0);
        Store.save(); closeModal(); Router.go("funds", { pid: pid, year: year }, true); toast("预算已保存", "ok");
      } }
    ] });
  }

  function exportFunds(p, bs, year) {
    var head = ["预算科目", "预算金额(万元)"];
    for (var mi = 1; mi <= 12; mi++) head.push(mi + "月(元)");
    head.push("年度合计(元)", "占预算比例%");
    var rows = [[p.name + " - " + year + "年度经费表"], head];
    FUND_ROWS.forEach(function (row) {
      var budget = fundRowBudget(bs, row);
      var line = [row.name, budget.toFixed(2)];
      var yearSum = 0;
      for (var mon = 1; mon <= 12; mon++) { var v = fundRowMonthVal(p.id, year, mon, row); yearSum += v; line.push(v ? v.toFixed(2) : ""); }
      line.push(yearSum.toFixed(2));
      line.push(budget > 0 ? (yearSum / (budget * 10000) * 100).toFixed(1) + "%" : "—");
      rows.push(line);
    });
    var ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: head.length - 1 } }];
    var cols = [{ wch: 30 }, { wch: 12 }];
    for (var ci = 0; ci < 12; ci++) cols.push({ wch: 11 });
    cols.push({ wch: 13 }, { wch: 11 });
    ws["!cols"] = cols;
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, year + "年经费");
    download("经费表-" + p.name.slice(0, 20) + "-" + year + ".xlsx", wb);
    toast("已导出经费表", "ok");
  }

  /* ---- 创新小组联系人 ---- */
  // 数据结构：Store.data.contacts = { cols:[列名...], rows:[[单元格...]...] }
  function ensureContacts() {
    if (!Store.data.contacts || !Store.data.contacts.cols) Store.data.contacts = defaultContacts();
    return Store.data.contacts;
  }

  VIEWS.contacts = function () {
    var data = ensureContacts();
    var root = h("div");

    root.appendChild(h("div", { class: "toolbar" }, [
      h("div", { class: "left" }, h("div", { class: "section-title" }, "创新小组联系人")),
      h("div", { class: "right" }, [
        h("button", { class: "btn", onclick: function () {
          var row = data.cols.map(function () { return ""; });
          row[0] = String(data.rows.length + 1);
          data.rows.push(row);
          Router.go("contacts", {}, true);
        } }, "添加行"),
        h("button", { class: "btn", onclick: function () {
          var name = prompt("请输入新列名称：");
          if (!name || !name.trim()) return;
          data.cols.push(name.trim());
          data.rows.forEach(function (r) { r.push(""); });
          Router.go("contacts", {}, true);
        } }, "添加列"),
        h("button", { class: "btn btn-primary", onclick: function () { Store.save(); toast("已保存", "ok"); } }, "保存")
      ])
    ]));

    var wrap = h("div", { class: "table-wrap" });
    var tbl = h("table", { class: "tbl" });
    var trh = h("tr");
    data.cols.forEach(function (col, ci) {
      trh.appendChild(h("th", {}, h("div", { style: "display:flex;align-items:center;gap:4px" }, [
        h("input", { value: col, style: "width:100%;min-width:70px;border:1px solid transparent;background:transparent;font-weight:600", onchange: function (e) { data.cols[ci] = e.target.value; } }),
        h("button", { class: "btn btn-sm", title: "删除列", style: "padding:0 6px", onclick: function () {
          confirmModal("删除列", "确认删除列「" + col + "」？该列所有单元格数据将一并删除。", function () {
            data.cols.splice(ci, 1);
            data.rows.forEach(function (r) { r.splice(ci, 1); });
            Router.go("contacts", {}, true);
          });
        } }, "×")
      ])));
    });
    trh.appendChild(h("th", {}, "操作"));
    tbl.appendChild(h("thead", {}, trh));
    var tb = h("tbody");
    data.rows.forEach(function (row, ri) {
      var tr = h("tr");
      data.cols.forEach(function (col, ci) {
        tr.appendChild(h("td", {}, h("input", { value: row[ci] || "", style: "width:100%;border:none;background:transparent", onchange: function (e) { row[ci] = e.target.value; } })));
      });
      tr.appendChild(h("td", {}, h("button", { class: "btn btn-sm btn-danger", onclick: function () {
        confirmModal("删除行", "确认删除第 " + (ri + 1) + " 行？", function () {
          data.rows.splice(ri, 1);
          Router.go("contacts", {}, true);
        });
      } }, "删除")));
      tb.appendChild(tr);
    });
    tbl.appendChild(tb); wrap.appendChild(tbl); root.appendChild(wrap);

    root.appendChild(h("div", { class: "text-muted", style: "margin-top:8px;font-size:12px" }, "说明：修改单元格后请点击保存按钮生效。"));
    return root;
  };

  /* ---- 导入导出 ---- */
  VIEWS.io = function () {
    var root = h("div");
    root.appendChild(h("div", { class: "stat-grid" }, [
      h("div", { class: "card" }, [h("div", { class: "card-head" }, [h("h2", {}, "数据导出")]), h("div", { class: "card-body" }, [
        h("div", { class: "text-muted", style: "margin-bottom:10px" }, "导出为 Excel，与原模板字段一致，可直接归档/上报。"),
        h("div", { class: "row" }, [
          h("button", { class: "btn btn-primary", onclick: function () { exportAttendance(Store.data.settings.currentMonth, "all"); } }, "导出考勤表（本月）"),
          h("button", { class: "btn btn-primary", onclick: function () { exportMonthly(Store.data.settings.currentMonth); } }, "导出月度执行情况（本月）"),
          h("button", { class: "btn", onclick: exportAchievements }, "导出成果清单"),
          h("button", { class: "btn", onclick: exportAllJson }, "导出全量数据(JSON备份)")
        ])
      ])]),
      h("div", { class: "card" }, [h("div", { class: "card-head" }, [h("h2", {}, "数据导入")]), h("div", { class: "card-body" }, [
        h("div", { class: "text-muted", style: "margin-bottom:10px" }, "从模板 Excel 导入：按姓名匹配人员/项目，自动写入对应月份考勤或汇报。"),
        h("div", { class: "field" }, [h("label", {}, "导入考勤表（.xls/.xlsx）"), uploadInput(function (f) { importAttendance(f); })]),
        h("div", { class: "field" }, [h("label", {}, "导入月度执行情况（.xlsx）"), uploadInput(function (f) { importReport(f); })]),
        h("div", { class: "field" }, [h("label", {}, "恢复 JSON 备份"), uploadInput(function (f) { importJson(f); })])
      ])]),
      h("div", { class: "card" }, [h("div", { class: "card-head" }, [h("h2", {}, "系统设置")]), h("div", { class: "card-body" }, [
        h("div", { class: "field" }, [h("label", {}, "公司/分公司名称"), h("input", { id: "io_company", value: Store.data.settings.companyName || "", style: "width:100%", onchange: function (e) { Store.data.settings.companyName = e.target.value; Store.save(); el("companyName").textContent = e.target.value; } })]),
        h("div", { class: "row" }, [
          h("button", { class: "btn btn-primary", onclick: function () { Store.syncFromServer(function (ok) { if (ok) { syncTopbar(); Router.go("dashboard"); toast("已从服务器同步最新数据", "ok"); } else toast("同步失败，可能未连接服务器", "err"); }); } }, "从服务器同步数据"),
          h("button", { class: "btn", onclick: function () { var m = prompt("跳转到月份(YYYY-MM)", Store.data.settings.currentMonth); if (m) { Store.data.settings.currentMonth = monthKey(m); Store.save(); syncTopbar(); Router.go("dashboard"); } } }, "切换月份"),
          h("button", { class: "btn btn-danger", onclick: function () { confirmModal("重置数据", "将清空所有本地数据并恢复为模板初始状态，确认？", function () { Store.reset(); syncTopbar(); Router.go("dashboard"); toast("已重置", "ok"); }); } }, "重置为模板初始状态")
        ]),
        h("div", { class: "text-muted", style: "margin-top:10px;font-size:12px" }, "多人协作：所有数据通过服务器共享保存，多人访问同一地址可查看和编辑同一份数据。建议编辑前先点「从服务器同步数据」获取最新。")
      ])])
    ]));
    return root;
  };
  function uploadInput(onpick) {
    var inp = h("input", { type: "file", accept: ".xls,.xlsx,.json", style: "border:1px solid var(--border);border-radius:7px;padding:6px;width:100%;background:#fff" });
    inp.addEventListener("change", function () { if (inp.files[0]) onpick(inp.files[0]); inp.value = ""; });
    return inp;
  }

  /* ============================================================
   * Excel 导入导出（SheetJS）
   * ============================================================ */
  function s2ab(s) { var buf = new ArrayBuffer(s.length); var view = new Uint8Array(buf); for (var i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xff; return buf; }
  function download(name, wb) {
    var wbout = XLSX.write(wb, { bookType: "xlsx", type: "binary" });
    var blob = new Blob([s2ab(wbout)], { type: "application/octet-stream" });
    var a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name; a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
  }

  function exportAttendance(m, pid) {
    // 调用后端接口生成带完整样式的 Excel（字体、边框、合并单元格按模板格式）
    var url = "/api/export/attendance?month=" + encodeURIComponent(m) + "&pid=" + encodeURIComponent(pid || "all");
    var a = document.createElement("a");
    a.href = url;
    a.download = "研发人员考勤表-" + m + ".xlsx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    toast("已导出考勤表（按模板格式）", "ok");
  }

  function exportReport(m) { exportMonthly(m); }

  function exportAllJson() {
    var blob = new Blob([JSON.stringify(Store.data, null, 2)], { type: "application/json" });
    var a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = "科研项目管理-数据备份-" + new Date().toISOString().slice(0, 10) + ".json"; a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
    toast("已导出 JSON 备份", "ok");
  }

  function importAttendance(file) {
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var wb = XLSX.read(e.target.result, { type: "array" });
        var ws = wb.Sheets[wb.SheetNames[0]];
        var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
        // 表头在第4-6行，数据从第7行开始(0-index:6)。列：0序号 1姓名 2职责 3课题 4研发工时 5总工时 6..每日(rd,nrd)
        var monthSheet = wb.SheetNames[0];
        // 从模板标题行提取月份（第2行有时间：2026年8月）
        var targetMonth = Store.data.settings.currentMonth;
        for (var i = 0; i < Math.min(aoa.length, 6); i++) {
          var line = (aoa[i] || []).join(" ");
          var m = line.match(/(\d{4})\s*年\s*(\d{1,2})\s*月/);
          if (m) { targetMonth = m[1] + "-" + ("0" + m[2]).slice(-2); break; }
        }
        var n = daysInMonth(targetMonth);
        var imported = 0, unmatched = [];
        for (var r = 6; r < aoa.length; r++) {
          var row = aoa[r]; if (!row) continue;
          var name = String(row[1] || "").trim();
          if (!name || name === "合计") continue;
          var pe = findPersonByName(name);
          if (!pe) { unmatched.push(name); continue; }
          var att = ensureAtt(pe, targetMonth); att.days = {};
          for (var d = 1; d <= n; d++) {
            var ri = 6 + (d - 1) * 2, ni = 7 + (d - 1) * 2;
            var rd = Number(row[ri]) || 0, nrd = Number(row[ni]) || 0;
            att.days[String(d)] = { rd: rd, nonRd: nrd };
          }
          imported++;
        }
        Store.save();
        var msg = "已导入 " + imported + " 人 " + targetMonth + " 考勤";
        if (unmatched.length) msg += "；未匹配：" + unmatched.join("、") + "（请先在人员台账添加）";
        toast(msg, imported ? "ok" : "err");
        Router.go("attendance");
      } catch (err) { toast("导入失败：" + err.message, "err"); }
    };
    reader.readAsArrayBuffer(file);
  }
  function findPersonByName(name) {
    for (var i = 0; i < Store.data.personnel.length; i++) if (Store.data.personnel[i].name === name) return Store.data.personnel[i];
    return null;
  }
  function findProjectByName(name) {
    for (var i = 0; i < Store.data.projects.length; i++) if (Store.data.projects[i].name === name) return Store.data.projects[i];
    return null;
  }

  function importReport(file) {
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var wb = XLSX.read(e.target.result, { type: "array" });
        var ws = wb.Sheets[wb.SheetNames[0]];
        var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
        var target = Store.data.settings.currentMonth;
        var imported = 0, unmatched = [];
        for (var r = 1; r < aoa.length; r++) {
          var row = aoa[r]; if (!row) continue;
          var name = String(row[0] || "").trim();
          if (!name) continue;
          var p = findProjectByName(name);
          if (!p) { unmatched.push(name); continue; }
          var rep = ensureMonth(p, target);
          var period = String(row[1] || "");
          var pm = period.match(/(\d{4}\.\d{2})\s*-\s*(\d{4}\.\d{2})/);
          if (pm) { p.periodStart = pm[1].replace(".", "-"); p.periodEnd = pm[2].replace(".", "-"); }
          if (row[2]) p.leader = String(row[2]);
          rep.completedWork = String(row[3] || ""); rep.problems = String(row[4] || ""); rep.nextPlan = String(row[5] || "");
          rep.monthlyFunds = row[6] != null ? String(row[6]) : ""; rep.cumulativeFunds = row[7] != null ? String(row[7]) : "";
          imported++;
        }
        Store.save();
        var msg = "已导入 " + imported + " 个项目 " + target + " 汇报";
        if (unmatched.length) msg += "；未匹配：" + unmatched.join("、");
        toast(msg, imported ? "ok" : "err");
        Router.go("monthly");
      } catch (err) { toast("导入失败：" + err.message, "err"); }
    };
    reader.readAsArrayBuffer(file);
  }

  function importJson(file) {
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var data = JSON.parse(e.target.result);
        if (!data.projects || !data.personnel) throw new Error("格式不符");
        if (!data.achievements) data.achievements = [];
        Store.data = data; Store.save(); syncTopbar(); Router.go("dashboard"); toast("已恢复备份", "ok");
      } catch (err) { toast("恢复失败：" + err.message, "err"); }
    };
    reader.readAsText(file);
  }

  /* 通用上传触发 */
  function triggerUpload(handler, accept) {
    var inp = h("input", { type: "file", accept: accept || ".xls,.xlsx,.json", style: "display:none" });
    inp.addEventListener("change", function () { if (inp.files[0]) handler(inp.files[0]); inp.value = ""; });
    document.body.appendChild(inp); inp.click();
    setTimeout(function () { if (inp.parentNode) inp.parentNode.removeChild(inp); }, 5000);
  }

  /* 人员导入模板下载 */
  function downloadPersonnelTemplate() {
    var rows = [["姓名", "职责分工", "所属项目名称"]];
    Store.data.personnel.slice(0, 3).forEach(function (pe) {
      rows.push([pe.name, pe.responsibility || "", pe.projectId ? projectById(pe.projectId).name : ""]);
    });
    rows.push(["（示例行可删除）", "", ""]);
    rows.push(["张三", "算法开发", Store.data.projects[0] ? Store.data.projects[0].name : ""]);
    var ws = XLSX.utils.aoa_to_sheet(rows);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "人员清单");
    download("人员台账-导入模板.xlsx", wb);
    toast("模板已下载", "ok");
  }

  /* 人员Excel导入 */
  function importPersonnel(file) {
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var wb = XLSX.read(e.target.result, { type: "array" });
        var ws = wb.Sheets[wb.SheetNames[0]];
        var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
        var added = 0, updated = 0, noProj = 0;
        for (var r = 1; r < aoa.length; r++) {
          var row = aoa[r]; if (!row) continue;
          var name = String(row[0] || "").trim();
          if (!name || name.indexOf("示例") >= 0 || name.indexOf("（") === 0) continue;
          var resp = String(row[1] || "").trim();
          var projName = String(row[2] || "").trim();
          var pid = null;
          if (projName) { var p = findProjectByName(projName); if (p) pid = p.id; else noProj++; }
          var exist = findPersonByName(name);
          if (exist) { exist.responsibility = resp; exist.projectId = pid; updated++; }
          else { Store.data.personnel.push({ id: uid(), name: name, responsibility: resp, projectId: pid, attendance: {} }); added++; }
        }
        Store.save();
        var msg = "新增 " + added + " 人，更新 " + updated + " 人";
        if (noProj) msg += "；" + noProj + " 行项目未匹配（已设为未分配）";
        toast(msg, (added + updated) ? "ok" : "err");
        Router.go("personnel");
      } catch (err) { toast("导入失败：" + err.message, "err"); }
    };
    reader.readAsArrayBuffer(file);
  }

  /* 成果清单导出 */
  function exportAchievements() {
    var rows = [["序号", "类型", "名称/标题", "作者/发明人", "所属项目", "日期", "说明"]];
    Store.data.achievements.forEach(function (a, i) {
      rows.push([i + 1, a.type, a.title, a.authors, a.projectId ? projectById(a.projectId).name : "", a.date, a.description]);
    });
    var ws = XLSX.utils.aoa_to_sheet(rows);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "成果清单");
    download("研发成果清单-" + new Date().toISOString().slice(0, 10) + ".xlsx", wb);
    toast("已导出成果清单", "ok");
  }

  /* ---------- 确认弹窗 ---------- */
  function confirmModal(title, msg, onOk) {
    var body = h("div", {}, [h("div", { class: "muted-block" }, msg)]);
    openModal(title, body, { foot: [
      { text: "取消", onclick: closeModal },
      { text: "确认", kind: "btn-primary", onclick: function () { closeModal(); onOk(); } }
    ] });
  }

  /* ---------- 顶栏同步 ---------- */
  function syncTopbar() {
    var s = Store.data.settings;
    el("currentMonth").value = s.currentMonth;
    el("companyName").textContent = s.companyName;
  }

  /* ---------- 初始化 ---------- */
  function init() {
    Store.load();
    if (!Store.data.achievements) Store.data.achievements = [];
    if (!Store.data.fundBudgets) Store.data.fundBudgets = {};
    if (!Store.data.fundRecords) Store.data.fundRecords = {};
    if (!Store.data.contacts || !Store.data.contacts.cols) Store.data.contacts = defaultContacts();
    if (!Store.data.settings.signatures) Store.data.settings.signatures = {};
    syncTopbar();
    // 从服务器同步最新共享数据，成功后刷新视图
    Store.syncFromServer(function (ok) {
      if (ok) { syncTopbar(); Router.go(Router.view, Router.params); toast("已同步共享数据", "ok"); }
      else { Router.go("dashboard"); }
    });
    el("currentMonth").addEventListener("change", function (e) {
      var prev = Store.data.settings.currentMonth;
      Store.data.settings.currentMonth = monthKey(e.target.value); Store.save();
      if (prev !== Store.data.settings.currentMonth) toast("已切换到 " + Store.data.settings.currentMonth + "，其他月份数据未受影响", "ok");
      if (Router.view === "dashboard" || Router.view === "projects" || Router.view === "personnel") Router.go(Router.view);
      else Router.go(Router.view, Router.params);
    });
    Array.prototype.forEach.call(el("nav").querySelectorAll(".nav-item"), function (a) {
      a.addEventListener("click", function () { Router.go(a.dataset.view); });
    });
    el("modalClose").addEventListener("click", closeModal);
    el("modalMask").addEventListener("click", function (e) { if (e.target === el("modalMask")) closeModal(); });
    el("menuToggle").addEventListener("click", function () { el("sidebar").classList.toggle("open"); });
    el("navBack").addEventListener("click", Router.back);
    el("navFwd").addEventListener("click", Router.fwd);
    document.addEventListener("keydown", function (e) { if (e.key === "Escape") closeModal(); });
    Router.go("dashboard");
  }
  if (document.readyState !== "loading") init();
  else document.addEventListener("DOMContentLoaded", init);
})();
