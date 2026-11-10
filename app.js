/* منصة ارتكاز — متابعة أعمال المجالات — الواجهة */
(function () {
  'use strict';
  document.documentElement.lang = 'ar';
  document.documentElement.dir = 'rtl';

  var C = window.CONFIG, D = window.DOMAINS, IND = window.INDICATORS, API = window.API;
  var FM = !!(C.FORM_URL && String(C.FORM_URL).trim()); // وضع Microsoft Forms
  var FRAMED = (function () { try { return window.self !== window.top; } catch (e) { return true; } })(); // داخل إطار معاينة لا يسمح بتضمين Excel
  var $ = function (s, r) { return (r || document).querySelector(s); };
  var esc = function (s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  };
  // يحفظ ترتيب «1-4» داخل نص عربي
  function ltrNums(t) { return esc(t).replace(/(\d+\s*[-–]\s*\d+)/g, '<bdi dir="ltr">$1</bdi>'); }
  var LS = {
    get: function (k, d) { try { var v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
    set: function (k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} }
  };

  var state = {
    works: [], loaded: false, error: null,
    tab: null, indFilter: null, form: null,
    sub: 'timeline', f: { c: '', m: '', q: '', ind: '' },
    rep: { c: '', empty: true }, lb: null
  };

  /* ---------- أدوات التاريخ والعدّ ---------- */
  var LOC = 'ar-OM-u-nu-latn';
  function todayISO() { return new Date().toLocaleDateString('en-CA'); }
  function asDate(iso) { return new Date(String(iso).length === 10 ? iso + 'T00:00:00' : iso); }
  function fmtDate(iso) {
    if (!iso) return '';
    var d = asDate(iso);
    return isNaN(d) ? esc(iso) : d.toLocaleDateString(LOC, { day: 'numeric', month: 'long', year: 'numeric' });
  }
  function monthKey(iso) { return String(iso || '').slice(0, 7); }
  function monthLabel(k) { return asDate(k + '-01').toLocaleDateString(LOC, { month: 'long', year: 'numeric' }); }
  function daysSince(iso) { return Math.floor((Date.now() - new Date(iso).getTime()) / 86400000); }
  function relTime(iso) {
    if (!iso) return 'لا يوجد';
    var n = daysSince(iso);
    if (n <= 0) return 'اليوم';
    if (n === 1) return 'أمس';
    if (n === 2) return 'قبل يومين';
    if (n <= 10) return 'قبل ' + n + ' أيام';
    return 'قبل ' + n + ' يومًا';
  }
  function pl(n, w) { // w = [مفرد, مثنى, جمع, تمييز]
    if (n === 0) return 'لا ' + w[2];
    if (n === 1) return w[0] + ' واحد';
    if (n === 2) return w[1];
    if (n <= 10) return n + ' ' + w[2];
    return n + ' ' + w[3];
  }
  var W_WORK = ['عمل', 'عملان', 'أعمال', 'عملًا'];

  /* ---------- مساعدات البيانات ---------- */
  function dom(n) { return D[n - 1]; }
  function byDate(a, b) { return (b.date || '').localeCompare(a.date || '') || (b.created || '').localeCompare(a.created || ''); }
  function worksOf(n) { return state.works.filter(function (w) { return +w.committee === n; }).sort(byDate); }
  function covSet(list) { var s = new Set(); list.forEach(function (w) { s.add(w.indicator); }); return s; }
  function indCount(list, code) { return list.filter(function (w) { return w.indicator === code; }).length; }
  function lastEntry(list) { return list.reduce(function (m, w) { return (w.created || '') > m ? w.created : m; }, ''); }
  function status(n) {
    var l = worksOf(n);
    if (!l.length) return { k: 'idle', t: 'لم يبدأ الإدخال' };
    var d = daysSince(lastEntry(l));
    if (d > C.INACTIVE_DAYS) return { k: 'warn', t: 'لا إدخال منذ ' + d + ' يومًا' };
    return { k: 'ok', t: 'نشط' };
  }
  function safeUrl(u) { u = String(u || ''); return /^(https?:|data:image\/)/i.test(u) ? u : '#'; }
  function ext(name) { var m = /\.([a-z0-9]{1,5})$/i.exec(name || ''); return m ? m[1].toUpperCase() : 'ملف'; }

  /* ---------- الأيقونات ---------- */
  var ICONS = {
    home: '<path d="M3 11l9-7 9 7v9a1 1 0 0 1-1 1h-5v-6H9v6H4a1 1 0 0 1-1-1z"/>',
    d1: '<path d="M2 9l10-5 10 5-10 5z"/><path d="M6 11v5c3 2.2 9 2.2 12 0v-5"/><path d="M22 9v6"/>',
    d2: '<path d="M12 21v-9"/><path d="M12 12C12 7.5 8.5 5 4 5c0 4.8 3.4 7 8 7z"/><path d="M12 10c0-3.8 3-6 8-6 0 4-3.2 6-8 6z"/>',
    d3: '<rect x="3" y="4" width="18" height="12" rx="2"/><path d="M8 20l4-4 4 4"/><path d="M7 9h6M7 12h9"/>',
    d4: '<path d="M3 21h18"/><path d="M5 21V9.5L12 4l7 5.5V21"/><path d="M10 21v-5h4v5"/><path d="M9 11h.5M14.5 11h.5"/>',
    d5: '<path d="M5 21V4"/><path d="M5 4h11l-2 4 2 4H5"/>',
    eye: '<path d="M2 12s3.6-7 10-7 10 7 10 7-3.6 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/>',
    moon: '<path d="M20 14.5A8.5 8.5 0 1 1 9.5 4a7 7 0 0 0 10.5 10.5z"/>',
    sun: '<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/>',
    book: '<path d="M4 5a2 2 0 0 1 2-2h13v16H6a2 2 0 0 0-2 2z"/><path d="M4 21V5"/><path d="M8 7h7M8 11h5"/>',
    list: '<path d="M9 6h11M9 12h11M9 18h11"/><path d="M4 6h.5M4 12h.5M4 18h.5"/>',
    target: '<circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="5"/><circle cx="12" cy="12" r="1.2"/>',
    cal: '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/>',
    bell: '<path d="M6 16V11a6 6 0 0 1 12 0v5l2 2H4z"/><path d="M10 21h4"/>',
    clock: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
    refresh: '<path d="M20 11a8 8 0 1 0-2.3 5.7"/><path d="M20 4v7h-7"/>',
    arrow: '<path d="M15 6l-6 6 6 6"/>'
  };
  function ic(name) {
    return '<svg class="svg-ic" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + (ICONS[name] || '') + '</svg>';
  }
  function tile(name, dc) { return '<span class="itile"' + (dc ? ' style="--dc:' + dc + '"' : '') + '>' + ic(name) + '</span>'; }

  function renderNav() {
    var t = state.tab;
    var item = function (href, key, icon, label, dc) {
      var on = t === key;
      return '<a href="' + href + '" class="nav-i' + (on ? ' on' : '') + '"' + (on ? ' aria-current="page"' : '') + (dc ? ' style="--dc:' + dc + '"' : '') + '>' +
        '<span class="nav-ic">' + ic(icon) + '</span><span class="nav-t">' + label + '</span></a>';
    };
    $('#sideNav').innerHTML = item('#home', 'home', 'home', 'الرئيسية') +
      '<p class="nav-sec">المجالات</p>' +
      D.map(function (d) { return item('#c' + d.n, 'c' + d.n, 'd' + d.n, esc(d.name), 'var(--d' + d.n + ')'); }).join('') +
      '<p class="nav-sec">المتابعة</p>' + item('#mudeera', 'mudeera', 'eye', 'الفريق الإشرافي');
  }

  /* ---------- الهيكل العام ---------- */
  function renderChrome() {
    $('#schoolName').textContent = C.SCHOOL_NAME;
    var g = $('#schoolGrades'); if (g) g.innerHTML = ltrNums(C.SCHOOL_GRADES || '');
    var b = $('#demoBanner');
    if (API.isDemo && !FM) {
      b.hidden = false;
      b.innerHTML = '<div class="demo-in"><p><b>وضع تجريبي.</b> البيانات محفوظة في هذا المتصفح فقط، والأعمال المعروضة أمثلة. للحفظ المشترك ضعي رابط Apps Script في ملف <code>js/config.js</code>.</p>' +
        '<button type="button" class="btn small ghost-on" data-act="clearDemo">مسح البيانات التجريبية</button></div>';
    }
  }

  function render() {
    renderNav();
    var v = $('#view');
    if (!state.loaded) { v.innerHTML = '<div class="page"><p class="loading">جارٍ تحميل الأعمال…</p></div>'; return; }
    if (state.error) {
      v.innerHTML = '<div class="page"><div class="err-box"><h2>تعذّر تحميل الأعمال</h2><p>' + esc(state.error) +
        '</p><button class="btn primary" data-act="refresh">إعادة المحاولة</button></div></div>';
      return;
    }
    v.innerHTML = state.tab === 'home' ? homeHTML() : state.tab === 'guide' ? guideHTML() : state.tab === 'mudeera' ? (FM ? fmSupervisorHTML() : mudeeraHTML()) : (FM ? fmDomainHTML(+state.tab.slice(1)) : committeeHTML(+state.tab.slice(1)));
  }

  function route() {
    var h = location.hash.replace('#', '');
    if (!/^(c[1-5]|mudeera|guide)$/.test(h)) h = 'home';
    if (h !== state.tab) { state.indFilter = null; state.form = null; }
    state.tab = h;
    if (/^c/.test(h)) LS.set('tab', h);
    render();
    window.scrollTo(0, 0);
  }

  async function load() {
    if (FM) { state.works = []; state.error = null; state.loaded = true; render(); return; }
    try { state.works = await API.list(); state.error = null; }
    catch (e) { state.error = e.message || 'خطأ غير معروف'; }
    state.loaded = true;
    render();
  }

  /* ---------- بطاقة العمل ---------- */
  function attsHTML(w) {
    var a = w.attachments || [], l = w.links || [];
    if (!a.length && !l.length) return '';
    var imgs = a.filter(function (x) { return x.thumb; });
    var files = a.filter(function (x) { return !x.thumb; });
    return '<div class="atts">' +
      imgs.map(function (x) {
        return '<button type="button" class="thumb" data-act="lb" data-w="' + esc(w.id) + '" data-a="' + esc(x.id) + '" aria-label="عرض ' + esc(x.name) + '">' +
          '<img src="' + esc(safeUrl(x.thumb)) + '" alt="" loading="lazy"><span class="thumb-alt">' + esc(x.name) + '</span></button>';
      }).join('') +
      files.map(function (x) {
        return '<a class="fchip" href="' + esc(safeUrl(x.url)) + '" target="_blank" rel="noopener"><span class="ext">' + esc(ext(x.name)) + '</span><span class="fname">' + esc(x.name) + '</span></a>';
      }).join('') +
      l.map(function (u) {
        var host = u; try { host = new URL(u).hostname.replace(/^www\./, ''); } catch (e) {}
        return '<a class="fchip" href="' + esc(safeUrl(u)) + '" target="_blank" rel="noopener"><span class="ext">رابط</span><span class="fname" dir="ltr">' + esc(host) + '</span></a>';
      }).join('') + '</div>';
  }

  function workCard(w, opts) {
    var ind = IND[w.indicator], d = dom(+w.committee);
    return '<article class="work" style="--dc:var(--d' + w.committee + ')">' +
      '<div class="work-meta"><span class="code" title="' + esc(ind ? ind.name : '') + '">' + esc(w.indicator) + '</span>' +
      (opts && opts.showCommittee && d ? '<span class="cchip">' + esc(d.committee) + '</span>' : '') +
      '<time datetime="' + esc(w.date) + '">' + fmtDate(w.date) + '</time></div>' +
      '<h3>' + esc(w.title) + '</h3>' +
      (ind ? '<p class="ind-name">' + esc(ind.name) + '</p>' : '') +
      (w.desc ? '<p class="desc">' + esc(w.desc) + '</p>' : '') +
      attsHTML(w) +
      '<div class="work-foot"><span>أدخلته: <b>' + esc(w.author) + '</b>' + (w.updated ? ' · عُدّل ' + relTime(w.updated) : '') + '</span>' +
      (opts && opts.editable ? '<button type="button" class="btn small ghost" data-act="edit" data-id="' + esc(w.id) + '">تعديل</button>' : '') +
      '</div></article>';
  }

  /* ---------- الصفحة الرئيسية ---------- */
  var BACK = '<a href="#home" class="back">' + ic('arrow').replace('<svg', '<svg style="transform:scaleX(-1)"') + 'الرئيسية</a>';
  var HERO_ART = '<div class="hero-logos"><div class="hero-logo"><img src="moe-logo.png" alt="شعار وزارة التعليم" width="266" height="320"></div>' +
    '<div class="hero-logo vision"><img src="vision2040.png" alt="شعار رؤية عُمان 2040" onerror="this.parentNode.remove()"></div></div>';

  function homeHTML() {
    var mine = LS.get('tab', '');
    var total = state.works.length;
    var hero = '<section class="hero"><div class="hero-txt"><p class="eyebrow light">منصة ارتكاز</p>' +
      '<h1>أهلًا بكِ في متابعة مجالات التقويم المدرسي</h1><p>اختاري المجال لتسجيل أعماله المنجزة وربطها بمؤشرات تقويم الأداء.</p>' +
      '<div class="hero-actions">' + (FM ? '<a class="btn primary" href="' + esc(C.FORM_URL) + '" target="_blank" rel="noopener">' + ic('list') + 'تسجيل عمل جديد</a>' : '<span class="hero-meta">' + pl(total, W_WORK) + ' مسجلة حتى الآن</span>') +
      '<a class="btn guide" href="#guide">' + ic('book') + 'دليل الاستخدام</a></div></div>' + HERO_ART + '</section>';
    var cards = D.map(function (d) {
      var l = worksOf(d.n);
      return '<a href="#c' + d.n + '" class="hcard" style="--dc:var(--d' + d.n + ');--tint:var(--t' + d.n + ')">' +
        '<span class="hcard-top">' + tile('d' + d.n) + (mine === 'c' + d.n ? '<span class="hcard-mine">مجالي</span>' : '<span class="hcard-n">' + d.n + '</span>') + '</span>' +
        '<span class="hcard-name">' + esc(d.committee) + '</span>' +
        '<span class="hcard-def">' + esc(d.def) + '</span>' +
        (FM ? '<span class="hcard-meta">' + d.standards.length + ' معايير · ' + d.count + ' مؤشرًا</span>' : '<span class="hcard-meta"><b>' + l.length + '</b> ' + (l.length === 1 ? 'عمل' : l.length === 2 ? 'عملان' : l.length >= 3 && l.length <= 10 ? 'أعمال' : 'عملًا') + ' · ' + d.count + ' مؤشرًا</span>') +
        '<span class="hcard-go" aria-hidden="true">' + ic('arrow') + '</span></a>';
    }).join('');
    cards += '<a href="#mudeera" class="hcard hcard-md"><span class="hcard-top">' + tile('eye', 'var(--brand-soft)') + '</span>' +
      '<span class="hcard-name">الفريق الإشرافي</span><span class="hcard-meta">متابعة المجالات والتقرير</span>' +
      '<span class="hcard-go" aria-hidden="true">' + ic('arrow') + '</span></a>';
    return '<div class="page">' + hero + '<h2 class="sec-title">مجالات تقويم أداء المدارس</h2><div class="hgrid">' + cards + '</div></div>';
  }

  /* ---------- صفحة المجال ---------- */
  function committeeHTML(n) {
    var d = dom(n), list = worksOf(n), cov = covSet(list);
    var shown = state.indFilter ? list.filter(function (w) { return w.indicator === state.indFilter; }) : list;
    var ordinal = ['الأول', 'الثاني', 'الثالث', 'الرابع', 'الخامس'][n - 1];

    var head = '<section class="cm-head" style="--dc:var(--d' + n + ');--tint:var(--t' + n + ')">' + BACK +
      '<div class="cm-head-in"><div class="cm-title">' + tile('d' + n) + '<div><p class="eyebrow">المجال ' + ordinal + '</p><h1>' + esc(d.committee) + '</h1>' +
      '<p class="sub">' + d.standards.length + ' معايير · ' + d.count + ' مؤشرًا</p></div></div>' +
      '<p class="cm-def">' + esc(d.def) + '</p>' +
      '<dl class="cm-stats"><div><dt>' + ic('list') + 'الأعمال المسجلة</dt><dd>' + list.length + '</dd></div>' +
      '<div><dt>' + ic('target') + 'مؤشرات لها أعمال</dt><dd>' + cov.size + '<small> / ' + d.count + '</small></dd></div>' +
      '<div><dt>' + ic('clock') + 'آخر إدخال</dt><dd class="txt">' + relTime(lastEntry(list)) + '</dd></div></dl></div></section>';

    var toolbar = '<div class="toolbar">' +
      (state.form ? '' : '<button type="button" class="btn primary" data-act="add"><span class="plus" aria-hidden="true"></span>تسجيل عمل جديد</button>') +
      (state.indFilter ? '<span class="filter-chip">المؤشر ' + esc(state.indFilter) + '<button type="button" data-act="clearInd" aria-label="إلغاء التصفية">×</button></span>' : '') +
      '</div>';

    var works = shown.length ? shown.map(function (w) { return workCard(w, { editable: true }); }).join('') :
      '<div class="empty"><p>' + (state.indFilter ? 'لا أعمال مسجلة لهذا المؤشر بعد.' : 'لم تُسجَّل أعمال لهذا المجال بعد.') + '</p>' +
      (state.form ? '' : '<p>ابدئي بزر «تسجيل عمل جديد».</p>') + '</div>';

    var aside = '<aside class="cm-aside"><details open><summary><h2>مؤشرات المجال</h2><span class="hint">اضغطي مؤشرًا لعرض أعماله</span></summary>' +
      d.standards.map(function (s) {
        return '<div class="std"><h3><span>' + s.code + '</span> ' + esc(s.name) + '</h3><ul>' +
          s.indicators.map(function (i) {
            var c = indCount(list, i.code);
            return '<li><button type="button" class="ind' + (state.indFilter === i.code ? ' on' : '') + (c ? '' : ' zero') + '" data-act="ind" data-ind="' + i.code + '">' +
              '<span class="ind-code">' + i.code + '</span><span class="ind-text">' + esc(i.name) + '</span><span class="ind-n">' + c + '</span></button></li>';
          }).join('') + '</ul></div>';
      }).join('') + '</details></aside>';

    return '<div class="page">' + head + '<div class="cm-grid" style="--dc:var(--d' + n + ')"><div class="cm-main">' + toolbar +
      (state.form ? formHTML(n) : '') + '<div class="works">' + works + '</div></div>' + aside + '</div></div>';
  }

  /* ---------- وضع Microsoft Forms: صفحة المجال ---------- */
  function fmDomainHTML(n) {
    var d = dom(n);
    var ordinal = ['الأول', 'الثاني', 'الثالث', 'الرابع', 'الخامس'][n - 1];
    var head = '<section class="cm-head" style="--dc:var(--d' + n + ');--tint:var(--t' + n + ')">' + BACK +
      '<div class="cm-head-in"><div class="cm-title">' + tile('d' + n) + '<div><p class="eyebrow">المجال ' + ordinal + '</p><h1>' + esc(d.committee) + '</h1>' +
      '<p class="sub">' + d.standards.length + ' معايير · ' + d.count + ' مؤشرًا</p></div></div></div>' +
      '<p class="cm-def">' + esc(d.def) + '</p></section>';
    var steps = '<section class="fm-card" style="--dc:var(--d' + n + ')">' +
      '<div class="fm-card-h">' + tile('list') + '<div><h2>تسجيل عمل منجز</h2><p class="fm-sub">التسجيل في نموذج Microsoft Forms ببريد الوزارة</p></div></div>' +
      '<ol class="fm-steps">' +
      '<li>اضغطي <b>تسجيل عمل جديد</b>، فيفتح النموذج في صفحة جديدة.</li>' +
      '<li>سجّلي الدخول ببريد الوزارة إن طُلب منكِ، واسمكِ يُسجَّل تلقائيًا.</li>' +
      '<li>اختاري في السؤال الأول: <b>' + esc(d.name) + '</b>، ثم المؤشر المرتبط بالعمل.</li>' +
      '<li>اكتبي العنوان والتاريخ والوصف، وارفعي الصور أو الملفات، ثم <b>إرسال</b>.</li>' +
      '</ol>' +
      '<div class="fm-actions"><a class="btn primary" href="' + esc(C.FORM_URL) + '" target="_blank" rel="noopener"><span class="plus" aria-hidden="true"></span>تسجيل عمل جديد</a>' +
      '<span class="fm-note">يمكنكِ تعديل ما أرسلتِه لاحقًا من رابط التعديل الذي يظهر بعد الإرسال.</span></div>' +
      '</section>';
    var works = '';
    if (C.WORKBOOK_BASE) {
      var sh = encodeURIComponent("'D" + n + "'!A1");
      var viewUrl = C.WORKBOOK_BASE + '&action=view&activeCell=' + sh;
      var embedUrl = C.WORKBOOK_BASE + '&action=embedview&wdAllowInteractivity=False&Item=' + encodeURIComponent("'D" + n + "'!A1:G300") +
        '&wdHideGridlines=True&wdHideHeaders=True&wdDownloadButton=False&wdInConfigurator=True';
      works = '<section class="fm-card fm-works" style="--dc:var(--d' + n + ')">' +
        '<div class="fm-card-h">' + tile('eye') + '<div><h2>أعمال المجال</h2><p class="fm-sub">كل ما سُجّل في هذا المجال، الأحدث أولًا، ويتحدث تلقائيًا. يُفتح ببريد الوزارة.</p></div></div>' +
        (FRAMED ? '' : '<div class="fm-embed"><iframe src="' + esc(embedUrl) + '" title="أعمال ' + esc(d.committee) + '" loading="lazy"></iframe></div>') +
        '<div class="fm-actions"><a class="btn ghost" href="' + esc(viewUrl) + '" target="_blank" rel="noopener">' + ic('eye') + 'فتح القائمة في Excel</a></div>' +
        '</section>';
    }
    var aside = '<aside class="cm-aside fm-inds"><details open><summary><h2>مؤشرات المجال</h2><span class="hint">ستختارين واحدًا منها في النموذج</span></summary>' +
      d.standards.map(function (s) {
        return '<div class="std"><h3><span>' + s.code + '</span> ' + esc(s.name) + '</h3><ul>' +
          s.indicators.map(function (i) {
            return '<li><div class="ind static"><span class="ind-code">' + i.code + '</span><span class="ind-text">' + esc(i.name) + '</span></div></li>';
          }).join('') + '</ul></div>';
      }).join('') + '</details></aside>';
    return '<div class="page">' + head + '<div class="cm-grid" style="--dc:var(--d' + n + ')"><div class="cm-main">' + steps + works + '</div>' + aside + '</div></div>';
  }

  /* ---------- دليل الاستخدام ---------- */
  function guideHTML() {
    function card(icon, title, body) {
      return '<section class="fm-card"><div class="fm-card-h">' + tile(icon) + '<div><h2>' + title + '</h2></div></div>' + body + '</section>';
    }
    function ul(items) { return '<ul class="fm-steps">' + items.map(function (x) { return '<li>' + x + '</li>'; }).join('') + '</ul>'; }
    function ol(items) { return '<ol class="fm-steps">' + items.map(function (x) { return '<li>' + x + '</li>'; }).join('') + '</ol>'; }
    var head = '<section class="hero md-hero"><div class="hero-txt">' + BACK.replace('class="back"', 'class="back on-dark"') +
      '<p class="eyebrow light">منصة ارتكاز</p><h1>دليل الاستخدام</h1>' +
      '<p>تُسجَّل أعمال المجالات الخمسة عبر نموذج واحد، ويتابعها الفريق الإشرافي من هذه المنصة.</p>' +
      '<div class="hero-actions"><a class="btn guide" href="guide.pdf" download="دليل استخدام منصة ارتكاز.pdf">' + ic('book') + 'تنزيل الدليل PDF</a></div></div>' + HERO_ART + '</section>';
    var who = card('eye', 'من يستخدم المنصة', ul([
      '<b>مسؤولو المجالات والمعلمات:</b> تسجيل الأعمال المنجزة عبر زر «تسجيل عمل جديد».',
      '<b>الفريق الإشرافي (المديرة):</b> الاطلاع والمتابعة فقط، من صفحة «الفريق الإشرافي».',
      'كل الأعمال محفوظة في حساب المدرسة <bdi dir="ltr">sc3212@moe.om</bdi>.']));
    var reg = card('list', 'تسجيل عمل جديد', ol([
      'اضغط «تسجيل عمل جديد» في الرئيسية أو في صفحة المجال.',
      'سجّل الدخول بحساب الوزارة <bdi dir="ltr">(@moe.om)</bdi> إن طُلب منك، فيُسجَّل اسمك تلقائيًا.',
      'اختر <b>المجال</b> ثم «التالي».',
      'اختر <b>المؤشر</b> الذي يخدمه العمل؛ تظهر مؤشرات المجال المختار وحدها.',
      'اكتب عنوان العمل وتاريخ التنفيذ (مطلوبان)، ثم وصفًا مختصرًا، وأرفق الصور والملفات (حتى 10 ملفات، 10 ميجابايت للملف) أو الروابط.',
      'اضغط «إرسال»، فيصل العمل فورًا إلى النتائج ولوحة المتابعة.']) +
      '<p class="fm-sub">العمل الذي يخدم أكثر من مؤشر يُسجَّل مرة لكل مؤشر.</p>');
    var sup = card('cal', 'المتابعة للفريق الإشرافي', ul([
      '<b>نتائج النموذج:</b> ماذا سُجّل ومن سجّله؛ كل عمل بتفاصيله وصوره ومرفقاته.',
      '<b>لوحة المتابعة في Excel:</b> أين نقف وما الناقص؛ إجمالي الأعمال، أعمال الشهر، رسم المجالات، رسم الأشهر الستة، والمؤشرات الـ 78.',
      'المؤشر الذي أمامه <b>0</b> في لوحة Excel لم يُسجَّل له عمل بعد.',
      'يصل إلى بريد المدرسة إشعار مع كل عمل جديد، واللوحة تتحدث وحدها.']));
    var faq = card('bell', 'أسئلة شائعة', ul([
      '<b>أخطأت في عمل أرسلته؟</b> أبلغ الإدارة لتحذفه من النتائج، ثم سجّله من جديد.',
      '<b>لا يقبل النموذج رفع الملفات؟</b> تأكد أنك داخل بحساب الوزارة، وأن الملف أقل من 10 ميجابايت.',
      '<b>الصور:</b> يُفضَّل ألا تُكتب أسماء الطلبة، وألا تُرفع صور تظهر فيها وجوههم بوضوح.',
      '<b>لا تفتح النتائج أو لوحة Excel؟</b> تُفتحان بحساب المدرسة فقط، حمايةً للبيانات.']));
    return '<div class="page md">' + head + who + reg + sup + faq + '</div>';
  }

  /* ---------- وضع Microsoft Forms: الفريق الإشرافي ---------- */
  function fmSupervisorHTML() {
    var head = '<section class="hero md-hero"><div class="hero-txt">' + BACK.replace('class="back"', 'class="back on-dark"') +
      '<p class="eyebrow light">اطلاع فقط</p><h1>الفريق الإشرافي</h1>' +
      '<p>متابعة المجالات الخمسة والتقرير. الأعمال محفوظة في حساب المدرسة، وتُفتح ببريد المدرسة.</p></div>' + HERO_ART + '</section>';
    var excel = C.EXCEL_URL && String(C.EXCEL_URL).trim();
    var cards = '<div class="fm-links">' +
      '<a class="fm-link k1" href="' + esc(C.FORM_RESULTS_URL || C.FORM_URL) + '" target="_blank" rel="noopener">' + tile('eye', 'var(--d1)') +
        '<b>نتائج النموذج</b><span>عدد الأعمال، ورسم توزيعها على المجالات والمؤشرات، وكل عمل بتفاصيله وصوره.</span><em>يفتح في Forms ←</em></a>' +
      (excel ? '<a class="fm-link k2" href="' + esc(C.EXCEL_URL) + '" target="_blank" rel="noopener">' : '<div class="fm-link k2 off">') + tile('cal', 'var(--d3)') +
        '<b>لوحة المتابعة في Excel</b><span>الأعمال خلال الأشهر، وتوزيعها على المجالات، والمؤشرات التي لم يُسجَّل لها عمل بعد.</span><em>' + (excel ? 'يفتح في Excel ←' : 'قيد التجهيز') + '</em>' + (excel ? '</a>' : '</div>') +
      '</div>';
    var note = '<section class="fm-card"><div class="fm-card-h">' + tile('bell') + '<div><h2>إشعار بكل عمل جديد</h2>' +
      '<p class="fm-sub">يصل إلى بريد المدرسة إشعار مع كل عمل تسجّله معلمة، فيه المجال والمؤشر والعنوان.</p></div></div></section>';
    return '<div class="page md">' + head + cards + note + '</div>';
  }

  /* ---------- نموذج الإدخال ---------- */
  function newForm() {
    return { editId: null, files: [], existing: [],
      draft: { author: LS.get('author', ''), title: '', date: todayISO(), indicator: state.indFilter || '', desc: '', links: '' } };
  }

  function filesListHTML() {
    var f = state.form;
    var ex = f.existing.map(function (a) { return '<li class="old"><span class="ext">' + esc(ext(a.name)) + '</span><span class="fname">' + esc(a.name) + '</span><small>مرفق محفوظ</small></li>'; });
    var nw = f.files.map(function (file, i) {
      return '<li><span class="ext">' + esc(ext(file.name)) + '</span><span class="fname">' + esc(file.name) + '</span><small>' + (file.size / 1048576).toFixed(1) + ' م.ب</small>' +
        '<button type="button" data-act="rmfile" data-i="' + i + '" aria-label="إزالة ' + esc(file.name) + '">×</button></li>';
    });
    return ex.concat(nw).join('');
  }

  function formHTML(n) {
    var f = state.form, v = f.draft, d = dom(n);
    return '<form class="wform" id="wform" novalidate>' +
      '<div class="wform-head"><h2>' + (f.editId ? 'تعديل العمل' : 'تسجيل عمل جديد') + '</h2>' +
      '<button type="button" class="btn ghost small" data-act="cancel">إلغاء</button></div>' +
      '<div class="grid2">' +
      '<label class="fld"><span>اسم المُدخِلة <i>*</i></span><input id="f-author" name="author" maxlength="80" autocomplete="name" value="' + esc(v.author) + '"' + (f.editId ? ' disabled' : '') + '></label>' +
      '<label class="fld"><span>تاريخ التنفيذ <i>*</i></span><input type="date" id="f-date" name="date" value="' + esc(v.date) + '"></label></div>' +
      '<label class="fld"><span>عنوان العمل <i>*</i></span><input id="f-title" name="title" maxlength="200" placeholder="مثال: تحليل نتائج الاختبارات القصيرة للصف الرابع" value="' + esc(v.title) + '"></label>' +
      '<label class="fld"><span>المؤشر المرتبط <i>*</i></span><select id="f-indicator" name="indicator"><option value="">اختاري المؤشر…</option>' +
      d.standards.map(function (s) {
        return '<optgroup label="' + s.code + ' ' + esc(s.name) + '">' + s.indicators.map(function (i) {
          return '<option value="' + i.code + '"' + (v.indicator === i.code ? ' selected' : '') + '>' + i.code + ' — ' + esc(i.name) + '</option>';
        }).join('') + '</optgroup>';
      }).join('') + '</select></label>' +
      '<label class="fld"><span>وصف مختصر</span><textarea id="f-desc" name="desc" rows="4" maxlength="3000" placeholder="ماذا نُفّذ؟ من شارك؟ وما النتيجة؟ (أرقام ونسب بلا أسماء طلبة)">' + esc(v.desc) + '</textarea></label>' +
      '<div class="fld"><span>الصور والملفات</span><label class="drop" for="f-files"><input type="file" id="f-files" multiple accept="image/*,.pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx">' +
      '<b>اختيار صور أو ملفات</b><small>الصور تُصغَّر تلقائيًا · الملف حتى 10 م.ب · حتى 10 مرفقات</small></label>' +
      '<ul class="flist" id="flist">' + filesListHTML() + '</ul></div>' +
      '<label class="fld"><span>روابط (فيديو أو غيره) — رابط في كل سطر</span><textarea id="f-links" name="links" rows="2" dir="ltr" placeholder="https://">' + esc(v.links) + '</textarea></label>' +
      '<p class="ferr" id="ferr" role="alert" hidden></p>' +
      '<div class="factions"><button class="btn primary" type="submit" id="fsubmit">' + (f.editId ? 'حفظ التعديل' : 'حفظ العمل') + '</button><span class="fprog" id="fprog" aria-live="polite"></span></div>' +
      '</form>';
  }

  function toBase64(file) {
    return new Promise(function (res, rej) {
      var r = new FileReader();
      r.onload = function () { res(String(r.result).split(',')[1]); };
      r.onerror = function () { rej(new Error('تعذّرت قراءة الملف ' + file.name)); };
      r.readAsDataURL(file);
    });
  }
  function compressImage(file, max, q) {
    return new Promise(function (res, rej) {
      var url = URL.createObjectURL(file), img = new Image();
      img.onload = function () {
        var s = Math.min(1, max / Math.max(img.width, img.height));
        var w = Math.round(img.width * s), h = Math.round(img.height * s);
        var c = document.createElement('canvas'); c.width = w; c.height = h;
        var x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, w, h); x.drawImage(img, 0, 0, w, h);
        URL.revokeObjectURL(url);
        res({ name: file.name.replace(/\.[^.]+$/, '') + '.jpg', mime: 'image/jpeg', data: c.toDataURL('image/jpeg', q).split(',')[1] });
      };
      img.onerror = function () { URL.revokeObjectURL(url); rej(new Error('img')); };
      img.src = url;
    });
  }
  async function prepareFile(file) {
    if (/^image\//.test(file.type) && !/gif|svg/.test(file.type)) {
      try { return await compressImage(file, API.isDemo ? 900 : 1600, 0.82); } catch (e) { /* يُرفع كما هو */ }
    }
    if (file.size > 10 * 1048576) throw new Error('الملف «' + file.name + '» أكبر من 10 ميجابايت. أضيفيه رابطًا بدلًا من رفعه.');
    return { name: file.name, mime: file.type || 'application/octet-stream', data: await toBase64(file) };
  }

  async function submitForm() {
    var f = state.form, v = f.draft, n = +state.tab.slice(1);
    var errEl = $('#ferr'), btn = $('#fsubmit'), prog = $('#fprog');
    function fail(msg) { errEl.textContent = msg; errEl.hidden = false; btn.disabled = false; prog.textContent = ''; }
    errEl.hidden = true;
    if (!f.editId && !v.author.trim()) return fail('اكتبي اسمكِ في خانة «اسم المُدخِلة».');
    if (!v.title.trim()) return fail('اكتبي عنوان العمل.');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(v.date)) return fail('اختاري تاريخ التنفيذ.');
    if (!v.indicator) return fail('اختاري المؤشر المرتبط بالعمل.');
    var lines = v.links.split(/\n+/).map(function (s) { return s.trim(); }).filter(Boolean);
    var bad = lines.filter(function (s) { return !/^https?:\/\/\S+$/i.test(s); });
    if (bad.length) return fail('هذا السطر ليس رابطًا صحيحًا: ' + bad[0] + ' — يجب أن يبدأ بـ https://');
    if (f.existing.length + f.files.length > 10) return fail('الحد الأقصى 10 مرفقات للعمل الواحد.');

    btn.disabled = true;
    if (!f.editId) LS.set('author', v.author.trim());
    try {
      var files = [];
      for (var i = 0; i < f.files.length; i++) {
        prog.textContent = 'جارٍ تجهيز المرفق ' + (i + 1) + ' من ' + f.files.length + '…';
        files.push(await prepareFile(f.files[i]));
      }
      var total = files.reduce(function (s, x) { return s + x.data.length * 0.75; }, 0);
      if (total > 30 * 1048576) throw new Error('حجم المرفقات معًا أكبر من 30 ميجابايت. قسّميها على أكثر من عمل.');
      prog.textContent = files.length ? 'جارٍ رفع المرفقات وحفظ العمل…' : 'جارٍ الحفظ…';
      var work = { committee: n, indicator: v.indicator, title: v.title.trim(), date: v.date, desc: v.desc.trim(), author: v.author.trim(), links: lines };
      if (f.editId) await API.update(f.editId, work, files); else await API.add(work, files);
      var wasEdit = !!f.editId;
      state.form = null;
      await load();
      toast(wasEdit ? 'حُفظ التعديل' : 'حُفظ العمل');
    } catch (e) { fail(e.message || 'تعذّر الحفظ.'); }
  }

  /* ---------- الفريق الإشرافي ---------- */
  function mudeeraHTML() {
    var all = state.works.slice().sort(byDate);
    var covered = covSet(all).size;
    var mk = monthKey(todayISO());
    var thisMonth = all.filter(function (w) { return monthKey(w.date) === mk; }).length;
    var stalled = D.filter(function (d) { return status(d.n).k !== 'ok'; }).length;

    var lastWeek = all.filter(function (w) { return w.created && daysSince(w.created) < 7; }).length;
    var head = '<section class="hero md-hero"><div class="hero-txt">' + BACK.replace('class="back"', 'class="back on-dark"') +
      '<p class="eyebrow light">اطلاع فقط</p><h1>الفريق الإشرافي</h1>' +
      '<p>متابعة المجالات الخمسة وتغطيتها لمؤشرات تقويم الأداء، والتقرير.</p>' +
      '<button type="button" class="btn primary" data-act="refresh">' + ic('refresh') + 'تحديث البيانات</button></div>' + HERO_ART + '</section>';

    var kpis = '<div class="kpis">' +
      kpi(all.length, 'إجمالي الأعمال', '+' + lastWeek + ' هذا الأسبوع', 'list', 1) +
      kpi(covered + '<small> / 78</small>', 'مؤشرات لها أعمال', Math.round(covered / 78 * 100) + '٪ من المؤشرات', 'target', 2) +
      kpi(thisMonth, 'أعمال هذا الشهر', monthLabel(mk), 'cal', 3) +
      kpi(stalled, 'مجالات تحتاج متابعة', stalled ? 'متوقفة أو لم تبدأ' : 'كل المجالات نشطة', 'bell', 4) + '</div>';
    var charts = '<div class="charts"><section class="panel"><header class="panel-h"><h2>الأعمال خلال الأشهر الستة الأخيرة</h2></header>' + barChart() + '</section>' +
      '<section class="panel"><header class="panel-h"><h2>توزيع الأعمال على المجالات</h2></header>' + donutChart() + '</section></div>';

    var cards = '<div class="ccards">' + D.map(function (d) {
      var l = worksOf(d.n), st = status(d.n), cv = covSet(l).size;
      return '<button type="button" class="ccard' + (state.f.c == d.n && state.sub === 'timeline' ? ' on' : '') + '" data-act="fc" data-c="' + d.n + '" style="--dc:var(--d' + d.n + ')">' +
        '<span class="ccard-top"><span class="num">' + d.n + '</span><span class="pill ' + st.k + '">' + st.t + '</span></span>' +
        '<span class="ccard-name">' + esc(d.committee) + '</span>' +
        '<span class="ccard-row"><span>' + pl(l.length, W_WORK) + '</span><span>آخر إدخال: ' + relTime(lastEntry(l)) + '</span></span>' +
        '<span class="bar" aria-hidden="true"><i style="width:' + (cv / d.count * 100).toFixed(1) + '%"></i></span>' +
        '<span class="ccard-cov">' + cv + ' من ' + d.count + ' مؤشرًا لها أعمال</span></button>';
    }).join('') + '</div>';

    var subs = [['timeline', 'السجل الزمني'], ['coverage', 'تغطية المؤشرات'], ['report', 'التقرير']];
    var subnav = '<div class="subnav" role="tablist">' + subs.map(function (s) {
      return '<button type="button" role="tab" aria-selected="' + (state.sub === s[0]) + '" class="sub' + (state.sub === s[0] ? ' on' : '') + '" data-act="sub" data-sub="' + s[0] + '">' + s[1] + '</button>';
    }).join('') + '</div>';

    var body = state.sub === 'coverage' ? coverageHTML() : state.sub === 'report' ? reportHTML() : timelineHTML();
    return '<div class="page md">' + head + kpis + charts + '<h2 class="sec-title">المجالات</h2>' + cards + subnav + '<div id="subview">' + body + '</div></div>';
  }
  function kpi(v, label, sub, icon, k) {
    return '<div class="kpi k' + k + '">' + tile(icon) + '<span class="kpi-l">' + label + '</span><b>' + v + '</b><span class="kpi-s">' + sub + '</span></div>';
  }

  /* ---------- الرسوم ---------- */
  function monthsBack(n) {
    var out = [], now = new Date();
    for (var i = n - 1; i >= 0; i--) {
      var x = new Date(now.getFullYear(), now.getMonth() - i, 1);
      out.push(x.getFullYear() + '-' + String(x.getMonth() + 1).padStart(2, '0'));
    }
    return out;
  }
  function barChart() {
    var ms = monthsBack(6);
    var counts = ms.map(function (m) { return state.works.filter(function (w) { return monthKey(w.date) === m; }).length; });
    var step = Math.max(1, Math.ceil(Math.max.apply(null, counts.concat([4])) / 4)), top = step * 4;
    var W = 380, H = 220, x0 = 12, x1 = W - 34, y0 = 178, y1 = 22, slot = (x1 - x0) / 6, bw = Math.min(30, slot * 0.5);
    var fills = ['var(--b1)', 'var(--b2)', 'var(--b3)', 'var(--b4)', 'var(--b5)', 'var(--b6)'];
    var y = function (v) { return y0 - (v / top) * (y0 - y1); };
    var g = '';
    for (var v = 0; v <= top; v += step) {
      g += '<line x1="' + x0 + '" x2="' + x1 + '" y1="' + y(v) + '" y2="' + y(v) + '" class="ch-grid"/>' +
        '<text x="' + (W - 6) + '" y="' + (y(v) + 4) + '" class="ch-ax" text-anchor="end">' + v + '</text>';
    }
    var bars = counts.map(function (c, i) {
      var cx = x1 - slot * (i + 0.5); // الأقدم يمينًا
      var h = Math.max(c ? 6 : 0, y0 - y(c));
      var lbl = asDate(ms[i] + '-01').toLocaleDateString(LOC, { month: 'long' });
      return (c ? '<rect x="' + (cx - bw / 2) + '" y="' + (y0 - h) + '" width="' + bw + '" height="' + h + '" rx="8" style="fill:' + fills[i] + '"/>' +
        '<text x="' + cx + '" y="' + (y0 - h - 7) + '" class="ch-val" text-anchor="middle">' + c + '</text>' : '') +
        '<text x="' + cx + '" y="' + (y0 + 20) + '" class="ch-ax" text-anchor="middle">' + lbl + '</text>';
    }).join('');
    return '<div class="chart"><svg viewBox="0 0 ' + W + ' ' + H + '" role="img" aria-label="عدد الأعمال في كل شهر خلال الأشهر الستة الأخيرة">' + g + bars + '</svg></div>';
  }
  function donutChart() {
    var counts = D.map(function (d) { return worksOf(d.n).length; });
    var total = counts.reduce(function (a, b) { return a + b; }, 0);
    var r = 62, C2 = 2 * Math.PI * r, acc = 0, segs = '';
    if (!total) segs = '<circle cx="90" cy="90" r="' + r + '" class="ch-empty" fill="none" stroke-width="30"/>';
    else counts.forEach(function (c, i) {
      if (!c) return;
      var len = c / total * C2, gap = counts.filter(Boolean).length > 1 ? 3 : 0;
      segs += '<circle cx="90" cy="90" r="' + r + '" fill="none" stroke-width="30" style="stroke:var(--d' + (i + 1) + ')" stroke-dasharray="' +
        Math.max(0.1, len - gap) + ' ' + (C2 - Math.max(0.1, len - gap)) + '" stroke-dashoffset="' + (-acc) + '" transform="rotate(-90 90 90)"/>';
      acc += len;
    });
    var legend = '<ul class="legend">' + D.map(function (d, i) {
      var pct = total ? Math.round(counts[i] / total * 100) : 0;
      return '<li style="--dc:var(--d' + d.n + ')"><span class="dot"></span><span class="lg-n">' + esc(d.name) + '</span><b>' + pct + '٪</b></li>';
    }).join('') + '</ul>';
    return '<div class="donut"><svg viewBox="0 0 180 180" role="img" aria-label="نسبة أعمال كل مجال">' + segs +
      '<text x="90" y="88" class="ch-big" text-anchor="middle">' + total + '</text><text x="90" y="110" class="ch-ax" text-anchor="middle">' + (total === 1 ? 'عمل' : total === 2 ? 'عملان' : total >= 3 && total <= 10 ? 'أعمال' : 'عملًا') + '</text></svg>' + legend + '</div>';
  }

  function filtered() {
    var f = state.f, q = f.q.trim();
    return state.works.filter(function (w) {
      if (f.c && +w.committee !== +f.c) return false;
      if (f.m && monthKey(w.date) !== f.m) return false;
      if (f.ind && w.indicator !== f.ind) return false;
      if (q && (w.title + ' ' + w.desc + ' ' + w.author + ' ' + w.indicator).indexOf(q) < 0) return false;
      return true;
    }).sort(byDate);
  }

  function timelineHTML() {
    var f = state.f;
    var months = Array.from(new Set(state.works.map(function (w) { return monthKey(w.date); }))).filter(Boolean).sort().reverse();
    return '<div class="filters">' +
      '<label class="fld compact"><span>المجال</span><select id="ft-c"><option value="">كل المجالات</option>' +
      D.map(function (d) { return '<option value="' + d.n + '"' + (f.c == d.n ? ' selected' : '') + '>' + esc(d.committee) + '</option>'; }).join('') + '</select></label>' +
      '<label class="fld compact"><span>الشهر</span><select id="ft-m"><option value="">كل الأشهر</option>' +
      months.map(function (m) { return '<option value="' + m + '"' + (f.m === m ? ' selected' : '') + '>' + monthLabel(m) + '</option>'; }).join('') + '</select></label>' +
      '<label class="fld compact grow"><span>بحث</span><input id="ft-q" type="search" placeholder="عنوان، وصف، اسم…" value="' + esc(f.q) + '"></label>' +
      '</div><div id="tl-results">' + tlResults() + '</div>';
  }

  function tlResults() {
    var f = state.f, list = filtered();
    var chips = '';
    if (f.ind) chips += '<span class="filter-chip">المؤشر ' + esc(f.ind) + ' — ' + esc(IND[f.ind] ? IND[f.ind].name : '') + '<button type="button" data-act="clearFind" aria-label="إلغاء">×</button></span>';
    var head = '<div class="tl-count"><span>' + pl(list.length, W_WORK) + '</span>' + chips + '</div>';
    if (!list.length) return head + '<div class="empty"><p>لا أعمال تطابق هذا الاختيار.</p></div>';
    var groups = {}, order = [];
    list.forEach(function (w) { var k = monthKey(w.date); if (!groups[k]) { groups[k] = []; order.push(k); } groups[k].push(w); });
    return head + order.map(function (k) {
      return '<section class="tl-month"><h3>' + monthLabel(k) + '<small>' + pl(groups[k].length, W_WORK) + '</small></h3><div class="works two">' +
        groups[k].map(function (w) { return workCard(w, { showCommittee: true }); }).join('') + '</div></section>';
    }).join('');
  }

  function coverageHTML() {
    return '<p class="cov-legend"><span class="sw has"></span>مؤشر عليه أعمال <span class="sw none"></span>مؤشر بلا أعمال بعد · اضغطي أي مؤشر لعرض أعماله</p>' +
      D.map(function (d) {
        var l = worksOf(d.n), cv = covSet(l).size;
        return '<section class="cov-dom" style="--dc:var(--d' + d.n + ')"><header><h3><span class="num">' + d.n + '</span>' + esc(d.name) + '</h3>' +
          '<span class="cov-frac"><b>' + cv + '</b> من ' + d.count + '</span></header>' +
          d.standards.map(function (s) {
            return '<h4>' + s.code + ' ' + esc(s.name) + '</h4><div class="tiles">' + s.indicators.map(function (i) {
              var c = indCount(l, i.code);
              return '<button type="button" class="tile' + (c ? '' : ' empty') + '" data-act="find" data-ind="' + i.code + '" title="' + esc(i.name) + '">' +
                '<span class="tile-code">' + i.code + '</span><span class="tile-name">' + esc(i.name) + '</span><span class="tile-n">' + c + '</span></button>';
            }).join('') + '</div>';
          }).join('') + '</section>';
      }).join('');
  }

  function reportHTML() {
    var r = state.rep;
    return '<div class="filters rp-controls">' +
      '<label class="fld compact"><span>النطاق</span><select id="rp-c"><option value="">المجالات الخمسة</option>' +
      D.map(function (d) { return '<option value="' + d.n + '"' + (r.c == d.n ? ' selected' : '') + '>المجال ' + d.n + ': ' + esc(d.name) + '</option>'; }).join('') + '</select></label>' +
      '<label class="check"><input type="checkbox" id="rp-empty"' + (r.empty ? ' checked' : '') + '><span>إظهار المؤشرات التي بلا أعمال</span></label>' +
      '<button type="button" class="btn primary" data-act="print">طباعة / حفظ PDF</button></div>' +
      '<div id="rp-body">' + reportBody() + '</div>';
  }

  function reportBody() {
    var r = state.rep;
    var doms = r.c ? [dom(+r.c)] : D;
    return '<div class="report"><div class="rp-title"><img class="rp-logo" src="logo.png" alt=""><p>' + esc(C.SCHOOL_NAME) + ' ' + ltrNums(C.SCHOOL_GRADES || '') + '</p>' +
      '<h2>أدلة أداء المدرسة حسب مجالات التقويم</h2><p>صدر في ' + fmtDate(todayISO()) + '</p></div>' +
      doms.map(function (d) {
        var l = worksOf(d.n);
        return '<section class="rp-dom" style="--dc:var(--d' + d.n + ')"><h3>المجال ' + d.n + ': ' + esc(d.name) +
          '<small>' + covSet(l).size + ' من ' + d.count + ' مؤشرًا لها أعمال</small></h3>' +
          d.standards.map(function (s) {
            var rows = s.indicators.map(function (i) {
              var ws = l.filter(function (w) { return w.indicator === i.code; });
              if (!ws.length && !r.empty) return '';
              return '<tr' + (ws.length ? '' : ' class="none"') + '><td class="rp-ind"><b>' + i.code + '</b> ' + esc(i.name) + '</td><td>' +
                (ws.length ? '<ol>' + ws.map(function (w) {
                  var na = (w.attachments || []).length + (w.links || []).length;
                  return '<li><b>' + esc(w.title) + '</b><span class="rp-meta">' + fmtDate(w.date) + ' · ' + esc(w.author) + (na ? ' · ' + na + ' مرفقات' : '') + '</span></li>';
                }).join('') + '</ol>' : 'لا أعمال مسجلة') + '</td></tr>';
            }).join('');
            if (!rows) return '';
            return '<h4>المعيار ' + s.code + ': ' + esc(s.name) + '</h4><div class="tbl"><table><thead><tr><th>المؤشر</th><th>الأعمال المسجلة</th></tr></thead><tbody>' + rows + '</tbody></table></div>';
          }).join('') + '</section>';
      }).join('') + '</div>';
  }

  /* ---------- عارض الصور ---------- */
  function openLightbox(wid, aid) {
    var w = state.works.find(function (x) { return String(x.id) === wid; });
    var a = w && (w.attachments || []).find(function (x) { return String(x.id) === aid; });
    if (!a) return;
    var big = String(a.thumb).replace(/sz=w\d+/, 'sz=w1600');
    var box = $('#lightbox');
    box.innerHTML = '<div class="lb-in" role="dialog" aria-modal="true" aria-label="' + esc(a.name) + '"><img src="' + esc(safeUrl(big)) + '" alt="' + esc(a.name) + '">' +
      '<div class="lb-bar"><span>' + esc(a.name) + '</span>' +
      (/^https?:/.test(a.url) ? '<a class="btn small ghost-on" href="' + esc(a.url) + '" target="_blank" rel="noopener">فتح في Drive</a>' : '') +
      '<button type="button" class="btn small ghost-on" data-act="closelb">إغلاق</button></div></div>';
    box.hidden = false;
    $('[data-act="closelb"]', box).focus();
  }

  /* ---------- رسالة عابرة ---------- */
  var toastT;
  function toast(msg) {
    var t = $('#toast'); t.textContent = msg; t.classList.add('show');
    clearTimeout(toastT); toastT = setTimeout(function () { t.classList.remove('show'); }, 2600);
  }

  /* ---------- الأحداث ---------- */
  document.addEventListener('click', function (e) {
    var t = e.target.closest('[data-act]');
    if (!t) { if (e.target.id === 'lightbox') $('#lightbox').hidden = true; return; }
    var a = t.dataset.act;
    if (a === 'add') { state.form = newForm(); render(); focusForm(); }
    else if (a === 'cancel') { state.form = null; render(); }
    else if (a === 'edit') {
      var w = state.works.find(function (x) { return String(x.id) === t.dataset.id; });
      if (!w) return;
      state.form = { editId: w.id, files: [], existing: w.attachments || [],
        draft: { author: w.author, title: w.title, date: w.date, indicator: w.indicator, desc: w.desc || '', links: (w.links || []).join('\n') } };
      render(); focusForm();
    }
    else if (a === 'ind') { state.indFilter = state.indFilter === t.dataset.ind ? null : t.dataset.ind; render(); }
    else if (a === 'clearInd') { state.indFilter = null; render(); }
    else if (a === 'rmfile') { state.form.files.splice(+t.dataset.i, 1); $('#flist').innerHTML = filesListHTML(); }
    else if (a === 'fc') {
      state.sub = 'timeline';
      state.f = { c: state.f.c == t.dataset.c && state.sub === 'timeline' ? '' : t.dataset.c, m: '', q: '', ind: '' };
      render(); scrollToSub();
    }
    else if (a === 'find') { state.sub = 'timeline'; state.f = { c: '', m: '', q: '', ind: t.dataset.ind }; render(); scrollToSub(); }
    else if (a === 'clearFind') { state.f.ind = ''; $('#tl-results').innerHTML = tlResults(); }
    else if (a === 'sub') { state.sub = t.dataset.sub; render(); }
    else if (a === 'print') { window.print(); }
    else if (a === 'refresh') { state.loaded = false; render(); load(); }
    else if (a === 'clearDemo') { API.clearDemo(); state.loaded = false; render(); load(); toast('أُعيدت البيانات التجريبية إلى بدايتها'); }
    else if (a === 'lb') { openLightbox(t.dataset.w, t.dataset.a); }
    else if (a === 'closelb') { $('#lightbox').hidden = true; }
    else if (a === 'theme') { setTheme(isDark() ? 'light' : 'dark'); }
  });

  document.addEventListener('keydown', function (e) { if (e.key === 'Escape') $('#lightbox').hidden = true; });

  document.addEventListener('input', function (e) {
    var el = e.target;
    if (el.form && el.form.id === 'wform' && el.name && state.form) state.form.draft[el.name] = el.value;
    if (el.id === 'ft-q') { state.f.q = el.value; $('#tl-results').innerHTML = tlResults(); }
  });

  document.addEventListener('change', function (e) {
    var el = e.target;
    if (el.id === 'f-files') {
      var add = Array.from(el.files || []);
      state.form.files = state.form.files.concat(add).slice(0, Math.max(0, 10 - state.form.existing.length));
      el.value = '';
      $('#flist').innerHTML = filesListHTML();
    } else if (el.form && el.form.id === 'wform' && el.name && state.form) {
      state.form.draft[el.name] = el.value;
    } else if (el.id === 'ft-c') { state.f.c = el.value; render(); }
    else if (el.id === 'ft-m') { state.f.m = el.value; $('#tl-results').innerHTML = tlResults(); }
    else if (el.id === 'rp-c') { state.rep.c = el.value; $('#rp-body').innerHTML = reportBody(); }
    else if (el.id === 'rp-empty') { state.rep.empty = el.checked; $('#rp-body').innerHTML = reportBody(); }
  });

  document.addEventListener('submit', function (e) {
    if (e.target.id === 'wform') { e.preventDefault(); submitForm(); }
  });

  // صورة لم تُحمَّل (مثلًا: مشاركة Drive مقيّدة) ← يظهر اسمها بدلها
  document.addEventListener('error', function (e) {
    var t = e.target;
    if (t.tagName === 'IMG' && t.closest('.thumb')) t.closest('.thumb').classList.add('broken');
  }, true);

  function focusForm() {
    var f = $('#wform'); if (!f) return;
    f.scrollIntoView({ behavior: 'smooth', block: 'start' });
    var first = $('#f-title', f); if (first) setTimeout(function () { first.focus({ preventScroll: true }); }, 300);
  }
  function scrollToSub() { var s = $('.subnav'); if (s) s.scrollIntoView({ behavior: 'smooth', block: 'start' }); }

  /* ---------- الوضع الداكن ---------- */
  var mq = window.matchMedia ? window.matchMedia('(prefers-color-scheme: dark)') : null;
  function isDark() {
    var t = document.documentElement.getAttribute('data-theme');
    return t ? t === 'dark' : !!(mq && mq.matches);
  }
  function syncThemeBtn() {
    var d = isDark(), b = $('#themeBtn');
    if (!b) return;
    b.classList.toggle('dark', d);
    b.setAttribute('aria-pressed', String(d));
    $('#themeTxt').textContent = d ? 'الوضع الفاتح' : 'الوضع الداكن';
    var ti = $('#themeIc'); if (ti) ti.innerHTML = ic(d ? 'sun' : 'moon');
  }
  function setTheme(t) {
    document.documentElement.setAttribute('data-theme', t);
    try { localStorage.setItem('lijan-theme', t); } catch (e) {}
    syncThemeBtn();
  }
  if (mq && mq.addEventListener) mq.addEventListener('change', syncThemeBtn);

  window.addEventListener('hashchange', route);
  renderChrome();
  syncThemeBtn();
  route();
  load();
})();
