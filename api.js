/* طبقة البيانات: الوضع التجريبي (المتصفح) أو Apps Script (الحفظ المشترك) */
window.API = (function () {
  var URL_ = String((window.CONFIG && window.CONFIG.API_URL) || '').trim();
  var isDemo = !URL_;
  var KEY = 'lijan-demo-works-v1';

  function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }

  function demoLoad() {
    try { var v = localStorage.getItem(KEY); if (v) return JSON.parse(v); } catch (e) {}
    var s = seed(); demoSave(s); return s;
  }
  function demoSave(list) {
    try { localStorage.setItem(KEY, JSON.stringify(list)); return true; } catch (e) { return false; }
  }
  function demoAtts(files) {
    return files.map(function (f) {
      var url = 'data:' + f.mime + ';base64,' + f.data;
      return { id: uid(), name: f.name, mime: f.mime, url: url, thumb: /^image\//.test(f.mime) ? url : '' };
    });
  }
  var FULL = 'امتلأت مساحة التخزين التجريبية في هذا المتصفح. قلّلي الصور أو امسحي البيانات التجريبية.';

  async function list() {
    if (isDemo) return demoLoad();
    var r = await fetch(URL_ + '?action=list', { cache: 'no-store' });
    var j = await r.json();
    if (!j.ok) throw new Error(j.error || 'تعذّر جلب الأعمال.');
    return j.works;
  }

  async function post(body) {
    var r;
    try {
      // بلا ترويسة Content-Type: يُرسل نصًّا عاديًا فلا يحتاج Apps Script طلبًا تمهيديًا
      r = await fetch(URL_, { method: 'POST', body: JSON.stringify(body) });
    } catch (e) {
      throw new Error('تعذّر الاتصال. تحققي من الإنترنت ثم حاولي مجددًا.');
    }
    var j;
    try { j = await r.json(); } catch (e) { throw new Error('وصل ردّ غير مفهوم من الخادم. حاولي مجددًا بعد قليل.'); }
    if (!j.ok) throw new Error(j.error || 'تعذّر الحفظ.');
    return j.work;
  }

  async function add(work, files) {
    if (isDemo) {
      var l = demoLoad();
      var w = Object.assign({}, work, { id: uid(), created: new Date().toISOString(), updated: '', attachments: demoAtts(files) });
      l.push(w);
      if (!demoSave(l)) throw new Error(FULL);
      return w;
    }
    return post({ action: 'add', work: work, files: files });
  }

  async function update(id, work, files) {
    if (isDemo) {
      var l = demoLoad();
      var w = l.find(function (x) { return x.id === id; });
      if (!w) throw new Error('لم يُعثر على العمل.');
      ['indicator', 'title', 'date', 'desc', 'links'].forEach(function (k) { w[k] = work[k]; });
      w.attachments = (w.attachments || []).concat(demoAtts(files));
      w.updated = new Date().toISOString();
      if (!demoSave(l)) throw new Error(FULL);
      return w;
    }
    return post({ action: 'update', id: id, work: work, files: files });
  }

  function clearDemo() { try { localStorage.removeItem(KEY); } catch (e) {} }

  /* أمثلة للوضع التجريبي فقط — تختفي عند ربط Apps Script */
  function seed() {
    var day = 86400000;
    function iso(daysAgo) { return new Date(Date.now() - daysAgo * day).toLocaleDateString('en-CA'); }
    function at(daysAgo) { return new Date(Date.now() - daysAgo * day + 3600000).toISOString(); }
    var rows = [
      [1, '1.1.1', 'تحليل نتائج الاختبارات النهائية للفصل الأول', 5, 'رئيسة المجال', 'مقارنة نسب الإتقان في المواد الأساسية بالعام السابق، وتحديد الصفوف التي تحتاج خططًا علاجية.'],
      [1, '1.3.6', 'مسابقة القارئ الصغير', 12, 'عضو المجال', 'شارك فيها 46 طالبًا من الصفوف 2–4، وعُرضت تلخيصاتهم في ممر المبنى.'],
      [1, '1.2.2', 'متابعة التقدم في مهارات القراءة والحساب', 25, 'رئيسة المجال', 'سجل متابعة شهري لطلبة الصف الثالث بعد تطبيق برنامج الدعم.'],
      [2, '2.2.1', 'فعالية يوم العلم العُماني', 2, 'عضو المجال', 'طابور صباحي وفقرات وطنية وورشة رسم للعلم بمشاركة الصفوف جميعها.'],
      [2, '2.3.2', 'حملة زراعة أشجار في ساحة المدرسة', 9, 'رئيسة المجال', 'زُرعت 20 شتلة محلية، ووُزّعت مسؤولية سقيها على مجموعات الطلبة.'],
      [2, '2.1.1', 'برنامج سفراء القيم', 30, 'عضو المجال', 'اختيار سفير للقيم في كل صف وتكريم أسبوعي للسلوك الإيجابي.'],
      [3, '3.3.3', 'ورشة توظيف المنصات التعليمية في الحصة', 1, 'رئيسة المجال', 'ورشة داخلية لمعلمات المجال الأول عن توظيف السبورة التفاعلية والمنصات.'],
      [3, '3.5.3', 'خطط علاجية بناءً على نتائج التقويم', 6, 'عضو المجال', 'إعداد خطط علاجية وإثرائية فردية لطلبة الصف الرابع في الرياضيات.'],
      [3, '3.2.1', 'زيارات تبادلية حول إدارة زمن التعلم', 15, 'عضو المجال', '8 زيارات تبادلية مع استمارة ملاحظة موحدة.'],
      [4, '4.1.1', 'تنفيذ خطة الإخلاء الوهمي', 21, 'رئيسة المجال', 'زمن الإخلاء 3 دقائق و40 ثانية، مع تقرير بالملاحظات ومقترحات التحسين.'],
      [4, '4.3.1', 'إذاعة مدرسية عن حقوق الطفل', 28, 'عضو المجال', 'إذاعة أعدها طلبة الصف الرابع مع بوسترات في الممرات.']
    ];
    return rows.map(function (r, i) {
      return { id: 'demo' + i, committee: r[0], indicator: r[1], title: r[2], date: iso(r[3]), created: at(r[3]),
        author: r[4], desc: r[5], attachments: [], links: [], updated: '' };
    });
  }

  return { isDemo: isDemo, list: list, add: add, update: update, clearDemo: clearDemo };
})();
