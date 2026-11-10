/* إعدادات الموقع — هذا هو الملف الوحيد الذي تحتاج تعديله */
window.CONFIG = {
  // رابط تطبيق الويب من Apps Script (ينتهي بـ /exec).
  // اتركه فارغًا ليعمل الموقع في الوضع التجريبي (البيانات محفوظة في المتصفح فقط).
  API_URL: "",

  // نموذج Microsoft Forms (حساب المدرسة sc3212@moe.om).
  // ما دام FORM_URL مملوءًا يعمل الموقع بنموذج Forms: التسجيل يتم في النموذج، والمتابعة في Forms وExcel.
  FORM_URL: "https://forms.cloud.microsoft/r/RRNb0EZ3hD",
  // صفحة نتائج النموذج (تفتح بحساب المدرسة فقط)
  FORM_RESULTS_URL: "https://forms.cloud.microsoft/Pages/DesignPageV2.aspx?subpage=design&id=Xcu0BEHMH0C9nUyooxpcL67TxYphdTlItOCekvkfi79UQzNLUEZaRERKR0hXTzhPTzIyMzRFR0o4US4u&analysis=true",
  // رابط ملف Excel للوحة المتابعة (يُضاف بعد تجهيزه)
  // أساس رابط ملف الردود (لعرض أوراق المجالات D1..D5 داخل صفحات المجالات)
  WORKBOOK_BASE: "https://moeom-my.sharepoint.com/personal/sc3212_moe_om/_layouts/15/Doc.aspx?sourcedoc=%7BD3297C9C-D329-417F-A878-7CE9B4738823%7D",
  EXCEL_URL: "https://moeom-my.sharepoint.com/:x:/r/personal/sc3212_moe_om/_layouts/15/Doc.aspx?sourcedoc=%7BD3297C9C-D329-417F-A878-7CE9B4738823%7D&action=view",

  SCHOOL_NAME: "مدرسة أبو عبالي الساحل للتعليم الأساسي",
  SCHOOL_GRADES: "الصفوف (1-4)",
  ACADEMIC_YEAR: "2026/2027",

  // بعد كم يومًا بلا إدخال يظهر المجال متوقفًا في صفحة الفريق الإشرافي
  INACTIVE_DAYS: 14
};
