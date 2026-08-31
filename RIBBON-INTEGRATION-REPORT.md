# تقرير إعادة تنظيم Nawa OCR Office — Ribbon

## الملخص

أعيد تنظيم الإضافة لتكون **Office Add-in حقيقية** يكون Ribbon فيها نقطة الدخول الأساسية في Word وExcel. يعتمد التنفيذ على manifest بصيغة add-in-only مع `VersionOverrides` وأوامر Ribbon من نوع `ShowTaskpane`. تظل المعالجة داخل Nawa OCR Core، بينما يتولى Office.js إدراج النتائج في مستند Word أو مصنف Excel.

توضح Microsoft أن Office Add-ins تتكون من manifest وتطبيق ويب مستضاف، ويمكنها إضافة أزرار Ribbon وTask Pane والتفاعل مع المستندات عبر Office.js [1]. كما توصي Microsoft باستخدام HTTPS، وتذكر أن خط أساس عملاء Windows هو Office 2016 أو أحدث مع مراعاة مجموعات المتطلبات المتاحة في العميل [2].

## ما تم تغييره

يظهر اسم **Nawa OCR** كمجموعة داخل Home tab في Word وExcel. أضيفت الأزرار التالية لكل مضيف: **استيراد وتحويل** لبدء اختيار PDF أو صورة والتحويل، **مراجعة** لفتح مسار المراجعة، و**إعدادات** لفتح الإعدادات. تفتح الأزرار Task Pane نفسها أو مساراتها عبر `?panel=review` و`?panel=settings`، لذلك لا تعتمد الوظيفة على فتح قائمة الإضافات يدويًا.

في Word، يدعم مسار الإدراج موضع المؤشر ونهاية المستند واستبدال التحديد النصي الأول، مع إنشاء فقرات وجداول Word حقيقية من خطة OCR. في Excel، يدعم الإدراج عند الخلية النشطة أو في ورقة جديدة، ويحسب نقطة البداية من الخلية المحددة. تُقرأ المنطقة المستهدفة قبل الكتابة، وإذا كانت غير فارغة يتوقف المسار برسالة واضحة ولا يحذف البيانات تلقائيًا.

## مصفوفة التوافق

| العميل | Ribbon في manifest | نتيجة اختبار محلي | الإعلان الحالي |
|---|---:|---:|---|
| Microsoft 365 Word على Windows | نعم | غير مختبر داخل Word | متوقع الدعم، ويجب تنفيذ اختبار Windows |
| Microsoft 365 Excel على Windows | نعم | غير مختبر داخل Excel | متوقع الدعم، ويجب تنفيذ اختبار Windows |
| Office 2021 Word/Excel | نعم من حيث manifest | غير مختبر | يحتاج تحققًا فعليًا |
| Office 2019 Word/Excel | نعم من حيث manifest | غير مختبر | يحتاج تحققًا فعليًا |
| Office 2016 Word/Excel | fallback موجود، Ribbon الحديثة تعتمد على العميل | غير مختبر | لا إعلان نهائي قبل اختبار API sets |
| Office 2013 وأقدم | غير مستهدف | غير مختبر | لا دعم معلن؛ قد تحتاج VSTO/COM على Windows |
| Office 2007/2010 | غير مستهدف | غير مختبر | خارج خط الأساس الموثق لـ Office Add-ins |
| Word/Excel للويب أو Mac | يعتمد على العميل والاستضافة | غير مختبر | يحتاج HTTPS واختبارًا منفصلًا |

## الاختبارات المحلية

نجح فحص TypeScript، ونجحت اختبارات العقد الخاصة بالـ manifest والإدراج والحواجز، إضافة إلى مجموعة اختبارات المشروع. لا يمثل ذلك اختبارًا بديلاً عن فتح Word وExcel المكتبيين وإعادة فتح DOCX وXLSX؛ هذه الخطوة تحتاج جهاز Windows وOffice فعليين.

## التثبيت والتجربة

يُشغّل خادم الواجهة عبر HTTPS، ثم يُحمّل `manifest.xml` بالطريقة الرسمية المناسبة للعميل، مثل Sideloading للتطوير أو Integrated Apps/كتالوج المؤسسة للنشر. بعد ظهور Nawa OCR في Home tab، يبدأ المستخدم من **استيراد وتحويل**، يختار المصدر، يراجع المعاينة، ثم يدرج النتيجة مباشرة. لا يتم تعديل أو حذف ملف PDF أو الصورة الأصلية.

## القيود

لم تُبنَ طبقة VSTO/COM للإصدارات القديمة، ولم تُجرَ اختبارات فعلية على Office 2007 أو 2010 أو 2013 أو 2016 أو 2019 أو 2021 أو Microsoft 365 داخل Windows. لذلك لا يصح إعلان عبارة «جميع إصدارات Office». كما أن عنوان HTTPS الحالي عنوان معاينة مؤقت للتطوير وليس عنوان إنتاج دائم.

## المراجع

[1]: https://learn.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins "Office Add-ins platform overview — Microsoft Learn"
[2]: https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/requirements-for-running-office-add-ins "Requirements for running Office Add-ins — Microsoft Learn"
