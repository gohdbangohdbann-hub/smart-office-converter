# مصفوفة توافق Nawa OCR Office

## النطاق المعتمد

Nawa OCR Office هو **Office Add-in مبني على Office.js**. يحتوي manifest على `VersionOverrides` وأوامر Ribbon من نوع `ShowTaskpane`، بينما تبقى المعالجة وواجهة المستخدم داخل تطبيق ويب مستضاف عبر HTTPS. هذا يطابق نموذج Office Add-ins الذي يجمع manifest وتطبيق ويب ويستخدم Office.js للتفاعل مع Word وExcel [1].

وفق وثيقة Microsoft الخاصة بمتطلبات التشغيل، يكون خط الأساس لعملاء Windows هو Office 2016 أو أحدث، مع مراعاة مجموعات المتطلبات ونسخة Office.js المتاحة في العميل [2]. لم تُجرَ اختبارات فعلية داخل Word أو Excel على هذه البيئة Linux، ولذلك تُفصل الأهلية النظرية عن نتيجة الاختبار العملي.

| العميل | Ribbon في manifest | اختبار عملي في هذه البيئة | الحالة الموثقة |
|---|---:|---:|---|
| Microsoft 365 Word على Windows | نعم، `VersionOverrides` | لا | متوقع الدعم مع اختبار Windows فعلي مطلوب |
| Microsoft 365 Excel على Windows | نعم، `VersionOverrides` | لا | متوقع الدعم مع اختبار Windows فعلي مطلوب |
| Office 2021 Word/Excel | نعم من حيث manifest | لا | يتطلب تحققًا فعليًا على الجهاز |
| Office 2019 Word/Excel | نعم من حيث manifest | لا | يتطلب تحققًا فعليًا على الجهاز |
| Office 2016 Word/Excel | fallback manifest موجود، Ribbon الحديثة غير مختبرة | لا | لا يُعلن دعمًا نهائيًا قبل اختبار العميل ومجموعة API |
| Office 2013 وأقدم | غير معلن كهدف لهذه النسخة | لا | خارج نطاق الدعم المعلن؛ قد تحتاج طبقة VSTO/COM منفصلة |
| Office 2007/2010 | غير معلن كهدف لهذه النسخة | لا | غير مدعوم بهذه البنية؛ لا توجد نتيجة اختبار |
| Word/Excel للويب | يمكن تشغيل واجهة الإضافة حسب الاستضافة والمتطلبات | لا | يتطلب HTTPS ونشرًا واختبارًا خاصًا بالويب |
| Word/Excel على Mac | يعتمد على دعم العميل وOffice.js | لا | يتطلب اختبارًا مستقلًا؛ لا يُستنتج من اختبار Windows |

## ما تم تنفيذه

يظهر اسم **Nawa OCR** كمجموعة داخل Home tab في Word وExcel. يحتوي كل مضيف على أزرار **استيراد وتحويل** و**مراجعة** و**إعدادات**، وتفتح هذه الأزرار Task Pane نفسها أو مساراتها بمعامل `panel`. Task Pane ليست نقطة الوصول الوحيدة المقصودة؛ Ribbon هي نقطة الدخول الأساسية.

في Word، يتيح مسار الإدراج الحالي موضع المؤشر، نهاية المستند، أو استبدال التحديد النصي الأول، مع إنشاء فقرات وجداول Word حقيقية من خطة OCR. في Excel، يتيح الإدراج عند الخلية النشطة أو في ورقة جديدة، ويحسب النطاق من `rowIndex` و`columnIndex` للخلية المحددة. قبل الكتابة، تُقرأ المنطقة المستهدفة؛ إذا كانت غير فارغة يتوقف الإدراج برسالة تحذير، ولا تُحذف البيانات تلقائيًا.

## حدود صريحة

لم يتم في هذه المرحلة تشغيل Word أو Excel المكتبيين أو إعادة فتح ملفات DOCX/XLSX داخل Windows. كما لم تُبنَ طبقة VSTO/COM، لأنها مسار Windows منفصل يحتاج مشروعًا وأداة تثبيت واختبارات خاصة. لا يجوز اعتبار الجدول السابق مصادقة توافق نهائية قبل تنفيذ runbook على نسخ Office المستهدفة.

## مراجع

[1]: https://learn.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins "Office Add-ins platform overview — Microsoft Learn"
[2]: https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/requirements-for-running-office-add-ins "Requirements for running Office Add-ins — Microsoft Learn"
