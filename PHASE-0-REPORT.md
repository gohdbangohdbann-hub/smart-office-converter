# PHASE 0 REPORT — Nawa OCR Office مقابل مواصفات Smart Office Converter

## الحالة

**Status: Completed — Phase 0 فقط.** تمت مراجعة المواصفات المرفقة ومقارنتها ببنية Nawa OCR Office الحالية. لم تبدأ Phase 1 الجديدة، ولم تُضف وظائف تجريبية أو معادلات مصطنعة أو بيانات OCR وهمية.

> القرار الرئيسي: نحافظ على إضافة Office.js الحالية كمسار حديث لـ Word وExcel، ونفصل مستقبلًا مسارًا أصليًا بـ VSTO/.NET Framework إذا كان دعم Office 2010 و2013 شرطًا إلزاميًا. تحويل المشروع الحالي إلى DLL ليس ترقية مباشرة، بل إعادة بناء تقنية مختلفة.

## 1. النتيجة التنفيذية

المشروع الحالي يحقق جزءًا كبيرًا من النواة المطلوبة: Task Pane واحدة لـ Word وExcel، استيراد PDF والصور، خطط Word وExcel موحدة، RTL/LTR، دعم العربية والفرنسية والإنجليزية في الواجهة، معاينة قبل التصدير، تصدير DOCX/XLSX، إدراج عبر Word.run وExcel.run، مراجعة وثقة، حماية من التغيير الصامت، معالجة جماعية، إعدادات محلية، وسياسة معالجة محلية دون رفع تلقائي.

الجزء غير المكتمل جوهريًا هو طبقة **الفهم الدلالي والاستدلال الحسابي**. الكود الحالي يصنّف الخلايا ويتحقق من بعض التناقضات الحسابية، لكنه لا يستنتج صيغ Excel موثوقة مثل `=B2*C2` ولا يعرضها كمقترحات بعتبات ثقة. لذلك يجب إبقاء هذه الوظيفة خارج الإدراج التلقائي حتى تُبنى طبقة Formula Intelligence مستقلة ومختبرة.

## 2. قرار التقنية والتوافق

| المجال | القرار | السبب |
|---|---|---|
| Word وExcel الحديثة | الإبقاء على Office Add-in الحالية المبنية على React وOffice.js وmanifest.xml | تحافظ على Task Pane، المعاينة، وإمكانية تشغيل واجهة واحدة في Word وExcel. |
| Office 2010 و2013 | عدم إعلان دعم فعلي في المسار الحالي؛ إنشاء مسار VSTO/.NET مستقل إذا كان الدعم إلزاميًا | وثائق Microsoft الحالية تذكر أن تشغيل Office Add-ins على Windows يتطلب عادة Office 2016 أو أحدث، مع اختلاف الدعم حسب Requirement Set. [1] |
| Office 2016 و2019 و2021 وMicrosoft 365 | اعتمادها كأهداف التوافق الأولية، مع اختبار فعلي لكل إصدار متاح | Microsoft تعتمد Requirement Sets للتحقق من توفر API، ولا يكفي اسم إصدار Office وحده. [2] |
| DLL أصلية | لا تُنشأ من الكود الحالي | DLL تتطلب VSTO/COM وRegistry وOffice Interop وثقة وتوقيعًا، وهي إعادة بناء لا تحويلًا آليًا. [3] |
| EXE | استخدامه كمثبت أو تطبيق مساعد، لا كبديل عن manifest | Office Add-in تحتاج ملفات الواجهة وmanifest على خادم HTTPS، ثم نشر manifest عبر catalog أو integrated apps portal أو sideload مناسب. [1] |
| OCR الحقيقي | مزود مستقل خلف عقد OCRProvider | يسمح بالإبقاء على الوضع المحلي، وإضافة Azure أو مزود آخر دون ربط الواجهة بالمزود. |

توصي هذه المرحلة بعدم إهدار الاستقرار الحالي بمحاولة جعل Office 2010 يعمل قسرًا عبر Office.js. إذا كان شرط Office 2010 غير قابل للتفاوض، فالمسار الصحيح هو منتج مزدوج: Office.js للإصدارات الحديثة، وVSTO/.NET Framework 4.x لـ Office 2010–2013، مع مشاركة Core Engine بعقد مستقلة قدر الإمكان.

## 3. ما يجب إبقاؤه كما هو

| الميزة الحالية | التقييم | القرار |
|---|---|---|
| زر Ribbon «استيراد وتحويل» في Word وExcel | يعمل منطقيًا عبر manifest ومعاملات host | إبقاء وتحسين التثبيت والاختبار المكتبي. |
| Task Pane واحدة | مناسبة لتقليل تكرار الكود | إبقاء مع فصل أوضح بين Word وExcel adapters. |
| العربية وRTL/LTR وواجهة AR/FR/EN | نواة صحيحة ومختبرة محليًا | إبقاء وتوسيع تغطية النصوص وعدم خلط اللغات. |
| WordDocumentPlan وExcelWorkbookPlan | عقد موحدة قبل الإدراج والتصدير | إبقاء وجعلها مصدرًا لطبقة Review وFormula Suggestions. |
| معاينة DOCX/XLSX | موجودة مع مراحل تحميل وتكبير وتنقل | إبقاء وتحسين ربطها بمصادر الخلايا. |
| التحقق والثقة والمراجعة | يمنع الإدراج الصامت ويعرض المشكلات | إبقاء، وإضافة نوع مستقل لصيغة مقترحة وتناقض حسابي. |
| معالجة الملفات محليًا والسجلات الآمنة | منسجمة مع Privacy Mode | إبقاء وعدم تسجيل النص الكامل أو bytes. |
| Batch Processing | موجود للمخرجات والتحويلات | إبقاء وإضافة تقرير إجمالي قابل للتصدير لاحقًا. |
| تصدير DOCX وXLSX | موجود ويُستخدم من نفس الخطط | إبقاء مع تحسين التنسيقات الرقمية والصيغ بعد بناء Formula Engine. |

## 4. ما يحتاج إلى تحسين قبل اعتباره تجاريًا

أولًا، يجب فصل نتائج النظام إلى أربع مراحل صريحة: **Extraction** لما قرأه OCR، و**Interpretation** لمعنى الخلية أو الحقل، و**Inference** للصيغة المقترحة، و**Validation** للتحقق قبل **Output**. النموذج الحالي قوي في Extraction وValidation المحدود، لكنه يحتاج عقودًا مستقلة لـ Interpretation وInference.

ثانيًا، يجب تحسين Excel Formatting. التصنيف الحالي يدعم النص والرقم والتاريخ والنسبة والعملة بصورة محافظة، لكن خطة Excel تحتاج إلى توسيع `numberFormat` و`currency` و`date`، وربط القيم الرقمية بالتنسيق دون تحويل غير قابل للمراجعة. يجب أن يبقى النص الأصلي محفوظًا إلى جانب القيمة المطبّعة.

ثالثًا، يجب تحسين المصدر القابل للتتبع. توجد مراجع مصدر في تقرير التحقق وبعض polygon/page metadata، لكن الواجهة لا تعرض بعد عند تحديد خلية: الصفحة، الجدول، الصف، العمود، والنص الأصلي. هذه ميزة مطلوبة قبل الاعتماد التجاري.

رابعًا، يجب تحسين دورة PDF. المسار المحلي الحالي يتحقق من PDF ويحافظ على bytes لمزود Layout، لكنه لا يقدم استخراجًا نصيًا أصليًا كاملًا كمسار مستقل قبل OCR. سيتم إنشاء PDF Strategy لاحقًا: استخراج النص الأصلي أولًا عندما يكون موثوقًا، ثم OCR للصفحات الممسوحة أو غير القابلة للاستخراج.

خامسًا، يجب تحسين التوزيع. حزمة Windows الحالية تحتوي مصادر Inno Setup وPowerShell وتعليمات عربية، لكنها ليست EXE مبنيًا؛ كما أن Office Add-in لا تصبح مسجلة داخل Word وExcel بمجرد نسخ manifest. يلزم اختبار تثبيت Windows حقيقي، وتوقيع الحزمة، واختيار آلية نشر manifest المناسبة: sideload للتجربة أو Centralized Deployment للمؤسسة.

## 5. ما يجب إنشاؤه في مراحل لاحقة

| المكوّن الجديد | الهدف | سياسة الأمان |
|---|---|---|
| `FormulaInference` | اقتراح SUM وAVERAGE والضرب والخصم والضريبة والإجمالي عند وجود دليل | لا إدراج تلقائي إلا عند ثقة عالية؛ المتوسط للمراجعة؛ المنخفض اقتراح فقط. |
| `FormulaValidation` | مقارنة الصيغة بالقيم المستخرجة وإظهار القيمة المتوقعة والفرق | لا تعديل صامت للقيمة الأصلية. |
| `SemanticTableModel` | تمثيل Quantity وUnit Price وSubtotal وTax وTotal وغيرها | كل تفسير يحتفظ بالنص الأصلي والثقة والمصدر. |
| `SourceTracePanel` | فتح موضع المصدر عند تحديد خلية أو فقرة | لا يحفظ صورة أو مستندًا دائمًا دون طلب. |
| `OriginalVsExtracted` | مقارنة المصدر مع الناتج وتحديد الاختلافات | يعتمد على bytes مؤقتة أو مصدر يختاره المستخدم. |
| `DataCleaning` | مسافات وفواصل وأرقام ووحدات وتكرارات | كل تعديل يظهر في سجل Before/After. |
| `History` | سجل محلي للعملية والنتيجة والتحذيرات | تخزين metadata فقط افتراضيًا، دون نص المستند. |
| `Cancellation` | إلغاء التحويلات الطويلة والدفعات | يجب ألا يترك ملفات مؤقتة. |
| `VSTO Compatibility Track` | دعم Office 2010 و2013 عند اعتماد هذا الشرط | مشروع .NET مستقل، Registry وTrust وتوقيع واختبارات Windows. |
| `Signed Installer` | مثبت EXE احترافي مع Repair وUninstall وOffice detection | توقيع Authenticode وشروط تشغيل واضحة. |

## 6. المعمارية المقترحة

```text
Office Ribbon / Task Pane
        |
        v
Office Adapters
  WordAdapter | ExcelAdapter
        |
        v
Review + Preview + Output Contracts
        |
        +------------------+
        |                  |
        v                  v
Extraction Layer      Interpretation Layer
OCR/PDF/Image         Types, fields, table semantics
        |                  |
        +--------+---------+
                 v
       Formula Inference Layer
                 |
                 v
       Validation + Confidence
                 |
                 v
          Word / Excel Output

Provider Boundary:
OCRProvider | Optional Azure | Future Providers
Privacy Boundary:
Local Processing | Optional Cloud | Secure Logs
Distribution Boundary:
Office.js manifest | VSTO legacy track | Signed installer
```

يحافظ هذا التصميم على النواة الحالية بدل استبدالها. تتم إضافة Formula Intelligence وSource Trace وPDF Strategy كطبقات مستقلة، ولا تُخلط مع OCR أو Office insertion.

## 7. مصفوفة التوافق الأولية

| البيئة | وضع المشروع الحالي | الحالة المطلوبة قبل الإعلان بالدعم |
|---|---|---|
| Word/Excel 2010 | غير مضمون مع Office.js الحالي | قرار VSTO مستقل واختبار Windows فعلي. |
| Word/Excel 2013 | غير مضمون مع Office.js الحالي | اختبار Requirement Sets أو مسار VSTO. |
| Office 2016 | مرشح للتشغيل | sideload، إدراج Word وExcel، إعادة فتح الملفات، واختبار الواجهات. |
| Office 2019 | مرشح للتشغيل | نفس اختبارات Office 2016 مع فحص API differences. |
| Office 2021 | مرشح للتشغيل | اختبار تثبيت وإدراج وتصدير. |
| Microsoft 365 | الهدف الحديث الأساسي | اختبار Windows حديث مع Office.js cache وRibbon. |
| Office Web | ليس هدف التوزيع المحلي الحالي | اختبار منفصل إذا أُضيف لاحقًا. |

لا توجد في هذه المرحلة أرقام دقة OCR؛ فالدقة تحتاج ملفات benchmark حقيقية وground truth. كما لم يُدّع تشغيل EXE أو sideload داخل Windows لأن بيئة التنفيذ الحالية Linux ولا تحتوي Office Desktop.

## 8. خطة التنفيذ بعد موافقة المستخدم

ستبدأ Phase 1 من النواة المشتركة وعقود `SemanticTableModel` و`FormulaSuggestion`، ثم Phase 2 بتحسين Word، ثم Phase 3 ببناء Excel Intelligence، ثم Phase 4 بمحرك المعادلات والتحقق، ثم Phase 5 بالمراجعة والمصادر، ثم Phase 6 بالتوزيع والاختبارات المكتبية. لا ينبغي البدء بالمثبت النهائي قبل تثبيت آلية نشر manifest واختيار نطاق Office المدعوم.

## 9. ملاحظات صريحة

المشروع الحالي **ليس Demo شكليًا** في مسارات المعاينة والخطط والإدراج المحلي والاختبارات، لكنه أيضًا **ليس بعد منتج OCR تجاريًا مكتملًا** لأن مزود OCR الحقيقي، وFormula Inference، واختبارات Office Desktop، ودعم Office 2010، والتوقيع النهائي للمثبت ما تزال خارج التحقق المحلي. هذا الفصل ضروري حتى لا تُنسب للمشروع دقة أو توافق لم يتم قياسهما.

## المراجع

[1]: https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/requirements-for-running-office-add-ins "Microsoft Learn — Requirements for running Office Add-ins"

[2]: https://learn.microsoft.com/en-us/office/dev/add-ins/develop/office-versions-and-requirement-sets "Microsoft Learn — Office versions and requirement sets"

[3]: https://learn.microsoft.com/en-us/visualstudio/vsto/deploying-a-vsto-solution-by-using-windows-installer?view=visualstudio "Microsoft Learn — Deploying a VSTO Solution Using Windows Installer"
