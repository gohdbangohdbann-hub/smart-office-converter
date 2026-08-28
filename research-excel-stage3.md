# بحث المرحلة الثالثة — Excel JavaScript API

## المصادر الرسمية

1. [Get Excel worksheet ranges using the JavaScript API](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-ranges-get): يستخدم Excel JavaScript API كائن Range لقراءة وكتابة وتنسيق الخلايا. يمكن الوصول إلى نطاق بعنوان مثل A1:D4 ثم تعيين values كمصفوفة ثنائية الأبعاد، مع تنفيذ العمليات داخل Excel.run ثم context.sync().

2. [Create, read, and manage tables with the Excel JavaScript API](https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-tables): يدعم Excel JavaScript API إنشاء جدول من نطاق، إضافة صفوف، قراءة البيانات، وإتاحة الفرز والتصفية والتنسيق. النموذج الصحيح للنتيجة هو خلايا حقيقية داخل Worksheet وليس صورة أو نصًا مدمجًا في خلية واحدة.

## القرار المعماري

ستنتج طبقة Excel خطة مستقلة عن Excel.js تحتوي على worksheets وجداول ومصفوفات خلايا typed values، ثم ينفذ adapter منفصل Excel.run عمليات إضافة Worksheet، تعيين Range.values، تطبيق التنسيق، وإنشاء Table عند تحقق وجود رأس. الدمج لن يُنفذ إلا إذا كانت إحداثيات المصدر تؤكده؛ وإلا تحفظ الخلايا منفصلة دون تخمين.

سيحافظ المصنف على الأصفار الافتتاحية كنص، ولن يحول التاريخ أو العملة أو الصيغة إلا عندما يكون النمط واضحًا. عند إدراج النتيجة يستخدم Excel.run وRange.values/format بدل إدراج صورة أو سلسلة كاملة في خلية واحدة.
