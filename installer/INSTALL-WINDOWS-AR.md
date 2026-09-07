# تثبيت Nawa OCR Office على Windows

هذه الحزمة تثبت ملف الإضافة وتجهزه للاستخدام داخل **Word وExcel**. الإضافة الحالية هي Office Add-in تعمل من خلال ملف `manifest.xml`، ولذلك لا تُسجل كـ DLL أصلية داخل Office.

## التثبيت

شغّل `Nawa-OCR-Office-Setup.ps1` باستخدام PowerShell. إذا منع Windows تشغيل السكربت، افتح PowerShell داخل مجلد الحزمة وشغّل:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\Nawa-OCR-Office-Setup.ps1
```

بعد ذلك افتح Word أو Excel، ثم انتقل إلى:

`Insert` ← `Get Add-ins` ← `Manage My Add-ins` ← `Upload My Add-in`

اختر الملف الذي جهزه المثبّت في المسار:

`%LOCALAPPDATA%\NawaOCROffice\manifest.xml`

بعد تثبيت الإضافة سيظهر زر **استيراد وتحويل** في شريط Word أو Excel. عند الضغط عليه تفتح نافذة Nawa OCR ويظهر اختيار الملف تلقائيًا، ثم تُحدد الوجهة حسب التطبيق الحالي.

## ملاحظة مهمة

يتطلب ظهور زر الإضافة داخل شريط Office تنفيذ خطوة Upload My Add-in مرة واحدة لأن Office يفرض تسجيل manifest من داخل Office أو عبر إدارة المؤسسة. يمكن لمسؤول Microsoft 365 استخدام Centralized Deployment إذا أراد توزيعها على عدة أجهزة.

هذه الحزمة لا ترفع الملفات تلقائيًا إلى خدمة سحابية. كما أن إنشاء ملف EXE النهائي يحتاج تشغيل Inno Setup على Windows؛ ملف `Nawa-OCR-Office-Setup.iss` مرفق لتوليد المثبّت.
