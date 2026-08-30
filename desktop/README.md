# NumberToWords Desktop

هذا المجلد يضيف غلاف Windows مستقلًا لتطبيق Nawa OCR Office. يعمل الغلاف على تحميل `dist/index.js` محليًا داخل Electron، ثم يفتح الواجهة نفسها المستخدمة في المتصفح وWord وExcel. محرك تحويل الأرقام موجود في `shared/number-to-words.ts` ولا يكرر داخل الغلاف.

## تشغيل التطوير

```bash
pnpm build
pnpm desktop:dev
```

## إنتاج Windows

```bash
pnpm desktop:dist
```

ينتج Electron Builder ملفًا محمولًا أو مثبت NSIS في `dist/`، ويستخدم اسم المنتج `NumberToWords`. الهدف الافتراضي هو Windows x64، ويمكن إضافة `--ia32` أو `--arm64` عند الحاجة إلى بناء منفصل. لا يتطلب التطبيق المستقل Office لتشغيل التحويل والنسخ والإعدادات، بينما تتطلب وظائف الإدراج في Word وExcel وجود Office والـ Add-in محمّلًا داخل التطبيق المضيف.

أُعد ملف التغليف في `electron-builder.yml` مع عزل Electron (`contextIsolation` و`nodeIntegration: false`). إنتاج ملف Windows نهائي من Linux قد يحتاج إلى Wine أو تشغيل أمر التغليف على Windows؛ لذلك يجب اختبار الملف الناتج على Windows قبل توزيعه، كما يجب توقيعه بشهادة موثوقة في بيئة الإصدار.
