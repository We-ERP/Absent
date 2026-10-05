# Absenteeism Dashboard — © Seif 2026

داشبورد ثابت (HTML/CSS/JS) بيشتغل على GitHub Pages، والـ Backend هو Google Apps Script بيقرا ويكتب في الشيتات:
`Codes SHR` · `Codes absent` · `STR`

## 1) إعداد الـ Apps Script (الباك إند)
1. افتح الشيت > Extensions > Apps Script، وانسخ محتوى `Code.gs` مكان الكود القديم.
2. Project Settings > Script properties > أضف `ADMIN_PASSWORD` = كلمة السر بتاعتك.
3. Deploy > Manage deployments > Edit > **New version** (Execute as: Me — Who has access: Anyone).
4. لو الرابط اتغير، حدّث `API` في أول `app.js`.

## 2) الرفع على GitHub
1. أنشئ Repository جديد وارفع الملفات: `index.html` · `style.css` · `app.js` · `Code.gs` · `README.md`
2. Settings > Pages > Branch: `main` / root > Save.
3. افتح الرابط وسجّل دخول بكلمة سر الأدمن.

## طريقة الاستخدام
- **الداشبورد**: ارفع ملف الاسكدول (RD) — بيتعالج بنفس منطق الـ VBA (حذف أول 20 صف، ID/Date/Duration/Code، ربط STR، حالة Shrinkage/Absent). النسبة = Absent ÷ Total Scheduled Time.
- **تصدير Final**: بيطلع الأعمدة A:F فقط (ID, Date, Duration, Agent Name, TL, Code).
- **Codes SHR / Codes absent / STR**: تعديل مباشر في الخلية، إضافة، حذف، ورفع ملف بيستبدل الجدول كله مباشرة في جوجل شيت (أول صف في الملف = هيدر).
- أي كود مش موجود في شيتات الأكواد بيظهر `#N/A` وبيتعرض تحت الرفع علشان تضيفه.

> لو عمود المدة في الاسكدول مش J أو الهيدر مش 20 صف، عدّلهم من خانات الإعدادات في الداشبورد.

## المستخدمين والصلاحيات (v2)
- اليوزر الرئيسي دايماً `admin` وكلمة السر = `ADMIN_PASSWORD` (كل الصلاحيات).
- من تاب **المستخدمين** الأدمن بيضيف/يعدّل/يحذف يوزرات ويحدد صلاحياتهم: `تشغيل` · `تصدير` · `تعديل الأكواد` · `تعديل STR` · `إدارة المستخدمين`.
- اليوزر العادي (تشغيل + تصدير) بيرفع الاسكدول ويشوف الفيو ويصدّر بس.
- الباسوردات متخزنة Hash (SHA-256) في شيت `Users` المخفي، والصلاحيات متتأكد منها في السيرفر مش في المتصفح.
- بعد لصق `Code.gs` الجديد: Deploy > Manage deployments > Edit > **New version**.
