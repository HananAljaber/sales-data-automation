<div dir="rtl" align="right">

# أتمتة معالجة بيانات مبيعات متجر ملابس

**Sales Data Automation | Python • Pandas • HTML • Excel**

| معلومات المشروع | التفاصيل |
|---|---|
| **تاريخ تنفيذ المشروع** | **12–13 نوفمبر 2026** |
| **الفترة التي تغطيها بيانات المبيعات** | **1 يناير 2025 – 30 يونيو 2026** (18 شهرًا) |
| **مصدر البيانات** | 20 ملف HTML |
| **عدد السجلات بعد التجميع** | 4,862 سجلًا |
| **عدد الأعمدة** | 14 عمودًا |
| **الأدوات المستخدمة** | Python، Pandas، Jupyter Notebook، Excel |
| **العمل المنفّذ** | استخراج البيانات، فحصها، تنظيفها، التحقق من جودتها، وتصديرها إلى Excel |

## فكرة المشروع

توجد بيانات مبيعات متجر ملابس داخل **20 ملف HTML**، تغطي فترة **سنة ونصف**. بدلًا من فتح كل ملف على حدة ونسخ بياناته يدويًا، استخدمتُ **Python** ومكتبة **Pandas** لقراءة الملفات وتجميعها في جدول واحد.

بعد جمع البيانات، فحصتُ الأعمدة واكتشفتُ ثلاث مشكلات:

<ul dir="rtl" align="right">
  <li><strong>التواريخ:</strong> وجدتُ 3 سجلات بتنسيق تاريخ مختلف عن بقية البيانات.</li>
  <li><strong>الأسعار:</strong> وجدتُ 4 سجلات يحتوي فيها سعر الوحدة على رمز العملة «ر.س» مع الرقم.</li>
  <li><strong>رموز العملاء:</strong> وجدتُ 5 سجلات ينقصها رمز العميل.</li>
</ul>

عالجتُ هذه المشكلات، ثم تحققتُ من النتائج قبل تصدير البيانات إلى **Excel**.

**النتيجة:** جدول موحّد يضم **4,862 سجل مبيعات** بدلًا من التعامل مع 20 تقريرًا منفصلًا.

## كيف نُفّذ المشروع؟

اتبعتُ مراحل **ETL**:

<ul dir="rtl" align="right">
  <li><strong>Extract — استخراج:</strong> قراءة جداول المبيعات من ملفات HTML ودمجها.</li>
  <li><strong>Transform — تحويل وتنظيف:</strong> توحيد التواريخ والأسعار، ومعالجة رموز العملاء المفقودة، وتحسين أسماء الأعمدة.</li>
  <li><strong>Load — تحميل:</strong> تصدير البيانات المعالجة إلى ملف Excel.</li>
</ul>

وأضفتُ فحص جودة البيانات **Data Profiling** قبل التنظيف، والتحقق من النتائج **Data Validation** بعده؛ حتى أتعرف على المشكلة قبل تعديلها وأتأكد من حلّها بعد التعديل.

---

## 01 — Extract | استخراج وتجميع بيانات المبيعات

كان مصدر البيانات مجلدًا يحتوي على **20 ملف HTML**. جميع الملفات تعرض نفس أعمدة المبيعات، لكن سجلاتها موزعة على أشهر مختلفة؛ بعض الأشهر مقسّمة إلى ملفين.

استخدمتُ حلقة `for` لقراءة كل ملف بدل كتابة أمر منفصل لكل ملف، ثم `pd.concat()` لجمع الجداول فوق بعضها في DataFrame واحد.

</div>

<div dir="ltr" align="left">
<pre dir="ltr"><code class="language-python">from pathlib import Path
import pandas as pd

# المجلد الذي يحتوي على ملفات المبيعات
html_folder_path = Path("HTML_Files")

# اختيار ملفات HTML فقط وترتيب أسمائها
html_files = sorted(html_folder_path.glob("*.html"))

# قائمة للاحتفاظ بالجداول المستخرجة
sales_dataframes = []

for file_path in html_files:
    # قراءة الجدول الأول من كل صفحة HTML
    sales_df = pd.read_html(str(file_path), encoding="utf-8")[0]
    sales_dataframes.append(sales_df)

# جمع جميع الصفوف في جدول واحد
combined_sales_df = pd.concat(
    sales_dataframes,
    ignore_index=True
)

print("عدد الملفات:", len(html_files))
print("حجم البيانات:", combined_sales_df.shape)</code></pre>
</div>

<div dir="rtl" align="right">

**مخرجات التشغيل:**

- عدد الملفات: **20**
- حجم البيانات: **4,862 صفًا × 14 عمودًا**

</div>

<div dir="rtl" align="right">

**النتيجة:** قراءة **20 ملفًا** وتجميع **4,862 صفًا و14 عمودًا**. الخيار `ignore_index=True` يعيد ترقيم صفوف الجدول بعد الدمج؛ ولا يكرّر رؤوس الأعمدة داخل البيانات.

### ماذا تعني الأعمدة؟

| اسم العمود في الملفات | معناه |
|---|---|
| التاريخ | تاريخ البيع |
| رقم الفاتورة | رقم الفاتورة؛ قد يظهر في أكثر من صف إذا كانت تحتوي عدة منتجات |
| رقم البند | رقم يميّز كل بند مبيعات |
| الفرع | فرع المتجر |
| رمز العميل | الرقم التعريفي للعميل |
| رمز المنتج | الرقم التعريفي للمنتج |
| اسم المنتج | اسم قطعة الملابس أو المنتج |
| الفئة | تصنيف المنتج |
| الكمية | عدد القطع في البند |
| سعر الوحدة | سعر بيع القطعة قبل الخصم |
| الخصم % | نسبة الخصم على البند |
| إجمالي السطر | قيمة البند المسجّلة بعد الخصم |
| تكلفة الوحدة | تكلفة القطعة على المتجر |
| حالة الطلب | حالة الطلب المسجّلة في النظام |

---

## 02 — Data Profiling | فحص البيانات واكتشاف المشكلات

بعد جمع الملفات، لم أبدأ بحذف القيم أو تغييرها مباشرة. فحصتُ البيانات أولًا لمعرفة ما إذا كانت هناك قيم مفقودة أو صفوف مكررة أو تنسيقات مختلفة.

### 2.1 فحص حجم البيانات وأنواع الأعمدة

</div>

<div dir="ltr" align="left">
<pre dir="ltr"><code class="language-python">print("حجم البيانات:", combined_sales_df.shape)
print("أسماء الأعمدة:", combined_sales_df.columns.tolist())
print("أنواع البيانات:")
print(combined_sales_df.dtypes)</code></pre>
</div>

<div dir="rtl" align="right">

أظهر الفحص **4,862 سجلًا و14 عمودًا**. ولاحظتُ أن عمود `سعر الوحدة` ليس من نوع رقمي موحّد، مما جعلني أراجع القيم الموجودة داخله.

### 2.2 اكتشاف القيم المفقودة

استخدمتُ دالة تعرض عدد القيم المفقودة ونسبتها في كل عمود:

</div>

<div dir="ltr" align="left">
<pre dir="ltr"><code class="language-python">def missing_values_per_columns(data):
    missing_count = data.isna().sum()
    missing_percentage = (data.isna().mean() * 100).round(3)

    missing_report = pd.DataFrame({
        "عدد القيم المفقودة": missing_count,
        "النسبة المئوية": missing_percentage
    })
    display(missing_report)

missing_values_per_columns(combined_sales_df)</code></pre>
</div>

<div dir="rtl" align="right">

**مخرجات التشغيل:**

| العمود | عدد القيم المفقودة |
|---|---:|
| رمز العميل | 5 |
| بقية الأعمدة | 0 |

</div>

<div dir="rtl" align="right">

**النتيجة الأولى:** اكتشفتُ **5 سجلات** فيها `رمز العميل` مفقود. ولم تكن هناك قيم فارغة ظاهرة في التاريخ أو السعر؛ لكن هذا لا يعني أن تنسيق كل القيم صحيح.

### 2.3 فحص الصفوف المتطابقة

</div>

<div dir="ltr" align="left">
<pre dir="ltr"><code class="language-python">print("الصفوف المتطابقة بالكامل:", combined_sales_df.duplicated().sum())</code></pre>
</div>

<div dir="rtl" align="right">

**مخرجات التشغيل:**

<pre dir="rtl" align="right"><code>الصفوف المتطابقة بالكامل: 0</code></pre>

</div>

<div dir="rtl" align="right">

**النتيجة:** لا توجد صفوف متطابقة بالكامل. وهذا يختلف عن فحص تكرار `رقم البند` الذي أجريته لاحقًا.

### 2.4 كيف اكتشفتُ التواريخ المكتوبة بتنسيق مختلف؟

معظم التواريخ مكتوبة بصيغة **يوم/شهر/سنة**، مثل `03/05/2025`. اختبرتُ قراءة كل تاريخ بهذا التنسيق، ثم عرضتُ السجلات التي لم تتوافق معه:

</div>

<div dir="ltr" align="left">
<pre dir="ltr"><code class="language-python"># اختبار التواريخ دون تعديل البيانات الأصلية
date_check = pd.to_datetime(
    combined_sales_df["التاريخ"],
    format="%d/%m/%Y",
    errors="coerce"
)

invalid_dates_df = combined_sales_df.loc[
    date_check.isna() &amp; combined_sales_df["التاريخ"].notna(),
    ["رقم البند", "التاريخ"]
]

print("عدد التواريخ مختلفة التنسيق:", len(invalid_dates_df))
display(invalid_dates_df)</code></pre>
</div>

<div dir="rtl" align="right">

**مخرجات التشغيل:**

<pre dir="rtl" align="right"><code>عدد التواريخ مختلفة التنسيق: 3</code></pre>

</div>

<div dir="rtl" align="right">

**النتيجة: 3 سجلات** مختلفة التنسيق:

| رقم الصف | القيمة الأصلية | المشكلة |
|---:|---|---|
| 972 | `2025-05-03` | مكتوب سنة-شهر-يوم |
| 2438 | `2025-11-03` | مكتوب سنة-شهر-يوم |
| 3920 | `2026-04-02` | مكتوب سنة-شهر-يوم |

هذه التواريخ ليست خاطئة في ذاتها؛ إنما **طريقة كتابتها مختلفة** عن التنسيق المتوقع. استخدمتُ `errors="coerce"` حتى تظهر القيم التي لم تُقرأ بنجاح على أنها `NaT` في **نتيجة الاختبار فقط**.

### 2.5 كيف اكتشفتُ الأسعار المكتوبة بطريقة مختلفة؟

عندما حاولتُ قراءة أسعار الوحدات كأرقام، اكتشفتُ أن بعض القيم تحتوي على **`ر.س`** داخل النص:

</div>

<div dir="ltr" align="left">
<pre dir="ltr"><code class="language-python"># اختبار الأسعار دون تعديل البيانات الأصلية
price_check = pd.to_numeric(
    combined_sales_df["سعر الوحدة"],
    errors="coerce"
)

invalid_prices_df = combined_sales_df.loc[
    price_check.isna() &amp; combined_sales_df["سعر الوحدة"].notna(),
    ["رقم البند", "سعر الوحدة"]
]

print("عدد الأسعار مختلفة التنسيق:", len(invalid_prices_df))
display(invalid_prices_df)</code></pre>
</div>

<div dir="rtl" align="right">

**مخرجات التشغيل:**

<pre dir="rtl" align="right"><code>عدد الأسعار مختلفة التنسيق: 4</code></pre>

</div>

<div dir="rtl" align="right">

**النتيجة: 4 سجلات**:

| رقم الصف | السعر كما ورد في الملف |
|---:|---|
| 714 | `ر.س 115.00` |
| 1927 | `ر.س 140.00` |
| 3071 | `ر.س 112.20` |
| 4558 | `ر.س 127.50` |

الأسعار نفسها موجودة، لكن **رمز العملة يمنع تحويلها مباشرة إلى أرقام**.

### 2.6 عرض المشكلات في تقرير واحد

حتى أرى نتائج اختبار التاريخ والسعر معًا، استخدمتُ نسخة مؤقتة من البيانات، دون تغيير القيم الأصلية:

</div>

<div dir="ltr" align="left">
<pre dir="ltr"><code class="language-python">profiling_sales_df = combined_sales_df.copy()
profiling_sales_df["التاريخ"] = date_check
profiling_sales_df["سعر الوحدة"] = price_check

missing_values_per_columns(profiling_sales_df)</code></pre>
</div>

<div dir="rtl" align="right">

| العمود | ما كشفه الاختبار |
|---|---:|
| التاريخ | 3 قيم لم تتوافق مع صيغة التاريخ المحددة |
| سعر الوحدة | 4 قيم لم يمكن تحويلها مباشرة إلى أرقام |
| رمز العميل | 5 قيم مفقودة من الأصل |

**ملاحظة مهمة:** القيم الـ3 والـ4 لم تكن مفقودة في الملفات؛ ظهرت كقيم غير قابلة للتفسير أثناء الاختبار. وهذا سبب أهمية مراجعتها قبل الحذف أو التعويض.

---

## 03 — Transform | تنظيف البيانات وتحويلها

بعد تحديد المشكلات، بدأتُ معالجتها. أستخدم في الأكواد التالية نسخة من البيانات حتى تبقى القيم المستخرجة من الملفات متاحة للمقارنة:

</div>

<div dir="ltr" align="left">
<pre dir="ltr"><code class="language-python">cleaned_sales_df = combined_sales_df.copy()</code></pre>
</div>

<div dir="rtl" align="right">

### 3.1 توحيد التواريخ

</div>

<div dir="ltr" align="left">
<pre dir="ltr"><code class="language-python">cleaned_sales_df["التاريخ"] = pd.to_datetime(
    cleaned_sales_df["التاريخ"],
    format="mixed",
    dayfirst=True,
    errors="coerce"
)</code></pre>
</div>

<div dir="rtl" align="right">

- `format="mixed"` يتيح قراءة أكثر من شكل للتاريخ.
- `dayfirst=True` يفضّل تفسير اليوم قبل الشهر في التواريخ الملتبسة.
- بعد التحويل، راجعتُ القيم للتأكد من أنها تمثل الأيام الصحيحة، ولا سيما التواريخ الثلاثة التي تبدأ بالسنة.

### 3.2 تنظيف سعر الوحدة

</div>

<div dir="ltr" align="left">
<pre dir="ltr"><code class="language-python">cleaned_sales_df["سعر الوحدة"] = pd.to_numeric(
    cleaned_sales_df["سعر الوحدة"]
    .astype(str)
    .str.replace("ر.س", "", regex=False)
    .str.strip(),
    errors="coerce"
)</code></pre>
</div>

<div dir="rtl" align="right">

في هذا الكود حوّلتُ القيم إلى نص مؤقتًا باستخدام `astype(str)`، ثم أزلتُ رمز العملة باستخدام `str.replace()`، وحذفتُ المسافات الزائدة باستخدام `str.strip()`. أخيرًا حوّلتُ القيم إلى أرقام باستخدام `pd.to_numeric()`.

مثلًا، أصبحت القيمة `ر.س 115.00` تساوي **`115.0`**، وأصبح بالإمكان استخدامها في الحسابات.

### 3.3 تعويض رموز العملاء المفقودة

لم أضع رموزًا عشوائية بدل القيم المفقودة. بحثتُ عن **نفس رقم الفاتورة**، وتحققتُ من رمز العميل في بنود الفاتورة الأخرى. إذا كانت جميع الرموز المعروفة في الفاتورة متفقة على رمز واحد، استخدمتُه لملء القيمة المفقودة.

</div>

<div dir="ltr" align="left">
<pre dir="ltr"><code class="language-python">missing_customer_rows = cleaned_sales_df.index[
    cleaned_sales_df["رمز العميل"].isna()
]

for row_index in missing_customer_rows:
    invoice_number = cleaned_sales_df.loc[row_index, "رقم الفاتورة"]

    matching_customer_ids = cleaned_sales_df.loc[
        (cleaned_sales_df["رقم الفاتورة"] == invoice_number)
        &amp; cleaned_sales_df["رمز العميل"].notna(),
        "رمز العميل"
    ].unique()

    # التعويض فقط إذا كان هناك رمز معروف واحد غير متعارض
    if len(matching_customer_ids) == 1:
        cleaned_sales_df.loc[row_index, "رمز العميل"] = matching_customer_ids[0]</code></pre>
</div>

<div dir="rtl" align="right">

**النتيجة:** أمكن استرجاع رموز العملاء الخمسة من فواتيرها، دون حذف سجلات المبيعات.

### 3.4 استخدام أسماء أعمدة أوضح

غيّرتُ اسمين حتى يعرف قارئ الجدول معنى المبلغ المسجل:

| الاسم قبل التعديل | الاسم بعد التعديل |
|---|---|
| `سعر الوحدة` | `سعر بيع الوحدة` |
| `إجمالي السطر` | `صافي مبيعات المنتج` |

</div>

<div dir="ltr" align="left">
<pre dir="ltr"><code class="language-python">cleaned_sales_df = cleaned_sales_df.rename(columns={
    "سعر الوحدة": "سعر بيع الوحدة",
    "إجمالي السطر": "صافي مبيعات المنتج"
})</code></pre>
</div>

<div dir="rtl" align="right">

يقصد بـ**صافي مبيعات المنتج** في هذا الجدول قيمة **بند الفاتورة بعد الخصم**؛ وليس بالضرورة صافي إيرادات المتجر بعد المرتجعات.

---

## 04 — Data Validation | التأكد من صحة النتائج

بعد التنظيف، أجريتُ اختبارات للتأكد من أن القيم سليمة، وأن المعالجة لم تُفقدنا سجلات مبيعات.

### 4.1 مقارنة عدد الصفوف والقيم المفقودة

</div>

<div dir="ltr" align="left">
<pre dir="ltr"><code class="language-python">print("الصفوف قبل التنظيف:", len(combined_sales_df))
print("الصفوف بعد التنظيف:", len(cleaned_sales_df))

missing_values_per_columns(cleaned_sales_df)
print(cleaned_sales_df.dtypes)</code></pre>
</div>

<div dir="rtl" align="right">

**النتيجة:** بقي عدد الصفوف **4,862**، ولم تبقَ قيم مفقودة في البيانات بعد المعالجة.

### 4.2 فحص الكميات والأسعار والخصومات وأرقام البنود

</div>

<div dir="ltr" align="left">
<pre dir="ltr"><code class="language-python">validation_report = pd.DataFrame([{
    "كميات صفرية أو سالبة": (cleaned_sales_df["الكمية"] &lt;= 0).sum(),
    "أسعار سالبة": (cleaned_sales_df["سعر بيع الوحدة"] &lt; 0).sum(),
    "خصومات سالبة": (cleaned_sales_df["الخصم %"] &lt; 0).sum(),
    "خصومات فوق 100%": (cleaned_sales_df["الخصم %"] &gt; 100).sum(),
    "أرقام بنود مكررة": cleaned_sales_df["رقم البند"].duplicated().sum()
}])

display(validation_report)</code></pre>
</div>

<div dir="rtl" align="right">

**النتيجة:** جميع الأعداد في التقرير كانت **صفرًا**؛ أي لم تظهر مخالفات للشروط التي اختبرتها.

### 4.3 التأكد من حساب صافي مبيعات المنتج

تحققتُ من أن المبلغ المسجّل في كل بند يتوافق مع الكمية وسعر البيع ونسبة الخصم:

**صافي مبيعات المنتج = الكمية × سعر بيع الوحدة × (1 − نسبة الخصم ÷ 100)**

</div>

<div dir="ltr" align="left">
<pre dir="ltr"><code class="language-python">expected_sales = (
    cleaned_sales_df["الكمية"]
    * cleaned_sales_df["سعر بيع الوحدة"]
    * (1 - cleaned_sales_df["الخصم %"] / 100)
).round(2)

sales_difference = (
    expected_sales - cleaned_sales_df["صافي مبيعات المنتج"]
).abs().round(2)

# فحص الفروق التي تتجاوز هللة واحدة
incorrect_sales_df = cleaned_sales_df.loc[
    (sales_difference &gt; 0.01)
    | expected_sales.isna()
    | cleaned_sales_df["صافي مبيعات المنتج"].isna()
]

print("السجلات بفارق أكبر من هللة:", len(incorrect_sales_df))</code></pre>
</div>

<div dir="rtl" align="right">

**مخرجات التشغيل:**

<pre dir="rtl" align="right"><code>السجلات بفارق أكبر من هللة: 0</code></pre>

</div>

<div dir="rtl" align="right">

**النتيجة:** لم يظهر أي سجل بفارق يزيد على **0.01 ريال**. ووجدتُ **17 سجلًا** بفارق هللة واحدة عند المقارنة بعد التقريب؛ لذلك لا يصح القول إن كل المبالغ متطابقة حرفيًا. لم أغيّر المبالغ المسجلة بسبب فرق التقريب.

---

## 05 — Load | حفظ النتيجة في Excel

بعد تنظيف البيانات وفحصها، صدّرتُ جدول المبيعات الموحّد إلى ملف Excel:

</div>

<div dir="ltr" align="left">
<pre dir="ltr"><code class="language-python">cleaned_sales_df.to_excel(
    "data ETL.xlsx",
    index=False
)</code></pre>
</div>

<div dir="rtl" align="right">

الخيار `index=False` يمنع إضافة عمود إضافي لأرقام الصفوف داخل ملف Excel.

## ملخص النتائج

| ما تم تنفيذه | النتيجة |
|---|---|
| قراءة ملفات المبيعات | 20 ملف HTML |
| تجميع البيانات | 4,862 سجلًا في 14 عمودًا |
| اكتشاف اختلافات التاريخ | 3 سجلات، ثم توحيد تنسيقها |
| اكتشاف اختلافات السعر | 4 سجلات، ثم إزالة رمز العملة وتحويل السعر إلى رقم |
| معالجة رموز العملاء | استرجاع 5 رموز من بنود الفواتير المطابقة |
| التحقق من القيم المفقودة | 0 بعد التنظيف |
| التحقق من القيم التجارية وأرقام البنود | لا مخالفات للقواعد التي تم اختبارها |
| مراجعة حسابات المبيعات | 17 فارق هللة، ولا يوجد فرق أكبر من هللة |
| حفظ النتيجة | ملف Excel موحّد |

**تاريخ تنفيذ المشروع: 12–13 نوفمبر 2026.**  
**فترة البيانات المستخدمة: يناير 2025 – يونيو 2026.**

</div>
