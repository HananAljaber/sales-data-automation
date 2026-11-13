# أتمتة تجميع وتنظيف بيانات مبيعات متجر ملابس

**Sales Data Automation | HTML → Python & Pandas → Excel**

يستعرض هذا المشروع طريقة التعامل مع تقارير المبيعات الموزّعة على عدة ملفات HTML: استخراج الجداول وتجميعها باستخدام Python وPandas، ثم اكتشاف مشكلات جودة البيانات، ومعالجتها، والتحقق من النتائج، وتصدير البيانات إلى ملف Excel موحّد.

## نظرة عامة

| العنصر | التفاصيل |
|---|---|
| مجال البيانات | مبيعات متجر ملابس |
| المصدر | 20 ملف HTML |
| الفترة | يناير 2025 – يونيو 2026 (18 شهرًا) |
| السجلات بعد التجميع | **4,862** |
| عدد الأعمدة | **14** |
| الأدوات | **Python، Pandas، Jupyter Notebook، HTML، Excel** |
| المخرج | ملف Excel موحّد باسم `data ETL.xlsx` |

## المشكلة والهدف

كانت بيانات المبيعات موزّعة على **20 ملف HTML**. الهدف هو استبدال جمع الصفوف يدويًا بمسار برمجي موحّد، مع التأكد من فهم الاختلاف بين **البيانات المفقودة** و**القيم الموجودة ولكن المكتوبة بتنسيق غير مناسب**.

## الأعمدة المستخدمة

| العمود الأصلي | المعنى |
|---|---|
| التاريخ | تاريخ عملية البيع |
| رقم الفاتورة | معرّف الفاتورة، وقد يتكرر إذا احتوت الفاتورة على عدة بنود |
| رقم البند | معرّف بند المبيعات |
| الفرع | الفرع الذي نُفذت فيه العملية |
| رمز العميل | معرّف العميل |
| رمز المنتج | معرّف المنتج |
| اسم المنتج | اسم المنتج |
| الفئة | تصنيف المنتج |
| الكمية | عدد الوحدات |
| سعر الوحدة | سعر بيع الوحدة قبل الخصم |
| الخصم % | نسبة الخصم المطبقة على البند |
| إجمالي السطر | القيمة المسجلة للبند بعد الخصم |
| تكلفة الوحدة | تكلفة القطعة على المتجر |
| حالة الطلب | حالة الطلب |

بعد التنظيف استخدمتُ الاسمين الأوضح: **`سعر بيع الوحدة`** بدل `سعر الوحدة`، و**`صافي مبيعات المنتج`** بدل `إجمالي السطر` (المقصود قيمة بند البيع بعد الخصم، وليس بالضرورة صافي مبيعات المتجر بعد المرتجعات).

---

## 01 — Extract | استخراج ودمج الملفات

أقرأ جميع ملفات HTML وأستخرج الجدول الأول من كل ملف، ثم أجمع الجداول رأسيًا في DataFrame واحد. استخدمت أسماء متغيرات توضّح المقصود بكل خطوة.

```python
import os
import pandas as pd

# غيّري المسار إذا كان مجلد HTML_Files في موقع آخر
html_folder_path = "HTML_Files"

html_file_names = sorted([
    file_name for file_name in os.listdir(html_folder_path)
    if file_name.lower().endswith(".html")
])

sales_dataframes = []

for file_name in html_file_names:
    file_path = os.path.join(html_folder_path, file_name)
    sales_df = pd.read_html(file_path, encoding="utf-8")[0]
    sales_dataframes.append(sales_df)

combined_sales_df = pd.concat(sales_dataframes, ignore_index=True)

print("عدد ملفات HTML:", len(html_file_names))
print("حجم البيانات المجمّعة:", combined_sales_df.shape)
```

**النتيجة من التطبيق:** قراءة 20 ملفًا وتجميع **4,862 صفًا × 14 عمودًا**. استخدام `ignore_index=True` يُعيد ترقيم صفوف الجدول المجمّع دون جعل أسماء الأعمدة تتكرر كصفوف.

---

## 02 — Data Profiling | فحص البيانات قبل التعديل

في هذه المرحلة أفحص البيانات كما وردت من المصدر، **دون تعديل قيمها**. ولذلك تُحفظ نتائج اختبارات التحويل في متغيرات مستقلة.

### 2.1 حجم البيانات والأنواع والقيم المفقودة

```python
print("Dataset Shape:", combined_sales_df.shape)
print(combined_sales_df.columns.tolist())
print(combined_sales_df.dtypes)

# تقرير القيم المفقودة: العدد والنسبة لكل عمود
def missing_values_per_columns(data):
    missing_count = data.isna().sum()
    missing_percent = (data.isna().mean() * 100).round(3)

    report = pd.DataFrame({
        "Missing Count": missing_count,
        "Missing %": missing_percent
    })
    display(report)

missing_values_per_columns(combined_sales_df)

# فحص الصفوف المتطابقة بالكامل
print("Duplicate Rows:", combined_sales_df.duplicated().sum())
```

**نتيجة الفحص الأولي:** خمس قيم مفقودة في `رمز العميل`، ولا توجد صفوف متطابقة بالكامل. في هذه النقطة لم يظهر سعر أو تاريخ مفقود لأن القيم غير المتوافقة كانت **موجودة كنصوص**.

### 2.2 كيف اكتشفتُ التواريخ المكتوبة بتنسيق مختلف؟

التنسيق المعتاد في الملفات هو `DD/MM/YYYY`. اختبرت إمكانية قراءة كل تاريخ بهذا التنسيق باستخدام `pd.to_datetime()`، ثم عرضت القيم التي فشل تفسيرها.

```python
# اختبار تحويل مؤقت؛ لا يغيّر عمود التاريخ الأصلي
date_check = pd.to_datetime(
    combined_sales_df["التاريخ"],
    format="%d/%m/%Y",
    errors="coerce"
)

invalid_dates_df = combined_sales_df.loc[
    date_check.isna() & combined_sales_df["التاريخ"].notna(),
    ["رقم البند", "التاريخ"]
]

print("التواريخ مختلفة التنسيق:", len(invalid_dates_df))
display(invalid_dates_df)
```

**النتيجة: 3 سجلات** مختلفة التنسيق:

| فهرس الصف | التاريخ الأصلي |
|---:|---|
| 972 | `2025-05-03` |
| 2438 | `2025-11-03` |
| 3920 | `2026-04-02` |

**التفسير:** التواريخ أعلاه مكتوبة بشكل `YYYY-MM-DD` بدل `DD/MM/YYYY`. فشل اختبار التنسيق لا يعني أن هذه التواريخ غير صحيحة في ذاتها.

### 2.3 كيف اكتشفتُ الأسعار غير القابلة للتحويل المباشر؟

ظهر نوع `سعر الوحدة` على أنه `object` بدل نوع رقمي. اختبرت التحويل إلى أرقام، ثم عرضت القيم التي فشل تحويلها.

```python
# اختبار تحويل مؤقت؛ لا يغيّر أسعار المصدر
price_check = pd.to_numeric(
    combined_sales_df["سعر الوحدة"],
    errors="coerce"
)

invalid_prices_df = combined_sales_df.loc[
    price_check.isna() & combined_sales_df["سعر الوحدة"].notna(),
    ["رقم البند", "سعر الوحدة"]
]

print("الأسعار مختلفة التنسيق:", len(invalid_prices_df))
display(invalid_prices_df)
```

**النتيجة: 4 أسعار** تحتوي على رمز العملة ضمن النص:

| فهرس الصف | السعر الأصلي |
|---:|---|
| 714 | `ر.س 115.00` |
| 1927 | `ر.س 140.00` |
| 3071 | `ر.س 112.20` |
| 4558 | `ر.س 127.50` |

### 2.4 فحص نتائج التحويل المؤقت دون المساس بالمصدر

جمعتُ نتائج الاختبارين في نسخة منفصلة، ثم استدعيت الدالة نفسها التي تفحص القيم المفقودة. بذلك صار الفرق واضحًا بين القيم المفقودة في الأصل والقيم التي **فشلت أثناء محاولة تحويلها**.

```python
profiling_sales_df = combined_sales_df.copy()
profiling_sales_df["التاريخ"] = date_check
profiling_sales_df["سعر الوحدة"] = price_check

missing_values_per_columns(profiling_sales_df)
```

| الحقل | القيم المفقودة في المصدر | القيم الناتجة عن فشل اختبار التنسيق |
|---|---:|---:|
| التاريخ | 0 | 3 |
| سعر الوحدة | 0 | 4 |
| رمز العميل | 5 | — |

> `errors="coerce"` يحوّل القيم التي لا يمكن تفسيرها إلى `NaT` للتاريخ أو `NaN` للسعر **في نتيجة الاختبار**؛ لا يعني أن هذه القيم كانت فارغة في ملفات HTML.

---

## 03 — Transform | تنظيف البيانات وتحويلها

أُجريت المعالجة على نسخة قابلة للتعديل من البيانات المجمّعة. الأكواد التالية مرتّبة لعرض نفس خطوات المعالجة بصورة واضحة.

```python
cleaned_sales_df = combined_sales_df.copy()
```

### 3.1 توحيد التواريخ

```python
cleaned_sales_df["التاريخ"] = pd.to_datetime(
    cleaned_sales_df["التاريخ"],
    format="mixed",
    dayfirst=True,
    errors="coerce"
)
```

استخدمت `format="mixed"` لقراءة أكثر من تنسيق، مع تفضيل اليوم أولًا للتواريخ التي تبدأ باليوم. تحققتُ من النتائج لأن `dayfirst=True` ليس ضمانًا صارمًا لتفسير كل تاريخ ملتبس.

### 3.2 إزالة رمز العملة وتحويل الأسعار إلى أرقام

```python
cleaned_sales_df["سعر الوحدة"] = pd.to_numeric(
    cleaned_sales_df["سعر الوحدة"]
    .astype(str)
    .str.replace("ر.س", "", regex=False)
    .str.strip(),
    errors="coerce"
)
```

**لماذا هذه الأوامر؟** `astype(str)` يسمح بتطبيق عمليات النص على الأرقام والنصوص معًا، و`str.replace()` يزيل `ر.س` من **داخل القيمة**، و`str.strip()` يزيل المسافات، ثم `pd.to_numeric()` يحوّل الناتج إلى رقم. هذا يتجنب فقدان الأسعار الرقمية عند تطبيق `.str.replace()` مباشرة على عمود مختلط الأنواع.

### 3.3 استرجاع رموز العملاء المفقودة من بنود الفواتير

بحثتُ عن بند آخر في **الفاتورة نفسها** يحمل رمز عميل معلومًا، ثم عوّضتُ الرمز فقط عندما وجدت **رمزًا واحدًا غير متعارض**. لا يُخترع رمز عند غياب دليل كافٍ.

```python
missing_customer_rows = cleaned_sales_df.index[
    cleaned_sales_df["رمز العميل"].isna()
]

for row_index in missing_customer_rows:
    invoice_number = cleaned_sales_df.loc[row_index, "رقم الفاتورة"]

    customer_ids = cleaned_sales_df.loc[
        (cleaned_sales_df["رقم الفاتورة"] == invoice_number)
        & cleaned_sales_df["رمز العميل"].notna(),
        "رمز العميل"
    ].unique()

    if len(customer_ids) == 1:
        cleaned_sales_df.loc[row_index, "رمز العميل"] = customer_ids[0]

print("رموز عملاء مفقودة بعد المعالجة:",
      cleaned_sales_df["رمز العميل"].isna().sum())
```

**النتيجة المسجلة:** استرجاع الرموز الخمسة من بنود فواتيرها، وعدم بقاء رموز عملاء مفقودة بعد المعالجة.

### 3.4 إعادة تسمية عمودين لتوضيح معناهما

```python
cleaned_sales_df = cleaned_sales_df.rename(columns={
    "سعر الوحدة": "سعر بيع الوحدة",
    "إجمالي السطر": "صافي مبيعات المنتج"
})
```

---

## 04 — Data Validation | التحقق من جودة البيانات

### 4.1 التحقق من القيم المفقودة وعدد الصفوف

```python
print("الصفوف الأصلية:", len(combined_sales_df))
print("الصفوف بعد التنظيف:", len(cleaned_sales_df))
missing_values_per_columns(cleaned_sales_df)
```

### 4.2 اختبار القيم التجارية غير المنطقية والتكرار

```python
validation_report = pd.DataFrame([{
    "Invalid Quantity": (cleaned_sales_df["الكمية"] <= 0).sum(),
    "Negative Price": (cleaned_sales_df["سعر بيع الوحدة"] < 0).sum(),
    "Negative Discount": (cleaned_sales_df["الخصم %"] < 0).sum(),
    "Discount Above 100%": (cleaned_sales_df["الخصم %"] > 100).sum(),
    "Duplicate Item IDs": cleaned_sales_df["رقم البند"].duplicated().sum()
}])

display(validation_report)
```

**النتائج المسجلة:** صفر كميات غير موجبة، وصفر أسعار سالبة، وصفر خصومات سالبة أو تتجاوز 100%، وصفر أرقام بنود مكررة. هذه النتائج تخص القواعد التي جرى اختبارها فقط.

### 4.3 التحقق من حساب «صافي مبيعات المنتج»

المعادلة المستخدمة لبند المبيعات قبل أي مرتجعات أو رسوم إضافية:

**صافي قيمة البند = الكمية × سعر بيع الوحدة × (1 − الخصم % ÷ 100)**

```python
expected_sales = (
    cleaned_sales_df["الكمية"]
    * cleaned_sales_df["سعر بيع الوحدة"]
    * (1 - cleaned_sales_df["الخصم %"] / 100)
).round(2)

sales_difference = (
    expected_sales - cleaned_sales_df["صافي مبيعات المنتج"]
).abs().round(2)

incorrect_sales_df = cleaned_sales_df.loc[
    (sales_difference > 0.01)
    | expected_sales.isna()
    | cleaned_sales_df["صافي مبيعات المنتج"].isna()
]

print("السجلات بفارق يتجاوز هللة واحدة:", len(incorrect_sales_df))
```

**النتيجة المسجلة:** لم تظهر سجلات بفروق **أكبر من هللة واحدة** بعد تقريب الفرق إلى منزلتين. هذا **لا يعني التطابق التام في كل سجل**؛ ظهرت فروق هللة مرتبطة بالتقريب، ولم أعدّل المبالغ المسجلة بسببها.

---

## 05 — Load | تصدير النتيجة إلى Excel

بعد إتمام المعالجة والفحوصات، جرى تصدير الجدول إلى Excel دون إضافة عمود فهرس غير مطلوب:

```python
cleaned_sales_df.to_excel("data ETL.xlsx", index=False)
```

## ملخص ما تم تطبيقه

- استخراج جداول المبيعات من **20 ملف HTML** ودمج **4,862 سجلًا**.
- اكتشاف اختلاف **3 تواريخ** و**4 أسعار** بالتجربة دون إتلاف البيانات الأصلية.
- توحيد تنسيقات التواريخ والأسعار واسترجاع **5 رموز عملاء** من فواتيرها.
- إعادة تسمية عمودين لتوضيح المعنى المحاسبي.
- اختبار التكرارات، وقواعد العمل، وحساب صافي قيمة بند البيع، ومراجعة فروق التقريب.
- تصدير نتيجة المعالجة إلى Excel.
