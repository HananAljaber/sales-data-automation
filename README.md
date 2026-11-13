<a id="arabic-version"></a>

<div dir="rtl" align="right">

# أتمتة معالجة بيانات مبيعات متجر ملابس

**اللغة:** العربية | [English ↓](#english-version)

**Sales Data Automation | Python · Pandas · HTML · Excel**

| المعلومة | التفاصيل |
|---|---|
| **تاريخ تنفيذ المشروع** | **12–13 نوفمبر 2026** |
| **الفترة التي تغطيها المبيعات** | من يناير 2025 إلى يونيو 2026 (18 شهرًا) |
| **ملفات المصدر** | 20 ملف HTML |
| **البيانات بعد التجميع** | 4,862 صفًا و14 عمودًا |
| **الأدوات المستخدمة** | Python، Pandas، Jupyter Notebook، Excel |

## فكرة المشروع

كانت بيانات مبيعات متجر ملابس موزّعة على 20 ملف HTML. استخدمتُ Python ومكتبة Pandas لقراءة الملفات وتجميع بياناتها في جدول واحد، ثم فحصتُ البيانات وعالجتُ المشكلات التي ظهرت قبل حفظ النتيجة في Excel.

خلال الفحص اكتشفتُ:

<ul dir="rtl" align="right">
<li><strong>3 تواريخ</strong> مكتوبة بتنسيق مختلف عن بقية التواريخ.</li>
<li><strong>4 أسعار</strong> تحتوي على رمز العملة «ر.س» داخل الخلية.</li>
<li><strong>5 رموز عملاء</strong> مفقودة، وتم العثور عليها في بنود أخرى من الفواتير نفسها.</li>
</ul>

## 01 — Extract | استخراج البيانات وتجميعها

قرأتُ ملفات HTML باستخدام حلقة `for`، ووضعتُ الجداول المستخرجة في قائمة، ثم جمعتُها باستخدام `pd.concat()`.

**الكود الذي استخدمتُه:**

</div>

<div dir="ltr" align="left">

```python
# Extract

html_files =os.listdir("C:\\Users\\hanan\\Downloads\\HTML_Files")
sales_dataframes=[]
for file_name in html_files:
    if file_name.lower().endswith(".html"):
        sales_df=pd.read_html(file_name)[0]

        sales_dataframes.append(sales_df)

# load data
combined_sales_df =pd.concat(sales_dataframes,ignore_index=True)
```

</div>

<div dir="rtl" align="right">

**النتيجة:** بعد تشغيل فحص حجم البيانات ظهر `Dataset Shape: (4862, 14)`؛ أي 4,862 سجلًا و14 عمودًا.

**ملاحظة تشغيل:** مسار `C:\Users\hanan\Downloads\HTML_Files` ومسار قراءة الملف باسمه وحده هما كما استُخدما في دفتر العمل المحلي. لم أغيّرهما هنا إلى كود آخر.

## 02 — Data Profiling | فحص جودة البيانات

### 2.1 حجم البيانات وأسماء الأعمدة وأنواعها

فحصتُ عدد الصفوف والأعمدة وأسماءها وأنواعها قبل تعديل البيانات.

</div>

<div dir="ltr" align="left">

```python
# Number of rows and columns
print("Dataset Shape:", combined_sales_df.shape)
# Column names
print("\nColumn Names:")
print(combined_sales_df.columns.tolist())
# Data types
print("\nData Types:")
print(combined_sales_df.dtypes)
```

</div>

<div dir="rtl" align="right">

**مخرجات التشغيل:** `Dataset Shape: (4862, 14)`، وظهر عمود `التاريخ` كنص (`str`) وعمود `سعر الوحدة` بنوع (`object`).

### 2.2 فحص القيم المفقودة

كتبتُ دالة تحسب عدد القيم المفقودة ونسبتها في كل عمود، ثم طبّقتها على البيانات المجمّعة.

</div>

<div dir="ltr" align="left">

```python
# Missing values per column
def miising_values_per_columns(data):
      missing_values = data.isna().sum()
      # Missing values percentage
      missing_percentage =( data.isna().mean()*100).round(3)
      # Data quality summary
      missing_values_report = pd.DataFrame({
      "Missing Count":missing_values,
            "Missing %":missing_percentage
      })
      display(missing_values_report)
      
miising_values_per_columns(combined_sales_df)
```

</div>

<div dir="rtl" align="right">

**النتيجة قبل التنظيف:** 5 قيم مفقودة في عمود `رمز العميل` (0.103%)، ولم تظهر قيم مفقودة في التواريخ والأسعار قبل اختبار تنسيقهما.

### 2.3 فحص الصفوف المكررة

استخدمتُ `duplicated()` للتحقق من وجود صفوف مكررة في البيانات.

</div>

<div dir="ltr" align="left">

```python
# Count fully duplicated rows
duplicate_rows_count = combined_sales_df.duplicated().sum()
print("Duplicate Rows:", duplicate_rows_count)
# Display duplicated records for review
duplicate_records_df = combined_sales_df[
    combined_sales_df.duplicated(keep=False)
]

duplicate_records_df
```

</div>

<div dir="rtl" align="right">

**مخرجات التشغيل:** `Duplicate Rows: 0`، وظهر جدول فارغ عند عرض السجلات المكررة.

### 2.4 اكتشاف التواريخ المختلفة

جرّبتُ تحويل التواريخ وفق الصيغة المتوقعة **يوم/شهر/سنة**. القيم التي لم يستطع Pandas تفسيرها ظهرت في نتيجة الفحص بوصفها قيمًا غير صالحة لهذا التنسيق، ثم عرضتُ سجلاتها الأصلية.

</div>

<div dir="ltr" align="left">

```python
# 1. Detect rows with inconsistent date formats
date_check = pd.to_datetime(
    combined_sales_df["التاريخ"],
    format="%d/%m/%Y",
    errors="coerce"
)
invalid_dates_df = combined_sales_df.loc[
    date_check.isna(),
    ["رقم البند", "التاريخ"]
]
print("عدد التواريخ مختلفة التنسيق:", len(invalid_dates_df))
invalid_dates_df
```

</div>

<div dir="rtl" align="right">

**مخرجات التشغيل: عدد التواريخ مختلفة التنسيق: 3**

| رقم الصف | رقم البند | التاريخ الأصلي |
|---|---|---|
| 972 | س000986 | `2025-05-03` |
| 2438 | س002486 | `2025-11-03` |
| 3920 | س004137 | `2026-04-02` |

هذه التواريخ تبدأ بالسنة بدل اليوم. **لم أعدّل البيانات الأصلية أثناء هذا الفحص.**

### 2.5 اكتشاف الأسعار المختلفة

اختبرتُ إمكانية تحويل `سعر الوحدة` إلى رقم باستخدام `pd.to_numeric()`. ثم عرضتُ الأسعار التي فشل تحويلها مباشرة.

</div>

<div dir="ltr" align="left">

```python
# 2. Detect rows with non-convertible prices

# Try converting unit price to numeric
price_check = pd.to_numeric(
    combined_sales_df["سعر الوحدة"],
    errors="coerce"
)

# Extract rows whose price failed conversion
invalid_prices_df = combined_sales_df.loc[
    price_check.isna(),
    ["رقم البند", "سعر الوحدة"]
]

print("عدد الأسعار مختلفة التنسيق:", len(invalid_prices_df))
display(invalid_prices_df)
```

</div>

<div dir="rtl" align="right">

**مخرجات التشغيل: عدد الأسعار مختلفة التنسيق: 4**

| رقم الصف | رقم البند | السعر كما ظهر في الملف |
|---|---|---|
| 714 | س000818 | `ر.س 115.00` |
| 1927 | س002157 | `ر.س 140.00` |
| 3071 | س003185 | `ر.س 112.20` |
| 4558 | س004828 | `ر.س 127.50` |

كانت المشكلة وجود **رمز العملة مع السعر**؛ لذلك لم تُقرأ هذه القيم كأرقام مباشرة.

### 2.6 فحص نتيجة الاختبارين دون تغيير الجدول الأساسي

أنشأتُ نسخة مخصصة للفحص، وأضفتُ إليها نتيجتي اختبار التاريخ والسعر، ثم استدعيتُ دالة فحص القيم المفقودة نفسها.

</div>

<div dir="ltr" align="left">

```python
profiling_sales_df = combined_sales_df.copy()

# Add date-test results
profiling_sales_df["التاريخ"] = date_check

# Add price-test results
profiling_sales_df["سعر الوحدة"] = price_check
      
miising_values_per_columns(profiling_sales_df)
```

</div>

<div dir="rtl" align="right">

**النتيجة في نسخة الفحص:**

| العمود | العدد |
|---|---:|
| التاريخ | 3 |
| سعر الوحدة | 4 |
| رمز العميل | 5 |

هذه الأرقام تجمع بين القيم المفقودة أصلًا، والقيم التي لم تتوافق مع التنسيق المطلوب في الاختبار.

## 03 — Transform | تنظيف البيانات وتحويلها

### 3.1 توحيد سعر الوحدة والتواريخ

استخدمتُ `str.replace()` لإزالة رمز العملة، و`pd.to_numeric()` لتحويل الأسعار إلى أرقام. واستخدمتُ `pd.to_datetime()` لمعالجة صيغ التواريخ المختلفة.

**الكود كما كتبته في الدفتر:**

</div>

<div dir="ltr" align="left">

```python
combined_sales_df["سعر الوحدة"] = pd.to_numeric(combined_sales_df["سعر الوحدة"]    .astype(str)
    .str.replace("ر.س", "", regex=False)
    .str.strip(),errors="coerce")

combined_sales_df["التاريخ"] = pd.to_datetime(combined_sales_df["التاريخ"], 
    errors="coerce",
     format="mixed",
    dayfirst=True,)
combined_sales_df.dtypes
```

</div>

<div dir="rtl" align="right">

**مخرجات التشغيل:** ظهر نوع `التاريخ` على أنه `datetime64[us]`، ونوع `سعر الوحدة` على أنه `float64`.

### 3.2 استرجاع رموز العملاء المفقودة

بحثتُ عن الفواتير التي ينقص أحد بنودها رمز العميل، ثم عن بند آخر في الفاتورة نفسها يحمل الرمز. استخدمتُ ذلك الرمز لتعبئة الخانة الناقصة في البيانات.

**الكود الذي نفذتُه:**

</div>

<div dir="ltr" align="left">

```python
mis =combined_sales_df[combined_sales_df["رمز العميل"].isna()]


# Invoices missing a customer ID
missing_invoices = combined_sales_df.loc[
    combined_sales_df["رمز العميل"].isna(),
    "رقم الفاتورة"
].unique()

# # Display all line items of these invoices
invoice_check_df = combined_sales_df.loc[
    combined_sales_df["رقم الفاتورة"].isin(missing_invoices),
    ["رقم الفاتورة", "رقم البند", "رمز العميل"]
]



for i in mis.index:
    invoice_number = mis.loc[i,"رقم الفاتورة"]
  
    customer_ids = combined_sales_df.loc[
        (combined_sales_df["رقم الفاتورة"] == invoice_number)
        & (combined_sales_df["رمز العميل"].notna()),
        "رمز العميل"
    ].unique()
 
    combined_sales_df.loc[i,"رمز العميل"] =customer_ids[0]
combined_sales_df[combined_sales_df["رمز العميل"].isna()]
```

</div>

<div dir="rtl" align="right">

**مخرجات التشغيل:** لم تبقَ سجلات ذات رمز عميل مفقود في الجدول الناتج.

### 3.3 إعادة تسمية الأعمدة

عدّلتُ اسم `إجمالي السطر` إلى `صافي مبيعات المنتج`، واسم `سعر الوحدة` إلى `سعر بيع الوحدة` ليكون معنى العمودين أوضح.

</div>

<div dir="ltr" align="left">

```python
combined_sales_df = combined_sales_df.rename(
    columns={
        "إجمالي السطر": "صافي مبيعات المنتج",
        "سعر الوحدة": "سعر بيع الوحدة"
    }
)
combined_sales_df
```

</div>

<div dir="rtl" align="right">

### 3.4 إعادة فحص القيم المفقودة

استخدمتُ دالة الفحص التي كتبتُها سابقًا مرة أخرى. **يُراعى تنفيذ هذا الفحص بعد معالجة رموز العملاء**؛ لأن النتيجة المحفوظة في Jupyter تعكس وقت تشغيل الخلية، وقد تختلف عند تشغيل الخلايا بترتيب آخر.

</div>

<div dir="ltr" align="left">

```python
miising_values_per_columns(combined_sales_df)
```

</div>

<div dir="rtl" align="right">

**النتيجة المحفوظة في الدفتر:** صفر قيم مفقودة في جميع الأعمدة بعد إتمام المعالجات.

## 04 — Data Validation | التحقق من صحة النتائج

### 4.1 مقارنة صافي المبيعات المحسوب بالمسجّل

حسبتُ صافي مبيعات المنتج من الكمية وسعر البيع ونسبة الخصم، وقارنتُه بالقيمة الموجودة في البيانات. كان فحصي يبحث عن الفروق **الأكبر من 0.01 ريال** أو القيم التي فشل حسابها.

</div>

<div dir="ltr" align="left">

```python
# Calculate expected net sales
expected_sales = (
    combined_sales_df["الكمية"]
    * combined_sales_df["سعر بيع الوحدة"]
    * (1 - combined_sales_df["الخصم %"] / 100)
).round(2)

# Compare expected and recorded sales
sales_difference = (
    expected_sales
    - combined_sales_df["صافي مبيعات المنتج"]
).abs().round(2)
# Find records with calculation errors
incorrect_sales_df = combined_sales_df.loc[
    (sales_difference > 0.01)
    | expected_sales.isna()
    | combined_sales_df["صافي مبيعات المنتج"].isna()
]

print("عدد السجلات غير المتطابقة:", len(incorrect_sales_df))
incorrect_sales_df[
        [
            "رقم البند",
            "الكمية",
            "سعر بيع الوحدة",
            "الخصم %",
            "صافي مبيعات المنتج"
        ]
]
```

</div>

<div dir="rtl" align="right">

**مخرجات التشغيل:** `عدد السجلات غير المتطابقة: 0` وفق الشرط الذي استخدمتُه في الكود. لا يعني هذا بالضرورة عدم وجود أي فرق أصغر من ذلك.

### 4.2 فحص قواعد الكميات والأسعار والخصومات وأرقام البنود

تحققتُ من الكميات غير الموجبة، والأسعار السالبة، والخصومات السالبة، والخصومات التي تتجاوز 100%، وأرقام البنود المكررة.

**الكود كما نفذتُه، دون تغيير شروط الفحص:**

</div>

<div dir="ltr" align="left">

```python
# 1. Zero or negative quantities

invalid_quantity =(combined_sales_df["الكمية"] <= 0).sum()
# 2. Zero or negative prices

invalid_price = (combined_sales_df["سعر بيع الوحدة"] <0).sum()
# 3. Negative discounts

negative_discount=     (combined_sales_df["الخصم %"] <0).sum()
# 4. Discounts above 100%

high_discount =     (combined_sales_df["الخصم %"] >100).sum()

# Count of duplicate line-item IDs
duplicate_item_count = (
    combined_sales_df["رقم البند"].duplicated().sum()
)
duplicate_item_count
# Create Validation Report
validation_report = pd.DataFrame([{
    "Invalid Quantity": invalid_quantity,
    "Invalid Price": invalid_price,
    "Negative Discount": negative_discount,
    "Discount Above 100%": high_discount,
    "duplicate_item_count":duplicate_item_count
}])
validation_report
```

</div>

<div dir="rtl" align="right">

**مخرجات التشغيل:** كانت جميع القيم الخمس **0**.

**ملاحظة:** يفحص شرط السعر (`< 0`) الأسعار السالبة فقط، ولا يشمل الأسعار الصفرية.

### 4.3 عرض البنود ذات الأرقام المكررة

عرضتُ جميع السجلات التي تشترك في رقم بند، إن وجدت.

</div>

<div dir="ltr" align="left">

```python
# Display all records sharing the same line-item ID
duplicate_items_df = combined_sales_df.loc[
    combined_sales_df["رقم البند"].duplicated(keep=False),
    ["رقم الفاتورة", "رقم البند", "اسم المنتج", "الفرع"]
]

duplicate_items_df
```

</div>

<div dir="rtl" align="right">

**مخرجات التشغيل:** جدول فارغ؛ لم تظهر أرقام بنود مكررة في الفحص.

## 05 — Load | حفظ البيانات في Excel

بعد معالجة البيانات وإجراء الفحوصات، حفظتُ الناتج في ملف Excel.

</div>

<div dir="ltr" align="left">

```python
combined_sales_df.to_excel("data ETL.xlsx", index=False)
```

</div>

<div dir="rtl" align="right">

**اسم الملف الناتج:** `data ETL.xlsx`.

## ملخص ما نفذتُه

| المهمة | النتيجة |
|---|---|
| استخراج وتجميع ملفات HTML | 20 ملفًا |
| عدد الصفوف والأعمدة | 4,862 × 14 |
| الصفوف المكررة | 0 |
| تواريخ بتنسيق مختلف | 3 |
| أسعار تحتوي على رمز العملة | 4 |
| رموز عملاء مفقودة | 5 قبل المعالجة، و0 بعدها |
| اختلافات الحساب أكبر من 0.01 ريال وفق الكود | 0 |
| مخالفات قواعد الفحص المكتوبة | 0 |
| الحفظ النهائي | `data ETL.xlsx` |

</div>

---

<a id="english-version"></a>

# Clothing Store Sales Data Automation

**Language:** [العربية ↑](#arabic-version) | English

**Sales Data Automation | Python · Pandas · HTML · Excel**

| Project detail | Information |
|---|---|
| **Project date** | **November 12–13, 2026** |
| **Sales data period** | January 2025 – June 2026 (18 months) |
| **Source files** | 20 HTML files |
| **Consolidated dataset** | 4,862 rows and 14 columns |
| **Tools** | Python, Pandas, Jupyter Notebook, Excel |

## Project Overview

The clothing store's sales records were spread across 20 HTML files. I used Python and Pandas to extract the tables, combine them into a single dataset, identify data-quality issues, and prepare the final Excel output.

During the checks, I found:

- **3 dates** formatted differently from the other dates.
- **4 prices** that included the currency symbol `ر.س` in the cell.
- **5 missing customer IDs**, which I recovered from other line items belonging to the same invoices.

## 01 — Extract | Reading and Combining the Files

I read the HTML files with a `for` loop, stored each extracted table in a list, and merged them using `pd.concat()`.

**Python code:**

<div dir="ltr" align="left">

```python
# Extract

html_files =os.listdir("C:\\Users\\hanan\\Downloads\\HTML_Files")
sales_dataframes=[]
for file_name in html_files:
    if file_name.lower().endswith(".html"):
        sales_df=pd.read_html(file_name)[0]

        sales_dataframes.append(sales_df)

# load data
combined_sales_df =pd.concat(sales_dataframes,ignore_index=True)
```

</div>

**Result:** The dataset shape was `Dataset Shape: (4862, 14)`, corresponding to 4,862 sales rows and 14 columns.

**Note:** The local folder path must point to the HTML files on the machine where the notebook is run.

## 02 — Data Profiling | Reviewing Data Quality

### 2.1 Dataset shape, column names, and data types

I checked the number of rows and columns, the column names, and their data types before applying any transformations.

<div dir="ltr" align="left">

```python
# Number of rows and columns
print("Dataset Shape:", combined_sales_df.shape)
# Column names
print("\nColumn Names:")
print(combined_sales_df.columns.tolist())
# Data types
print("\nData Types:")
print(combined_sales_df.dtypes)
```

</div>

**Output:** `Dataset Shape: (4862, 14)`. The `التاريخ` (date) column was read as `str`, while `سعر الوحدة` (unit price) was of type `object`.

### 2.2 Missing-value assessment

I wrote a function to count missing values and calculate their percentage in each column, then ran it on the combined dataset.

<div dir="ltr" align="left">

```python
# Missing values per column
def miising_values_per_columns(data):
      missing_values = data.isna().sum()
      # Missing values percentage
      missing_percentage =( data.isna().mean()*100).round(3)
      # Data quality summary
      missing_values_report = pd.DataFrame({
      "Missing Count":missing_values,
            "Missing %":missing_percentage
      })
      display(missing_values_report)
      
miising_values_per_columns(combined_sales_df)
```

</div>

**Before cleaning:** Five values were missing from `رمز العميل` (customer ID), representing 0.103% of the rows. No dates or prices were missing at this point, before checking whether their formats were valid.

### 2.3 Duplicate-row check

I used `duplicated()` to check for repeated rows.

<div dir="ltr" align="left">

```python
# Count fully duplicated rows
duplicate_rows_count = combined_sales_df.duplicated().sum()
print("Duplicate Rows:", duplicate_rows_count)
# Display duplicated records for review
duplicate_records_df = combined_sales_df[
    combined_sales_df.duplicated(keep=False)
]

duplicate_records_df
```

</div>

**Output:** `Duplicate Rows: 0`. The table of duplicate records was empty.

### 2.4 Detecting inconsistent date formats

I tried to interpret every date using the expected **day/month/year** format. Dates that did not match that format were flagged by the test, and I displayed the original affected records.

<div dir="ltr" align="left">

```python
# 1. Detect rows with inconsistent date formats
date_check = pd.to_datetime(
    combined_sales_df["التاريخ"],
    format="%d/%m/%Y",
    errors="coerce"
)
invalid_dates_df = combined_sales_df.loc[
    date_check.isna(),
    ["رقم البند", "التاريخ"]
]
print("عدد التواريخ مختلفة التنسيق:", len(invalid_dates_df))
invalid_dates_df
```

</div>

**Output: 3 dates with a different format**

| Row index | Line-item ID | Original date |
|---|---|---|
| 972 | س000986 | `2025-05-03` |
| 2438 | س002486 | `2025-11-03` |
| 3920 | س004137 | `2026-04-02` |

These dates start with the year rather than the day. The original data was **not modified** during this check.

### 2.5 Detecting prices with a different format

I tested whether `سعر الوحدة` (unit price) could be converted directly to a number with `pd.to_numeric()`, then displayed the values that failed.

<div dir="ltr" align="left">

```python
# 2. Detect rows with non-convertible prices

# Try converting unit price to numeric
price_check = pd.to_numeric(
    combined_sales_df["سعر الوحدة"],
    errors="coerce"
)

# Extract rows whose price failed conversion
invalid_prices_df = combined_sales_df.loc[
    price_check.isna(),
    ["رقم البند", "سعر الوحدة"]
]

print("عدد الأسعار مختلفة التنسيق:", len(invalid_prices_df))
display(invalid_prices_df)
```

</div>

**Output: 4 prices with a different format**

| Row index | Line-item ID | Original value |
|---|---|---|
| 714 | س000818 | `ر.س 115.00` |
| 1927 | س002157 | `ر.س 140.00` |
| 3071 | س003185 | `ر.س 112.20` |
| 4558 | س004828 | `ر.س 127.50` |

The prices included a **currency symbol in the same cell as the number**, preventing direct numerical conversion.

### 2.6 Reviewing conversion results without changing the source data

I created a temporary copy of the dataset, inserted the date and price test results into that copy, and reused the missing-value checking function.

<div dir="ltr" align="left">

```python
profiling_sales_df = combined_sales_df.copy()

# Add date-test results
profiling_sales_df["التاريخ"] = date_check

# Add price-test results
profiling_sales_df["سعر الوحدة"] = price_check
      
miising_values_per_columns(profiling_sales_df)
```

</div>

**Results of the temporary profiling dataset:**

| Column | Count |
|---|---:|
| Date (`التاريخ`) | 3 |
| Unit price (`سعر الوحدة`) | 4 |
| Customer ID (`رمز العميل`) | 5 |

These figures include values that were originally missing and values that failed the format tests.

## 03 — Transform | Cleaning and Converting Data

### 3.1 Cleaning unit prices and converting dates

I used `str.replace()` to remove the currency symbol, `pd.to_numeric()` to convert prices into numbers, and `pd.to_datetime()` to handle different date formats.

**Python code:**

<div dir="ltr" align="left">

```python
combined_sales_df["سعر الوحدة"] = pd.to_numeric(combined_sales_df["سعر الوحدة"]    .astype(str)
    .str.replace("ر.س", "", regex=False)
    .str.strip(),errors="coerce")

combined_sales_df["التاريخ"] = pd.to_datetime(combined_sales_df["التاريخ"], 
    errors="coerce",
     format="mixed",
    dayfirst=True,)
combined_sales_df.dtypes
```

</div>

**Output:** The date column became `datetime64[us]`, and the unit-price column became `float64`.

### 3.2 Recovering missing customer IDs

I identified invoices with a missing customer ID, searched their other line items for the existing customer ID, and filled the missing value with the corresponding ID from the same invoice.

**Python code:**

<div dir="ltr" align="left">

```python
mis =combined_sales_df[combined_sales_df["رمز العميل"].isna()]


# Invoices missing a customer ID
missing_invoices = combined_sales_df.loc[
    combined_sales_df["رمز العميل"].isna(),
    "رقم الفاتورة"
].unique()

# # Display all line items of these invoices
invoice_check_df = combined_sales_df.loc[
    combined_sales_df["رقم الفاتورة"].isin(missing_invoices),
    ["رقم الفاتورة", "رقم البند", "رمز العميل"]
]



for i in mis.index:
    invoice_number = mis.loc[i,"رقم الفاتورة"]
  
    customer_ids = combined_sales_df.loc[
        (combined_sales_df["رقم الفاتورة"] == invoice_number)
        & (combined_sales_df["رمز العميل"].notna()),
        "رمز العميل"
    ].unique()
 
    combined_sales_df.loc[i,"رمز العميل"] =customer_ids[0]
combined_sales_df[combined_sales_df["رمز العميل"].isna()]
```

</div>

**Output:** No missing customer IDs remained in the resulting table.

### 3.3 Renaming columns

I renamed `إجمالي السطر` (line amount) to `صافي مبيعات المنتج` (product net sales) and `سعر الوحدة` (unit price) to `سعر بيع الوحدة` (unit selling price), making the column meanings clearer.

<div dir="ltr" align="left">

```python
combined_sales_df = combined_sales_df.rename(
    columns={
        "إجمالي السطر": "صافي مبيعات المنتج",
        "سعر الوحدة": "سعر بيع الوحدة"
    }
)
combined_sales_df
```

</div>

### 3.4 Checking missing values again

I reused the same missing-value function after correcting the records, to confirm the result.

<div dir="ltr" align="left">

```python
miising_values_per_columns(combined_sales_df)
```

</div>

**Saved notebook result:** Zero missing values remained across the columns after cleaning.

## 04 — Data Validation | Checking the Results

### 4.1 Comparing calculated and recorded net sales

I calculated each product line's net sales using quantity, unit selling price, and discount percentage, then compared that amount with the recorded value. The test flagged differences **greater than SAR 0.01** and values that could not be calculated.

<div dir="ltr" align="left">

```python
# Calculate expected net sales
expected_sales = (
    combined_sales_df["الكمية"]
    * combined_sales_df["سعر بيع الوحدة"]
    * (1 - combined_sales_df["الخصم %"] / 100)
).round(2)

# Compare expected and recorded sales
sales_difference = (
    expected_sales
    - combined_sales_df["صافي مبيعات المنتج"]
).abs().round(2)
# Find records with calculation errors
incorrect_sales_df = combined_sales_df.loc[
    (sales_difference > 0.01)
    | expected_sales.isna()
    | combined_sales_df["صافي مبيعات المنتج"].isna()
]

print("عدد السجلات غير المتطابقة:", len(incorrect_sales_df))
incorrect_sales_df[
        [
            "رقم البند",
            "الكمية",
            "سعر بيع الوحدة",
            "الخصم %",
            "صافي مبيعات المنتج"
        ]
]
```

</div>

**Output:** `عدد السجلات غير المتطابقة: 0` (0 mismatching records) under the condition used in the notebook. This does not rule out differences smaller than or equal to the threshold.

### 4.2 Checking business rules and duplicate item IDs

I checked for non-positive quantities, negative selling prices, negative discounts, discounts above 100%, and repeated line-item IDs.

**Python code:**

<div dir="ltr" align="left">

```python
# 1. Zero or negative quantities

invalid_quantity =(combined_sales_df["الكمية"] <= 0).sum()
# 2. Zero or negative prices

invalid_price = (combined_sales_df["سعر بيع الوحدة"] <0).sum()
# 3. Negative discounts

negative_discount=     (combined_sales_df["الخصم %"] <0).sum()
# 4. Discounts above 100%

high_discount =     (combined_sales_df["الخصم %"] >100).sum()

# Count of duplicate line-item IDs
duplicate_item_count = (
    combined_sales_df["رقم البند"].duplicated().sum()
)
duplicate_item_count
# Create Validation Report
validation_report = pd.DataFrame([{
    "Invalid Quantity": invalid_quantity,
    "Invalid Price": invalid_price,
    "Negative Discount": negative_discount,
    "Discount Above 100%": high_discount,
    "duplicate_item_count":duplicate_item_count
}])
validation_report
```

</div>

**Output:** All five checks returned **0**.

**Note:** The price condition in the code is `< 0`, so it checks negative prices only, not zero prices.

### 4.3 Displaying records with repeated line-item IDs

I displayed all records that shared a line-item ID, if any.

<div dir="ltr" align="left">

```python
# Display all records sharing the same line-item ID
duplicate_items_df = combined_sales_df.loc[
    combined_sales_df["رقم البند"].duplicated(keep=False),
    ["رقم الفاتورة", "رقم البند", "اسم المنتج", "الفرع"]
]

duplicate_items_df
```

</div>

**Output:** The table was empty. No repeated line-item IDs appeared in this check.

## 05 — Load | Exporting to Excel

After cleaning the data and checking the results, I saved the consolidated dataset as an Excel file.

<div dir="ltr" align="left">

```python
combined_sales_df.to_excel("data ETL.xlsx", index=False)
```

</div>

**Output file:** `data ETL.xlsx`.

## Summary of Results

| Task | Result |
|---|---|
| HTML files read and combined | 20 files |
| Data size | 4,862 rows × 14 columns |
| Duplicate rows | 0 |
| Inconsistently formatted dates | 3 |
| Unit prices containing the currency symbol | 4 |
| Missing customer IDs | 5 before cleaning, 0 afterward |
| Net-sales discrepancies over SAR 0.01 under the notebook check | 0 |
| Business-rule violations under the checks used | 0 |
| Final export | `data ETL.xlsx` |
