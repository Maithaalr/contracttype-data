import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# =========================================================
# إعداد الصفحة
# =========================================================
st.set_page_config(
    page_title="نظام تحليل العقود",
    page_icon="📊",
    layout="wide"
)

# =========================================================
# CSS
# =========================================================
st.markdown("""
<style>
    .main {
        direction: rtl;
    }

    html, body, [class*="css"] {
        direction: rtl;
        text-align: right;
    }

    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }

    div[data-testid="stMetric"] {
        background-color: #f8f9fa;
        border: 1px solid #e9ecef;
        padding: 18px;
        border-radius: 12px;
        text-align: center;
    }

    div[data-testid="stMetricLabel"] {
        justify-content: center;
    }

    div[data-testid="stMetricValue"] {
        text-align: center;
    }

    .main-title {
        font-size: 32px;
        font-weight: 700;
        margin-bottom: 5px;
    }

    .subtitle {
        color: #6c757d;
        font-size: 16px;
        margin-bottom: 25px;
    }

    .section-title {
        font-size: 21px;
        font-weight: 700;
        margin-top: 25px;
        margin-bottom: 15px;
    }
</style>
""", unsafe_allow_html=True)


# =========================================================
# دوال مساعدة
# =========================================================

def clean_column_names(df):
    """تنظيف أسماء الأعمدة"""
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.replace("\n", " ", regex=False)
    )
    return df


def find_column(df, possible_names):
    """
    البحث عن العمود حتى لو كان اسمه مختلفاً بشكل بسيط
    """
    normalized = {
        str(col).strip().replace(" ", ""): col
        for col in df.columns
    }

    for name in possible_names:
        key = name.strip().replace(" ", "")
        if key in normalized:
            return normalized[key]

    return None


def convert_numeric(df, column):
    """تحويل العمود إلى أرقام"""
    if column and column in df.columns:
        df[column] = (
            df[column]
            .astype(str)
            .str.replace(",", "", regex=False)
            .str.replace("AED", "", regex=False)
            .str.replace("درهم", "", regex=False)
            .str.strip()
        )

        df[column] = pd.to_numeric(
            df[column],
            errors="coerce"
        ).fillna(0)

    return df


def format_money(value):
    """تنسيق المبالغ"""
    try:
        return f"{value:,.2f}"
    except:
        return "0.00"


def create_excel(
    filtered_df,
    contract_type,
    total_contracts,
    total_cost,
    total_basic,
    total_supplementary,
    total_nature,
    top_entity,
    top_category,
    category_counts
):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        # -----------------------------------------
        # Sheet 1: العقود
        # -----------------------------------------
        filtered_df.to_excel(
            writer,
            sheet_name="العقود",
            index=False
        )

        # -----------------------------------------
        # Sheet 2: الإحصائيات
        # -----------------------------------------

        summary = pd.DataFrame({
            "المؤشر": [
                "نوع العقد",
                "إجمالي عدد العقود",
                "أكثر جهة",
                "أكثر فئة وظيفية",
                "إجمالي الراتب الأساسي",
                "إجمالي التكميلي",
                "إجمالي بدل طبيعة العمل",
                "إجمالي التكلفة الشهرية"
            ],
            "القيمة": [
                contract_type,
                total_contracts,
                top_entity,
                top_category,
                total_basic,
                total_supplementary,
                total_nature,
                total_cost
            ]
        })

        summary.to_excel(
            writer,
            sheet_name="الإحصائيات",
            index=False,
            startrow=0
        )

        # توزيع الفئات
        if not category_counts.empty:

            category_export = category_counts.reset_index()
            category_export.columns = [
                "الفئة الوظيفية",
                "عدد العقود"
            ]

            category_export.to_excel(
                writer,
                sheet_name="الإحصائيات",
                index=False,
                startrow=len(summary) + 3
            )

        # -----------------------------------------
        # تنسيق Excel
        # -----------------------------------------

        workbook = writer.book

        for sheet_name in workbook.sheetnames:

            worksheet = workbook[sheet_name]

            worksheet.sheet_view.rightToLeft = True

            # عرض الأعمدة
            for column_cells in worksheet.columns:

                max_length = 0

                column_letter = column_cells[0].column_letter

                for cell in column_cells:

                    try:
                        if cell.value is not None:
                            length = len(str(cell.value))

                            if length > max_length:
                                max_length = length

                    except:
                        pass

                adjusted_width = min(max_length + 4, 40)

                worksheet.column_dimensions[
                    column_letter
                ].width = adjusted_width

            # تثبيت الصف الأول
            worksheet.freeze_panes = "A2"

    output.seek(0)

    return output


# =========================================================
# العنوان
# =========================================================

st.markdown(
    '<div class="main-title">نظام تحليل العقود</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle">'
    'تحليل أنواع العقود والتكاليف والإحصائيات بصورة تلقائية'
    '</div>',
    unsafe_allow_html=True
)


# =========================================================
# رفع ملف Excel
# =========================================================

uploaded_file = st.file_uploader(
    "ارفع ملف بيانات العقود",
    type=["xlsx", "xls"]
)


if uploaded_file is None:

    st.info(
        "يرجى رفع ملف Excel الذي يحتوي على بيانات العقود للبدء."
    )

    st.stop()


# =========================================================
# قراءة الملف
# =========================================================

try:

    excel_file = pd.ExcelFile(uploaded_file)

    sheet_names = excel_file.sheet_names

    # إذا كان هناك أكثر من Sheet
    if len(sheet_names) > 1:

        selected_sheet = st.selectbox(
            "اختر شيت البيانات",
            sheet_names
        )

    else:

        selected_sheet = sheet_names[0]

    df = pd.read_excel(
        uploaded_file,
        sheet_name=selected_sheet
    )

except Exception as e:

    st.error(
        f"حدث خطأ أثناء قراءة الملف: {e}"
    )

    st.stop()


# =========================================================
# تنظيف البيانات
# =========================================================

df = clean_column_names(df)

# حذف الصفوف الفارغة بالكامل
df = df.dropna(how="all")


# =========================================================
# اكتشاف الأعمدة
# =========================================================

contract_type_col = find_column(
    df,
    [
        "نوع العقد",
        "نوع عقد",
        "العقد",
        "Contract Type"
    ]
)

entity_col = find_column(
    df,
    [
        "الجهة",
        "الدائرة",
        "الجهة الحكومية",
        "اسم الجهة",
        "Entity"
    ]
)

category_col = find_column(
    df,
    [
        "الفئة الوظيفية",
        "الفئة",
        "Job Category"
    ]
)

basic_salary_col = find_column(
    df,
    [
        "الراتب الأساسي",
        "الراتب الاساسي",
        "أساسي",
        "Basic Salary"
    ]
)

supplementary_col = find_column(
    df,
    [
        "التكميلي",
        "الراتب التكميلي",
        "Supplementary"
    ]
)

nature_allowance_col = find_column(
    df,
    [
        "بدل طبيعة عمل",
        "بدل طبيعة العمل",
        "Nature Allowance"
    ]
)

monthly_cost_col = find_column(
    df,
    [
        "التكلفة الشهرية",
        "إجمالي التكلفة الشهرية",
        "التكلفة",
        "Monthly Cost"
    ]
)


# =========================================================
# التأكد من وجود نوع العقد
# =========================================================

if contract_type_col is None:

    st.error(
        """
        لم أجد عمود **نوع العقد** في الملف.

        يجب أن يحتوي ملف البيانات الرئيسي على عمود باسم:
        **نوع العقد**

        مثال:
        عقد خاص / عقد مؤقت / عقد خبير / عقد مستشار
        """
    )

    st.write("الأعمدة الموجودة في الملف:")

    st.write(list(df.columns))

    st.stop()


# =========================================================
# تنظيف نوع العقد
# =========================================================

df[contract_type_col] = (
    df[contract_type_col]
    .astype(str)
    .str.strip()
)

df = df[
    ~df[contract_type_col].isin(
        ["", "nan", "None"]
    )
]


# =========================================================
# تحويل الأعمدة المالية
# =========================================================

for col in [
    basic_salary_col,
    supplementary_col,
    nature_allowance_col,
    monthly_cost_col
]:

    df = convert_numeric(df, col)


# =========================================================
# اختيار نوع العقد
# =========================================================

st.markdown(
    '<div class="section-title">اختيار نوع العقد</div>',
    unsafe_allow_html=True
)

contract_types = sorted(
    df[contract_type_col]
    .dropna()
    .unique()
    .tolist()
)

selected_contract = st.selectbox(
    "نوع العقد",
    contract_types
)


# =========================================================
# الفلترة الأساسية
# =========================================================

filtered_df = df[
    df[contract_type_col] == selected_contract
].copy()


# =========================================================
# حساب التكلفة الشهرية إذا لم تكن موجودة
# =========================================================

if monthly_cost_col is None:

    available_salary_columns = [
        col
        for col in [
            basic_salary_col,
            supplementary_col,
            nature_allowance_col
        ]
        if col is not None
    ]

    if available_salary_columns:

        filtered_df["التكلفة الشهرية المحسوبة"] = (
            filtered_df[
                available_salary_columns
            ].sum(axis=1)
        )

        monthly_cost_col = "التكلفة الشهرية المحسوبة"


# =========================================================
# الإحصائيات
# =========================================================

total_contracts = len(filtered_df)


# الراتب الأساسي
if basic_salary_col:
    total_basic = filtered_df[
        basic_salary_col
    ].sum()
else:
    total_basic = 0


# التكميلي
if supplementary_col:
    total_supplementary = filtered_df[
        supplementary_col
    ].sum()
else:
    total_supplementary = 0


# بدل طبيعة العمل
if nature_allowance_col:
    total_nature = filtered_df[
        nature_allowance_col
    ].sum()
else:
    total_nature = 0


# التكلفة الشهرية
if monthly_cost_col:
    total_cost = filtered_df[
        monthly_cost_col
    ].sum()
else:
    total_cost = (
        total_basic
        + total_supplementary
        + total_nature
    )


# =========================================================
# أكثر جهة
# =========================================================

if entity_col and not filtered_df.empty:

    entity_data = filtered_df[
        entity_col
    ].dropna()

    if not entity_data.empty:

        top_entity = (
            entity_data
            .value_counts()
            .idxmax()
        )

    else:

        top_entity = "غير متوفر"

else:

    top_entity = "غير متوفر"


# =========================================================
# أكثر فئة وظيفية
# =========================================================

if category_col and not filtered_df.empty:

    category_data = filtered_df[
        category_col
    ].dropna()

    if not category_data.empty:

        category_counts = (
            category_data
            .value_counts()
        )

        top_category = category_counts.idxmax()

    else:

        category_counts = pd.Series(dtype=int)

        top_category = "غير متوفر"

else:

    category_counts = pd.Series(dtype=int)

    top_category = "غير متوفر"


# =========================================================
# Dashboard
# =========================================================

st.markdown("---")

st.markdown(
    f'<div class="section-title">'
    f'ملخص {selected_contract}'
    f'</div>',
    unsafe_allow_html=True
)


# الصف الأول
col1, col2, col3, col4 = st.columns(4)

with col1:

    st.metric(
        "إجمالي العقود",
        f"{total_contracts:,}"
    )

with col2:

    st.metric(
        "إجمالي التكلفة الشهرية",
        f"{format_money(total_cost)} د.إ"
    )

with col3:

    st.metric(
        "أكثر جهة",
        top_entity
    )

with col4:

    st.metric(
        "أكثر فئة وظيفية",
        top_category
    )


# =========================================================
# تفاصيل التكلفة
# =========================================================

st.markdown(
    '<div class="section-title">'
    'تفاصيل التكلفة الشهرية'
    '</div>',
    unsafe_allow_html=True
)

cost1, cost2, cost3 = st.columns(3)

with cost1:

    st.metric(
        "إجمالي الراتب الأساسي",
        f"{format_money(total_basic)} د.إ"
    )

with cost2:

    st.metric(
        "إجمالي التكميلي",
        f"{format_money(total_supplementary)} د.إ"
    )

with cost3:

    st.metric(
        "إجمالي بدل طبيعة العمل",
        f"{format_money(total_nature)} د.إ"
    )


# =========================================================
# توزيع الفئات الوظيفية
# =========================================================

if not category_counts.empty:

    st.markdown(
        '<div class="section-title">'
        'توزيع العقود حسب الفئة الوظيفية'
        '</div>',
        unsafe_allow_html=True
    )

    category_df = (
        category_counts
        .reset_index()
    )

    category_df.columns = [
        "الفئة الوظيفية",
        "عدد العقود"
    ]

    chart_col, table_col = st.columns([2, 1])

    with chart_col:

        st.bar_chart(
            category_df.set_index(
                "الفئة الوظيفية"
            )
        )

    with table_col:

        st.dataframe(
            category_df,
            use_container_width=True,
            hide_index=True
        )


# =========================================================
# بيانات العقود
# =========================================================

st.markdown(
    '<div class="section-title">'
    'بيانات العقود'
    '</div>',
    unsafe_allow_html=True
)

st.caption(
    f"يتم عرض {len(filtered_df):,} سجل من نوع «{selected_contract}»"
)

st.dataframe(
    filtered_df,
    use_container_width=True,
    hide_index=True
)


# =========================================================
# تصدير Excel
# =========================================================

st.markdown("---")

excel_file = create_excel(
    filtered_df=filtered_df,
    contract_type=selected_contract,
    total_contracts=total_contracts,
    total_cost=total_cost,
    total_basic=total_basic,
    total_supplementary=total_supplementary,
    total_nature=total_nature,
    top_entity=top_entity,
    top_category=top_category,
    category_counts=category_counts
)


st.download_button(
    label="⬇️ تحميل تقرير Excel",
    data=excel_file,
    file_name=f"تقرير_{selected_contract}.xlsx",
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
    use_container_width=True
)
