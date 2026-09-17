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
# تنسيق الصفحة RTL
# =========================================================

st.markdown("""
<style>

    html, body, [class*="css"] {
        direction: rtl;
        text-align: right;
    }

    .main {
        direction: rtl;
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


# =========================================================
# إنشاء ملف Excel
# =========================================================

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
    category_counts,
    entity_counts
):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        # =================================================
        # Sheet 1: العقود
        # =================================================

        filtered_df.to_excel(
            writer,
            sheet_name="العقود",
            index=False
        )


        # =================================================
        # Sheet 2: الإحصائيات الرئيسية
        # =================================================

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


        # =================================================
        # توزيع العقود حسب الفئة الوظيفية
        # =================================================

        category_start_row = len(summary) + 3

        if not category_counts.empty:

            category_export = category_counts.reset_index()

            category_export.columns = [
                "الفئة الوظيفية",
                "عدد العقود"
            ]

            # إضافة الإجمالي
            category_total = pd.DataFrame({
                "الفئة الوظيفية": ["الإجمالي"],
                "عدد العقود": [
                    category_export["عدد العقود"].sum()
                ]
            })

            category_export = pd.concat(
                [
                    category_export,
                    category_total
                ],
                ignore_index=True
            )

            category_export.to_excel(
                writer,
                sheet_name="الإحصائيات",
                index=False,
                startrow=category_start_row
            )


        # =================================================
        # عدد الموظفين حسب الدائرة
        # =================================================

        entity_start_row = (
            category_start_row
            + len(category_counts)
            + 5
        )

        if not entity_counts.empty:

            entity_export = entity_counts.reset_index()

            entity_export.columns = [
                "الدائرة / الجهة",
                "عدد الموظفين"
            ]

            # إضافة الإجمالي
            entity_total = pd.DataFrame({

                "الدائرة / الجهة": [
                    "الإجمالي"
                ],

                "عدد الموظفين": [
                    entity_export[
                        "عدد الموظفين"
                    ].sum()
                ]

            })

            entity_export = pd.concat(
                [
                    entity_export,
                    entity_total
                ],
                ignore_index=True
            )

            entity_export.to_excel(
                writer,
                sheet_name="الإحصائيات",
                index=False,
                startrow=entity_start_row
            )


        # =================================================
        # تنسيق ملف Excel
        # =================================================

        workbook = writer.book

        from openpyxl.styles import (
            Font,
            PatternFill,
            Alignment,
            Border,
            Side
        )


        # لون العناوين
        header_fill = PatternFill(
            fill_type="solid",
            fgColor="D9EAF7"
        )

        header_font = Font(
            bold=True
        )

        total_fill = PatternFill(
            fill_type="solid",
            fgColor="E2F0D9"
        )

        thin_border = Border(
            left=Side(
                style="thin",
                color="D9D9D9"
            ),
            right=Side(
                style="thin",
                color="D9D9D9"
            ),
            top=Side(
                style="thin",
                color="D9D9D9"
            ),
            bottom=Side(
                style="thin",
                color="D9D9D9"
            )
        )


        for sheet_name in workbook.sheetnames:

            worksheet = workbook[sheet_name]

            # RTL
            worksheet.sheet_view.rightToLeft = True

            # تثبيت الصف
            worksheet.freeze_panes = "A2"


            # ---------------------------------------------
            # تنسيق الخلايا
            # ---------------------------------------------

            for row in worksheet.iter_rows():

                for cell in row:

                    cell.alignment = Alignment(
                        horizontal="right",
                        vertical="center"
                    )

                    if cell.value is not None:

                        cell.border = thin_border


            # ---------------------------------------------
            # تنسيق الصف الأول
            # ---------------------------------------------

            for cell in worksheet[1]:

                cell.fill = header_fill
                cell.font = header_font


            # ---------------------------------------------
            # البحث عن أي صف يحتوي على الإجمالي
            # ---------------------------------------------

            for row in worksheet.iter_rows():

                if any(
                    cell.value == "الإجمالي"
                    for cell in row
                ):

                    for cell in row:

                        cell.fill = total_fill
                        cell.font = Font(
                            bold=True
                        )


            # ---------------------------------------------
            # عرض الأعمدة
            # ---------------------------------------------

            for column_cells in worksheet.columns:

                max_length = 0

                column_letter = (
                    column_cells[0]
                    .column_letter
                )

                for cell in column_cells:

                    try:

                        if cell.value is not None:

                            length = len(
                                str(cell.value)
                            )

                            if length > max_length:

                                max_length = length

                    except:
                        pass


                adjusted_width = min(
                    max_length + 5,
                    45
                )

                worksheet.column_dimensions[
                    column_letter
                ].width = adjusted_width


    output.seek(0)

    return output


# =========================================================
# عنوان النظام
# =========================================================

st.markdown(
    '<div class="main-title">'
    'نظام تحليل العقود'
    '</div>',
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
# قراءة ملف Excel
# =========================================================

try:

    excel_file = pd.ExcelFile(
        uploaded_file
    )

    sheet_names = (
        excel_file.sheet_names
    )


    # إذا كان الملف يحتوي على أكثر من Sheet
    if len(sheet_names) > 1:

        selected_sheet = st.selectbox(
            "اختر شيت البيانات",
            sheet_names
        )

    else:

        selected_sheet = (
            sheet_names[0]
        )


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
df = df.dropna(
    how="all"
)


# =========================================================
# اكتشاف الأعمدة تلقائياً
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
        "اسم الدائرة",
        "الدائرة الحكومية",
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
        "اجمالي التكلفة الشهرية",
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

        يجب أن يحتوي ملف البيانات على عمود باسم:
        **نوع العقد**
        """
    )

    st.write(
        "الأعمدة الموجودة في الملف:"
    )

    st.write(
        list(df.columns)
    )

    st.stop()


# =========================================================
# تنظيف عمود نوع العقد
# =========================================================

df[contract_type_col] = (
    df[contract_type_col]
    .astype(str)
    .str.strip()
)


df = df[
    ~df[contract_type_col].isin(
        [
            "",
            "nan",
            "None"
        ]
    )
]


# =========================================================
# تحويل الأعمدة المالية إلى أرقام
# =========================================================

for col in [
    basic_salary_col,
    supplementary_col,
    nature_allowance_col,
    monthly_cost_col
]:

    df = convert_numeric(
        df,
        col
    )


# =========================================================
# اختيار نوع العقد
# =========================================================

st.markdown(
    '<div class="section-title">'
    'اختيار نوع العقد'
    '</div>',
    unsafe_allow_html=True
)


contract_types = sorted(
    df[
        contract_type_col
    ]
    .dropna()
    .unique()
    .tolist()
)


selected_contract = st.selectbox(
    "نوع العقد",
    contract_types
)


# =========================================================
# فلترة البيانات حسب نوع العقد
# =========================================================

filtered_df = df[
    df[contract_type_col]
    == selected_contract
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

        filtered_df[
            "التكلفة الشهرية المحسوبة"
        ] = (

            filtered_df[
                available_salary_columns
            ]
            .sum(axis=1)

        )

        monthly_cost_col = (
            "التكلفة الشهرية المحسوبة"
        )


# =========================================================
# الإحصائيات الرئيسية
# =========================================================

total_contracts = len(
    filtered_df
)


# الراتب الأساسي
if basic_salary_col:

    total_basic = (
        filtered_df[
            basic_salary_col
        ].sum()
    )

else:

    total_basic = 0


# التكميلي
if supplementary_col:

    total_supplementary = (
        filtered_df[
            supplementary_col
        ].sum()
    )

else:

    total_supplementary = 0


# بدل طبيعة العمل
if nature_allowance_col:

    total_nature = (
        filtered_df[
            nature_allowance_col
        ].sum()
    )

else:

    total_nature = 0


# التكلفة الشهرية
if monthly_cost_col:

    total_cost = (
        filtered_df[
            monthly_cost_col
        ].sum()
    )

else:

    total_cost = (
        total_basic
        + total_supplementary
        + total_nature
    )


# =========================================================
# أكثر جهة / دائرة
# =========================================================

if entity_col and not filtered_df.empty:

    entity_data = (
        filtered_df[
            entity_col
        ]
        .dropna()
        .astype(str)
        .str.strip()
    )

    entity_data = entity_data[
        entity_data != ""
    ]


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
# عدد الموظفين حسب كل دائرة
# =========================================================

if entity_col and not filtered_df.empty:

    entity_counts = (
        filtered_df[
            entity_col
        ]
        .dropna()
        .astype(str)
        .str.strip()
    )

    entity_counts = entity_counts[
        entity_counts != ""
    ]

    entity_counts = (
        entity_counts
        .value_counts()
    )

else:

    entity_counts = (
        pd.Series(dtype=int)
    )


# =========================================================
# أكثر فئة وظيفية
# =========================================================

if category_col and not filtered_df.empty:

    category_data = (
        filtered_df[
            category_col
        ]
        .dropna()
        .astype(str)
        .str.strip()
    )

    category_data = category_data[
        category_data != ""
    ]


    if not category_data.empty:

        category_counts = (
            category_data
            .value_counts()
        )

        top_category = (
            category_counts
            .idxmax()
        )

    else:

        category_counts = (
            pd.Series(dtype=int)
        )

        top_category = (
            "غير متوفر"
        )

else:

    category_counts = (
        pd.Series(dtype=int)
    )

    top_category = (
        "غير متوفر"
    )


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


# =========================================================
# بطاقات الإحصائيات
# =========================================================

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


cost1, cost2, cost3 = (
    st.columns(3)
)


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
# توزيع العقود حسب الفئة الوظيفية
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


    # إضافة الإجمالي للجدول
    category_total = pd.DataFrame({

        "الفئة الوظيفية": [
            "الإجمالي"
        ],

        "عدد العقود": [
            category_df[
                "عدد العقود"
            ].sum()
        ]

    })


    category_df_display = (
        pd.concat(
            [
                category_df,
                category_total
            ],
            ignore_index=True
        )
    )


    category_chart_col, category_table_col = (
        st.columns([2, 1])
    )


    with category_chart_col:

        st.bar_chart(
            category_df.set_index(
                "الفئة الوظيفية"
            )
        )


    with category_table_col:

        st.dataframe(
            category_df_display,
            use_container_width=True,
            hide_index=True
        )


# =========================================================
# عدد الموظفين حسب الدائرة
# =========================================================

if not entity_counts.empty:

    st.markdown(
        '<div class="section-title">'
        'عدد الموظفين حسب الدائرة'
        '</div>',
        unsafe_allow_html=True
    )


    entity_df = (
        entity_counts
        .reset_index()
    )


    entity_df.columns = [
        "الدائرة / الجهة",
        "عدد الموظفين"
    ]


    # إضافة الإجمالي
    entity_total = pd.DataFrame({

        "الدائرة / الجهة": [
            "الإجمالي"
        ],

        "عدد الموظفين": [
            entity_df[
                "عدد الموظفين"
            ].sum()
        ]

    })


    entity_df_display = (
        pd.concat(
            [
                entity_df,
                entity_total
            ],
            ignore_index=True
        )
    )


    entity_chart_col, entity_table_col = (
        st.columns([2, 1])
    )


    with entity_chart_col:

        st.bar_chart(
            entity_df.set_index(
                "الدائرة / الجهة"
            )
        )


    with entity_table_col:

        st.dataframe(
            entity_df_display,
            use_container_width=True,
            hide_index=True
        )

# =========================================================
# تحليل بدل طبيعة العمل حسب الدائرة
# =========================================================

if nature_allowance_col and entity_col:

    # فقط الموظفين الذين يستلمون بدل طبيعة عمل
    # أي قيمة 0 أو Blank يتم تجاهلها
    nature_analysis_df = filtered_df[
        filtered_df[nature_allowance_col] > 0
    ].copy()

    if not nature_analysis_df.empty:

        # إجمالي عدد المستلمين
        nature_employee_count = len(nature_analysis_df)

        # إجمالي تكلفة بدل طبيعة العمل
        nature_total_cost = (
            nature_analysis_df[
                nature_allowance_col
            ].sum()
        )

        # التجميع حسب الدائرة
        nature_entity_df = (
            nature_analysis_df
            .groupby(entity_col)
            .agg(
                عدد_الموظفين=(
                    nature_allowance_col,
                    "count"
                ),
                إجمالي_التكلفة=(
                    nature_allowance_col,
                    "sum"
                )
            )
            .reset_index()
            .sort_values(
                "إجمالي_التكلفة",
                ascending=False
            )
        )

        # تغيير اسم عمود الدائرة للعرض
        nature_entity_df = nature_entity_df.rename(
            columns={
                entity_col: "الدائرة / الجهة"
            }
        )

        # العنوان
        st.markdown(
            '<div class="section-title">'
            'تحليل بدل طبيعة العمل'
            '</div>',
            unsafe_allow_html=True
        )

        # بطاقات الإجمالي
        nature_col1, nature_col2 = st.columns(2)

        with nature_col1:

            st.metric(
                "عدد الموظفين المستلمين لبدل طبيعة العمل",
                f"{nature_employee_count:,}"
            )

        with nature_col2:

            st.metric(
                "إجمالي تكلفة بدل طبيعة العمل",
                f"{format_money(nature_total_cost)} د.إ"
            )

        # جدول حسب الدائرة
        st.dataframe(
            nature_entity_df,
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
    f"""
    يتم عرض {len(filtered_df):,}
    سجل من نوع «{selected_contract}»
    """
)


st.dataframe(
    filtered_df,
    use_container_width=True,
    hide_index=True
)


# =========================================================
# إنشاء Excel
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

    category_counts=category_counts,

    entity_counts=entity_counts

)


# =========================================================
# تحميل Excel
# =========================================================

st.download_button(

    label="⬇️ تحميل تقرير Excel",

    data=excel_file,

    file_name=(
        f"تقرير_{selected_contract}.xlsx"
    ),

    mime=(
        "application/"
        "vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),

    use_container_width=True

)
