import streamlit as st
import pandas as pd
import io

# إعداد الصفحة
st.set_page_config(page_title="سيستم الإنتاج الذكي - New Egypt Gold", layout="wide")

st.title("📊 نظام تحليل المبيعات وإدارة طلبات الإنتاج")
st.write("ارفع ملف الإكسيل وحدد الأعمدة لبدء التحليل")

# رفع البيانات
uploaded_file = st.file_uploader("اختر ملف Excel", type=['xlsx'])

if uploaded_file:
    # قراءة الملف ومعالجة تكرار أسماء الأعمدة تلقائياً
    df = pd.read_excel(uploaded_file)
    
    # حل مشكلة الأسماء المتكررة
    df.columns = [f"{c}_{i}" if df.columns.duplicated()[i] else c for i, c in enumerate(df.columns)]
    all_columns = list(df.columns)
    
    st.success("✅ تم رفع الملف بنجاح")
    
    # واجهة اختيار الأعمدة
    st.info("💡 حدد الأعمدة الصحيحة من ملفك:")
    col_a, col_b, col_c = st.columns(3)
    
    with col_a:
        cat_col = st.selectbox("عمود (المجموعة الصنفية):", all_columns, key="cat")
    with col_b:
        sales_col = st.selectbox("عمود (كمية المبيعات):", all_columns, key="sales")
    with col_c:
        stock_col = st.selectbox("عمود (المخزون الحالي):", all_columns, key="stock")

    # زر التشغيل - لضمان عدم حدوث خطأ قبل الاختيار
    if st.button("🚀 تشغيل التحليل وحساب الكميات"):
        # تنظيف البيانات وتحويلها لأرقام
        df[sales_col] = pd.to_numeric(df[sales_col], errors='coerce').fillna(0)
        df[stock_col] = pd.to_numeric(df[stock_col], errors='coerce').fillna(0)
        
        # معادلة الإنتاج: (البيع - المخزون) وبحد أدنى صفر
        df['الكمية المطلوبة'] = (df[sales_col] - df[stock_col]).clip(lower=0)
        
        # تحديد الحالة
        def get_status(row):
            if row[stock_col] < 0.2 * row[sales_col]: return '⚠️ طلب إنتاج عاجل'
            elif row[stock_col] > 2 * row[sales_col]: return '🧊 مخزون زائد (راكد)'
            return '✅ مستقر'
            
        df['الحالة'] = df.apply(get_status, axis=1)

        # عرض النتائج في مربعات إحصائية
        crit = df[df['الحالة'] == '⚠️ طلب إنتاج عاجل']
        over = df[df['الحالة'] == '🧊 مخزون زائد (راكد)']
        
        m1, m2, m3 = st.columns(3)
        m1.metric("إجمالي الموديلات", len(df))
        m2.metric("موديلات ناقصة", len(crit))
        m3.metric("موديلات زائدة", len(over))

        st.divider()
        
        # عرض الجدول النهائي
        st.subheader("📋 نتائج التحليل التفصيلية")
        show_cols = [cat_col, sales_col, stock_col, 'الكمية المطلوبة', 'الحالة']
        st.dataframe(df[show_cols], use_container_width=True)

        # تجهيز ملف التحميل
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
        
        st.download_button(
            label="📥 تحميل التقرير النهائي (Excel)",
            data=output.getvalue(),
            file_name="Production_Plan.xlsx",
            mime="application/vnd.ms-excel"
        )

        # إرسال واتساب
        st.divider()
        st.subheader("📱 إرسال النتائج")
        # دفتر العناوين
        contacts = {
            "مدير المصنع": "201012345678", 
            "المكتب الفني": "201234567890",
            "رقم آخر": "custom"
        }
        choice = st.selectbox("إرسال إلى:", list(contacts.keys()))
        phone = st.text_input("ادخل الرقم:") if contacts[choice] == "custom" else contacts[choice]
        
        if st.button("ارسل الآن عبر WhatsApp"):
            msg = f"تقرير New Egypt Gold:\n- مطلوب إنتاج: {len(crit)} صنف.\n- أصناف راكدة: {len(over)} صنف."
            import urllib.parse
            link = f"https://wa.me/{phone}?text={urllib.parse.quote(msg)}"
            st.markdown(f"[✅ اضغط هنا لفتح واتساب والإرسال]({link})")
