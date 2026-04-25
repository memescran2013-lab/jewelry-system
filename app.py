import streamlit as st
import pandas as pd
import io

# إعداد الصفحة
st.set_page_config(page_title="سيستم طلبات الإنتاج الذكي", layout="wide")

st.title("📊 نظام تحليل المبيعات وإدارة طلبات الإنتاج")
st.write("ارفع ملف الإكسيل وحدد الأعمدة لبدء التحليل")

# رفع البيانات
uploaded_file = st.file_uploader("اختر ملف Excel", type=['xlsx'])

if uploaded_file:
    # قراءة الملف
    df = pd.read_excel(uploaded_file)
    all_columns = list(df.columns)
    
    st.success("✅ تم رفع الملف بنجاح")
    
    # واجهة اختيار الأعمدة ديناميكياً
    st.info("💡 برجاء تحديد الأعمدة المقابلة لكل مسمى ليبدأ البرنامج بالعمل:")
    col_a, col_b, col_c = st.columns(3)
    
    with col_a:
        cat_col = st.selectbox("عمود (المجموعة الصنفية):", all_columns)
    with col_b:
        sales_col = st.selectbox("عمود (كمية المبيعات):", all_columns)
    with col_c:
        stock_col = st.selectbox("عمود (المخزون الحالي):", all_columns)

    # تنفيذ المنطق البرمجي بناءً على الاختيارات
    if st.button("تشغيل التحليل الآن"):
        # تحويل الأعمدة المختارة لأرقام (في حال وجود نصوص بالخطأ)
        df[sales_col] = pd.to_numeric(df[sales_col], errors='coerce').fillna(0)
        df[stock_col] = pd.to_numeric(df[stock_col], errors='coerce').fillna(0)
        
        # الحسابات
        df['الكمية المطلوبة'] = df.apply(lambda row: max(0, row[sales_col] - row[stock_col]), axis=1)
        
        def check_status(row):
            if row[stock_col] < 0.2 * row[sales_col]: return 'طلب إنتاج عاجل'
            elif row[stock_col] > 2 * row[sales_col]: return 'مخزون زائد'
            return 'مستقر'
            
        df['الحالة'] = df.apply(check_status, axis=1)

        # عرض النتائج
        critical_low = df[df['الحالة'] == 'طلب إنتاج عاجل']
        overstock = df[df['الحالة'] == 'مخزون زائد']
        
        c1, c2, c3 = st.columns(3)
        c1.metric("إجمالي الموديلات", len(df))
        c2.metric("عجز (إنتاج عاجل)", len(critical_low))
        c3.metric("مخزون راكد", len(overstock))

        st.divider()
        st.subheader("📋 تقرير الحالة العامة")
        st.dataframe(df[[cat_col, sales_col, stock_col, 'الكمية المطلوبة', 'الحالة']])

        # تصدير الملف
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='التحليل')
        
        st.download_button(
            label="📥 تحميل النتائج كملف Excel",
            data=buffer.getvalue(),
            file_name="Production_Analysis.xlsx",
            mime="application/vnd.ms-excel"
        )
        
        # --- جزء الواتساب ---
        st.divider()
        st.subheader("📱 إرسال الطلبية")
        contacts = {"مدير المصنع": "201012345678", "إضافة رقم آخر": "custom"}
        selected_contact = st.selectbox("ارسل إلى:", list(contacts.keys()))
        
        target_phone = st.text_input("الرقم:") if contacts[selected_contact] == "custom" else contacts[selected_contact]
        
        if st.button("ارسل عبر واتساب"):
            msg = f"تقرير المصنع:\n- عدد العجز: {len(critical_low)}\n- الرواكد: {len(overstock)}"
            import urllib.parse
            st.markdown(f"[✅ اضغط للإرسال](https://wa.me/{target_phone}?text={urllib.parse.quote(msg)})")
