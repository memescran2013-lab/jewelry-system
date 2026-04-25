import streamlit as st
import pandas as pd
import io
import urllib.parse

# إعدادات الصفحة بشعار "نيو إيجيبت جولد"
st.set_page_config(page_title="نظام الإنتاج - New Egypt Gold", layout="wide")

st.title("🏭 نظام تحليل الماستر داتا وإدارة طلبات الإنتاج")
st.info("ارفع ملف الإكسيل الخاص بك، وحدد الأعمدة المطلوبة لبدء الحسابات التلقائية.")

# رفع الملف
uploaded_file = st.file_uploader("اختر ملف الماستر داتا (Excel)", type=['xlsx'])

if uploaded_file:
    try:
        # قراءة البيانات ومعالجة الأسماء المتكررة
        df = pd.read_excel(uploaded_file)
        df.columns = [str(c).strip() for c in df.columns]
        all_cols = list(df.columns)
        
        st.success("✅ تم رفع الملف بنجاح")
        
        # واجهة اختيار الأعمدة (عشان يشتغل على أي ملف)
        st.subheader("⚙️ إعدادات الربط مع الماستر داتا")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            sales_col = st.selectbox("حدد عمود (المبيعات):", all_cols)
        with col2:
            stock_col = st.selectbox("حدد عمود (المخزون/الرصيد):", all_cols)
        with col3:
            name_col = st.selectbox("حدد عمود (اسم الصنف):", all_cols)

        # زر التشغيل
        if st.button("🚀 تشغيل التحليل وحساب 'ط جديد'"):
            # تحويل البيانات لأرقام
            df[sales_col] = pd.to_numeric(df[sales_col], errors='coerce').fillna(0)
            df[stock_col] = pd.to_numeric(df[stock_col], errors='coerce').fillna(0)
            
            # --- المنهجية الاحترافية ---
            # حساب "ط جديد" (الكمية المطلوبة)
            df['ط جديد'] = (df[sales_col] - df[stock_col]).clip(lower=0)
            
            # تحديد الحالة
            def get_status(row):
                if row[stock_col] < (0.2 * row[sales_col]): return '⚠️ طلب إنتاج عاجل'
                elif row[stock_col] > (2.0 * row[sales_col]): return '🧊 مخزون زائد'
                return '✅ مستقر'
            
            df['الحالة'] = df.apply(get_status, axis=1)

            # عرض المؤشرات
            critical = df[df['الحالة'] == '⚠️ طلب إنتاج عاجل']
            c1, c2, c3 = st.columns(3)
            c1.metric("إجمالي الأصناف", len(df))
            c2.metric("أصناف تحتاج إنتاج", len(critical))
            c3.metric("إجمالي قطع 'ط جديد'", int(df['ط جديد'].sum()))

            st.divider()
            
            # عرض الجدول بنفس مسمياتك الأصلية + الأعمدة الجديدة
            st.subheader("📋 تقرير الماستر داتا المحدث")
            st.dataframe(df, use_container_width=True)

            # تصدير الملف بنفس الهيكل
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Production_Plan')
            
            st.download_button(
                label="📥 تحميل ملف الماستر داتا المحدث (Excel)",
                data=buffer.getvalue(),
                file_name="New_Egypt_Gold_MasterData.xlsx",
                mime="application/vnd.ms-excel"
            )

            # نظام الواتساب
            st.divider()
            st.subheader("📱 إرسال النتائج")
            phone = st.text_input("اكتب رقم الواتساب (مثلاً 2010...):", "2010") # حط رقمك الافتراضي هنا
            
            if st.button("ارسل عبر واتساب"):
                msg = f"تقرير نيو إيجيبت جولد:\n- أصناف عجز: {len(critical)}\n- إجمالي مطلوب: {int(df['ط جديد'].sum())} قطعة."
                st.markdown(f"[✅ اضغط للإرسال](https://wa.me/{phone}?text={urllib.parse.quote(msg)})")

    except Exception as e:
        st.error(f"حدث خطأ: {e}")
