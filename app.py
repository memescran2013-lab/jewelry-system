import streamlit as st
import pandas as pd
import io

# إعداد الصفحة
st.set_page_config(page_title="سيستم طلبات الإنتاج - New Egypt Gold", layout="wide")

st.title("📊 نظام تحليل المبيعات وإدارة طلبات الإنتاج")
st.write("ارفع ملف الإكسيل الخاص بك لتحليل المخزون واقتراح الكميات المطلوبة.")

# رفع البيانات
uploaded_file = st.file_uploader("اختر ملف Excel", type=['xlsx'])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    # التأكد من وجود الأعمدة المطلوبة
    required_columns = ['المجموعة الصنفية', 'كمية المبيعات', 'المخزون الحالي']
    if all(col in df.columns for col in required_columns):
        
        # منطق الحساب: الكمية المطلوبة
        df['الكمية المطلوبة'] = df.apply(lambda row: max(0, row['كمية المبيعات'] - row['المخزون الحالي']), axis=1)
        
        # تحديد الحالة
        def check_status(row):
            if row['المخزون الحالي'] < 0.2 * row['كمية المبيعات']:
                return 'طلب إنتاج عاجل'
            elif row['المخزون الحالي'] > 2 * row['كمية المبيعات']:
                return 'مخزون زائد'
            return 'مستقر'
            
        df['الحالة'] = df.apply(check_status, axis=1)

        # عرض ملخص سريع
        critical_low = df[df['الحالة'] == 'طلب إنتاج عاجل']
        overstock = df[df['الحالة'] == 'مخزون زائد']
        
        col1, col2, col3 = st.columns(3)
        col1.metric("إجمالي الموديلات", len(df))
        col2.metric("عجز (إنتاج عاجل)", len(critical_low))
        col3.metric("مخزون راكد", len(overstock))

        # اختيار المجموعة الصنفية
        st.divider()
        category = st.selectbox("فلتر حسب المجموعة الصنفية:", ["الكل"] + list(df['المجموعة الصنفية'].unique()))
        
        filtered_df = df if category == "الكل" else df[df['المجموعة الصنفية'] == category]
        
        st.subheader(f"تحليل أصناف: {category}")
        st.dataframe(filtered_df.style.highlight_max(axis=0, subset=['كمية المبيعات'], color='#d4edda'))

        # التوب 15
        st.subheader("🔝 أفضل 15 موديل مبيعاً")
        st.table(filtered_df.nlargest(15, 'كمية المبيعات')[['المجموعة الصنفية', 'كمية المبيعات', 'المخزون الحالي', 'الكمية المطلوبة']])

        # تصدير الملف
        st.divider()
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='طلبات الإنتاج')
        
        st.download_button(
            label="📥 تحميل ملف طلبات الإنتاج (Excel)",
            data=buffer.getvalue(),
            file_name="Production_Orders.xlsx",
            mime="application/vnd.ms-excel"
        )

        # واتساب
        st.subheader("📱 إرسال تقرير سريع")
        phone = st.text_input("رقم الواتساب (بمفتاح الدولة مثلا 2010...)")
        if st.button("تجهيز رسالة الواتساب"):
            msg = f"تقرير المصنع: عدد الموديلات المطلوب إنتاجها فوراً هو {len(critical_low)} موديل."
            whatsapp_url = f"https://wa.me/{phone}?text={msg}"
            st.write(f"[اضغط هنا لفتح واتساب وإرسال التقرير]({whatsapp_url})")
            
    else:
        st.error(f"تأكد أن ملف الإكسيل يحتوي على الأعمدة: {required_columns}")
