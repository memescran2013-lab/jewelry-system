import streamlit as st
import pandas as pd
import io
import urllib.parse

# إعدادات واجهة New Egypt Gold
st.set_page_config(page_title="نظام التخطيط - New Egypt Gold", layout="wide")

st.markdown("""
    <style>
    .main { text-align: right; }
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #007bff; color: white; }
    </style>
    """, unsafe_allow_config=True)

st.title("🏭 نظام إدارة طلبات الإنتاج - الماستر داتا")
st.info("ارفع ملف الإكسيل (الماستر داتا) وسيقوم النظام بحساب الطلبات بناءً على مبيعاتك ورصيدك الحالي.")

uploaded_file = st.file_uploader("اختر ملف الماستر داتا (Excel)", type=['xlsx'])

if uploaded_file:
    try:
        # قراءة الملف - بنفترض إن البيانات في أول شيت
        df = pd.read_excel(uploaded_file)
        
        # تنظيف أسماء الأعمدة من أي مسافات مخفية
        df.columns = df.columns.str.strip()
        
        # قائمة بالأعمدة الأساسية المطلوبة من ملفك (حسب الملف المرفق)
        required_cols = ['SKU', 'ItemName', 'ItemGroup', 'مبيعات', 'رصيد']
        
        # التأكد من وجود الأعمدة
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            st.error(f"⚠️ الملف ينقصه الأعمدة التالية: {missing}")
            st.stop()

        # تحويل البيانات لأرقام لضمان دقة الحسابات
        df['مبيعات'] = pd.to_numeric(df['مبيعات'], errors='coerce').fillna(0)
        df['رصيد'] = pd.to_numeric(df['رصيد'], errors='coerce').fillna(0)

        # --- منهجية الحساب الاحترافية ---
        # 1. حساب "ط جديد" (الكمية المطلوبة لسد الفجوة)
        df['ط جديد'] = (df['مبيعات'] - df['رصيد']).clip(lower=0)
        
        # 2. تحديد الحالة (عاجل / راكد)
        def get_status(row):
            if row['رصيد'] < (0.2 * row['مبيعات']): return '⚠️ طلب إنتاج عاجل'
            elif row['رصيد'] > (2.0 * row['مبيعات']): return '🧊 مخزون زائد'
            return '✅ مستقر'
        
        df['الحالة'] = df.apply(get_status, axis=1)

        # --- العرض الإحصائي ---
        critical = df[df['الحالة'] == '⚠️ طلب إنتاج عاجل']
        overstock = df[df['الحالة'] == '🧊 مخزون زائد']
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("إجمالي الأصناف", len(df))
        c2.metric("عجز عاجل", len(critical))
        c3.metric("رصيد زائد", len(overstock))
        c4.metric("إجمالي المطلوب", f"{int(df['ط جديد'].sum())} قطعة")

        st.divider()

        # --- الجدول النهائي (بنفس مسميات الماستر داتا) ---
        st.subheader("📋 جدول تحليل الإنتاج")
        # هنعرض الماستر داتا ومعاها الحسابات الجديدة
        display_cols = ['SKU', 'ItemName', 'ItemGroup', 'مبيعات', 'رصيد', 'ط جديد', 'الحالة']
        st.dataframe(df[display_cols], use_container_width=True)

        # --- تصدير النتائج بنفس الهيكل ---
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Master_Data_Plan')
        
        st.download_button(
            label="📥 تحميل ملف الماستر داتا المعدل",
            data=buffer.getvalue(),
            file_name="New_Egypt_Gold_Plan.xlsx",
            mime="application/vnd.ms-excel"
        )

        # --- نظام الواتساب المرن ---
        st.divider()
        st.subheader("📱 إرسال الطلبية")
        
        contacts = {
            "رئيس قسم تنسيق المبيعات": "201002928684",
            "المكتب الفني": "201203004455",
            "إضافة رقم جديد": "custom"
        }
        
        sel_contact = st.selectbox("إرسال إلى:", list(contacts.keys()))
        
        if contacts[sel_contact] == "custom":
            phone = st.text_input("اكتب الرقم (مثال: 2010...)")
        else:
            phone = contacts[sel_contact]
            st.info(f"الرقم المختار: {phone}")

        if st.button("ارسل التقرير عبر WhatsApp"):
            if phone:
                report_msg = (
                    f"🛑 *تقرير إنتاج New Egypt Gold*\n"
                    f"--------------------------\n"
                    f"✅ إجمالي المطلوب إنتاجه: {int(df['ط جديد'].sum())} قطعة\n"
                    f"⚠️ عدد الموديلات العاجلة: {len(critical)}\n"
                    f"--------------------------\n"
                    f"تم تحديث ملف الماستر داتا.. برجاء المراجعة."
                )
                encoded_msg = urllib.parse.quote(report_msg)
                st.markdown(f"### [✅ اضغط هنا للإرسال لواتساب]({f'https://wa.me/{phone}?text={encoded_msg}'})")

    except Exception as e:
        st.error(f"حدث خطأ أثناء معالجة الملف: {e}")
