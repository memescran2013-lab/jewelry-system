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
       # --- جزء الواتساب المطور (دفتر عناوين) ---
        st.divider()
        st.subheader("📱 إرسال الطلبية لجهات الاتصال")

        # قائمة بأسماء وأرقام ثابتة (تقدر تغير الأسماء والأرقام دي براحتك هنا)
        contacts = {
            "محمود رضا": "201002928684",
            "احمد محمود": "201234567890",
            "ورشة الإنتاج": "201122334455",
            "إضافة رقم آخر": "custom"
        }

        # اختيار الشخص من القائمة
        selected_contact = st.selectbox("اختر الشخص المراد الإرسال إليه:", list(contacts.keys()))

        # لو اختار "رقم آخر" يظهر له مكان يكتب فيه الرقم
        if contacts[selected_contact] == "custom":
            target_phone = st.text_input("اكتب الرقم الجديد (بمفتاح الدولة مثلاً 2010...):")
        else:
            target_phone = contacts[selected_contact]
            st.info(f"سيتم الإرسال إلى رقم: {target_phone}")

        if st.button("إرسال التقرير الآن"):
            if target_phone:
                # تجهيز الرسالة
                msg = f"تقرير طلبية إنتاج - New Egypt Gold\n"
                msg += f"---------------------------\n"
                msg += f"✅ إجمالي الموديلات: {len(df)}\n"
                msg += f"⚠️ موديلات عجز (إنتاج عاجل): {len(critical_low)}\n"
                msg += f"📦 موديلات مخزون زائد: {len(overstock)}\n"
                msg += f"---------------------------\n"
                msg += f"يرجى مراجعة ملف الإكسيل المرفق."

                import urllib.parse
                encoded_msg = urllib.parse.quote(msg)
                whatsapp_url = f"https://wa.me/{target_phone}?text={encoded_msg}"
                
                st.markdown(f"### [✅ اضغط هنا لفتح الواتساب والإرسال]({whatsapp_url})")
            else:
                st.warning("برجاء التأكد من كتابة الرقم أولاً.")
            
    else:
        st.error(f"تأكد أن ملف الإكسيل يحتوي على الأعمدة: {required_columns}")
