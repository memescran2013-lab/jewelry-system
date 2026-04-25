import streamlit as st
import pandas as pd
import io
import urllib.parse

# إعدادات الصفحة
st.set_page_config(page_title="سيستم نيو إيجيبت جولد - التخطيط الشامل", layout="wide")

st.title("🏭 نظام التخطيط المتكامل (ماستر داتا)")
st.write("التحليل بناءً على مبيعات الفروع، المخزون، وحالة الإنتاج بالمصنع.")

# رفع الملف
uploaded_file = st.file_uploader("ارفع ملف الإكسيل الأخير", type=['xlsx'])

if uploaded_file:
    try:
        # قراءة البيانات
        df = pd.read_excel(uploaded_file)
        df.columns = [str(c).strip() for c in df.columns] # تنظيف الأسماء

        # --- تعريف الأعمدة بناءً على ملفك الأخير ---
        # المسميات في ملفك: SKU, ItemName, طلب الفروع, تحت التشغيل كمية, مبيعات, رصيد
        
        col_map = {
            'sku': 'SKU',
            'name': 'ItemName',
            'branch_order': 'طلب الفروع',
            'wip': 'تحت التشغيل كمية',
            'sales': 'مبيعات',
            'stock': 'رصيد'
        }

        # التأكد من وجود الأعمدة أو السماح للمستخدم باختيار البديل
        st.sidebar.header("⚙️ ضبط الربط")
        actual_cols = {}
        for key, default in col_map.items():
            actual_cols[key] = st.sidebar.selectbox(f"حدد عمود {default}:", df.columns, index=list(df.columns).index(default) if default in df.columns else 0)

        if st.button("📊 تشغيل التحليل المجمع"):
            # تحويل البيانات لأرقام
            for key in ['branch_order', 'wip', 'sales', 'stock']:
                df[actual_cols[key]] = pd.to_numeric(df[actual_cols[key]], errors='coerce').fillna(0)

            # --- معادلة القرار الموثوقة ---
            # صافي المطلوب = (طلب الفروع) - (الرصيد الحالي + اللي تحت التشغيل في المصنع)
            df['إجمالي المتوفر'] = df[actual_cols['stock']] + df[actual_cols['wip']]
            df['القرار (ط جديد)'] = (df[actual_cols['branch_order']] - df['إجمالي المتوفر']).clip(lower=0)
            
            # تحليل حركة الصنف (سريع / راكد)
            def get_movement(row):
                if row[actual_cols['sales']] > row[actual_cols['stock']]: return "🔥 سريع جداً"
                if row[actual_cols['sales']] == 0: return "🧊 راكد"
                return "✅ مستقر"
            
            df['حركة الموديل'] = df.apply(get_movement, axis=1)

            # عرض النتائج
            st.success("✅ تم بناء الخطة بناءً على أرصدة الفروع والمصنع")
            
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("إجمالي قطع 'ط جديد'", int(df['القرار (ط جديد)'].sum()))
            m2.metric("أصناف تحت التشغيل", int(df[actual_cols['wip']].sum()))
            m3.metric("موديلات مطلوبة فوراً", len(df[df['القرار (ط جديد)'] > 0]))
            m4.metric("موديلات سريعة البيع", len(df[df['حركة الموديل'] == "🔥 سريع جداً"]))

            st.divider()

            # عرض الجدول بالبيانات المختصرة والمهمة
            st.subheader("📋 تقرير التخطيط المقترح")
            display_df = df[[actual_cols['sku'], actual_cols['name'], actual_cols['sales'], 
                             actual_cols['stock'], actual_cols['wip'], actual_cols['branch_order'], 
                             'القرار (ط جديد)', 'حركة الموديل']]
            
            st.dataframe(display_df, use_container_width=True)

            # تحميل الملف المعدل
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='خطة الإنتاج')
            
            st.download_button("📥 تحميل التقرير النهائي (Excel)", output.getvalue(), 
                               "New_Egypt_Gold_Final_Plan.xlsx", "application/vnd.ms-excel")

            # واتساب
            st.divider()
            if st.button("📱 إرسال ملخص الطلبية للمدير"):
                msg = f"تقرير نيو إيجيبت جولد:\n- إجمالي المطلوب: {int(df['القرار (ط جديد)'].sum())} قطعة.\n- تحت التشغيل بالمصنع: {int(df[actual_cols['wip']].sum())} قطعة.\n- الموديلات السريعة: {len(df[df['حركة الموديل']=='🔥 سريع جداً'])}"
                link = f"https://wa.me/201012345678?text={urllib.parse.quote(msg)}"
                st.markdown(f"### [✅ اضغط هنا للإرسال عبر واتساب]({link})")

    except Exception as e:
        st.error(f"حدث خطأ في قراءة الملف: {e}")
