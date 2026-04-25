import streamlit as st
import pandas as pd
import io
import urllib.parse

# إعداد الصفحة وتصميم فخم
st.set_page_config(page_title="نظام التخطيط الذكي - برو", layout="wide")

# تصميم الواجهة باللون الأسود والذهبي (فخامة)
st.markdown("""
    <style>
    .main { background-color: #1a1a1a; color: #ffffff; }
    .stMetric { background-color: #262626; padding: 15px; border-radius: 10px; border: 1px solid #d4af37; }
    div[data-testid="stExpander"] { border: 1px solid #d4af37; }
    .stButton>button { 
        background: linear-gradient(45deg, #d4af37, #f9f295); 
        color: black; border: none; font-weight: bold; width: 100%; height: 3em;
    }
    h1, h2, h3 { color: #d4af37 !important; }
    .stDataFrame { border: 1px solid #444; }
    </style>
    """, unsafe_allow_config=True)

st.title("💎 نظام إدارة وتخطيط طلبات الإنتاج الاحترافي")
st.write("التحليل الشامل للمبيعات، الأرصدة، وخطوط الإنتاج")

uploaded_file = st.file_uploader("📂 ارفع ملف الماستر داتا (Excel)", type=['xlsx'])

if uploaded_file:
    try:
        # قراءة البيانات وتنظيفها
        df = pd.read_excel(uploaded_file)
        df.columns = [str(c).strip() for c in df.columns]

        # الأعمدة الأساسية بناءً على ملفك الأخير
        # (طلب الفروع، كمية الفروع، طلب المصنع، كمية المصنع، مسلم عدد، مسلم كمية، مبيعات، رصيد، تحت التشغيل عدد، تحت التشغيل كمية)
        
        cols_to_convert = [
            'طلب الفروع', 'كمية الفروع', 'طلب المصنع', 'كمية المصنع', 
            'مسلم عدد', 'مسلم كمية', 'تحت التشغيل عدد', 'تحت التشغيل كمية', 
            'مبيعات', 'رصيد'
        ]

        # التأكد من وجود الأعمدة وتحويلها لأرقام
        for col in cols_to_convert:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            else:
                st.error(f"❌ العمود '{col}' مفقود في الملف.")
                st.stop()

        # --- المحرك الذكي للقرار ---
        
        # 1. حساب الاحتياج الكلي (فروع + مصنع)
        df['إجمالي الطلب عدد'] = df['طلب الفروع'] + df['طلب المصنع']
        df['إجمالي الطلب كمية'] = df['كمية الفروع'] + df['كمية المصنع']

        # 2. حساب المتوفر (رصيد + اللي لسه بيتعمل في المصنع)
        df['المتوفر فعلياً عدد'] = df['رصيد'] + df['تحت التشغيل عدد']
        df['المتوفر فعلياً كمية'] = df['رصيد'] * (df['كمية المصنع']/df['طلب المصنع'].replace(0,1)) + df['تحت التشغيل كمية']

        # 3. القرار النهائي (الطلب الجديد)
        df['الطلب النهائي (عدد)'] = (df['إجمالي الطلب عدد'] - df['المتوفر فعلياً عدد']).clip(lower=0)
        df['الطلب النهائي (كمية)'] = (df['إجمالي الطلب كمية'] - df['المتوفر فعلياً كمية']).clip(lower=0)

        # 4. معالجة الحالة الخاصة (بايع ومخزونه قليل وغير مطلوب)
        # لو مبيعاته أكتر من الرصيد، والطلب عليه 0 -> السيستم يقترح إنتاج "عينة أمان"
        def get_action_status(row):
            if row['الطلب النهائي (عدد)'] > 0: return "🚀 إنتاج فوراً"
            if row['مبيعات'] > row['رصيد'] and row['طلب الفروع'] == 0: return "⚠️ صنف بايع (توفير مخزون)"
            if row['مبيعات'] == 0 and row['رصيد'] > 5: return "🧊 راكد (توقف)"
            return "✅ مستقر"

        df['حالة القرار'] = df.apply(get_action_status, axis=1)

        # --- شاشة العرض (Dashboard) ---
        st.subheader("📊 ملخص حالة العمليات")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("إجمالي الموديلات", len(df))
        m2.metric("مطلوب إنتاجه (عدد)", int(df['الطلب النهائي (عدد)'].sum()))
        m3.metric("مطلوب إنتاجه (وزن)", f"{df['الطلب النهائي (كمية)'].sum():.2f} جم")
        m4.metric("أصناف تحت التشغيل", int(df['تحت التشغيل عدد'].sum()))

        st.divider()

        # عرض الجدول الاحترافي
        st.subheader("📝 جدول الماستر داتا النهائي")
        view_cols = [
            'SKU', 'ItemName', 'مبيعات', 'رصيد', 'تحت التشغيل عدد', 
            'طلب الفروع', 'الطلب النهائي (عدد)', 'الطلب النهائي (كمية)', 'حالة القرار'
        ]
        
        # تنسيق الجدول بألوان بناءً على الحالة
        st.dataframe(df[view_cols].sort_values(by='الطلب النهائي (عدد)', ascending=False), use_container_width=True)

        # تحميل النتائج
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='خطة الإنتاج النهائية')
        
        st.download_button(
            label="📥 تحميل ملف القرار النهائي (الماستر داتا المحدثة)",
            data=buffer.getvalue(),
            file_name="Final_Production_Order.xlsx",
            mime="application/vnd.ms-excel"
        )

        # واتساب فخم
        st.divider()
        if st.button("📱 إرسال ملخص الطلبية للمكتب الفني"):
            msg = (f"🛑 *تقرير طلبية الإنتاج النهائي*\n"
                   f"--------------------------\n"
                   f"📦 إجمالي العدد المطلوب: {int(df['الطلب النهائي (عدد)'].sum())} قطعة\n"
                   f"⚖️ إجمالي الوزن المطلوب: {df['الطلب النهائي (كمية)'].sum():.2f} جرام\n"
                   f"🏭 أصناف تحت التشغيل: {int(df['تحت التشغيل عدد'].sum())} قطعة\n"
                   f"⚠️ أصناف بايعة بدون طلب: {len(df[df['حالة القرار']=='⚠️ صنف بايع (توفير مخزون)'])}")
            link = f"https://wa.me/201012345678?text={urllib.parse.quote(msg)}"
            st.markdown(f"### [✅ إرسال عبر WhatsApp]({link})")

    except Exception as e:
        st.error(f"حدث خطأ تقني: {e}")
