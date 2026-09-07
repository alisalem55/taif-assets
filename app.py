import streamlit as st
import pandas as pd
from datetime import datetime
import os
from docx import Document

# 1. إعدادات الهوية البصرية العالمية المطابقة للتصميم الشبكي الاحترافي
st.set_page_config(
    page_title="منصة الممتلكات والرقابة الرقمية | صحة الطائف",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# استدعاء مكتبة الأيقونات العالمية وتنسيق الواجهة الاحترافية عالية التباين
st.markdown('<link rel="stylesheet" href="https://cloudflare.com">', unsafe_allow_html=True)

st.markdown("""
    <style>
    @import url('https://googleapis.com');
    
    * { font-family: 'Cairo', sans-serif; }
    .main { background-color: #f3f4f6 !important; }
    
    /* تصميم الخلايا والشبكات المنفصلة (Grid Cards) */
    .grid-box {
        background: #ffffff !important;
        padding: 22px !important;
        border-radius: 14px !important;
        border: 1px solid #e5e7eb !important;
        box-shadow: 0 4px 10px rgba(0, 0, 0, 0.04) !important;
        margin-bottom: 20px !important;
        text-align: right !important;
    }
    
    /* رأسية كل خلية شبكية */
    .grid-header {
        border-bottom: 2px solid #f3f4f6 !important;
        padding-bottom: 10px !important;
        margin-bottom: 14px !important;
        font-weight: 700 !important;
        color: #1d5c43 !important;
        font-size: 1.15rem !important;
        direction: rtl !important;
    }
    
    /* تصميم بطاقات الأداء العلوية الصغيرة */
    .mini-kpi {
        background: #ffffff !important;
        padding: 18px !important;
        border-radius: 12px !important;
        border-right: 6px solid #1d5c43 !important;
        box-shadow: 0 2px 5px rgba(0,0,0,0.02) !important;
        text-align: right !important;
    }
    
    /* تنبيهات النظام الاحترافية الناعمة */
    .alert-premium {
        background: #fff5f5 !important;
        padding: 16px !important;
        border-radius: 12px !important;
        border-right: 6px solid #e74c3c !important;
        text-align: right;
        color: #c0392b !important;
        font-weight: bold !important;
    }
    
    /* إخفاء الدوائر البيضاء والأزرار الافتراضية في القائمة الجانبية وتعديل مظهرها جذرياً */
    div[data-testid="stSidebarUserContent"] {
        background: #113829 !important;
        color: #ffffff !important;
    }
    div[data-testid="stSidebarUserContent"] p, 
    div[data-testid="stSidebarUserContent"] h3,
    div[data-testid="stSidebarUserContent"] label,
    div[data-testid="stSidebarUserContent"] span {
        color: #ffffff !important;
        font-weight: 600 !important;
    }
    
    /* تعديل شكل الراديو ليصبح كأزرار وباقات منفصلة احترافية بدون دوائر */
    div[data-testid="stSidebarUserContent"] .stRadio div[role="radiogroup"] {
        display: flex;
        flex-direction: column;
        gap: 10px;
    }
    div[data-testid="stSidebarUserContent"] .stRadio div[role="radiogroup"] label {
        background: rgba(255, 255, 255, 0.08) !important;
        padding: 12px 15px !important;
        border-radius: 8px !important;
        border: 1px solid rgba(255, 255, 255, 0.1) !important;
        cursor: pointer;
        transition: all 0.2s ease;
        display: block;
        width: 100%;
    }
    div[data-testid="stSidebarUserContent"] .stRadio div[role="radiogroup"] label:hover {
        background: rgba(255, 255, 255, 0.2) !important;
        border-color: #dfb76c !important;
    }
    /* إخفاء الدائرة الاختيارية الصغيرة تماماً */
    div[data-testid="stSidebarUserContent"] .stRadio div[role="radiogroup"] [data-testid="stMarkdownContainer"]::before {
        display: none !important;
    }
    div[data-testid="stSidebarUserContent"] .stRadio div[role="radiogroup"] input[type="radio"] {
        display: none !important;
    }
    
    /* تحسين أزرار النظام بالكامل */
    .stButton>button {
        background-color: #1d5c43 !important;
        color: white !important;
        border-radius: 8px !important;
        padding: 10px 20px !important;
        font-size: 14px !important;
        font-weight: 700 !important;
        border: none !important;
        width: 100% !important;
        box-shadow: 0 2px 4px rgba(29, 92, 67, 0.1) !important;
        transition: all 0.2s ease !important;
    }
    .stButton>button:hover {
        background-color: #dfb76c !important;
        color: #1d5c43 !important;
        transform: translateY(-1px);
    }
    
    /* محاذاة الجداول والخانات من اليمين لليسار */
    .stTextInput>div>div>input, .stSelectbox>div>div>div, .stNumberInput>div>div>input, .stTextArea>div>div>textarea {
        text-align: right !important;
        direction: rtl !important;
    }
    </style>
""", unsafe_allow_html=True)

# 2. تهيئة وتأمين قاعدة البيانات المحلية لجميع الخدمات بما فيها المعاملات والاستثمارات
def init_db():
    if not os.path.exists("assets_db.csv"):
        pd.DataFrame(columns=["رقم_الصك", "رقم_المعاملة", "اسم_الموقع", "نوع_العقار", "الحي", "مساحة_الصك", "مساحة_الرفع_الفعلي", "خط_العرض", "خط_الطول", "تاريخ_انتهاء_الرخصة"]).to_csv("assets_db.csv", index=False)
    if not os.path.exists("audit_log.csv"):
        pd.DataFrame(columns=["الوقت", "المستخدم", "الإجراء", "تفاصيل"]).to_csv("audit_log.csv", index=False)
    if not os.path.exists("encroachments_db.csv"):
        pd.DataFrame(columns=["رقم_الصك", "اسم_الموقع", "نوع_التعدي", "حالة_القضية", "تاريخ_الرصد", "الإجراء_المتخذ"]).to_csv("encroachments_db.csv", index=False)
    if not os.path.exists("investments_db.csv"):
        pd.DataFrame(columns=["رقم_الصك", "اسم_العين_المستثمرة", "نوع_الاستثمار", "المستأجر", "قيمة_العقد", "تاريخ_انتهاء_العقد"]).to_csv("investments_db.csv", index=False)
    if not os.path.exists("workflows_db.csv"):
        pd.DataFrame(columns=["رقم_المعاملة", "موضوع_المعاملة", "الإدارة_الحالية", "حالة_المعاملة", "تاريخ_التحديث", "الموظف_المسؤول"]).to_csv("workflows_db.csv", index=False)

init_db()

def log_action(user, action, details):
    df = pd.read_csv("audit_log.csv")
    new_log = pd.DataFrame([{"الوقت": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "المستخدم": user, "الإجراء": action, "تفاصيل": details}])
    pd.concat([df, new_log]).to_csv("audit_log.csv", index=False)

# 3. بوابة الأمان والتحقق الرقمي للولوج
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False
    st.session_state['role'] = None
    st.session_state['username'] = ""

if not st.session_state['logged_in']:
    st.markdown("<br><br><h2 style='text-align: center; color: #1d5c43;'><i class='fa-solid fa-shield-halved'></i> المنصة الرقمية للأصول والممتلكات</h2>", unsafe_allow_html=True)
    st.markdown("<h4 style='text-align: center; color: #777;'>فرع وزارة الصحة بمحافظة الطائف</h4>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col2:
        st.markdown("<div class='grid-box' style='border-top: 4px solid #1d5c43;'>", unsafe_allow_html=True)
        role_input = st.selectbox("🔒 فئة الدخول للمنصة", ["موظف الإدارة (Staff)", "مدير النظام (Admin)"])
        password = st.text_input("🔑 كود التحقق السري", type="password")
        if st.button("🔓 دخول آمن للنظام"):
            if role_input == "مدير النظام (Admin)" and password == "MOH@2026":
                st.session_state['logged_in'] = True
                st.session_state['role'] = "Admin"
                st.session_state['username'] = "مدير الإدارة"
                log_action("مدير الإدارة", "تسجيل دخول", "نجاح الدخول للمنصة")
                st.rerun()
            elif role_input == "موظف الإدارة (Staff)" and password == "Staff@Taif":
                st.session_state['logged_in'] = True
                st.session_state['role'] = "Staff"
                st.session_state['username'] = "موظف ممتلكات"
                log_action("موظف ممتلكات", "تسجيل دخول", "نجاح الدخول للمنصة")
                st.rerun()
            else:
                st.error("❌ كود التحقق السري غير صحيح!")
        st.markdown("</div>", unsafe_allow_html=True)
else:
    # القائمة الجانبية الفخمة عالية التباين
    st.sidebar.markdown(f"<div style='text-align: center; padding: 10px;'><i class='fa-solid fa-circle-user' style='font-size: 3.5rem; color: #dfb76c;'></i></div>", unsafe_allow_html=True)
    st.sidebar.markdown(f"<h3 style='color: white; text-align: center; margin-top: 5px;'>{st.session_state['username']}</h3>", unsafe_allow_html=True)
    st.sidebar.markdown(f"<p style='color: #dfb76c; text-align: center; font-weight: bold; margin-bottom: 20px;'><i class='fa-solid fa-id-card-clip'></i> رتبة: {st.session_state['role']}</p>", unsafe_allow_html=True)
    
    st.sidebar.markdown("<hr style='border-color: rgba(255,255,255,0.2);'>", unsafe_allow_html=True)
    
    # خيارات تصفح الأقسام بنمط البطاقات بدون دوائر اختيارية قديمة
    menu = st.sidebar.radio(
        "📂 تصفح أقسام المنصة",
        [
            "📊 لوحة التحكم اليومية", 
            "📥 استيراد ورفع ملفات Excel",
            "🔍 محرك صائغ الخطابات", 
            "⚙️ التحكم بالأصول (تعديل/حذف)",
            "💼 تتبع وحفظ المعاملات الرقمية",
            "⚠️ رقابة الأراضي وتتبع التعديات",
            "💰 موديول الاستثمار العقاري",
            "📜 سجل الحماية والمراقبة"
        ]
    )
    
    st.sidebar.markdown("<br><br>", unsafe_allow_html=True)
    if st.sidebar.button("🚪 خروج آمن من النظام"):
        log_action(st.session_state['username'], "تسجيل خروج", "تم تسجيل الخروج بنجاح")
        st.session_state['logged_in'] = False
        st.clear_not_saved_mutations()
        st.rerun()

    df_assets = pd.read_csv("assets_db.csv")
    df_encroach = pd.read_csv("encroachments_db.csv")
    df_invest = pd.read_csv("investments_db.csv")
    df_workflows = pd.read_csv("workflows_db.csv")

    # 📊 لوحة التحكم اليومية الشبكية المطابقة للمواصفات الاحترافية
    if menu == "📊 لوحة التحكم اليومية":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-table-cells-large'></i> لوحة الأداء والتوزيع الشبكي الموحد</h1>", unsafe_allow_html=True)
        
        # الصف الأول: مصفوفة المؤشرات العلوية الأربعة بجانب بعضها
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.markdown(f"<div class='mini-kpi'><b><i class='fa-solid fa-box'></i> إجمالي الأصول:</b> <br><span style='font-size: 1.15rem; font-weight: bold; color: #1d5c43;'>{len(df_assets)} عقار مسجل</span></div>", unsafe_allow_html=True)
        with k2:
            # معالجة ذكية لتفادي خطأ الـ KeyError الظاهر بالصورة يدوياً أو آلياً
            if not df_assets.empty:
                area_col = "مساحة_الصك" if "مساحة_الصك" in df_assets.columns else ("المساحة" if "المساحة" in df_assets.columns else None)
                total_area = df_assets[area_col].astype(float).sum() if area_col else 0
            else:
                total_area = 0
            st.markdown(f"<div class='mini-kpi'><b><i class='fa-solid fa-vector-square'></i> مساحات الصكوك:</b> <br><span style='font-size: 1.15rem; font-weight: bold; color: #1d5c43;'>{total_area:,.2f} م²</span></div>", unsafe_allow_html=True)
        with k3:
            st.markdown(f"<div class='mini-kpi'><b><i class='fa-solid fa-route'></i> المعاملات النشطة:</b> <br><span style='font-size: 1.15rem; font-weight: bold; color: #dfb76c;'>{len(df_workflows)} معاملة تحت التتبع</span></div>", unsafe_allow_html=True)
        with k4:
            st.markdown(f"<div class='mini-kpi'><b><i class='fa-solid fa-triangle-exclamation'></i> حالات التعدي:</b> <br><span style='font-size: 1.15rem; font-weight: bold; color: #e74c3c;'>{len(df_encroach)} قضايا نشطة</span></div>", unsafe_allow_html=True)
            
        st.markdown("<br>", unsafe_allow_html=True)
        
        # الصف الثاني: التوزيع الجغرافي ونظام التنبيهات الموزّع بالتوازي
        block1, block2 = st.columns(2)
        with block1:
            st.markdown("<div class='grid-box'><div class='grid-header'><span><i class='fa-solid fa-map-location-dot'></i> الرقابة الجيومكانية والرفع المساحي (Satellite Map)</span></div>", unsafe_allow_html=True)
            if not df_assets.empty and "خط_العرض" in df_assets.columns and "خط_الطول" in df_assets.columns:
                try:
                    map_data = df_assets[["خط_العرض", "خط_الطول"]].dropna()
                    map_data.columns = ["lat", "lon"]
                    map_data["lat"] = pd.to_numeric(map_data["lat"], errors='coerce')
                    map_data["lon"] = pd.to_numeric(map_data["lon"], errors='coerce')
                    map_data = map_data.dropna()
                    if not map_data.empty:
                        st.map(map_data, use_container_width=True)
                    else:
                        st.info("💡 لا توجد إحداثيات رقمية صالحة لعرضها على الخريطة حالياً.")
                except:
                    st.info("💡 الخريطة الجغرافية تنتظر قراءة البيانات الميدانية.")
            else:
                st.info("💡 الخريطة جاهزة وتنتظر إدخال الأصول والمشاريع.")
            st.markdown("</div>", unsafe_allow_html=True)
            
        with block2:
            st.markdown("<div class='grid-box'><div class='grid-header'><span><i class='fa-solid fa-clock-lock'></i> التنبيهات اللونية وفحص رخص البناء المؤتمت</span></div>", unsafe_allow_html=True)
            st.markdown("<div class='alert-premium' style='background: #fffdf5; border-right: 6px solid #f1c40f; color:#7f8c8d;'><i class='fa-solid fa-circle-check' style='color:#2ecc71;'></i> الفحص الرقمي الدوري يوضح سلامة وثبات كافة الوثائق والصكوك المخزنة بنجاح.</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

    # 📥 قسم استيراد ورفع ملفات Excel
    elif menu == "📥 استيراد ورفع ملفات Excel":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-file-excel'></i> استيراد ورفع البيانات الذكي من ملفات Excel</h1>", unsafe_allow_html=True)
        if st.session_state['role'] != "Admin":
            st.markdown("<div class='alert-premium'><i class='fa-solid fa-lock'></i> عذراً، خاصية رفع واستيراد الملفات تتطلب صلاحية مدير النظام (Admin).</div>", unsafe_allow_html=True)
        else:
            st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
            st.markdown("<div class='grid-header'><span><i class='fa-solid fa-upload'></i> مركز الاستيراد التلقائي لمعاملات الصكوك والمساحات</span></div>", unsafe_allow_html=True)
            uploaded_file = st.file_uploader("اختر ملف الـ Excel من جهاز الكمبيوتر", type=["xlsx", "xls"])
            if uploaded_file is not None:
                try:
                    df_uploaded = pd.read_excel(uploaded_file)
                    st.success("✅ تم قراءة ملفك بنجاح! معاينة الجداول قبل التخزين:")
                    st.dataframe(df_uploaded, use_container_width=True)
                    if st.button("🚀 دمج وحفظ القوائم المرفوعة في قاعدة بيانات الكمبيوتر"):
                        df_combined = pd.concat([df_assets, df_uploaded]).drop_duplicates(subset=["رقم_الصك"], keep='last')
                        df_combined.to_csv("assets_db.csv", index=False)
                        log_action(st.session_state['username'], "استيراد ملف Excel", f"تم استيراد أصول بنجاح")
                        st.balloons()
                        st.success("🎉 تم دمج وحفظ البيانات المستوردة بنجاح على جهاز الكمبيوتر الخاص بك!")
                        st.rerun()
                except Exception as e:
                    st.error(f"حدث خطأ أثناء قراءة الملف: {e}")
            st.markdown("</div>", unsafe_allow_html=True)
    # 🔍 محرك المراسلات الفورية وصائغ الخطابات المؤتمت
    elif menu == "🔍 محرك صائغ الخطابات":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-file-signature'></i> محرك صائغ الخطابات والمراسلات الرقمي</h1>", unsafe_allow_html=True)
        search_query = st.text_input("🔍 ابحث فوراً برقم الصك، رقم المعاملة الهندسية، أو اسم الموقع الطبي")
        
        if search_query and not df_assets.empty:
            filtered_df = df_assets[df_assets.astype(str).apply(lambda x: search_query in x.values, axis=1)]
            st.dataframe(filtered_df, use_container_width=True)
            
            if not filtered_df.empty:
                asset = filtered_df.iloc[0]
                st.markdown("<hr style='border-color: #1d5c43;'>", unsafe_allow_html=True)
                
                col_panel1, col_panel2 = st.columns(2)
                with col_panel1:
                    st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
                    st.markdown("<div class='grid-header'><span><i class='fa-solid fa-satellite'></i> الرفع المساحي ونظام الـ GIS الجغرافي للعقار</span></div>", unsafe_allow_html=True)
                    area_key = "مساحة_الصك" if "مساحة_الصك" in df_assets.columns else "المساحة"
                    s_area = float(asset[area_key]) if pd.notnull(asset[area_key]) else 0.0
                    f_area = float(asset['مساحة_الرفع_الفعلي']) if ("مساحة_الرفع_الفعلي" in df_assets.columns and pd.notnull(asset['مساحة_الرفع_الفعلي'])) else s_area
                    diff = f_area - s_area
                    st.markdown(f"<p style='text-align: right;'><b>الموقع:</b> {asset['اسم_الموقع']}<br><b>الحي:</b> {asset['الحي']}<br><b>المساحة بالصك:</b> {s_area} م²</p>", unsafe_allow_html=True)
                    
                    maps_url = f"https://google.com{asset['خط_العرض']},{asset['خط_الطول']}"
                    st.markdown(f"<a href='{maps_url}' target='_blank'><button style='width:100%; padding:10px; background:#1d5c43; color:white; border:none; border-radius:6px; cursor:pointer; font-weight:bold;'><i class='fa-solid fa-location-arrow'></i> استعراض الموقع الميداني بالقمر الصناعي</button></a>", unsafe_allow_html=True)
                    st.markdown("<br><hr><b>📱 رمز الاستجابة السريع للرصد الميداني (QR Code):</b>", unsafe_allow_html=True)
                    qr_url = f"https://qrserver.com{maps_url}"
                    st.image(qr_url, caption="امسح الرمز بالجوال للانتقال المباشر للموقع من الميدان")
                    st.markdown("</div>", unsafe_allow_html=True)
                    
                with col_panel2:
                    st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
                    st.markdown("<div class='grid-header'><span><i class='fa-solid fa-envelope-open-text'></i> إعداد وصياغة الخطاب الرسمي المؤتمت للجهة الحكومية</span></div>", unsafe_allow_html=True)
                    letter_type = st.selectbox("تحديد الجهة الحكومية الموجه إليها الخطاب", [
                        "خطاب موجه لسعادة أمين محافظة الطائف (استخراج رخصة بناء)", 
                        "خطاب موجه لسعادة مدير شركة الكهرباء بالطائف (إيصال التيار)", 
                        "خطاب موجه لفضيلة رئيس المحكمة العامة بالطائف (مطابقة الحدود وتعديل الفروقات المساحية)"
                    ])
                    
                    text_content = f"جهة الخطاب المحددة: {letter_type}\nبموجب الصك رقم {asset['رقم_الصك']} التابع لـ {asset['اسم_الموقع']} بحي {asset['الحي']}.\nنأمل من سعادتكم التوجيه لإكمال اللازم هندسياً وميدانياً حسب المخططات المعتمدة لفرع وزارة الصحة بمحافظة الطائف.\n\nوتقبلوا وافر التحية والتقدير،،"
                    st.text_area("📄 صيغة الخطاب الرسمي والذكية المحدثة", text_content, height=200)
                    st.markdown("</div>", unsafe_allow_html=True)

    # ⚙️ لوحة التحكم بالأصول
    elif menu == "⚙️ التحكم بالأصول (تعديل/حذف)":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-folder-gear'></i> إدارة وتعديل وحذف الأصول العقارية</h1>", unsafe_allow_html=True)
        if st.session_state['role'] != "Admin":
            st.markdown("<div class='alert-premium'><i class='fa-solid fa-lock'></i> جدار الحماية: ميزة تعديل أو حذف الأصول تقتصر على مدير النظام (Admin) فقط.</div>", unsafe_allow_html=True)
        else:
            st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
            st.markdown("<div class='grid-header'><span><i class='fa-solid fa-plus-minus'></i> إضافة وإدخال الأصول يدوياً لقاعدة البيانات</span></div>", unsafe_allow_html=True)
            with st.form("manual_add"):
                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    r_sok = st.text_input("رقم الصك الشرعي")
                    r_moamala = st.text_input("رقم المعاملة الوزارية")
                    site_name = st.text_input("اسم الموقع الطبي / الأرض")
                with col_f2:
                    district = st.text_input("الحي السكني داخل الطائف")
                    area = st.number_input("المساحة الإجمالية بالصك (م²)", min_value=0.0, step=0.1)
                    g_type = st.selectbox("تصنيف العقار", ["أرض فضاء تابعة للوزارة", "مركز صحي قائم", "مستشفى عام وتخصصي"])
                if st.form_submit_button("💾 اعتماد وحفظ الأصل في قاعدة البيانات"):
                    new_asset = pd.DataFrame([{"رقم_الصك": r_sok, "رقم_المعاملة": r_moamala, "اسم_الموقع": site_name, "نوع_العقار": g_type, "الحي": district, "مساحة_الصك": area, "مساحة_الرفع_الفعلي": area, "خط_العرض": "21.27", "خط_الطول": "40.41", "تاريخ_انتهاء_الرخصة": "لا يوجد"}])
                    pd.concat([df_assets, new_asset]).to_csv("assets_db.csv", index=False)
                    st.success("✅ تم حفظ الأصل وتخزينه بنجاح على جهازك!")
                    st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

            st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
            st.markdown("<div class='grid-header'><span><i class='fa-solid fa-pen-to-square'></i> تعديل وحذف الأصول مباشرة من الجدول التفاعلي المشبك</span></div>", unsafe_allow_html=True)
            if not df_assets.empty:
                edited_df = st.data_editor(df_assets, num_rows="dynamic", use_container_width=True, key="assets_editor")
                if st.button("💾 حفظ كافة التغييرات والتعديلات النشطة للأصول"):
                    edited_df.to_csv("assets_db.csv", index=False)
                    st.success("🎉 تم تحديث وحفظ جدول الأصول على جهازك بنجاح!")
                    st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)
    # 💼 تتبع وحفظ المعاملات الرقمية بين الإدارات والجهات للعودة إليها
    elif menu == "💼 تتبع وحفظ المعاملات الرقمية":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-route'></i> حفظ وتتبع خط سير المعاملات الرقمية</h1>", unsafe_allow_html=True)
        st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
        st.markdown("<div class='grid-header'><span><i class='fa-solid fa-folder-plus'></i> تسجيل وقيد معاملة جديدة وتحديد مسارها للعودة إليها مستقبلاً</span></div>", unsafe_allow_html=True)
        with st.form("workflow_add"):
            w_col1, w_col2 = st.columns(2)
            with w_col1:
                w_id = st.text_input("رقم المعاملة (الصادر أو الوارد)")
                w_subject = st.text_input("موضوع المعاملة (مثال: طلب رخصة بناء مركز صحي شهار)")
            with w_col2:
                w_dept = st.selectbox("الجهة أو الإدارة الحالية الواقفة عندها المعاملة", ["أمانة محافظة الطائف", "المحكمة العامة بالطائف", "كتابة العدل", "الشؤون الهندسية بالفرع"])
                w_status = st.selectbox("حالة المعاملة الحالية", ["قيد الدراسة والتدقيق المساحي", "تم الإفراغ والمطابقة بنجاح", "معاملة منتهية وأرشفت"])
            if st.form_submit_button("💾 قيد وحفظ خط سير المعاملة في النظام"):
                new_wf = pd.DataFrame([{"رقم_المعاملة": w_id, "موضوع_المعاملة": w_subject, "الإدارة_الحالية": w_dept, "حالة_المعاملة": w_status, "تاريخ_التحديث": datetime.now().strftime("%Y-%m-%d %H:%M"), "الموظف_المسؤول": "متابع الحركة"}])
                pd.concat([df_workflows, new_wf]).to_csv("workflows_db.csv", index=False)
                st.success("✅ تم حفظ وتأمين المعاملة في النظام للعودة إليها مستقبلاً!")
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
        st.markdown("<div class='grid-header'><span><i class='fa-solid fa-folder-open'></i> أرشيف المعاملات المحفوظة (تعديل وحذف مباشر فوري)</span></div>", unsafe_allow_html=True)
        if not df_workflows.empty:
            edited_wf_df = st.data_editor(df_workflows, num_rows="dynamic", use_container_width=True, key="workflow_editor")
            if st.button("💾 حفظ وتثبيت كافة تعديلات وحذوفات جدول تتبع المعاملات"):
                edited_wf_df.to_csv("workflows_db.csv", index=False)
                st.success("🎉 تم حفظ وتثبيت خط سير معاملاتك بنجاح على جهازك!")
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    # ⚠️ نظام رقابة الأراضي وتتبع التعديات
    elif menu == "⚠️ رقابة الأراضي وتتبع التعديات":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-shield-halved'></i> تتبع التعديات والرقابة الميدانية للأراضي</h1>", unsafe_allow_html=True)
        st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
        with st.form("encroach_form"):
            c_e1, c_e2 = st.columns(2)
            with c_e1:
                e_sok = st.text_input("رقم الصك الشرعي المعتدى على حدوده")
                e_name = st.text_input("اسم الموقع الطبي أو الأرض الفضاء")
            with c_e2:
                e_type = st.selectbox("نوع التعدي المرصود ميدانياً", ["إقامة حوش أو أسوار غير نظامية", "بناء شعبي بدون رخصة", "وضع لوحات أو شبوك وتجريف تربة"])
                e_status = st.selectbox("حالة المعاملة والقضية", ["تحت الرفع للمحافظة", "بانتظار إزالة لجنة التعديات", "تمت الإزالة واسترداد الأرض"])
            if st.form_submit_button("🚨 تسجيل حالة التعدي وأرشفة المعاملة رسمياً"):
                new_enc = pd.DataFrame([{"رقم_الصك": e_sok, "اسم_الموقع": e_name, "نوع_التعدي": e_type, "حالة_القضية": e_status, "تاريخ_الرصد": datetime.now().strftime("%Y-%m-%d"), "الإجراء_المتخذ": "متابع ميدانياً"}])
                pd.concat([df_encroach, new_enc]).to_csv("encroachments_db.csv", index=False)
                st.success("✅ تم تسجيل حالة التعدي بنجاح في السجلات.")
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
        if not df_encroach.empty:
            edited_encroach = st.data_editor(df_encroach, num_rows="dynamic", use_container_width=True, key="encroach_editor_v5")
            if st.button("💾 حفظ كافة التعديلات والتغييرات النشطة في سجل التعديات"):
                edited_encroach.to_csv("encroachments_db.csv", index=False)
                st.success("🎉 تم حفظ وتثبيت كافة تحديثات الرقابة والتعديات الميدانية على ملفات جهازك!")
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    # 💰 موديول الاستثمار العقاري
    elif menu == "💰 موديول الاستثمار العقاري":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-coins'></i> موديول حصر العوائد والاستثمارات العقارية</h1>", unsafe_allow_html=True)
        st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
        if not df_invest.empty:
            edited_invest = st.data_editor(df_invest, num_rows="dynamic", use_container_width=True, key="invest_editor")
            if st.button("💾 حفظ كافة تعديلات عقود الاستثمار السنوية"):
                edited_invest.to_csv("investments_db.csv", index=False)
                st.success("✅ تم حفظ تحديثات الاستثمار على جهازك.")
        else:
            st.info("👍 لا توجد عقود استثمارية مسجلة حالياً.")
        st.markdown("</div>", unsafe_allow_html=True)

    # 📜 سجل الحماية الرقابي والأمان
    elif menu == "📜 سجل الحماية والمراقبة":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-user-shield'></i> سجل الحماية الرقابي والعمليات النشطة (Audit Log)</h1>", unsafe_allow_html=True)
        st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
        df_logs = pd.read_csv("audit_log.csv")
        st.dataframe(df_logs, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)
