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
    
    /* تعديل جذري لتصميم القائمة الجانبية للتخلص من الدوائر والتداخل والخطوط الباهتة */
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
    
    /* تحويل الراديو الجانبي إلى بطاقات أزرار منفصلة جذابة بدون أي دوائر خيارات قديمة */
    div[data-testid="stSidebarUserContent"] .stRadio div[role="radiogroup"] {
        display: flex !important;
        flex-direction: column !important;
        gap: 12px !important;
    }
    div[data-testid="stSidebarUserContent"] .stRadio div[role="radiogroup"] label {
        background: rgba(255, 255, 255, 0.07) !important;
        padding: 14px 18px !important;
        border-radius: 10px !important;
        border: 1px solid rgba(255, 255, 255, 0.12) !important;
        cursor: pointer !important;
        transition: all 0.25s ease !important;
        display: block !important;
        width: 100% !important;
    }
    div[data-testid="stSidebarUserContent"] .stRadio div[role="radiogroup"] label:hover {
        background: rgba(255, 255, 255, 0.2) !important;
        border-color: #dfb76c !important;
        transform: translateX(-3px);
    }
    /* إلغاء الدائرة الافتراضية وحجبها بالكامل */
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

# 2. تهيئة وتأمين قاعدة البيانات المحلية لفرع الوزارة بالطائف
def init_db():
    required_cols = ["م", "المنشأة", "نوع_العقار", "حاله_العقار", "المحافظة", "مركز/حي/قرية", "المساحة", "رقم الصك", "تاريخ الصك", "خط الطول", "دائرة العرض"]
    if not os.path.exists("assets_db.csv"):
        pd.DataFrame(columns=required_cols).to_csv("assets_db.csv", index=False)
    if not os.path.exists("audit_log.csv"):
        pd.DataFrame(columns=["الوقت", "المستخدم", "الإجراء", "تفاصيل"]).to_csv("audit_log.csv", index=False)
    if not os.path.exists("encroachments_db.csv"):
        pd.DataFrame(columns=["رقم الصك", "المنشأة", "نوع_التعدي", "حالة_القضية", "تاريخ_الرصد", "الإجراء_المتخذ"]).to_csv("encroachments_db.csv", index=False)
    if not os.path.exists("investments_db.csv"):
        pd.DataFrame(columns=["رقم الصك", "اسم_العين_المستثمرة", "نوع_الاستثمار", "المستأجر", "قيمة_العقد", "تاريخ_انتهاء_العقد"]).to_csv("investments_db.csv", index=False)
    if not os.path.exists("workflows_db.csv"):
        pd.DataFrame(columns=["رقم_المعاملة", "موضوع_المعاملة", "الإدارة_الحالية", "حالة_المعاملة", "تاريخ_التحديث", "الموظف_المسؤول"]).to_csv("workflows_db.csv", index=False)

init_db()

def log_action(user, action, details):
    df = pd.read_csv("audit_log.csv")
    new_log = pd.DataFrame([{"الوقت": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "المستخدم": user, "الإجراء": action, "تفاصيل": details}])
    pd.concat([df, new_log]).to_csv("audit_log.csv", index=False)

# 3. بوابة الأمان والتحقق الرقمي للولوج للمنصة
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
    # القائمة الجانبية فائقة التباين والوضوح باللون الذهبي والأبيض الناصع
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
            "🔍 ملفات الأملاك والخطابات وبطاقات الوثائق المصورة", 
            "⚙️ التحكم بالأصول (تعديل/حذف)",
            "💼 تتبع وحفظ المعاملات الرقمية",
            "⚠️ رقابة الأراضي وتتبع التعديات",
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
    df_workflows = pd.read_csv("workflows_db.csv")

    # 📊 لوحة التحكم اليومية لـ 138 منشأة التابعة لصحة الطائف
    if menu == "📊 لوحة التحكم اليومية":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-table-cells-large'></i> لوحة الأداء والتوزيع الشبكي الموحد لممتلكات الصحة</h1>", unsafe_allow_html=True)
        
        # الصف الأول: مصفوفة المؤشرات العلوية الأربعة بجانب بعضها
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.markdown(f"<div class='mini-kpi'><b><i class='fa-solid fa-hospital'></i> إجمالي ممتلكات الصحة:</b> <br><span style='font-size: 1.15rem; font-weight: bold; color: #1d5c43;'>{len(df_assets)} منشأة وموقع</span></div>", unsafe_allow_html=True)
        with k2:
            if not df_assets.empty:
                area_col = "المساحة" if "المساحة" in df_assets.columns else ("مساحة_الصك" if "مساحة_الصك" in df_assets.columns else None)
                # تنظيف المساحات النصية وحسابها بدقة من الملف الموثق
                df_assets['clean_area'] = pd.to_numeric(df_assets[area_col].astype(str).str.replace(',', '').str.strip(), errors='coerce') if area_col else 0
                total_area = df_assets['clean_area'].sum()
            else:
                total_area = 0
            st.markdown(f"<div class='mini-kpi'><b><i class='fa-solid fa-vector-square'></i> إجمالي مساحات الأراضي:</b> <br><span style='font-size: 1.15rem; font-weight: bold; color: #1d5c43;'>{total_area:,.2f} م²</span></div>", unsafe_allow_html=True)
        with k3:
            st.markdown(f"<div class='mini-kpi'><b><i class='fa-solid fa-route'></i> المعاملات والخطابات:</b> <br><span style='font-size: 1.15rem; font-weight: bold; color: #dfb76c;'>{len(df_workflows)} قيد المتابعة</span></div>", unsafe_allow_html=True)
        with k4:
            st.markdown(f"<div class='mini-kpi'><b><i class='fa-solid fa-shield-halved'></i> التعديات المرصودة:</b> <br><span style='font-size: 1.15rem; font-weight: bold; color: #e74c3c;'>{len(df_encroach)} قضايا نشطة</span></div>", unsafe_allow_html=True)
            
        st.markdown("<br>", unsafe_allow_html=True)
        
        # الصف الثاني: التوزيع الجغرافي ونظام التنبيهات الموزّع بالتوازي (الشبكة الذكية المتطورة)
        block1, block2 = st.columns(2)
        with block1:
            st.markdown("<div class='grid-box'><div class='grid-header'><span><i class='fa-solid fa-map-location-dot'></i> الرقابة الجيومكانية والخرائط الميدانية الشاملة</span></div>", unsafe_allow_html=True)
            # استدعاء دوائر العرض وخطوط الطول الحقيقية من ملفك الموثق لعرض الـ 138 موقعاً جغرافياً
            if not df_assets.empty and "خط الطول" in df_assets.columns and "دائرة العرض" in df_assets.columns:
                try:
                    map_data = df_assets[["دائرة العرض", "خط الطول"]].dropna()
                    map_data.columns = ["lat", "lon"]
                    map_data["lat"] = pd.to_numeric(map_data["lat"], errors='coerce')
                    map_data["lon"] = pd.to_numeric(map_data["lon"], errors='coerce')
                    map_data = map_data.dropna()
                    if not map_data.empty:
                        st.map(map_data, use_container_width=True)
                    else:
                        st.info("💡 لا توجد إحداثيات رقمية صالحة لعرضها جغرافياً حالياً.")
                except:
                    st.info("💡 نظام الخرائط بانتظار تحديث الإحداثيات السليمة.")
            else:
                st.info("💡 الخريطة الرقمية جاهزة ومستعدة وتنتظر استيراد الأصول.")
            st.markdown("</div>", unsafe_allow_html=True)
            
        with block2:
            st.markdown("<div class='grid-box'><div class='grid-header'><span><i class='fa-solid fa-clock-lock'></i> التنبيهات الذكية وفحص النواقص العقارية لفرع الوزارة</span></div>", unsafe_allow_html=True)
            if not df_assets.empty:
                # تتبع وحصر العقارات التي لا تمتلك رقم صك أو مسجلة بدون صك آلياً من ملفك
                no_sok_count = len(df_assets[df_assets['رقم الصك'].astype(str).str.contains('لا يوجد|NULL|بدون', na=True, case=False)])
                st.markdown(f"<div class='alert-premium' style='background: #fffdf5; border-right: 6px solid #e67e22; color:#d35400;'><i class='fa-solid fa-triangle-exclamation'></i> نظام الفحص الآلي المطور رصد وجود {no_sok_count} منشأة وموقع صحي مدرج بحالة 'بدون صك' أو بيانات غير مكتملة تتطلب تحديث ومخاطبة أمانة الطائف بشكل عاجل.</div>", unsafe_allow_html=True)
            else:
                st.markdown("<div class='alert-premium'><i class='fa-solid fa-circle-check' style='color:#2ecc71;'></i> النظام مستقر وقاعدة البيانات فارغة تماماً وبانتظار رفع ملف الاستيراد لتوليد التنبيهات.</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

    # 📥 قسم استيراد ورفع ملفات Excel دفعة واحدة وقراءة الـ 138 منشأة كاملة
    elif menu == "📥 استيراد ورفع ملفات Excel":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-file-excel'></i> استيراد ورفع القوائم من ملفات Excel</h1>", unsafe_allow_html=True)
        if st.session_state['role'] != "Admin":
            st.markdown("<div class='alert-premium'><i class='fa-solid fa-lock'></i> عذراً، خاصية رفع واستيراد الملفات تتطلب صلاحية مدير النظام (Admin).</div>", unsafe_allow_html=True)
        else:
            st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
            st.markdown("<div class='grid-header'><span><i class='fa-solid fa-upload'></i> مركز القراءة الآلي لبيانات ممتلكات الصحة وتوليد قاعدة بيانات الكمبيوتر</span></div>", unsafe_allow_html=True)
            uploaded_file = st.file_uploader("اختر ملف ممتلكات وزارة الصحة الموثق بالطائف من جهازك", type=["xlsx", "xls", "csv"])
            if uploaded_file is not None:
                try:
                    df_uploaded = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
                    st.success(f"✅ تم قراءة ملفك بنجاح! تم رصد وعزل {len(df_uploaded)} منشأة عقارية وطبية.")
                    st.dataframe(df_uploaded.head(10), use_container_width=True)
                    if st.button("🚀 دمج وحفظ الـ 138 منشأة بالكامل وتخزينها دائمًا على جهازك"):
                        df_uploaded.to_csv("assets_db.csv", index=False)
                        log_action(st.session_state['username'], "استيراد ملف الأصول", f"تم استيراد بيان الممتلكات لـ {len(df_uploaded)} منشأة")
                        st.balloons()
                        st.success("🎉 تهانينا! تم حفظ وتوثيق كامل ملف الممتلكات الحقيقي بنجاح على جهاز الكمبيوتر الخاص بك وتنشيط الخرائط والمراسلات!")
                        st.rerun()
                except Exception as e:
                    st.error(f"حدث خطأ أثناء قراءة هيكلية ملف الـ Excel: {e}")
            st.markdown("</div>", unsafe_allow_html=True)
    # 🔍 محرك المراسلات وبطاقات الوثائق المصورة الأربعة لكل أرض (القسم المطلوب بدقة)
    elif menu == "🔍 ملفات الأملاك والخطابات وبطاقات الوثائق المصورة":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-folder-open'></i> ملفات الأصول وبطاقات الوثائق المصورة والخطابات</h1>", unsafe_allow_html=True)
        
        if df_assets.empty:
            st.info("💡 قاعدة البيانات فارغة حالياً، يرجى التوجه لقسم استيراد ملفات Excel أولاً لرفع الـ 138 منشأة وصك.")
        else:
            # حل مشكلة القوائم المنسدلة عبر برمجة واجهة اختيار سلسة وواضحة جداً وعزلها لتفادي التداخل
            st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
            st.markdown("<div class='grid-header'><span><i class='fa-solid fa-hospital-user'></i> اختر المنشأة أو الأرض من القائمة المستدعاة من ملفك</span></div>", unsafe_allow_html=True)
            
            # استخراج قائمة المنشآت الحقيقية الموثقة لفرع الوزارة
            facility_list = df_assets['المنشأة'].dropna().unique().tolist()
            selected_facility = st.selectbox("🏥 اضغط هنا لاختيار الموقع الطبي / الأرض الفضاء المستهدفة من القائمة المنسدلة:", facility_list, key="facility_dropdown_clean")
            st.markdown("</div>", unsafe_allow_html=True)
            
            if selected_facility:
                # جلب بيانات الأرض أو المستشفى المحدد بالكامل من ملف جهازك
                asset_data = df_assets[df_assets['المنشأة'] == selected_facility].iloc[0]
                
                st.markdown(f"<h2><i class='fa-solid fa-file-invoice-dollar'></i> الملف الفني العقاري لـ: {selected_facility}</h2>", unsafe_allow_html=True)
                
                # 🖼️ لوحة عرض بطاقات الوثائق المصورة الأربعة المنسقة شبكياً بالتوازي (Grid View) مطابق لطلبك
                row1_col1, row1_col2 = st.columns(2)
                row2_col1, row2_col2 = st.columns(2)
                
                # الخانة 1: صورة الصك الشرعي الموثق للعقار
                with row1_col1:
                    st.markdown("<div class='grid-box' style='border-top: 4px solid #dfb76c;'>", unsafe_allow_html=True)
                    st.markdown("<div class='grid-header'><span><i class='fa-solid fa-scroll'></i> 1. صـورة الصك الشرعي الموثق للعقار</span></div>", unsafe_allow_html=True)
                    sok_img = st.file_uploader(f"رفع/تحديث صورة الصك لـ {selected_facility}", type=["jpg", "png", "jpeg"], key=f"sok_up_{asset_data['رقم الصك']}")
                    if sok_img is not None:
                        st.image(sok_img, caption="معاينة صورة الصك الشرعي المعتمد برمجياً", use_container_width=True)
                    else:
                        st.info("📷 لم يتم رفع صورة الصك لهذا الموقع بعد، يمكنك سحب وإفلات صورة الملف هنا لتخزينها.")
                    st.markdown("</div>", unsafe_allow_html=True)
                    
                # الخانة 2: صورة رخصة البناء الهندسية المعمارية
                with row1_col2:
                    st.markdown("<div class='grid-box' style='border-top: 4px solid #1d5c43;'>", unsafe_allow_html=True)
                    st.markdown("<div class='grid-header'><span><i class='fa-solid fa-trowel-bricks'></i> 2. صـورة رخصة البناء الهندسية المعمارية</span></div>", unsafe_allow_html=True)
                    permit_img = st.file_uploader(f"رفع/تحديث صورة الرخصة لـ {selected_facility}", type=["jpg", "png", "jpeg"], key=f"permit_up_{asset_data['رقم الصك']}")
                    if permit_img is not None:
                        st.image(permit_img, caption="معاينة رخصة البناء الهندسية الصادرة"، use_container_width=True)
                    else:
                        st.info("📷 لم يتم رفع صورة رخصة البناء الهندسية للموقع بعد.")
                    st.markdown("</div>", unsafe_allow_html=True)
                    
                # الخانة 3: صورة الموقع الميداني الفعلي والرفع المساحي الجغرافي (Satellite)
                with row2_col1:
                    st.markdown("<div class='grid-box' style='border-top: 4px solid #58a6ff;'>", unsafe_allow_html=True)
                    st.markdown("<div class='grid-header'><span><i class='fa-solid fa-map-location-dot'></i> 3. صورة الرفع المساحي والموقع الميداني (رؤية الأقمار الصناعية)</span></div>", unsafe_allow_html=True)
                    site_img = st.file_uploader(f"رفع/تحديث المخطط الجغرافي لـ {selected_facility}", type=["jpg", "png", "jpeg"], key=f"site_up_{asset_data['رقم الصك']}")
                    if site_img is not None:
                        st.image(site_img, caption="معاينة صورة كروكي الموقع والرفع المساحي الفعلي"، use_container_width=True)
                    else:
                        # استدعاء الرابط المباشر من ملف بايثون لخرائط جوجل لعرض الموقع بشكل فوري إذا لم ترفع صورة
                        maps_url = f"https://google.com{asset_data['دائرة العرض']},{asset_data['خط الطول']}"
                        st.markdown(f"<a href='{maps_url}' target='_blank'><button style='width:100%; padding:10px; background:#1d5c43; color:white; border:none; border-radius:6px; cursor:pointer; font-weight:bold;'><i class='fa-solid fa-location-arrow'></i> استعراض الموقع جغرافياً على Google Maps</button></a>", unsafe_allow_html=True)
                    st.markdown("</div>", unsafe_allow_html=True)
                    
                # الخانة 4: صورة قرار الذرعة والقرار المساحي المعتمد من أمانة الطائف
                with row2_col2:
                    st.markdown("<div class='grid-box' style='border-top: 4px solid #143f2e;'>", unsafe_allow_html=True)
                    st.markdown("<div class='grid-header'><span><i class='fa-solid fa-receipt'></i> 4. صورة قـرار الذرعة والقرار المساحي المعتمد</span></div>", unsafe_allow_html=True)
                    zar_img = st.file_uploader(f"رفع/تحديث صورة قرار الذرعة لـ {selected_facility}", type=["jpg", "png", "jpeg"], key=f"zar_up_{asset_data['رقم الصك']}")
                    if zar_img is not None:
                        st.image(zar_img, caption="معاينة قرار الذرعة المعتمد والمطابق المساحي من الأمانة"، use_container_width=True)
                    else:
                        st.info("📷 لم يتم رفع صورة قرار الذرعة المعتمد من أمانة محافظة الطائف بعد.")
                    st.markdown("</div>", unsafe_allow_html=True)

                # 📝 محرك صائغ الخطابات التلقائي المبني أسفل بطاقات الصور
                st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
                st.markdown("<div class='grid-header'><span><i class='fa-solid fa-envelope-open-text'></i> محرك أتمتة الصياغة الفورية للخطابات الموجهة للجهات الحكومية</span></div>", unsafe_allow_html=True)
                
                letter_target = st.selectbox("✉️ حدد الجهة الحكومية الموجه إليها الخطاب برمجياً:", [
                    "خطاب موجه لسعادة أمين محافظة الطائف (استخراج رخصة بناء ومطابقة ذرعة)", 
                    "خطاب موجه لسعادة مدير شركة الكهرباء بالطائف (طلب إيصال التيار وتحديد المحول)", 
                    "خطاب موجه لفضيلة رئيس المحكمة العامة بالطائف (تحديث ومطابقة صك ممتلكات الصحة)"
                ], key="letter_target_dropdown")
                
                # صياغة الخطاب بشكل آلي بالكامل بناءً على بيانات الملف الموثق للطائف
                text_content = ""
                sok_num = asset_data['رقم الصك'] if pd.notnull(asset_data['رقم الصك']) else "لا يوجد"
                sok_date = asset_data['تاريخ الصك'] if pd.notnull(asset_data['تاريخ الصك']) else "لا يوجد"
                facility_area = asset_data['المساحة'] if pd.notnull(asset_data['المساحة']) else "غير محدد"
                facility_village = asset_data['مركز/حي/قرية'] if pd.notnull(asset_data['مركز/حي/قرية']) else "محافظة الطائف"
                
                if "أمين" in letter_target:
                    text_content = f"سعادة أمين محافظة الطائف\nالسلام عليكم ورحمة الله وبركاته،،\n\nتفيدكم فرع وزارة الصحة بمحافظة الطائف علماً بملكية الوزارة الرسمية للموقع المخصص لـ ({selected_facility}) والواقع بنطاق ({facility_village}) بموجب الصك الشرعي المقيد برقم ({sok_num}) وتاريخ ({sok_date}) بمساحة إجمالية قدرها ({facility_area} م²). ونظراً للمصلحة العامة وبدء التجهيزات الهندسية بالموقع، نأمل من سعادتكم التوجيه لمن يلزم لاعتماد الرفع المساحي وقرار الذرعة واستخراج رخصة بناء وفق الإحداثيات الجغرافية المعتمدة المرفقة ({asset_data['دائرة العرض']} , {asset_data['خط الطول']}).\n\nوتقبلوا خالص التحية والتقدير،،\n\nمدير فرع وزارة الصحة بمحافظة الطائف"
                elif "الكهرباء" in letter_target:
                    text_content = f"سعادة مدير شركة الكهرباء بمحافظة الطائف\nالسلام عليكم ورحمة الله وبركاته،،\n\nنظراً لجهوزية البدء الإنشائي والتشغيلي للموقع الطبي التابع للوزارة ({selected_facility}) والواقع في ({facility_village}) والمقام على الأرض ذات الصك رقم ({sok_num})، نأمل منكم الإيعاز للمختصين لطلب إيصال التيار الكهربائي وتحديد موقع محول الطاقة الفرعي حسب المخططات الهندسية المعمارية المعتمدة.\n\nوتقبلوا وافر التحية والتقدير،،\n\nمساعد مدير فرع الوزارة للشؤون الهندسية والمشاريع"
                else:
                    text_content = f"فضيلة رئيس المحكمة العامة بمحافظة الطائف\nالسلام عليكم ورحمة الله وبركاته،،\n\nتتقدم فرع وزارة الصحة بمحافظة الطائف بطلب تحديث وإصدار صك إلكتروني موحد ومطابقة مسحية للصك رقم ({sok_num}) وتاريخ ({sok_date}) العائد لملك الوزارة في موقع ({selected_facility}) بنطاق ({facility_village})، وذلك لمطابقة وتعديل الحدود الإنشائية والمساحية حسب أنظمة الأصول الحديثة للوزارة والرفع المساحي المرفق.\n\nوتقبلوا خالص التحية والتقدير،،"
                
                st.text_area("📄 صيغة الخطاب الرسمي والذكية المولد تلقائياً لفرع الوزارة:", text_content, height=200)
                
                doc = Document()
                doc.add_heading(letter_target, 0)
                doc.add_paragraph(text_content)
                doc.save("generated_letter.docx")
                with open("generated_letter.docx", "rb") as f:
                    st.download_button("📥 تنزيل الخطاب الآن كملف Word رسمي مجهز بالكامل للطباعة", f, file_name=f"خطاب_صحة_الطائف_{sok_num}.docx")
                st.markdown("</div>", unsafe_allow_html=True)
    # ⚙️ لوحة التحكم بالأصول (إدخال وتعديل وحذف يدوي بالكامل)
    elif menu == "⚙️ التحكم بالأصول (تعديل/حذف)":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-folder-gear'></i> إدارة وتعديل وحذف الأصول العقارية</h1>", unsafe_allow_html=True)
        
        if st.session_state['role'] != "Admin":
            st.markdown("<div class='alert-premium'><i class='fa-solid fa-lock'></i> جدار الحماية: ميزة حفظ أو تعديل أو حذف الأصول تقتصر على مدير النظام (Admin) فقط.</div>", unsafe_allow_html=True)
        else:
            st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
            st.markdown("<div class='grid-header'><span><i class='fa-solid fa-plus-minus'></i> إضافة وإدخال عقار جديد يدوياً لقاعدة بيانات جهازك</span></div>", unsafe_allow_html=True)
            with st.form("manual_add"):
                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    r_sok = st.text_input("رقم الصك الشرعي الجديد")
                    r_moamala = st.text_input("رقم المعاملة أو القيد")
                    site_name = st.text_input("اسم الموقع الطبي / الأرض")
                with col_f2:
                    district = st.text_input("المركز / الحي / القرية")
                    area = st.number_input("المساحة الإجمالية بالصك (م²)", min_value=0.0, step=0.1)
                    g_type = st.selectbox("تصنيف العقار الحالي", ["أرض فضاء التابعة للوزارة", "مركز صحي قائم", "مستشفى عام وتخصصي"])
                
                c_c1, c_c2 = st.columns(2)
                with c_c1:
                    lat = st.text_input("إحداثي دائرة العرض (Latitude) - مثل: 21.26")
                with c_c2:
                    lon = st.text_input("إحداثي خط الطول (Longitude) - مثل: 40.41")
                    
                if st.form_submit_button("💾 اعتماد وحفظ الأصل في قاعدة البيانات الحالية"):
                    if r_sok and site_name:
                        new_asset = pd.DataFrame([{"م": len(df_assets)+1, "المنشأة": site_name, "نوع_العقار": g_type, "حاله_العقار": "ملك", "المحافظة": "الطائف", "مركز/حي/قرية": district, "المساحة": area, "رقم الصك": r_sok, "تاريخ الصك": datetime.now().strftime("%Y-%m-%d"), "خط الطول": lon, "دائرة العرض": lat, "مساحة_الرفع_الفعلي": area}])
                        pd.concat([df_assets, new_asset]).to_csv("assets_db.csv", index=False)
                        log_action(st.session_state['username'], "إضافة أصل يدوياً", f"منشأة: {site_name}")
                        st.success("✅ تم حفظ الأصل وتخزينه بنجاح داخل قاعدة بيانات كمبيوترك!")
                        st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

            st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
            st.markdown("<div class='grid-header'><span><i class='fa-solid fa-pen-to-square'></i> تعديل وحذف الأصول والـ 138 منشأة مباشرة من الجدول الشبكي التفاعلي</span></div>", unsafe_allow_html=True)
            if not df_assets.empty:
                edited_df = st.data_editor(df_assets, num_rows="dynamic", use_container_width=True, key="assets_editor_v138")
                if st.button("💾 حفظ وتثبيت كافة التغييرات والتعديلات النشطة للأصول"):
                    edited_df.to_csv("assets_db.csv", index=False)
                    log_action(st.session_state['username'], "تعديل جدول الأصول", "تحديث قيود الجدول الحقيقي")
                    st.success("🎉 تم تحديث وحفظ جدول البيانات والـ 138 موقع بنجاح على جهازك!")
                    st.rerun()
            else:
                st.info("قاعدة البيانات فارغة حالياً، يرجى التوجه لرفع ملف الاستيراد.")
            st.markdown("</div>", unsafe_allow_html=True)

    # 💼 تتبع وحفظ المعاملات الرقمية وحركات خط السير للعودة إليها
    elif menu == "💼 تتبع وحفظ المعاملات الرقمية":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-route'></i> حفظ وتتبع خط سير المعاملات الرقمية</h1>", unsafe_allow_html=True)
        
        st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
        st.markdown("<div class='grid-header'><span><i class='fa-solid fa-folder-plus'></i> تسجيل وقيد حركة معاملة جديدة وأرشفتها للعودة إليها مستقبلاً</span></div>", unsafe_allow_html=True)
        with st.form("workflow_add"):
            w_col1, w_col2 = st.columns(2)
            with w_col1:
                w_id = st.text_input("رقم المعاملة (الصادر أو الوارد)")
                w_subject = st.text_input("موضوع المعاملة الأساسي (مثال: طلب رخصة بناء مركز صحي الوشحاء)")
            with w_col2:
                w_dept = st.selectbox("الجهة أو الإدارة الحكومية الحالية الواقفة عندها المعاملة:", ["أمانة محافظة الطائف", "المحكمة العامة بالطائف", "كتابة العدل بمحافظة الطائف", "إدارة الشؤون الهندسية بالفرع", "وزارة الصحة بالرياض"], key="dept_workflow_drop")
                w_status = st.selectbox("حالة المعاملة الميدانية الحالية:", ["قيد الدراسة والتدقيق المساحي بالذرعة", "بانتظار الاعتماد النهائي الهندسي", "تم الإفراغ والمطابقة بنجاح", "معاملة منتهية تم أرشفتها"], key="status_workflow_drop")
            
            if st.form_submit_button("💾 قيد وحفظ خط سير المعاملة بأرشيف المنصة"):
                if w_id and w_subject:
                    new_wf = pd.DataFrame([{"رقم_المعاملة": w_id, "موضوع_المعاملة": w_subject, "الإدارة_الحالية": w_dept, "حالة_المعاملة": w_status, "تاريخ_التحديث": datetime.now().strftime("%Y-%m-%d %H:%M"), "الموظف_المسؤول": st.session_state['username']}])
                    pd.concat([df_workflows, new_wf]).to_csv("workflows_db.csv", index=False)
                    log_action(st.session_state['username'], "قيد معاملة أرشيفية", f"معاملة رقم {w_id}")
                    st.success("✅ تم حفظ وتأمين المعاملة في الأرشيف الرقمي للعودة إليها ومراجعتها في أي وقت!")
                    st.rerun()
                else:
                    st.error("❌ يجب ملء الخانات الأساسية (رقم المعاملة والموضوع) لإتمام الحفظ السليم.")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
        st.markdown("<div class='grid-header'><span><i class='fa-solid fa-folder-open'></i> أرشيف المعاملات المحفوظة بالجهاز (تعديل وحذف مباشر فوري بضغطة زر)</span></div>", unsafe_allow_html=True)
        if not df_workflows.empty:
            edited_wf_df = st.data_editor(df_workflows, num_rows="dynamic", use_container_width=True, key="workflow_editor_active")
            if st.button("💾 حفظ وتثبيت كافة تعديلات وحذوفات جدول تتبع المعاملات"):
                edited_wf_df.to_csv("workflows_db.csv", index=False)
                st.success("🎉 تم حفظ وتثبيت خط سير معاملاتك بنجاح على جهاز كمبيوترك!")
                st.rerun()
        else:
            st.info("💡 الأرشيف فارغ حالياً. قم بقيد معاملاتك بالاستمارة العلوية للعودة إليها وحفظها دائماً.")
        st.markdown("</div>", unsafe_allow_html=True)
    # ⚠️ نظام رقابة الأراضي وتتبع التعديات (مصلح بالكامل للتعديل والحفظ المباشر كطلبك)
    elif menu == "⚠️ رقابة الأراضي وتتبع التعديات":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-shield-halved'></i> تتبع التعديات والرقابة الميدانية للأراضي</h1>", unsafe_allow_html=True)
        
        st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
        st.markdown("<div class='grid-header'><span><i class='fa-solid fa-triangle-exclamation'></i> رصد وإدخال حالة تعدي جديدة على أراضي ممتلكات الصحة بالطائف</span></div>", unsafe_allow_html=True)
        with st.form("encroach_form"):
            c_e1, c_e2 = st.columns(2)
            with c_e1:
                e_sok = st.text_input("رقم الصك الشرعي المعتدى على حدوده")
                e_name = st.text_input("اسم المنشأة أو الموقع الطبي أو الأرض الفضاء")
            with c_e2:
                e_type = st.selectbox("نوع التعدي المرصود ميدانياً:", ["إقامة حوش أو أسوار غير نظامية", "بناء شعبي بدون رخصة", "وضع لوحات أو شبوك وتجريف تربة", "تعدي جزئي على الحدود المساحية الإنشائية"], key="enc_type_drop")
                e_status = st.selectbox("حالة المعاملة والقضية:", ["تحت الرفع للمحافظة", "بانتظار إزالة لجنة التعديات", "منظورة لدى الجهات الأمنية والمحكمة", "تمت الإزالة واسترداد الأرض بالكامل"], key="enc_status_drop")
            e_action = st.text_area("الإجراء المتخذ هندسياً وقانونياً", placeholder="اكتب رقم وتاريخ الخطاب الموجه للمحافظة أو الجهة الأمنية...")
            
            if st.form_submit_button("🚨 تسجيل حالة التعدي وأرشفة المعاملة رسمياً"):
                if e_sok and e_name:
                    new_enc = pd.DataFrame([{"رقم الصك": e_sok, "المنشأة": e_name, "نوع_التعدي": e_type, "حالة_القضية": e_status, "تاريخ_الرصد": datetime.now().strftime("%Y-%m-%d"), "الإجراء_المتخذ": e_action}])
                    pd.concat([df_encroach, new_enc]).to_csv("encroachments_db.csv", index=False)
                    log_action(st.session_state['username'], "رصد تعدي", f"صك رقم {e_sok}")
                    st.success("✅ تم تسجيل حالة التعدي بنجاح في السجلات لحماية أراضي الدولة وصحة الطائف.")
                    st.rerun()
                else:
                    st.error("❌ يرجى ملء الحقول الأساسية لتوثيق التعدي الميداني.")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
        st.markdown("<div class='grid-header'><span><i class='fa-solid fa-clock-rotate-left'></i> سجل قضايا التعديات المرصودة (متاح للتعديل والحذف والحفظ الفوري والشامل كطلبك)</span></div>", unsafe_allow_html=True)
        if not df_encroach.empty:
            edited_encroach = st.data_editor(df_encroach, num_rows="dynamic", use_container_width=True, key="encroach_editor_vfinal")
            if st.button("💾 حفظ كافة التعديلات والتغييرات النشطة في سجل التعديات"):
                edited_encroach.to_csv("encroachments_db.csv", index=False)
                log_action(st.session_state['username'], "تعديل جدول التعديات", "تحديث قضايا الرقابة")
                st.success("🎉 تم حفظ وتثبيت كافة تحديثات الرقابة والتعديات الميدانية على ملفات جهازك بنجاح!")
                st.rerun()
        else:
            st.info("👍 لا توجد قضايا تعديات مسجلة. جميع الأراضي والممتلكات محمية ومستقرة.")
        st.markdown("</div>", unsafe_allow_html=True)

    # 📜 سجل الحماية الرقابي والأمان (Audit Log)
    elif menu == "📜 سجل الحماية والمراقبة":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-user-shield'></i> سجل الحماية الرقابي والعمليات النشطة (Audit Log)</h1>", unsafe_allow_html=True)
        st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
        df_logs = pd.read_csv("audit_log.csv")
        st.dataframe(df_logs, use_container_width=True)
        
        st.markdown("<br><hr>", unsafe_allow_html=True)
        if st.button("🚨 تصفير النظام وإلغاء كافة البيانات المخزنة (إجراء نهائي ومسح شامل)"):
            pd.DataFrame(columns=["م", "المنشأة", "نوع_العقار", "حاله_العقار", "المحافظة", "مركز/حي/قرية", "المساحة", "رقم الصك", "تاريخ الصك", "خط الطول", "دائرة العرض"]).to_csv("assets_db.csv", index=False)
            pd.DataFrame(columns=["رقم الصك", "المنشأة", "نوع_التعدي", "حالة_القضية", "تاريخ_الرصد", "الإجراء_المتخذ"]).to_csv("encroachments_db.csv", index=False)
            log_action(st.session_state['username'], "تصفير المنصة", "قام المسؤول بتصفير المنصة بالكامل")
            st.success("🔄 تمت إعادة تهيئة وتصفير جداول البيانات بنجاح تام!")
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)
