import streamlit as st
import pandas as pd
from datetime import datetime
import os
from docx import Document

# 1. إعدادات الهوية البصرية الرسمية لوزارة الصحة
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
    .main { background-color: #f4f7f6 !important; }
    
    /* تصميم بطاقات الأداء الرقمية الفاخرة */
    .card-luxury {
        background: white;
        padding: 24px;
        border-radius: 16px;
        box-shadow: 0 10px 25px rgba(0,0,0,0.05);
        border-top: 4px solid #dfb76c;
        border-right: 6px solid #1d5c43;
        transition: transform 0.3s ease, box-shadow 0.3s ease;
        text-align: right;
        margin-bottom: 20px;
    }
    .card-luxury:hover { 
        transform: translateY(-5px);
        box-shadow: 0 15px 35px rgba(29, 92, 67, 0.1);
    }
    .card-icon { font-size: 2.5rem; color: #1d5c43; margin-bottom: 10px; }
    
    /* تنبيهات النظام */
    .alert-premium {
        background: #fff5f5;
        padding: 18px;
        border-radius: 12px;
        border-right: 6px solid #e74c3c;
        text-align: right;
        color: #c0392b;
        font-weight: bold;
    }
    
    /* تحسين أزرار النظام بالكامل */
    .stButton>button {
        background: linear-gradient(135deg, #1d5c43 0%, #154431 100%) !important;
        color: #ffffff !important;
        border-radius: 10px !important;
        padding: 12px !important;
        font-size: 14px !important;
        font-weight: 700 !important;
        border: none !important;
        width: 100% !important;
        box-shadow: 0 4px 15px rgba(29, 92, 67, 0.2) !important;
    }
    .stButton>button:hover {
        background: linear-gradient(135deg, #dfb76c 0%, #c59b4e 100%) !important;
        color: #1d5c43 !important;
    }
    
    /* ضبط لون القائمة الجانبية للتخلص من التداخل والنقاط */
    div[data-testid="stSidebarUserContent"] {
        background: #113829 !important;
        color: #ffffff !important;
    }
    div[data-testid="stSidebarUserContent"] p, 
    div[data-testid="stSidebarUserContent"] h3,
    div[data-testid="stSidebarUserContent"] label,
    div[data-testid="stSidebarUserContent"] span {
        color: #ffffff !important;
    }
    div[data-widget="stSidebar"] .stRadio>label { color: #ffffff !important; }
    div[data-widget="stSidebar"] div[data-testid="stMarkdownContainer"] p { color: #ffffff !important; font-size: 15px; }
    
    /* محاذاة الجداول والخانات من اليمين لليسار */
    .stTextInput>div>div>input, .stSelectbox>div>div>div, .stNumberInput>div>div>input {
        text-align: right !important;
        direction: rtl !important;
    }
    </style>
""", unsafe_allow_html=True)

# 2. تهيئة وتأمين قاعدة البيانات المحلية وحمايتها بشكل قاطع من الأخطاء الفارغة
def init_db():
    required_cols = ["م", "المنشأة", "نوع_العقار", "حاله_العقار", "المحافظة", "مركز/حي/قرية", "المساحة", "رقم الصك", "تاريخ الصك", "خط الطول", "دائرة العرض"]
    # إذا كان الملف غير موجود أو فارغ تماماً، يتم بناؤه فوراً بالهيكلية الصحيحة لمنع خطأ EmptyDataError
    if not os.path.exists("assets_db.csv") or os.stat("assets_db.csv").st_size == 0:
        pd.DataFrame(columns=required_cols).to_csv("assets_db.csv", index=False)
    if not os.path.exists("encroachments_db.csv") or os.stat("encroachments_db.csv").st_size == 0:
        pd.DataFrame(columns=["رقم الصك", "المنشأة", "نوع_التعدي", "حالة_القضية", "تاريخ_الرصد", "الإجراء_المتخذ"]).to_csv("encroachments_db.csv", index=False)
    if not os.path.exists("workflows_db.csv") or os.stat("workflows_db.csv").st_size == 0:
        pd.DataFrame(columns=["رقم_المعاملة", "موضوع_المعاملة", "الإدارة_الحالية", "حالة_المعاملة", "تاريخ_التحديث", "الموظف_المسؤول"]).to_csv("workflows_db.csv", index=False)

init_db()

# 3. بوابة الأمان والتحقق الرقمي
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False
    st.session_state['role'] = None
    st.session_state['username'] = ""

if not st.session_state['logged_in']:
    st.markdown("<br><br><h2 style='text-align: center; color: #1d5c43;'><i class='fa-solid fa-shield-halved'></i> المنصة الرقمية للأصول والممتلكات</h2>", unsafe_allow_html=True)
    st.markdown("<h4 style='text-align: center; color: #777;'>فرع وزارة الصحة بمحافظة الطائف</h4>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col2:
        st.markdown("<div class='card-luxury' style='border-top: 4px solid #1d5c43;'>", unsafe_allow_html=True)
        role_input = st.selectbox("🔒 فئة الدخول للمنصة", ["مدير النظام (Admin)", "موظف الإدارة (Staff)"])
        password = st.text_input("🔑 كود التحقق السري", type="password")
        if st.button("🔓 دخول آمن للنظام"):
            if role_input == "مدير النظام (Admin)" and password == "MOH@2026":
                st.session_state['logged_in'] = True
                st.session_state['role'] = "Admin"
                st.session_state['username'] = "مدير الإدارة"
                st.rerun()
            elif role_input == "موظف الإدارة (Staff)" and password == "Staff@Taif":
                st.session_state['logged_in'] = True
                st.session_state['role'] = "Staff"
                st.session_state['username'] = "موظف ممتلكات"
                st.rerun()
            else:
                st.error("❌ كود التحقق السري غير صحيح!")
        st.markdown("</div>", unsafe_allow_html=True)
else:
    # القائمة الجانبية الفخمة وعالية التباين
    st.sidebar.markdown(f"<div style='text-align: center; padding: 10px;'><i class='fa-solid fa-circle-user' style='font-size: 3.5rem; color: #dfb76c;'></i></div>", unsafe_allow_html=True)
    st.sidebar.markdown(f"<h3 style='color: white; text-align: center; margin-top: 5px;'>{st.session_state['username']}</h3>", unsafe_allow_html=True)
    st.sidebar.markdown(f"<p style='color: #dfb76c; text-align: center; font-weight: bold; margin-bottom: 20px;'><i class='fa-solid fa-id-card-clip'></i> رتبة: {st.session_state['role']}</p>", unsafe_allow_html=True)
    
    st.sidebar.markdown("<hr style='border-color: rgba(255,255,255,0.2);'>", unsafe_allow_html=True)
    
    menu = st.sidebar.radio(
        "📂 تصفح أقسام المنصة",
        [
            "📊 لوحة التحكم اليومية", 
            "📥 استيراد ورفع ملفات Excel",
            "🔍 ملفات الأصول وبطاقات الوثائق المصورة", 
            "⚙️ التحكم بالأصول (تعديل/حذف)",
            "💼 تتبع وحفظ المعاملات الرقمية",
            "⚠️ رقابة الأراضي وتتبع التعديات"
        ]
    )
    
    st.sidebar.markdown("<br><br>", unsafe_allow_html=True)
    if st.sidebar.button("🚪 خروج آمن من النظام"):
        st.session_state['logged_in'] = False
        st.rerun()

    # القراءة الفورية والمحمية بحلقات الفحص الآمنة لمنع الـ EmptyDataError تماماً
    try:
        df_assets = pd.read_csv("assets_db.csv")
    except:
        df_assets = pd.DataFrame()
        
    try:
        df_encroach = pd.read_csv("encroachments_db.csv")
    except:
        df_encroach = pd.DataFrame()
        
    try:
        df_workflows = pd.read_csv("workflows_db.csv")
    except:
        df_workflows = pd.DataFrame()

    # 📊 لوحة التحكم اليومية المستقرة
    if menu == "📊 لوحة التحكم اليومية":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-chart-pie'></i> لوحة المؤشرات الرقمية والأداء اليومي للأصول</h1>", unsafe_allow_html=True)
        
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f"<div class='card-luxury'><div class='card-icon'><i class='fa-solid fa-hotel'></i></div><h3 style='color: #1d5c43;'>إجمالي المنشآت والعقارات</h3><h2>{len(df_assets) if not df_assets.empty else 0} موقع مقيد</h2></div>", unsafe_allow_html=True)
        with c2:
            if not df_assets.empty and "المساحة" in df_assets.columns:
                df_assets['clean_area'] = pd.to_numeric(df_assets["المساحة"].astype(str).str.replace(',', '', regex=True).str.replace(' ', '', regex=True).str.strip(), errors='coerce').fillna(0)
                total_area = df_assets['clean_area'].sum()
            else:
                total_area = 0
            st.markdown(f"<div class='card-luxury'><div class='card-icon'><i class='fa-solid fa-up-right-and-down-left-from-center'></i></div><h3 style='color: #1d5c43;'>المساحات الإجمالية المحمية</h3><h2>{total_area:,.2f} م²</h2></div>", unsafe_allow_html=True)
        with c3:
            st.markdown(f"<div class='card-luxury'><div class='card-icon'><i class='fa-solid fa-shield-cat'></i></div><h3 style='color: #1d5c43;'>التعديات المرصودة</h3><h2>{len(df_encroach) if not df_encroach.empty else 0} حالة تحت الإجراء</h2></div>", unsafe_allow_html=True)
            
        # تشغيل محرك الخريطة الجغرافية القديم الموثوق بناءً على أسماء حقول ملفك الحقيقية
        st.markdown("<br><h3 style='text-align: right;'><i class='fa-solid fa-map-marked-alt'></i> النطاق الجغرافي للأملاك والمشاريع (رؤية الأقمار الصناعية)</h3>", unsafe_allow_html=True)
        if not df_assets.empty and "خط الطول" in df_assets.columns and "دائرة العرض" in df_assets.columns:
            try:
                map_data = df_assets[["دائرة العرض", "خط الطول"]].dropna()
                map_data.columns = ["lat", "lon"]
                map_data["lat"] = pd.to_numeric(map_data["lat"].astype(str).str.replace('°', '', regex=True).str.strip(), errors='coerce')
                map_data["lon"] = pd.to_numeric(map_data["lon"].astype(str).str.replace('°', '', regex=True).str.strip(), errors='coerce')
                map_data = map_data.dropna()
                if not map_data.empty:
                    st.map(map_data, use_container_width=True)
                else:
                    st.info("💡 لا توجد إحداثيات رقمية صالحة لعرضها على الخريطة حالياً.")
            except:
                st.info("💡 الخريطة الرقمية تنتظر معالجة البيانات الجغرافية.")
        else:
            st.info("💡 الخريطة الرقمية جاهزة وتنتظر استيراد بيان ملف الـ Excel لتنشيط المواقع الـ 144 تلقائياً.")

    # 📥 قسم استيراد ورفع ملفات Excel
    elif menu == "📥 استيراد ورفع ملفات Excel":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-file-excel'></i> استيراد ورفع البيانات الذكي من ملفات Excel</h1>", unsafe_allow_html=True)
        if st.session_state['role'] != "Admin":
            st.markdown("<div class='alert-premium'><i class='fa-solid fa-lock'></i> عذراً، خاصية رفع واستيراد الملفات تتطلب صلاحية مدير النظام (Admin).</div>", unsafe_allow_html=True)
        else:
            st.markdown("<div class='card-luxury'>", unsafe_allow_html=True)
            st.markdown("<h4 style='text-align: right; color: #1d5c43;'>📥 مركز القراءة والرفع الآلي للملف الموثق:</h4>", unsafe_allow_html=True)
            uploaded_file = st.file_uploader("اختر ملف ممتلكات وزارة الصحة الموثق بالطائف من جهازك", type=["xlsx", "xls", "csv"])
            if uploaded_file is not None:
                try:
                    df_uploaded = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
                    st.success(f"✅ تم قراءة ملفك بنجاح! تم رصد وعزل {len(df_uploaded)} منشأة وموقع عقاري.")
                    st.dataframe(df_uploaded.head(10), use_container_width=True)
                    if st.button("🚀 اعتماد وحفظ القوائم المرفوعة في قاعدة بيانات خادم الويب دائمًا"):
                        df_uploaded.to_csv("assets_db.csv", index=False)
                        st.balloons()
                        st.success("🎉 تم حفظ وتثبيت كافة البيانات حية وبنفس مسميات عناوين ملفك الأصلي بنجاح تام!")
                        st.rerun()
                except Exception as e:
                    st.error(f"حدث خطأ أثناء قراءة الملف: {e}")
            st.markdown("</div>", unsafe_allow_html=True)
    # 🔍 ملفات الأصول وبطاقات الوثائق المصورة وصائغ الخطابات الرسمي المطور
    elif menu == "🔍 ملفات الأصول وبطاقات الوثائق المصورة":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-folder-open'></i> ملفات الأصول وبطاقات الوثائق المصورة والخطابات</h1>", unsafe_allow_html=True)
        
        if df_assets.empty or 'المنشأة' not in df_assets.columns:
            st.info("💡 قاعدة البيانات فارغة حالياً على الويب، يرجى التوجه أولاً لقسم '📥 استيراد ورفع ملفات Excel' لرفع بيان منشآتك وتنشيط الواجهة تلقائياً.")
        else:
            st.markdown("<div class='card-luxury'>", unsafe_allow_html=True)
            facility_list = df_assets['المنشأة'].dropna().unique().tolist()
            selected_facility = st.selectbox("🏥 اضغط هنا لاختيار المنشأة أو الأرض الفضاء المستهدفة لاستعراض ملفها الفني:", facility_list)
            st.markdown("</div>", unsafe_allow_html=True)
            
            if selected_facility:
                asset_data = df_assets[df_assets['المنشأة'] == selected_facility].iloc[0]
                st.markdown(f"<h2> الملف العقاري والوثائق الموثقة لـ: {selected_facility}</h2>", unsafe_allow_html=True)
                
                # عرض بطاقات المعاينة الأربعة الفاخرة المنسقة جرافيكياً بالتوازي كطلبك
                col_box1, col_box2 = st.columns(2)
                col_box3, col_box4 = st.columns(2)
                
                with col_box1:
                    st.markdown("<div class='card-luxury' style='border-top: 4px solid #dfb76c;'><h4>📜 1. صورة الصك الشرعي الموثق</h4>", unsafe_allow_html=True)
                    sok_img = st.file_uploader(f"رفع/تحديث صورة الصك لـ {selected_facility}", type=["jpg", "png", "jpeg"], key=f"sok_up_{asset_data['رقم الصك']}")
                    if sok_img is not None: st.image(sok_img, use_container_width=True)
                    else: st.info("📷 لم يتم رفع صورة الصك لهذا الموقع بعد.")
                    st.markdown("</div>", unsafe_allow_html=True)
                    
                with col_box2:
                    st.markdown("<div class='card-luxury' style='border-top: 4px solid #1d5c43;'><h4>🏗️ 2. صورة رخصة البناء الهندسية</h4>", unsafe_allow_html=True)
                    permit_img = st.file_uploader(f"رفع/تحديث صورة الرخصة لـ {selected_facility}", type=["jpg", "png", "jpeg"], key=f"perm_up_{asset_data['رقم الصك']}")
                    if permit_img is not None: st.image(permit_img, use_container_width=True)
                    else: st.info("📷 لم يتم رفع صورة رخصة البناء لهذا الموقع بعد.")
                    st.markdown("</div>", unsafe_allow_html=True)
                    
                with col_box3:
                    st.markdown("<div class='card-luxury' style='border-top: 4px solid #58a6ff;'><h4>🛰️ 3. صورة الرفع المساحي والموقع الميداني (الأقمار الصناعية)</h4>", unsafe_allow_html=True)
                    site_img = st.file_uploader(f"رفع/تحديث المخطط الجغرافي لـ {selected_facility}", type=["jpg", "png", "jpeg"], key=f"site_up_{asset_data['رقم الصك']}")
                    if site_img is not None: st.image(site_img, use_container_width=True)
                    else:
                        lat_val = str(asset_data['دائرة العرض']).replace('°','').strip() if 'دائرة العرض' in asset_data else '21.27'
                        lon_val = str(asset_data['خط الطول']).replace('°','').strip() if 'خط الطول' in asset_data else '40.41'
                        maps_url = f"https://google.com{lat_val},{lon_val}"
                        st.markdown(f"<a href='{maps_url}' target='_blank'><button style='width:100%; padding:10px; background:#1d5c43; color:white; border:none; border-radius:6px; cursor:pointer; font-weight:bold;'><i class='fa-solid fa-location-arrow'></i> استعراض الموقع الجغرافي الحي على Google Maps</button></a>", unsafe_allow_html=True)
                    st.markdown("</div>", unsafe_allow_html=True)
                    
                with col_box4:
                    st.markdown("<div class='card-luxury' style='border-top: 4px solid #143f2e;'><h4>📐 4. صورة قرار الذرعة والقرار المساحي مع أمانة الطائف</h4>", unsafe_allow_html=True)
                    zar_img = st.file_uploader(f"رفع/تحديث صورة قرار الذرعة لـ {selected_facility}", type=["jpg", "png", "jpeg"], key=f"zar_up_{asset_data['رقم الصك']}")
                    if zar_img is not None: st.image(zar_img, use_container_width=True)
                    else: st.info("📷 لم يتم رفع صورة قرار الذرعة المعتمد من الأمانة بعد.")
                    st.markdown("</div>", unsafe_allow_html=True)

                # محرك صائغ الخطابات الموثق بالمسميات القيادية الرسمية الجديدة المتكاملة
                st.markdown("<div class='card-luxury'>", unsafe_allow_html=True)
                st.markdown("<h4>✉️ محرك الصياغة الآلي المعتمد للخطابات الرسمية فرع الوزارة</h4>", unsafe_allow_html=True)
                letter_target = st.selectbox("حدد الجهة الحكومية المستهدفة بالخطاب:", [
                    "خطاب موجه لسعادة أمين محافظة الطائف (استخراج رخصة بناء ومطابقة ذرعة)", 
                    "خطاب موجه لسعادة مدير شركة الكهرباء بالطائف (طلب إيصال التيار وتحديد المحول)", 
                    "خطاب موجه لفضيلة رئيس المحكمة العامة بالطائف (تحديث ومطابقة صك ممتلكات الصحة)"
                ])
                
                sok_num = asset_data['رقم الصك'] if ('رقم الصك' in asset_data and pd.notnull(asset_data['رقم الصك'])) else "لا يوجد"
                sok_date = asset_data['تاريخ الصك'] if ('تاريخ الصك' in asset_data and pd.notnull(asset_data['تاريخ الصك'])) else "لا يوجد"
                facility_area = asset_data['المساحة'] if ('المساحة' in asset_data and pd.notnull(asset_data['المساحة'])) else "غير محدد"
                facility_village = asset_data['مركز/حي/قرية'] if ('مركز/حي/قرية' in asset_data and pd.notnull(asset_data['مركز/حي/قرية'])) else "محافظة الطائف"
                
                if "أمين" in letter_target:
                    text_content = f"سعادة أمين محافظة الطائف\nالسلام عليكم ورحمة الله وبركاته،،\n\nتفيدكم فرع وزارة الصحة بمحافظة الطائف علماً بملكية الوزارة الرسمية للموقع المخصص لـ ({selected_facility}) والواقع بنطاق ({facility_village}) بموجب الصك الشرعي رقم ({sok_num}) وتاريخ ({sok_date}) بمساحة قدرها ({facility_area} م²). نأمل التوجيه لمن يلزم لاعتماد الرفع المساحي وقرار الذرعة واستخراج رخصة بناء وفق الإحداثيات المرفقة.\n\nوتقبلوا خالص التحية والتقدير،،\n\nمدير فرع وزارة الصحة بمحافظة الطائف\nإدارة الأراضي والممتلكات"
                elif "الكهرباء" in letter_target:
                    text_content = f"سعادة مدير شركة الكهرباء بمحافظة الطائف\nالسلام عليكم ورحمة الله وبركاته،،\n\nنظراً لجهوزية البدء الإنشائي والتشغيلي للموقع الطبي التابع للوزارة ({selected_facility}) والمقام على الأرض ذات الصك رقم ({sok_num})، نأمل منكم الإيعاز للمختصين لطلب إيصال التيار الكهربائي وتحديد موقع محول الطاقة الفرعي.\n\nوتقبلوا وافر التحية والتقدير،،\n\nمدير فرع وزارة الصحة بمحافظة الطائف\nإدارة الأراضي والممتلكات"
                else:
                    text_content = f"فضيلة رئيس المحكمة العامة بمحافظة الطائف\nالسلام عليكم ورحمة الله وبركاته،،\n\nتتقدم فرع وزارة الصحة بمحافظة الطائف بطلب تحديث وإصدار صك إلكتروني موحد ومطابقة مسحية للصك رقم ({sok_num}) وتاريخ ({sok_date}) العائد لملك الوزارة في موقع ({selected_facility}) بنطاق ({facility_village}). نأمل التوجيه لمطابقة الحدود الإنشائية حسب الرفع المساحي المرفق.\n\nوتقبلوا خالص التحية والتقدير،،\n\nمدير فرع وزارة الصحة بمحافظة الطائف\nإدارة الأراضي والممتلكات"
                
                st.text_area("📄 صيغة الخطاب الرسمي والذكية المولد تلقائياً لفرع الوزارة:", text_content, height=200)
                doc = Document()
                doc.add_heading(letter_target, 0)
                doc.add_paragraph(text_content)
                doc.save("generated_letter.docx")
                with open("generated_letter.docx", "rb") as f:
                    st.download_button("📥 تنزيل الخطاب الآن كملف Word رسمي مجهز بالكامل للطباعة", f, file_name=f"خطاب_صحة_الطائف_{sok_num}.docx")
                st.markdown("</div>", unsafe_allow_html=True)
    # ⚙️ لوحة التحكم بالأصول
    elif menu == "⚙️ التحكم بالأصول (تعديل/حذف)":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-folder-gear'></i> إدارة وتعديل وحذف الأصول العقارية</h1>", unsafe_allow_html=True)
        if st.session_state['role'] != "Admin":
            st.markdown("<div class='alert-premium'><i class='fa-solid fa-lock'></i> جدار الحماية: ميزة حفظ أو تعديل أو حذف الأصول تقتصر على مدير النظام (Admin) فقط.</div>", unsafe_allow_html=True)
        else:
            st.markdown("<div class='card-luxury'>", unsafe_allow_html=True)
            with st.form("manual_add"):
                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    r_sok = st.text_input("رقم الصك الشرعي الجديد")
                    r_moamala = st.text_input("رقم المعاملة أو القيد")
                    site_name = st.text_input("اسم الموقع الطبي / الأرض")
                with col_f2:
                    district = st.text_input("المركز / الحي / القرية")
                    area = st.text_input("المساحة الإجمالية بالصك (م²)")
                    g_type = st.selectbox("تصنيف العقار الحالي", ["أرض فضاء التابعة للوزارة", "مركز صحي قائم", "مستشفى عام وتخصصي"])
                
                c_c1, c_c2 = st.columns(2)
                with c_c1: lat = st.text_input("إحداثي دائرة العرض (Latitude)")
                with c_c2: lon = st.text_input("إحداثي خط الطول (Longitude)")
                
                bypass_dup = st.checkbox("السماح برفع وتمرير رقم الصك وتكراره استثنائياً (موافقة الإدارة العليا)")
                
                if st.form_submit_button("💾 اعتماد وحفظ الأصل في قاعدة البيانات الحالية"):
                    has_dup = False
                    if not df_assets.empty and 'رقم الصك' in df_assets.columns:
                        if r_sok in df_assets['رقم الصك'].astype(str).values: has_dup = True
                            
                    if has_dup and not bypass_dup:
                        st.error("❌ تنبيه أمني عاجل: رقم الصك هذا مسجل مسبقاً في النظام! لا يمكن التكرار إلا بموافقة الإدارة العليا.")
                    else:
                        new_asset = pd.DataFrame([{"م": len(df_assets)+1, "المنشأة": site_name, "نوع_العقار": g_type, "حاله_العقار": "ملك", "المحافظة": "الطائف", "مركز/حي/قرية": district, "المساحة": area, "رقم الصك": r_sok, "تاريخ الصك": datetime.now().strftime("%Y-%m-%d"), "خط الطول": lon, "دائرة العرض": lat}])
                        updated_df = pd.concat([df_assets, new_asset]).reset_index(drop=True)
                        updated_df.to_csv("assets_db.csv", index=False)
                        st.success("✅ تم حفظ وتأمين الأصل الجديد بنجاح في قاعدة البيانات!")
                        st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

            st.markdown("<div class='card-luxury'>", unsafe_allow_html=True)
            if not df_assets.empty:
                edited_df = st.data_editor(df_assets, num_rows="dynamic", use_container_width=True, key="assets_editor_v138")
                if st.button("💾 حفظ كافة التغييرات والتعديلات النشطة للأصول العقارية"):
                    edited_df.to_csv("assets_db.csv", index=False)
                    st.success("🎉 تم تحديث وحفظ جدول البيانات والمنشآت بنجاح!")
                    st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

    # 💼 تتبع وحفظ المعاملات الرقمية
    elif menu == "💼 تتبع وحفظ المعاملات الرقمية":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-route'></i> حفظ وتتبع خط سير المعاملات الرقمية للأراضي</h1>", unsafe_allow_html=True)
        st.markdown("<div class='card-luxury'>", unsafe_allow_html=True)
        with st.form("workflow_add"):
            w_col1, w_col2 = st.columns(2)
            with w_col1:
                w_id = st.text_input("رقم المعاملة (الصادر أو الوارد)")
                w_subject = st.text_input("موضوع المعاملة الأساسي")
            with w_col2:
                w_dept = st.selectbox("الجهة أو الإدارة الحكومية الحالية عندها المعاملة:", ["أمانة محافظة الطائف", "المحكمة العامة بالطائف", "كتابة العدل بمحافظة الطائف", "إدارة الأراضي والممتلكات بالفرع"])
                w_status = st.selectbox("حالة المعاملة الميدانية الحالية:", ["قيد الدراسة والتدقيق المساحي بالذرعة", "تم الإفراغ والمطابقة بنجاح", "معاملة منتهية تم أرشفتها"])
            if st.form_submit_button("💾 قيد وحفظ خط سير المعاملة بأرشيف المنصة"):
                new_wf = pd.DataFrame([{"رقم_المعاملة": w_id, "موضوع_المعاملة": w_subject, "الإدارة_الحالية": w_dept, "حالة_المعاملة": w_status, "تاريخ_التحديث": datetime.now().strftime("%Y-%m-%d %H:%M"), "الموظف_المسؤول": st.session_state['username']}])
                updated_wf = pd.concat([df_workflows, new_wf]).reset_index(drop=True)
                updated_wf.to_csv("workflows_db.csv", index=False)
                st.success("✅ تم حفظ وتأمين المعاملة في الأرشيف الرقمي لفرع الوزارة للعودة إليها مستقبلاً!")
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='card-luxury'>", unsafe_allow_html=True)
        if not df_workflows.empty:
            edited_wf_df = st.data_editor(df_workflows, num_rows="dynamic", use_container_width=True, key="workflow_editor_active")
            if st.button("💾 حفظ كافة تعديلات وحذوفات جدول تتبع المعاملات الرقمية"):
                edited_wf_df.to_csv("workflows_db.csv", index=False)
                st.success("🎉 تم حفظ وتثبيت خط سير معاملاتك بنجاح على خادم الويب!")
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    # ⚠️ نظام رقابة الأراضي وتتبع التعديات
    elif menu == "⚠️ رقابة الأراضي وتتبع التعديات":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-shield-halved'></i> تتبع التعديات والرقابة الميدانية للأراضي</h1>", unsafe_allow_html=True)
        st.markdown("<div class='card-luxury'>", unsafe_allow_html=True)
        with st.form("encroach_form"):
            c_e1, c_e2 = st.columns(2)
            with c_e1:
                e_sok = st.text_input("رقم الصك الشرعي المعتدى على حدوده")
                e_name = st.text_input("اسم المنشأة أو الموقع الطبي")
            with c_e2:
                e_type = st.selectbox("نوع التعدي المرصود ميدانياً:", ["إقامة حوش أو أسوار غير نظامية", "بناء شعبي بدون رخصة", "وضع لوحات أو شبوك وتجريف تربة"])
                e_status = st.selectbox("حالة المعاملة والقضية:", ["تحت الرفع للمحافظة", "بانتظار إزالة لجنة التعديات", "تمت الإزالة واسترداد الأرض بالكامل"])
            e_action = st.text_area("الإجراء المتخذ هندسياً وقانونياً")
            if st.form_submit_button("🚨 تسجيل حالة التعدي وأرشفة المعاملة رسمياً"):
                new_enc = pd.DataFrame([{"رقم الصك": e_sok, "المنشأة": e_name, "نوع_التعدي": e_type, "حالة_القضية": e_status, "تاريخ_الرصد": datetime.now().strftime("%Y-%m-%d"), "الإجراء_المتخذ": e_action}])
                updated_enc = pd.concat([df_encroach, new_enc]).reset_index(drop=True)
                updated_enc.to_csv("encroachments_db.csv", index=False)
                st.success("✅ تم تسجيل حالة التعدي بنجاح في السجلات لحماية ممتلكات صحة الطائف.")
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='card-luxury'>", unsafe_allow_html=True)
        if not df_encroach.empty:
            edited_encroach = st.data_editor(df_encroach, num_rows="dynamic", use_container_width=True, key="encroach_editor_vfinal")
            if st.button("💾 حفظ كافة التعديلات والتغييرات النشطة في سجل التعديات والرقابة"):
                edited_encroach.to_csv("encroachments_db.csv", index=False)
                st.success("🎉 تم حفظ وتثبيت كافة تحديثات الرقابة والتعديات الميدانية بنجاح على خادم الويب!")
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)
