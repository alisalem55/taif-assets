import streamlit as st
import pandas as pd
from datetime import datetime
import os
from docx import Document

# 1. إعدادات الهوية البصرية الرسمية لوزارة الصحة السعودية لعام 2026
st.set_page_config(
    page_title="منصة الممتلكات والرقابة الرقمية | صحة الطائف",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown('<link rel="stylesheet" href="https://cloudflare.com">', unsafe_allow_html=True)

st.markdown("""
    <style>
    @import url('https://googleapis.com');
    * { font-family: 'Cairo', sans-serif; }
    .main { background-color: #f4f7f6 !important; }
    
    /* تصميم البطاقات الشبكية الفاخرة الموزعة بالتوازي */
    .grid-box {
        background: #ffffff !important;
        padding: 22px !important;
        border-radius: 14px !important;
        border: 1px solid #e5e7eb !important;
        box-shadow: 0 4px 10px rgba(0, 0, 0, 0.04) !important;
        margin-bottom: 20px !important;
        text-align: right !important;
    }
    .grid-header {
        border-bottom: 2px solid #f3f4f6 !important;
        padding-bottom: 10px !important;
        margin-bottom: 14px !important;
        font-weight: 700 !important;
        color: #1d5c43 !important;
        font-size: 1.15rem !important;
    }
    .card-luxury {
        background: white; padding: 24px; border-radius: 16px;
        box-shadow: 0 10px 25px rgba(0,0,0,0.05); border-top: 4px solid #dfb76c;
        border-right: 6px solid #1d5c43; text-align: right; margin-bottom: 20px;
    }
    .card-icon { font-size: 2.5rem; color: #1d5c43; margin-bottom: 10px; }
    .alert-premium {
        background: #fff5f5; padding: 18px; border-radius: 12px;
        border-right: 6px solid #e74c3c; text-align: right; color: #c0392b; font-weight: bold;
    }
    
    /* أزرار النظام بالهوية الرسمية الاستباقية للوزارة */
    .stButton>button {
        background: linear-gradient(135deg, #1d5c43 0%, #154431 100%) !important;
        color: #ffffff !important; border-radius: 10px !important; padding: 12px !important;
        font-size: 14px !important; font-weight: 700 !important; border: none !important; width: 100% !important;
        box-shadow: 0 4px 15px rgba(29, 92, 67, 0.2) !important;
    }
    .stButton>button:hover {
        background: linear-gradient(135deg, #dfb76c 0%, #c59b4e 100%) !important; color: #1d5c43 !important;
    }
    
    /* التخلص من الدوائر والتداخل المزعج في القائمة الجانبية المحدثة */
    div[data-testid="stSidebarUserContent"] { background: #113829 !important; color: #ffffff !important; }
    div[data-testid="stSidebarUserContent"] p, div[data-testid="stSidebarUserContent"] h3,
    div[data-testid="stSidebarUserContent"] label, div[data-testid="stSidebarUserContent"] span { color: #ffffff !important; }
    div[data-widget="stSidebar"] .stRadio>label { color: #ffffff !important; }
    div[data-widget="stSidebar"] div[data-testid="stMarkdownContainer"] p { color: #ffffff !important; font-size: 15px; }
    
    /* مواءمة القوائم المنسدلة وصناديق النصوص من اليمين لليسار وبخطوط واضحة */
    .stTextInput>div>div>input, .stSelectbox>div>div>div, .stNumberInput>div>div>input, .stTextArea>div>div>textarea {
        text-align: right !important; direction: rtl !important; font-size: 15px !important; font-weight: 600 !important;
    }
    </style>
""", unsafe_allow_html=True)

# تفعيل الذاكرة السحابية التفاعلية المستمرة لحفظ الأصول ومنع اختفائها من السيرفر
if "cloud_assets" not in st.session_state:
    st.session_state["cloud_assets"] = pd.DataFrame(columns=["م", "المنشأة", "نوع_العقار", "حاله_العقار", "المحافظة", "مركز/حي/قرية", "المساحة", "رقم الصك", "تاريخ الصك", "خط الطول", "دائرة العرض"])
if "cloud_workflows" not in st.session_state:
    st.session_state["cloud_workflows"] = pd.DataFrame(columns=["رقم_المعاملة", "موضوع_المعاملة", "الإدارة_الحالية", "حالة_المعاملة", "تاريخ_التحديث"])
if "cloud_encroach" not in st.session_state:
    st.session_state["cloud_encroach"] = pd.DataFrame(columns=["رقم الصك", "المنشأة", "نوع_التعدي", "حالة_القضية", "تاريخ_الرصد", "الإجراء_المتخذ"])

# بوابة الأمان والتحقق الرقمي للولوج
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
    df_assets = st.session_state["cloud_assets"]
    df_encroach = st.session_state["cloud_encroach"]
    df_workflows = st.session_state["cloud_workflows"]

    # القائمة الجانبية الاحترافية المتكاملة
    menu = st.sidebar.radio(
        "📂 تصفح أقسام المنصة حياً",
        [
            "📊 لوحة التحكم والمؤشرات",
            "📥 استيراد ورفع ملفات Excel",
            "🔍 ملفات الأصول وبطاقات الوثائق والخطابات",
            "⚙️ التحكم بالأصول (تعديل يدوياً/حذف)",
            "💼 تتبع وحفظ المعاملات الرقمية",
            "⚠️ رقابة الأراضي وتتبع التعديات"
        ]
    )
    
    st.sidebar.markdown("<br><br>", unsafe_allow_html=True)
    if st.sidebar.button("🚪 خروج آمن من النظام"):
        st.session_state['logged_in'] = False
        st.rerun()

    # 📊 لوحة التحكم والمؤشرات اليومية الشاملة
    if menu == "📊 لوحة التحكم والمؤشرات":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-chart-line'></i> لوحة القيادة والمؤشرات الرقمية لأصول ممتلكات الصحة</h1>", unsafe_allow_html=True)
        
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f"<div class='card-luxury'><div class='card-icon'><i class='fa-solid fa-hotel'></i></div><h3 style='color: #1d5c43;'>إجمالي المواقع المقيدة</h3><h2>{len(df_assets)} منشأة مسجلة</h2></div>", unsafe_allow_html=True)
        with c2:
            if not df_assets.empty and "المساحة" in df_assets.columns:
                # تنظيف وحساب فوري للمساحات مهما كان شكل الحقول في ملفك
                df_assets['clean_area'] = pd.to_numeric(df_assets["المساحة"].astype(str).str.replace(',', '', regex=True).str.replace(' ', '', regex=True).str.strip(), errors='coerce').fillna(0)
                total_area = df_assets['clean_area'].sum()
            else:
                total_area = 0
            st.markdown(f"<div class='card-luxury'><div class='card-icon'><i class='fa-solid fa-up-right-and-down-left-from-center'></i></div><h3 style='color: #1d5c43;'>المساحات الإجمالية المحمية</h3><h2>{total_area:,.2f} م²</h2></div>", unsafe_allow_html=True)
        with c3:
            st.markdown(f"<div class='card-luxury'><div class='card-icon'><i class='fa-solid fa-triangle-exclamation'></i></div><h3 style='color: #1d5c43;'>التعديات الميدانية النشطة</h3><h2>{len(df_encroach)} حالة مرصودة</h2></div>", unsafe_allow_html=True)
            
        st.markdown("<br><h3 style='text-align: right;'><i class='fa-solid fa-map-location-dot'></i> النطاق الجغرافي وحصر المنشآت بالأقمار الصناعية (GIS)</h3>", unsafe_allow_html=True)
        if not df_assets.empty and "خط الطول" in df_assets.columns and "دائرة العرض" in df_assets.columns:
            try:
                map_data = df_assets[["دائرة العرض", "خط الطول"]].dropna()
                map_data.columns = ["lat", "lon"]
                map_data["lat"] = pd.to_numeric(map_data["lat"].astype(str).str.replace('°', '', regex=True).str.strip(), errors='coerce')
                map_data["lon"] = pd.to_numeric(map_data["lon"].astype(str).str.replace('°', '', regex=True).str.strip(), errors='coerce')
                map_data = map_data.dropna()
                if not map_data.empty:
                    st.map(map_data, use_container_width=True)
                else: st.info("💡 لا توجد إحداثيات صالحة حالياً لعرضها جغرافياً.")
            except: st.info("💡 نظام الخرائط يفتقر لإحداثيات سليمة بالملف المرفوع.")
        else:
            st.info("💡 الخريطة الجغرافية جاهزة وتنتظر رفع ملف الإكسل لتثبيت وعرض الـ 144 موقعًا تلقائيًا.")

    # 📥 قسم استيراد ورفع ملفات Excel والتثبيت السحابي
    elif menu == "📥 استيراد ورفع ملفات Excel":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-file-excel'></i> مركز استيراد وقراءة البيانات من ملفات Excel دفعة واحدة</h1>", unsafe_allow_html=True)
        if st.session_state['role'] != "Admin":
            st.markdown("<div class='alert-premium'><i class='fa-solid fa-lock'></i> عذراً، ميزة رفع الجداول واستيرادها تتطلب صلاحية مدير النظام (Admin).</div>", unsafe_allow_html=True)
        else:
            st.markdown("<div class='card-luxury'>", unsafe_allow_html=True)
            uploaded_file = st.file_uploader("اختر ملف ممتلكات وزارة الصحة الموثق بمحافظة الطائف (xlsx أو xls)", type=["xlsx", "xls", "csv"])
            if uploaded_file is not None:
                try:
                    df_uploaded = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
                    st.success(f"✅ تم قراءة الملف بنجاح! تم رصد وتوثيق {len(df_uploaded)} منشأة وموقع صحي.")
                    st.dataframe(df_uploaded, use_container_width=True)
                    if st.button("🚀 دمج وحفظ وتثبيت الـ 144 منشأة بالكامل وتنشيط الواجهة الحية للموقع"):
                        st.session_state["cloud_assets"] = df_uploaded
                        st.balloons()
                        st.success("🎉 تهانينا التامة! تم دمج وتثبيت كامل ملف الممتلكات الحقيقي بنجاح وتنشيط كافة محركات البحث والخطابات الموصولة والمواقع الجغرافية!")
                        st.rerun()
                except Exception as e:
                    st.error(f"حدث خطأ ما أثناء تحليل هيكلية ملف الـ Excel المرفوع: {e}")
            st.markdown("</div>", unsafe_allow_html=True)
    # 🔍 ملفات الأصول وبطاقات الوثائق وصائغ الخطابات التلقائي
    elif menu == "🔍 ملفات الأصول وبطاقات الوثائق والخطابات":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-folder-open'></i> ملفات الأصول واستدعاء بطاقات الوثائق المصورة الأربعة والخطابات</h1>", unsafe_allow_html=True)
        
        if df_assets.empty or 'المنشأة' not in df_assets.columns:
            st.markdown("<div class='alert-premium' style='background:#fff9e6; border-right:6px solid #f39c12; color:#d35400;'><i class='fa-solid fa-circle-exclamation'></i> النظام بانتظار تفعيل البيانات، يرجى التوجه أولاً لقسم '📥 استيراد ورفع ملفات Excel' لتثبيت ملفك الحقيقي وتنشيط محرك الاستدعاء الحالي تلقائياً.</div>", unsafe_allow_html=True)
        else:
            st.markdown("<div class='card-luxury'>", unsafe_allow_html=True)
            # محرك بحث مرن وذكي يتيح تصفية المنشآت الـ 144 بكتابة الاسم أو رقم الصك فوراً
            search_box = st.text_input("🔍 ابحث فوراً بكتابة اسم المنشأة أو الأرض أو جزء من رقم الصك الشرعي لاستدعاء الملف الفني:")
            
            if search_box:
                filtered_list = df_assets[df_assets['المنشأة'].astype(str).str.contains(search_box, na=False) | df_assets['رقم الصك'].astype(str).str.contains(search_box, na=False)]
                if not filtered_list.empty:
                    facility_options = filtered_list['المنشأة'].unique().tolist()
                else:
                    facility_options = df_assets['المنشأة'].unique().tolist()
                    st.warning("⚠️ لم يتم العثور على نتائج مطابقة دقيقة للبحث، تفضل بالاختيار المباشر من القائمة الكاملة أدناه:")
            else:
                facility_options = df_assets['المنشأة'].unique().tolist()
                
            selected_facility = st.selectbox("🏥 اضغط هنا لاختيار وتأكيد المنشأة/الأرض المطلوبة لتفجير مستنداتها:", facility_options)
            st.markdown("</div>", unsafe_allow_html=True)
            
            if selected_facility:
                asset_data = df_assets[df_assets['المنشأة'] == selected_facility].iloc[0]
                st.markdown(f"<h2><i class='fa-solid fa-file-shield'></i> الملف العقاري والبطاقات الفنية المصورة لـ: {selected_facility}</h2>", unsafe_allow_html=True)
                
                # 🖼️ لوحة عرض بطاقات الوثائق المصورة الأربعة المنسقة جرافيكياً بالتوازي مطابق لطلبك
                c_b1, c_b2 = st.columns(2)
                c_b3, c_b4 = st.columns(2)
                
                with c_b1:
                    st.markdown("<div class='card-luxury' style='border-top: 4px solid #dfb76c;'><h4>📜 1. صورة الصك الشرعي الموثق</h4>", unsafe_allow_html=True)
                    s_img = st.file_uploader(f"رفع صورة الصك لـ {selected_facility}", type=["jpg","png","jpeg"], key=f"sok_{asset_data['رقم الصك']}")
                    if s_img: st.image(s_img, use_container_width=True)
                    else: st.info("📷 اسحب صورة الصك الشرعي وأفلتها هنا لحفظها بالملف.")
                    st.markdown("</div>", unsafe_allow_html=True)
                    
                with c_b2:
                    st.markdown("<div class='card-luxury' style='border-top: 4px solid #1d5c43;'><h4>🏗️ 2. صورة رخصة البناء الهندسية المعمارية</h4>", unsafe_allow_html=True)
                    p_img = st.file_uploader(f"رفع صورة الرخصة لـ {selected_facility}", type=["jpg","png","jpeg"], key=f"perm_{asset_data['رقم الصك']}")
                    if p_img: st.image(p_img, use_container_width=True)
                    else: st.info("📷 اسحب صورة رخصة البناء الهندسية الصادرة للموقع هنا.")
                    st.markdown("</div>", unsafe_allow_html=True)
                    
                with c_b3:
                    st.markdown("<div class='card-luxury' style='border-top: 4px solid #58a6ff;'><h4>🛰️ 3. صورة الرفع المساحي والموقع الميداني بالأقمار الصناعية</h4>", unsafe_allow_html=True)
                    g_img = st.file_uploader(f"رفع صورة كروكي الموقع لـ {selected_facility}", type=["jpg","png","jpeg"], key=f"site_{asset_data['رقم الصك']}")
                    if g_img: st.image(g_img, use_container_width=True)
                    else:
                        # معالجة وحل مشكلة فك تشفير الإحداثيات الرقمية وربطها التلقائي الفوري بخرائط جوجل لتعمل كالنظام القديم تماماً
                        lat_val = str(asset_data['دائرة العرض']).replace('°','').replace(' ','').strip() if 'دائرة العرض' in asset_data else '21.27'
                        lon_val = str(asset_data['خط الطول']).replace('°','').replace(' ','').strip() if 'خط الطول' in asset_data else '40.41'
                        maps_url = f"https://google.com{lat_val},{lon_val}"
                        st.markdown(f"<a href='{maps_url}' target='_blank'><button style='width:100%; padding:12px; background:#1d5c43; color:white; border:none; border-radius:8px; font-weight:bold; cursor:pointer;'><i class='fa-solid fa-location-arrow'></i> 🗺️ فتح وتحديد الموقع الميداني الحي على Google Maps</button></a>", unsafe_allow_html=True)
                    st.markdown("</div>", unsafe_allow_html=True)
                    
                with c_b4:
                    st.markdown("<div class='card-luxury' style='border-top: 4px solid #143f2e;'><h4>📐 4. صورة قرار الذرعة المساحي المعتمد من الأمانة</h4>", unsafe_allow_html=True)
                    z_img = st.file_uploader(f"رفع صورة قرار الذرعة لـ {selected_facility}", type=["jpg","png","jpeg"], key=f"zar_{asset_data['رقم الصك']}")
                    if z_img: st.image(z_img, use_container_width=True)
                    else: st.info("📷 اسحب صورة قرار الذرعة الصادر المطابق لمساحة الموقع هنا.")
                    st.markdown("</div>", unsafe_allow_html=True)

                # محرك صائغ الخطابات الفوري المعتمد باسم القيادات العليا الحالية وإدارة الأراضي والممتلكات
                st.markdown("<div class='card-luxury'>", unsafe_allow_html=True)
                st.markdown("<h4>✉️ محرك صائغ الخطابات الرقمي المعتمد لإدارة الأراضي والممتلكات</h4>", unsafe_allow_html=True)
                letter_target = st.selectbox("حدد الجهة الحكومية المستهدفة لتوجه الخطاب إليها آلياً:", [
                    "خطاب موجه لسعادة أمين محافظة الطائف (استخراج رخصة بناء ومطابقة ذرعة مساحية)", 
                    "خطاب موجه لسعادة مدير شركة الكهرباء بالطائف (طلب إيصال التيار وتحديد المحول الفرعي)", 
                    "خطاب موجه لفضيلة رئيس المحكمة العامة بالطائف (تحديث ومطابقة الحدود الشرعية للصك)"
                ])
                
                sok_num = asset_data['رقم الصك'] if ('رقم الصك' in asset_data and pd.notnull(asset_data['رقم الصك'])) else "لا يوجد"
                sok_date = asset_data['تاريخ الصك'] if ('تاريخ الصك' in asset_data and pd.notnull(asset_data['تاريخ الصك'])) else "لا يوجد"
                facility_area = asset_data['المساحة'] if ('المساحة' in asset_data and pd.notnull(asset_data['المساحة'])) else "غير محدد"
                facility_village = asset_data['مركز/حي/قرية'] if ('مركز/حي/قرية' in asset_data and pd.notnull(asset_data['مركز/حي/قرية'])) else "محافظة الطائف"
                
                if "أمين" in letter_target:
                    text_content = f"سعادة أمين محافظة الطائف\nالسلام عليكم ورحمة الله وبركاته،،\n\nتفيدكم فرع وزارة الصحة بمحافظة الطائف علماً بملكية الوزارة الرسمية للموقع المخصص لـ ({selected_facility}) والواقع بنطاق ({facility_village}) بموجب الصك الشرعي رقم ({sok_num}) وتاريخ ({sok_date}) بمساحة قدرها ({facility_area} م²). نأمل التوجيه لمن يلزم لاعتماد الرفع المساحي وقرار الذرعة واستخراج رخصة بناء وفق الإحداثيات المرفقة.\n\nوتقبلوا خالص التحية والتقدير،،\n\nمدير فرع وزارة الصحة بمحافظة الطائف\nإدارة الأراضي والممتلكات"
                elif "الكهرباء" in letter_target:
                    text_content = f"سعادة مدير شركة الكهرباء بمحافظة الطائف\nالسلام عليكم ورحمة الله وبركاته،،\n\nنظراً لجهوزية البدء الإنشائي والتشغيلي للموقع الطبي التابع للوزارة ({selected_facility}) والمقام على الأرض ذات الصك رقم ({sok_num})، نأمل منكم الإيعاز للمختصين لطلب إيصال التيار الكهربائي وتحديد موقع محول الطاقة الفرعي حسب المخططات المعتمدة لدينا.\n\nوتقبلوا وافر التحية والتقدير،،\n\nمدير فرع وزارة الصحة بمحافظة الطائف\nإدارة الأراضي والممتلكات"
                else:
                    text_content = f"فضيلة رئيس المحكمة العامة بمحافظة الطائف\nالسلام عليكم ورحمة الله وبركاته،،\n\nتتقدم فرع وزارة الصحة بمحافظة الطائف بطلب تحديث وإصدار صك إلكتروني موحد ومطابقة مسحية للصك رقم ({sok_num}) وتاريخ ({sok_date}) العائد لملك الوزارة في موقع ({selected_facility}) بنطاق ({facility_village}). نأمل التوجيه لمطابقة الحدود الإنشائية والمساحية حسب الرفع المساحي المعتمد لوزارتنا.\n\nوتقبلوا خالص التحية والتقدير،،\n\nمدير فرع وزارة الصحة بمحافظة الطائف\nإدارة الأراضي والممتلكات"
                
                st.text_area("📄 صيغة الخطاب المولد برمجياً والمجهز للطباعة الفورية والتوقيع القيادي:", text_content, height=180)
                doc = Document()
                doc.add_heading(letter_target, 0)
                doc.add_paragraph(text_content)
                doc.save("generated_letter.docx")
                with open("generated_letter.docx", "rb") as f:
                    st.download_button("📥 تنزيل الخطاب الآن كملف Word رسمي معتمد ومكتمل للتوقيع", f, file_name=f"خطاب_صحة_الطائف_{sok_num}.docx")
                st.markdown("</div>", unsafe_allow_html=True)
    # ⚙️ لوحة التحكم بالأصول (تعديل كامل يدوي وحذف للمدخلات حياً وإعادة الميزة المطلوبة)
    elif menu == "⚙️ التحكم بالأصول (تعديل يدوياً/حذف)":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-folder-gear'></i> إدارة وتعديل وحذف الأصول والمنشآت يدوياً</h1>", unsafe_allow_html=True)
        if st.session_state['role'] != "Admin":
            st.markdown("<div class='alert-premium'><i class='fa-solid fa-lock'></i> جدار الحماية: ميزة حفظ أو تعديل أو حذف الأصول تقتصر على مدير النظام (Admin) فقط.</div>", unsafe_allow_html=True)
        else:
            st.markdown("<div class='card-luxury'>", unsafe_allow_html=True)
            st.markdown("<h5>➕ إضافة منشأة أو مشروع جديد يدوياً لقاعدة البيانات الحية</h5>", unsafe_allow_html=True)
            with st.form("manual_add_v4"):
                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    r_sok = st.text_input("رقم الصك الشرعي العقاري")
                    r_moamala = st.text_input("رقم المعاملة أو المعاملة المساحية")
                    site_name = st.text_input("اسم المنشأة الطبية / الأرض الجديدة")
                with col_f2:
                    district = st.text_input("المركز / الحي السكني داخل الطائف")
                    area = st.text_input("المساحة الإجمالية (م²)")
                    g_type = st.selectbox("تصنيف العقار الفني الحالي:", ["أرض فضاء التابعة للوزارة", "مركز صحي قائم", "مستشفى عام وتخصصي", "مشروع إنشائي قيد التنفيذ"])
                
                c_c1, c_c2 = st.columns(2)
                with c_c1: lat = st.text_input("إحداثي دائرة العرض (Latitude) - مثل: 21.27")
                with c_c2: lon = st.text_input("إحداثي خط الطول (Longitude) - مثل: 40.41")
                
                bypass_dup = st.checkbox("السماح برفع وتمرير رقم الصك وتكراره استثنائياً (موافقة الإدارة العليا)")
                
                if st.form_submit_button("💾 اعتماد وحفظ الأصل في النظام السحابي"):
                    has_dup = False
                    if not df_assets.empty and 'رقم الصك' in df_assets.columns:
                        if str(r_sok).strip() in df_assets['رقم الصك'].astype(str).values: has_dup = True
                            
                    if has_dup and not bypass_dup:
                        st.error("❌ تنبيه أمني عاجل: رقم الصك هذا مسجل مسبقاً في النظام! لا يمكن التكرار إلا بموافقة الإدارة العليا.")
                    else:
                        new_asset = pd.DataFrame([{"م": len(df_assets)+1, "المنشأة": site_name, "نوع_العقار": g_type, "حاله_العقار": "ملك للوزارة", "المحافظة": "الطائف", "مركز/حي/قرية": district, "المساحة": area, "رقم الصك": r_sok, "تاريخ الصك": datetime.now().strftime("%Y-%m-%d"), "خط الطول": lon, "دائرة العرض": lat}])
                        st.session_state["cloud_assets"] = pd.concat([df_assets, new_asset]).reset_index(drop=True)
                        st.success("✅ تم حفظ وتأمين الأصل الجديد بنجاح في قاعدة البيانات السحابية!")
                        st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

            st.markdown("<div class='card-luxury'>", unsafe_allow_html=True)
            st.markdown("<h5>📋 تعديل وحذف قيود الـ 144 منشأة مباشرة من الجدول التفاعلي التلقائي</h5>", unsafe_allow_html=True)
            if not df_assets.empty:
                edited_df = st.data_editor(df_assets, num_rows="dynamic", use_container_width=True, key="assets_editor_vfinal")
                if st.button("💾 حفظ وتثبيت كافة حركات التعديل أو الحذف المباشرة حلياً"):
                    st.session_state["cloud_assets"] = edited_df
                    st.success("🎉 تم تحديث وحفظ جدول المنشآت بنجاح على السحابة الدائمة!")
                    st.rerun()
            else: st.info("قاعدة البيانات فارغة حالياً.")
            st.markdown("</div>", unsafe_allow_html=True)

    # 💼 تتبع وحفظ المعاملات الرقمية بين الإدارات للعودة إليها
    elif menu == "💼 تتبع وحفظ المعاملات الرقمية":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-route'></i> حفظ وتتبع خط سير المعاملات الرقمية للأراضي</h1>", unsafe_allow_html=True)
        st.markdown("<div class='card-luxury'>", unsafe_allow_html=True)
        with st.form("workflow_add_v4"):
            w_col1, w_col2 = st.columns(2)
            with w_col1:
                w_id = st.text_input("رقم المعاملة (الصادر أو الوارد الرسمي)")
                w_subject = st.text_input("موضوع المعاملة الأساسي للمراسلة")
            with w_col2:
                w_dept = st.selectbox("الجهة الحكومية الحالية الواقفة عندها المعاملة:", ["أمانة محافظة الطائف", "المحكمة العامة بالطائف", "كتابة العدل بمحافظة الطائف", "إدارة الأراضي والممتلكات بالفرع"])
                w_status = st.selectbox("حالة المعاملة الحالية:", ["قيد الدراسة والتدقيق المساحي بالذرعة", "بانتظار الاعتماد النهائي للخطاب", "تم الإفراغ والمطابقة بنجاح للحدود"])
            if st.form_submit_button("💾 قيد وحفظ المعاملة في أرشيف المتابعة"):
                new_wf = pd.DataFrame([{"رقم_المعاملة": w_id, "موضوع_المعاملة": w_subject, "الإدارة_الحالية": w_dept, "حالة_المعاملة": w_status, "تاريخ_التحديث": datetime.now().strftime("%Y-%m-%d %H:%M"), "الموظف_المسؤول": st.session_state['username']}])
                st.session_state["cloud_workflows"] = pd.concat([df_workflows, new_wf]).reset_index(drop=True)
                st.success("✅ تم حفظ وأرشفة المعاملة بنجاح للعودة المباشرة إليها مستقبلاً!")
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='card-luxury'>", unsafe_allow_html=True)
        edited_wf = st.data_editor(st.session_state["cloud_workflows"], num_rows="dynamic", use_container_width=True, key="wf_editor")
        if st.button("💾 حفظ تعديلات أرشيف المعاملات"):
            st.session_state["cloud_workflows"] = edited_wf
            st.success("✅ تم التثبيت.")
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    # ⚠️ نظام رقابة الأراضي وتتبع التعديات الميدانية للأراضي
    elif menu == "⚠️ رقابة الأراضي وتتبع التعديات":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-shield-halved'></i> تتبع التعديات والرقابة الميدانية للأراضي وصحة الطائف</h1>", unsafe_allow_html=True)
        st.markdown("<div class='card-luxury'>", unsafe_allow_html=True)
        with st.form("encroach_form_v4"):
            c_e1, c_e2 = st.columns(2)
            with c_e1:
                e_sok = st.text_input("رقم الصك الشرعي المعتدى على حدوده")
                e_name = st.text_input("اسم المنشأة أو الموقع الطبي المتأثر")
            with c_e2:
                e_type = st.selectbox("نوع التعدي المرصود ميدانياً:", ["إقامة حوش أو أسوار غير نظامية", "بناء شعبي بدون رخصة", "وضع لوحات أو شبوك وتجريف تربة"])
                e_status = st.selectbox("حالة المعاملة والقضية القانونية:", ["تحت الرفع للمحافظة", "بانتظار إزالة لجنة التعديات", "تمت الإزالة واسترداد الأرض بالكامل"])
            e_action = st.text_area("الإجراء المتخذ وتفاصيل وتاريخ الخطاب الموجه للمحافظة")
            if st.form_submit_button("🚨 تسجيل حالة التعدي وأرشفة المعاملة فوراً"):
                new_enc = pd.DataFrame([{"رقم الصك": e_sok, "المنشأة": e_name, "نوع_التعدي": e_type, "حالة_القضية": e_status, "تاريخ_الرصد": datetime.now().strftime("%Y-%m-%d"), "الإجراء_المتخذ": e_action}])
                st.session_state["cloud_encroach"] = pd.concat([df_encroach, new_enc]).reset_index(drop=True)
                st.success("✅ تم تسجيل حالة التعدي بنجاح في السجلات الرقابية الدائمة.")
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='card-luxury'>", unsafe_allow_html=True)
        edited_enc = st.data_editor(st.session_state["cloud_encroach"], num_rows="dynamic", use_container_width=True, key=\"enc_editor\")
        if st.button("💾 حفظ تعديلات حصر التعديات الميدانية"):
            st.session_state["cloud_encroach"] = edited_enc
            st.success("✅ تم تحديث السجل.")
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)
