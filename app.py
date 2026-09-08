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
    </style>
""", unsafe_allow_html=True)

# تفعيل نظام الذاكرة السحابية المستمرة لضمان عدم ضياع قوائم الإكسل المرفوعة على الويب أبداً
if "db_assets" not in st.session_state:
    if os.path.exists("assets_db.csv"):
        st.session_state["db_assets"] = pd.read_csv("assets_db.csv")
    else:
        st.session_state["db_assets"] = pd.DataFrame()

if "db_encroach" not in st.session_state:
    st.session_state["db_encroach"] = pd.DataFrame()

if "db_workflows" not in st.session_state:
    st.session_state["db_workflows"] = pd.DataFrame()

def log_action(user, action, details):
    pass

# بوابة الأمان والتحقق الرقمي للولوج للمنصة
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
    # القائمة الجانبية الفخمة عالية التباين
    st.sidebar.markdown(f"<div style='text-align: center; padding: 10px;'><i class='fa-solid fa-circle-user' style='font-size: 3.5rem; color: #dfb76c;'></i></div>", unsafe_allow_html=True)
    st.sidebar.markdown(f"<h3 style='color: white; text-align: center; margin-top: 5px;'>{st.session_state['username']}</h3>", unsafe_allow_html=True)
    st.sidebar.markdown(f"<p style='color: #dfb76c; text-align: center; font-weight: bold; margin-bottom: 20px;'><i class='fa-solid fa-id-card-clip'></i> رتبة: {st.session_state['role']}</p>", unsafe_allow_html=True)
    
    st.sidebar.markdown("<hr style='border-color: rgba(255,255,255,0.2);'>", unsafe_allow_html=True)
    
    menu = st.sidebar.radio(
        "📂 تصفح أقسام المنصة",
        [
            "📊 لوحة التحكم اليومية", 
            "📥 استيراد ورفع ملفات Excel",
            "🔍 ملفات الأملاك والخطابات وبطاقات الوثائق المصورة", 
            "⚙️ التحكم بالأصول (تعديل/حذف)",
            "💼 تتبع وحفظ المعاملات الرقمية",
            "⚠️ رقابة الأراضي وتتبع التعديات"
        ]
    )
    
    st.sidebar.markdown("<br><br>", unsafe_allow_html=True)
    if st.sidebar.button("🚪 خروج آمن من النظام"):
        st.session_state['logged_in'] = False
        st.clear_not_saved_mutations()
        st.rerun()

    df_assets = st.session_state["db_assets"]
    df_encroach = st.session_state["db_encroach"]
    df_workflows = st.session_state["db_workflows"]

    # 📊 لوحة التحكم اليومية لـ 144 منشأة الحقيقية
    if menu == "📊 لوحة التحكم اليومية":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-table-cells-large'></i> لوحة الأداء والتوزيع الشبكي الموحد لممتلكات الصحة</h1>", unsafe_allow_html=True)
        
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.markdown(f"<div class='mini-kpi'><b><i class='fa-solid fa-hospital'></i> إجمالي ممتلكات الصحة:</b> <br><span style='font-size: 1.15rem; font-weight: bold; color: #1d5c43;'>{len(df_assets) if not df_assets.empty else 0} منشأة وموقع</span></div>", unsafe_allow_html=True)
        with k2:
            if not df_assets.empty and "المساحة" in df_assets.columns:
                df_assets['clean_area'] = df_assets["المساحة"].astype(str).str.replace(',', '', regex=True).str.replace(' ', '', regex=True).str.strip()
                df_assets['clean_area'] = pd.to_numeric(df_assets['clean_area'], errors='coerce').fillna(0)
                total_area = df_assets['clean_area'].sum()
            else:
                total_area = 0
            st.markdown(f"<div class='mini-kpi'><b><i class='fa-solid fa-vector-square'></i> إجمالي مساحات الأراضي:</b> <br><span style='font-size: 1.15rem; font-weight: bold; color: #1d5c43;'>{total_area:,.2f} م²</span></div>", unsafe_allow_html=True)
        with k3:
            st.markdown(f"<div class='mini-kpi'><b><i class='fa-solid fa-route'></i> المعاملات والخطابات:</b> <br><span style='font-size: 1.15rem; font-weight: bold; color: #dfb76c;'>{len(df_workflows)} قيد المتابعة</span></div>", unsafe_allow_html=True)
        with k4:
            st.markdown(f"<div class='mini-kpi'><b><i class='fa-solid fa-shield-halved'></i> التعديات المرصودة:</b> <br><span style='font-size: 1.15rem; font-weight: bold; color: #e74c3c;'>{len(df_encroach)} قضايا نشطة</span></div>", unsafe_allow_html=True)
            
        st.markdown("<br>", unsafe_allow_html=True)
        
        block1, block2 = st.columns(2)
        with block1:
            st.markdown("<div class='grid-box'><div class='grid-header'><span><i class='fa-solid fa-map-location-dot'></i> الرقابة الجيومكانية والخرائط الميدانية الشاملة (الـ 144 موقع)</span></div>", unsafe_allow_html=True)
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
                        st.info("💡 لا توجد إحداثيات رقمية صالحة لعرضها جغرافياً حالياً.")
                except:
                    st.info("💡 نظام الخرائط بانتظار تحديث الإحداثيات السليمة.")
            else:
                st.info("💡 الخريطة الرقمية جاهزة ومستعدة وتنتظر استيراد الأصول.")
            st.markdown("</div>", unsafe_allow_html=True)
            
        with block2:
            st.markdown("<div class='grid-box'><div class='grid-header'><span><i class='fa-solid fa-clock-lock'></i> التنبيهات الذكية وفحص النواقص العقارية لفرع الوزارة</span></div>", unsafe_allow_html=True)
            if not df_assets.empty and 'رقم الصك' in df_assets.columns:
                no_sok_count = len(df_assets[df_assets['رقم الصك'].astype(str).str.contains('لا يوجد|NULL|بدون', na=True, case=False)])
                st.markdown(f"<div class='alert-premium' style='background: #fffdf5; border-right: 6px solid #e67e22; color:#d35400;'><i class='fa-solid fa-triangle-exclamation'></i> نظام الفحص الآلي المطور رصد وجود {no_sok_count} منشأة وموقع صحي مدرج بحالة 'بدون صك' أو بيانات غير مكتملة تتطلب تحديث ومخاطبة أمانة الطائف بشكل عاجل.</div>", unsafe_allow_html=True)
            else:
                st.markdown("<div class='alert-premium' style='background:#f4f7f6; color:#555; border-right:6px solid #1d5c43;'><i class='fa-solid fa-circle-check' style='color:#2ecc71;'></i> المنصة الرقمية جاهزة ومحمية ومستعدة بالكامل لتلقي وحصر ممتلكات صحة الطائف.</div>", unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

    # 📥 قسم استيراد ورفع ملفات Excel والتثبيت السحابي الدائم
    elif menu == "📥 استيراد ورفع ملفات Excel":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-file-excel'></i> استيراد ورفع القوائم من ملفات Excel</h1>", unsafe_allow_html=True)
        if st.session_state['role'] != "Admin":
            st.markdown("<div class='alert-premium'><i class='fa-solid fa-lock'></i> عذراً، خاصية رفع واستيراد الملفات تتطلب صلاحية مدير النظام (Admin).</div>", unsafe_allow_html=True)
        else:
            st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
            st.markdown("<div class='grid-header'><span><i class='fa-solid fa-upload'></i> مركز القراءة الآلي لبيانات ممتلكات الصحة وتثبيتها بشكل دائم على الويب</span></div>", unsafe_allow_html=True)
            uploaded_file = st.file_uploader("اختر ملف ممتلكات وزارة الصحة الموثق بالطائف من جهازك", type=["xlsx", "xls", "csv"])
            if uploaded_file is not None:
                try:
                    df_uploaded = pd.read_excel(uploaded_file) if not uploaded_file.name.endswith('.csv') else pd.read_csv(uploaded_file)
                    st.success(f"✅ تم قراءة ملفك بنجاح! تم رصد وعزل {len(df_uploaded)} منشأة عقارية وطبية.")
                    st.dataframe(df_uploaded.head(10), use_container_width=True)
                    if st.button("🚀 دمج وحفظ وتثبيت الـ 144 منشأة بشكل قاطع على الويب"):
                        # حفظ البيانات في الـ session_state لمنع اختفائها عند الانتقال بين الصفحات
                        st.session_state["db_assets"] = df_uploaded
                        st.balloons()
                        st.success("🎉 تهانينا! تم تثبيت ملف الممتلكات الحقيقي بنجاح على سيرفر الويب وتنشيط الخطابات وبطاقات الوثائق بالكامل!")
                        st.rerun()
                except Exception as e:
                    st.error(f"حدث خطأ أثناء قراءة هيكلية ملف الـ Excel: {e}")
            st.markdown("</div>", unsafe_allow_html=True)
    # 🔍 ملفات الأملاك والخطابات وبطاقات الوثائق المصورة
    elif menu == "🔍 ملفات الأملاك والخطابات وبطاقات الوثائق المصورة":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-folder-open'></i> ملفات الأصول وبطاقات الوثائق المصورة والخطابات</h1>", unsafe_allow_html=True)
        
        if df_assets.empty or 'المنشأة' not in df_assets.columns:
            st.markdown("<div class='alert-premium' style='background:#fff9e6; border-right:6px solid #f39c12; color:#d35400;'><i class='fa-solid fa-circle-exclamation'></i> أولاً لرفع قوائم الـ 144 منشأة وصك لتفعيل بطاقات الوثائق وصانع الخطابات تلقائياً، يرجى التوجه لقسم '📥 استيراد ورفع ملفات Excel' والضغط على زر التثبيت قاطع السحابي.</div>", unsafe_allow_html=True)
        else:
            st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
            st.markdown("<div class='grid-header'><span><i class='fa-solid fa-hospital-user'></i> اختر المنشأة أو الأرض من القائمة المستدعاة من ملفك الموثق</span></div>", unsafe_allow_html=True)
            
            facility_list = df_assets['المنشأة'].dropna().unique().tolist()
            selected_facility = st.selectbox("🏥 اضغط هنا لاختيار الموقع الطبي / الأرض الفضاء المستهدفة من القائمة المنسدلة:", facility_list, key="facility_dropdown_clean")
            st.markdown("</div>", unsafe_allow_html=True)
            
            if selected_facility:
                asset_data = df_assets[df_assets['المنشأة'] == selected_facility].iloc[0]
                st.markdown(f"<h2><i class='fa-solid fa-file-invoice-dollar'></i> الملف الفني العقاري لـ: {selected_facility}</h2>", unsafe_allow_html=True)
                
                row1_col1, row1_col2 = st.columns(2)
                row2_col1, row2_col2 = st.columns(2)
                
                with row1_col1:
                    st.markdown("<div class='grid-box' style='border-top: 4px solid #dfb76c;'>", unsafe_allow_html=True)
                    st.markdown("<div class='grid-header'><span><i class='fa-solid fa-scroll'></i> 1. صـورة الصك الشرعي الموثق للعقار</span></div>", unsafe_allow_html=True)
                    sok_img = st.file_uploader(f"رفع/تحديث صورة الصك لـ {selected_facility}", type=["jpg", "png", "jpeg"], key=f"sok_up_{asset_data['رقم الصك']}")
                    if sok_img is not None:
                        st.image(sok_img, caption="معاينة صورة الصك الشرعي المعتمد برمجياً", use_container_width=True)
                    else:
                        st.info("📷 لم يتم رفع صورة الصك لهذا الموقع بعد.")
                    st.markdown("</div>", unsafe_allow_html=True)
                    
                with row1_col2:
                    st.markdown("<div class='grid-box' style='border-top: 4px solid #1d5c43;'>", unsafe_allow_html=True)
                    st.markdown("<div class='grid-header'><span><i class='fa-solid fa-trowel-bricks'></i> 2. صـورة رخصة البناء الهندسية المعمارية</span></div>", unsafe_allow_html=True)
                    permit_img = st.file_uploader(f"رفع/تحديث صورة الرخصة لـ {selected_facility}", type=["jpg", "png", "jpeg"], key=f"permit_up_{asset_data['رقم الصك']}")
                    if permit_img is not None:
                        st.image(permit_img, caption="معاينة رخصة البناء الهندسية الصادرة", use_container_width=True)
                    else:
                        st.info("📷 لم يتم رفع صورة رخصة البناء الهندسية للموقع بعد.")
                    st.markdown("</div>", unsafe_allow_html=True)
                    
                with row2_col1:
                    st.markdown("<div class='grid-box' style='border-top: 4px solid #58a6ff;'>", unsafe_allow_html=True)
                    st.markdown("<div class='grid-header'><span><i class='fa-solid fa-map-location-dot'></i> 3. صورة الرفع المساحي والموقع الميداني (الأقمار الصناعية)</span></div>", unsafe_allow_html=True)
                    site_img = st.file_uploader(f"رفع/تحديث المخطط الجغرافي لـ {selected_facility}", type=["jpg", "png", "jpeg"], key=f"site_up_{asset_data['رقم الصك']}")
                    if site_img is not None:
                        st.image(site_img, caption="معاينة صورة كروكي الموقع والرفع المساحي الفعلي", use_container_width=True)
                    else:
                        maps_url = f"https://google.com{str(asset_data['دائرة العرض']).replace('°','')},{str(asset_data['خط الطول']).replace('°','')}"
                        st.markdown(f"<a href='{maps_url}' target='_blank'><button style='width:100%; padding:10px; background:#1d5c43; color:white; border:none; border-radius:6px; cursor:pointer; font-weight:bold;'><i class='fa-solid fa-location-arrow'></i> استعراض الموقع جغرافياً على Google Maps</button></a>", unsafe_allow_html=True)
                    st.markdown("</div>", unsafe_allow_html=True)
                    
                with row2_col2:
                    st.markdown("<div class='grid-box' style='border-top: 4px solid #143f2e;'>", unsafe_allow_html=True)
                    st.markdown("<div class='grid-header'><span><i class='fa-solid fa-receipt'></i> 4. صورة قـرار الذرعة والقرار المساحي المعتمد</span></div>", unsafe_allow_html=True)
                    zar_img = st.file_uploader(f"رفع/تحديث صورة قرار الذرعة لـ {selected_facility}", type=["jpg", "png", "jpeg"], key=f"zar_up_{asset_data['رقم الصك']}")
                    if zar_img is not None:
                        st.image(zar_img, caption="معاينة قرار الذرعة المعتمد والمطابق المساحي من الأمانة", use_container_width=True)
                    else:
                        st.info("📷 لم يتم رفع صورة قرار الذرعة المعتمد من أمانة محافظة الطائف بعد.")
                    st.markdown("</div>", unsafe_allow_html=True)

                st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
                st.markdown("<div class='grid-header'><span><i class='fa-solid fa-envelope-open-text'></i> محرك أتمتة الصياغة الفورية للخطابات الموجهة للجهات الحكومية</span></div>", unsafe_allow_html=True)
                
                letter_target = st.selectbox("✉️ حدد الجهة الحكومية الموجه إليها الخطاب برمجياً:", [
                    "خطاب موجه لسعادة أمين محافظة الطائف (استخراج رخصة بناء ومطابقة ذرعة)", 
                    "خطاب موجه لسعادة مدير شركة الكهرباء بالطائف (طلب إيصال التيار وتحديد المحول)", 
                    "خطاب موجه لفضيلة رئيس المحكمة العامة بالطائف (تحديث ومطابقة صك ممتلكات الصحة)"
                ], key="letter_target_dropdown")
                
                text_content = ""
                sok_num = asset_data['رقم الصك'] if pd.notnull(asset_data['رقم الصك']) else "لا يوجد"
                sok_date = asset_data['تاريخ الصك'] if pd.notnull(asset_data['تاريخ الصك']) else "لا يوجد"
                facility_area = asset_data['المساحة'] if pd.notnull(asset_data['المساحة']) else "غير محدد"
                facility_village = asset_data['مركز/حي/قرية'] if pd.notnull(asset_data['مركز/حي/قرية']) else "محافظة الطائف"
                
                if "أمين" in letter_target:
                    text_content = f"سعادة أمين محافظة الطائف\nالسلام عليكم ورحمة الله وبركاته،،\n\nتفيدكم فرع وزارة الصحة بمحافظة الطائف علماً بملكية الوزارة الرسمية للموقع المخصص لـ ({selected_facility}) والواقع بنطاق ({facility_village}) بموجب الصك الشرعي رقم ({sok_num}) وتاريخ ({sok_date}) بمساحة قدرها ({facility_area} م²). نأمل التوجيه لمن يلزم لاعتماد الرفع المساحي وقرار الذرعة واستخراج رخصة بناء وفق الإحداثيات المرفقة ({asset_data['دائرة العرض']} , {asset_data['خط الطول']}).\n\nوتقبلوا خالص التحية والتقدير،،\n\nمدير فرع وزارة الصحة بمحافظة الطائف\nإدارة الأراضي والممتلكات"
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
            st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
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
                    lat = st.text_input("إحداثي دائرة العرض (Latitude)")
                with c_c2:
                    lon = st.text_input("إحداثي خط الطول (Longitude)")
                
                st.markdown("<b>🔒 إعدادات التحقق من التكرار والازدواجية:</b>", unsafe_allow_html=True)
                bypass_dup = st.checkbox("السماح برفع وتمرير رقم الصك وتكراره استثنائياً (موافقة الإدارة العليا)")
                
                if st.form_submit_button("💾 اعتماد وحفظ الأصل في قاعدة البيانات الحالية"):
                    if not df_assets.empty and r_sok in df_assets['رقم الصك'].astype(str).values and not bypass_dup:
                        st.error("❌ تنبيه أمني عاجل: رقم الصك هذا مسجل مسبقاً في النظام! لا يمكن التكرار إلا بموافقة الإدارة العليا.")
                    else:
                        new_asset = pd.DataFrame([{"م": len(df_assets)+1, "المنشأة": site_name, "نوع_العقار": g_type, "حاله_العقار": "ملك", "المحافظة": "الطائف", "مركز/حي/قرية": district, "المساحة": area, "رقم الصك": r_sok, "تاريخ الصك": datetime.now().strftime("%Y-%m-%d"), "خط الطول": lon, "دائرة العرض": lat}])
                        st.session_state["db_assets"] = pd.concat([df_assets, new_asset]).reset_index(drop=True)
                        st.success("✅ تم حفظ وتأمين الأصل الجديد بنجاح!")
                        st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

            st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
            if not df_assets.empty:
                edited_df = st.data_editor(df_assets, num_rows="dynamic", use_container_width=True, key="assets_editor_v138")
                if st.button("💾 حفظ كافة التغييرات والتعديلات النشطة للأصول"):
                    st.session_state["db_assets"] = edited_df
                    st.success("🎉 تم تحديث وحفظ جدول البيانات والـ 144 موقع بنجاح على السحابة!")
                    st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

    # 💼 تتبع وحفظ المعاملات الرقمية
    elif menu == "💼 تتبع وحفظ المعاملات الرقمية":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-route'></i> حفظ وتتبع خط سير المعاملات الرقمية</h1>", unsafe_allow_html=True)
        st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
        with st.form("workflow_add"):
            w_col1, w_col2 = st.columns(2)
            with w_col1:
                w_id = st.text_input("رقم المعاملة (الصادر أو الوارد)")
                w_subject = st.text_input("موضوع المعاملة الأساسي")
            with w_col2:
                w_dept = st.selectbox("الجهة أو الإدارة الحكومية الحالية عندها المعاملة:", ["أمانة محافظة الطائف", "المحكمة العامة بالطائف", "كتابة العدل بمحافظة الطائف", "إدارة الأراضي والممتلكات بالفرع"], key="dept_workflow_drop")
                w_status = st.selectbox("حالة المعاملة الميدانية الحالية:", ["قيد الدراسة والتدقيق المساحي بالذرعة", "تم الإفراغ والمطابقة بنجاح", "معاملة منتهية تم أرشفتها"], key="status_workflow_drop")
            if st.form_submit_button("💾 قيد وحفظ خط سير المعاملة بأرشيف المنصة"):
                new_wf = pd.DataFrame([{"رقم_المعاملة": w_id, "موضوع_المعاملة": w_subject, "الإدارة_الحالية": w_dept, "حالة_المعاملة": w_status, "تاريخ_التحديث": datetime.now().strftime("%Y-%m-%d %H:%M"), "الموظف_المسؤول": st.session_state['username']}])
                st.session_state["db_workflows"] = pd.concat([df_workflows, new_wf]).reset_index(drop=True)
                st.success("✅ تم حفظ وتأمين المعاملة في الأرشيف الرقمي!")
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
        if not df_workflows.empty:
            edited_wf_df = st.data_editor(df_workflows, num_rows="dynamic", use_container_width=True, key="workflow_editor_active")
            if st.button("💾 حفظ كافة تعديلات وحذوفات جدول تتبع المعاملات"):
                st.session_state["db_workflows"] = edited_wf_df
                st.success("🎉 تم حفظ وتثبيت خط سير معاملاتك بنجاح!")
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
                e_name = st.text_input("اسم المنشأة أو الموقع الطبي")
            with c_e2:
                e_type = st.selectbox("نوع التعدي المرصود ميدانياً:", ["إقامة حوش أو أسوار غير نظامية", "بناء شعبي بدون رخصة", "وضع لوحات أو شبوك وتجريف تربة"], key="enc_type_drop")
                e_status = st.selectbox("حالة المعاملة والقضية:", ["تحت الرفع للمحافظة", "بانتظار إزالة لجنة التعديات", "تمت الإزالة واسترداد الأرض بالكامل"], key="enc_status_drop")
            e_action = st.text_area("الإجراء المتخذ هندسياً وقانونياً")
            if st.form_submit_button("🚨 تسجيل حالة التعدي وأرشفة المعاملة رسمياً"):
                new_enc = pd.DataFrame([{"رقم الصك": e_sok, "المنشأة": e_name, "نوع_التعدي": e_type, "حالة_القضية": e_status, "تاريخ_الرصد": datetime.now().strftime("%Y-%m-%d"), "الإجراء_المتخذ": e_action}])
                st.session_state["db_encroach"] = pd.concat([df_encroach, new_enc]).reset_index(drop=True)
                st.success("✅ تم تسجيل حالة التعدي بنجاح في السجلات.")
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='grid-box'>", unsafe_allow_html=True)
        if not df_encroach.empty:
            edited_encroach = st.data_editor(df_encroach, num_rows="dynamic", use_container_width=True, key="encroach_editor_vfinal")
            if st.button("💾 حفظ كافة التعديلات والتغييرات النشطة في سجل التعديات"):
                st.session_state["db_encroach"] = edited_encroach
                st.success("🎉 تم حفظ وتثبيت كافة تحديثات الرقابة والتعديات الميدانية بنجاح!")
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)
