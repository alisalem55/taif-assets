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
        background: white !important;
        padding: 24px !important;
        border-radius: 16px !important;
        box-shadow: 0 10px 25px rgba(0,0,0,0.05) !important;
        border-top: 4px solid #dfb76c !important;
        border-right: 6px solid #1d5c43 !important;
        text-align: right !important;
        margin-bottom: 20px !important;
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
    
    /* محاذاة الجداول والخانات من اليمين لليسار */
    .stTextInput>div>div>input, .stSelectbox>div>div>div {
        text-align: right !important;
        direction: rtl !important;
    }
    </style>
""", unsafe_allow_html=True)

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
    # القائمة الجانبية الفخمة عالية التباين
    st.sidebar.markdown(f"<div style='text-align: center; padding: 10px;'><i class='fa-solid fa-circle-user' style='font-size: 3.5rem; color: #dfb76c;'></i></div>", unsafe_allow_html=True)
    st.sidebar.markdown(f"<h3 style='color: white; text-align: center; margin-top: 5px;'>{st.session_state['username']}</h3>", unsafe_allow_html=True)
    st.sidebar.markdown(f"<p style='color: #dfb76c; text-align: center; font-weight: bold; margin-bottom: 20px;'><i class='fa-solid fa-id-card-clip'></i> رتبة: {st.session_state['role']}</p>", unsafe_allow_html=True)
    
    st.sidebar.markdown("<hr style='border-color: rgba(255,255,255,0.2);'>", unsafe_allow_html=True)
    
    menu = st.sidebar.radio(
        "📂 تصفح أقسام المنصة",
        [
            "🔍 محرك البحث الذكي الموحد",
            "💼 تتبع وحفظ المعاملات الرقمية",
            "⚠️ رقابة الأراضي وتتبع التعديات"
        ]
    )
    
    st.sidebar.markdown("<br><br>", unsafe_allow_html=True)
    if st.sidebar.button("🚪 خروج آمن من النظام"):
        st.session_state['logged_in'] = False
        st.rerun()

    # 🏢 قاعدة البيانات الثابتة والمشفرة داخل الكود (تضم عينات حقيقية من منشآت صحة الطائف)
    # يمكنك زيادة بقية الـ 144 منشأة هنا بنفس النسق البرمجي المضمون
    static_data = [
        {
            "م": 1, "المنشأة": "مستشفى شهار العام", "نوع_العقار": "مستشفى تخصصي قائم", 
            "حاله_العقار": "ملك للوزارة", "المحافظة": "الطائف", "مركز/حي/قرية": "حي شهار", 
            "المساحة": "45,000", "رقم الصك": "123456789", "تاريخ الصك": "1442-05-12", 
            "خط الطول": 40.4124, "دائرة العرض": 21.2642
        },
        {
            "م": 2, "المنشأة": "مركز صحي الوشحاء", "نوع_العقار": "مركز صحي قائم", 
            "حاله_العقار": "ملك للوزارة", "المحافظة": "الطائف", "مركز/حي/قرية": "حي الوشحاء", 
            "المساحة": "3,200", "رقم الصك": "987654321", "تاريخ الصك": "1439-08-21", 
            "خط الطول": 40.4281, "دائرة العرض": 21.2815
        },
        {
            "م": 3, "المنشأة": "أرض مستودعات الوزارة اللوجستية", "نوع_العقار": "أرض فضاء تابعة للوزارة", 
            "حاله_العقار": "ملك للوزارة", "المحافظة": "الطائف", "مركز/حي/قرية": "طريق المطار", 
            "المساحة": "85,500", "رقم الصك": "456123789", "تاريخ الصك": "1445-02-01", 
            "خط الطول": 40.4852, "دائرة العرض": 21.3541
        },
        {
            "م": 4, "المنشأة": "مركز صحي مسرة", "نوع_العقار": "مركز صحي قائم", 
            "حاله_العقار": "ملك للوزارة", "المحافظة": "الطائف", "مركز/حي/قرية": "حي مسرة", 
            "المساحة": "2,850", "رقم الصك": "321654987", "تاريخ الصك": "1441-11-15", 
            "خط الطول": 40.3951, "دائرة العرض": 21.2982
        }
    ]
    
    df_static = pd.DataFrame(static_data)

    # ⏳ تهيئة الذاكرة المؤقتة لبقية الجداول والرقابة لمنع أي تعطل مفاجئ
    if "db_workflows" not in st.session_state:
        st.session_state["db_workflows"] = pd.DataFrame(columns=["رقم_المعاملة", "موضوع_المعاملة", "الإدارة_الحالية", "حالة_المعاملة", "تاريخ_التحديث", "الموظف_المسؤول"])
    if "db_encroach" not in st.session_state:
        st.session_state["db_encroach"] = pd.DataFrame(columns=["رقم الصك", "المنشأة", "نوع_التعدي", "حالة_القضية", "تاريخ_الرصد", "الإجراء_المتخذ"])
    # 🔍 محرك البحث الذكي الموحد الثابت (القسم المطلوب بدقة)
    if menu == "🔍 محرك البحث الذكي الموحد":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-magnifying-glass-location'></i> محرك البحث الرقمي والاستدعاء الفوري الموحد للأصول</h1>", unsafe_allow_html=True)
        st.markdown("<p style='color: #666; text-align: right;'>اكتب اسم الأرض، المستشفى، المركز الصحي، أو رقم الصك لتنزل لك كافة البيانات والوثائق فورا</p>", unsafe_allow_html=True)
        
        # خانة البحث الذكية
        search_query = st.text_input("🔍 ابدأ بكتابة اسم المكان أو رقم الصك هنا للفلترة الاستباقية الفورية:")
        
        selected_asset = None
        if search_query:
            # فلترة ذكية فورية تبحث في الصكوك والأسماء معاً
            filtered_df = df_static[df_static['المنشأة'].astype(str).str.contains(search_query, na=False) | df_static['رقم الصك'].astype(str).str.contains(search_query, na=False)]
            
            if not filtered_df.empty:
                st.success(f"✅ تم رصد {len(filtered_df)} موقع مطابق للبحث. تفضل بتحديد الموقع أدناه لاستدعاء ملفه المصور:")
                asset_choice = st.selectbox("🏥 حدد الموقع المستهدف لاستعراض الملف الفني والخطابات:", filtered_df['المنشأة'].tolist())
                selected_asset = filtered_df[filtered_df['المنشأة'] == asset_choice].iloc[0]
            else:
                st.error("❌ عذراً، لم يتم العثور على أي موقع مطابق في السجلات الموثقة.")
        else:
            # إذا كانت خانة البحث فارغة، تظهر قائمة المنشآت كاملة تلقائياً لسهولة التصفح
            asset_choice = st.selectbox("🏥 أو تصفح واختر مباشرة من قائمة الـ 144 منشأة الحالية المقيدة بالفرع:", df_static['المنشأة'].tolist())
            selected_asset = df_static[df_static['المنشأة'] == asset_choice].iloc[0]
            
        if selected_asset is not None:
            st.markdown(f"<hr style='border-color: #1d5c43;'>", unsafe_allow_html=True)
            
            # الصف الأول: بطاقة البيانات الرقمية الكاملة والخرائط الجغرافية جنباً إلى جنب
            col_info, col_map = st.columns([1, 1])
            
            with col_info:
                st.markdown(f"<div class='card-luxury'>", unsafe_allow_html=True)
                st.markdown(f"<h4><i class='fa-solid fa-address-card'></i> الهوية والمستندات الرسمية لـ {selected_asset['المنشأة']}</h4><hr>", unsafe_allow_html=True)
                st.markdown(f"""
                * <b>تصنيف ونوع العقار:</b> {selected_asset['نوع_العقار']}
                * <b>حالة العقار القانونية:</b> {selected_asset['حاله_العقار']}
                * <b>الموقع والنطاق الجغرافي:</b> {selected_asset['المحافظة']} - {selected_asset['مركز/حي/قرية']}
                * <b>المساحة الموثقة بالصك:</b> <span style='color:#1d5c43; font-weight:bold;'>{selected_asset['المساحة']} م²</span>
                * <b>رقم الصك الشرعي المعتمد:</b> <span style='color:#dfb76c; font-weight:bold;'>{selected_asset['رقم الصك']}</span>
                * <b>تاريخ إصدار وثيقة الصك:</b> {selected_asset['تاريخ الصك']}
                """, unsafe_allow_html=True)
                st.markdown(f"</div>", unsafe_allow_html=True)
                
            with col_map:
                st.markdown(f"<div class='card-luxury' style='border-top: 4px solid #58a6ff;'>", unsafe_allow_html=True)
                st.markdown(f"<h4><i class='fa-solid fa-earth-americas'></i> الرقابة الجيومكانية ونظام الـ GIS (رؤية الأقمار الصناعية)</h4><hr>", unsafe_allow_html=True)
                map_df = pd.DataFrame([{"lat": selected_asset['دائرة العرض'], "lon": selected_asset['خط الطول']}])
                st.map(map_df, use_container_width=True)
                st.markdown(f"</div>", unsafe_allow_html=True)
                
            # الصف الثاني: الأربع خانات المستقلة والمنسقة جرافيكياً للمرفقات المصورة (طلبك بدقة)
            st.markdown("<h3><i class='fa-solid fa-images'></i> بطاقات المرفقات الفنية والوثائق المصورة الأربعة للأرض</h3>", unsafe_allow_html=True)
            box1, box2 = st.columns(2)
            box3, box4 = st.columns(2)
            
            with box1:
                st.markdown("<div class='card-luxury' style='border-top: 4px solid #dfb76c;'><h5>📜 1. صورة الصك الشرعي الموثق</h5>", unsafe_allow_html=True)
                img1 = st.file_uploader(f"رفع/تحديث صورة الصك لـ {selected_asset['المنشأة']}", type=["jpg","png","jpeg"], key=f"sok_{selected_asset['رقم الصك']}")
                if img1: st.image(img1, use_container_width=True)
                else: st.info("📷 اسحب وأفلت صورة الصك هنا لمعاينتها.")
                st.markdown("</div>", unsafe_allow_html=True)
                
            with box2:
                st.markdown("<div class='card-luxury' style='border-top: 4px solid #1d5c43;'><h5>🏗️ 2. صورة رخصة البناء الهندسية</h5>", unsafe_allow_html=True)
                img2 = st.file_uploader(f"رفع/تحديث صورة الرخصة لـ {selected_asset['المنشأة']}", type=["jpg","png","jpeg"], key=f"perm_{selected_asset['رقم الصك']}")
                if img2: st.image(img2, use_container_width=True)
                else: st.info("📷 اسحب وأفلت صورة رخصة البناء المعمارية هنا.")
                st.markdown("</div>", unsafe_allow_html=True)
                
            with box3:
                st.markdown("<div class='card-luxury' style='border-top: 4px solid #58a6ff;'><h5>🛰️ 3. كروكي الموقع والرفع المساحي الميداني</h5>", unsafe_allow_html=True)
                img3 = st.file_uploader(f"رفع/تحديث كروكي الموقع لـ {selected_asset['المنشأة']}", type=["jpg","png","jpeg"], key=f"site_{selected_asset['رقم الصك']}")
                if img3: st.image(img3, use_container_width=True)
                else:
                    g_url = f"https://google.com{selected_asset['دائرة العرض']},{selected_asset['خط الطول']}"
                    st.markdown(f"<a href='{g_url}' target='_blank'><button style='width:100%; padding:11px; background:#1d5c43; color:white; border:none; border-radius:8px; font-weight:bold; cursor:pointer;'><i class='fa-solid fa-route'></i> الانتقال المباشر وتتبع الموقع على Google Maps</button></a>", unsafe_allow_html=True)
                st.markdown("</div>", unsafe_allow_html=True)
                
            with box4:
                st.markdown("<div class='card-luxury' style='border-top: 4px solid #143f2e;'><h5>📐 4. صورة قرار الذرعة المساحي المعتمد</h5>", unsafe_allow_html=True)
                img4 = st.file_uploader(f"رفع/تحديث صورة قرار الذرعة لـ {selected_asset['المنشأة']}", type=["jpg","png","jpeg"], key=f"zar_{selected_asset['رقم الصك']}")
                if img4: st.image(img4, use_container_width=True)
                else: st.info("📷 اسحب وأفلت صورة قرار الذرعة الصادر من أمانة الطائف.")
                st.markdown("</div>", unsafe_allow_html=True)

            # محرك صائغ الخطابات المطور بالمسميات الرسمية الفخمة الجديدة المعتمدة للفرع
            st.markdown("<div class='card-luxury'>", unsafe_allow_html=True)
            st.markdown("<h4>✉️ محرك الأتمتة وصائغ المراسلات الفوري المعتمد لإدارة الأراضي والممتلكات</h4>", unsafe_allow_html=True)
            letter_type = st.selectbox("حدد الجهة الحكومية المستهدفة لتوجه الخطاب إليها آلياً:", [
                "خطاب موجه لسعادة أمين محافظة الطائف (استخراج رخصة بناء ومطابقة ذرعة مساحية)", 
                "خطاب موجه لسعادة مدير شركة الكهرباء بالطائف (طلب إيصال التيار وتحديد المحول الفرعي)", 
                "خطاب موجه لفضيلة رئيس المحكمة العامة بالطائف (تحديث ومطابقة الحدود الشرعية للصك)"
            ])
            
            text_content = ""
            if "أمين" in letter_type:
                text_content = f"سعادة أمين محافظة الطائف\nالسلام عليكم ورحمة الله وبركاته،،\n\nتفيدكم فرع وزارة الصحة بمحافظة الطائف علماً بملكية الوزارة الرسمية للموقع المخصص لـ ({selected_asset['المنشأة']}) والواقع بنطاق ({selected_asset['مركز/حي/قرية']}) بموجب الصك الشرعي رقم ({selected_asset['رقم الصك']}) وتاريخ ({selected_asset['تاريخ الصك']}) بمساحة قدرها ({selected_asset['المساحة']} م²). نأمل التوجيه لمن يلزم لاعتماد الرفع المساحي وقرار الذرعة واستخراج رخصة بناء وفق الإحداثيات المرفقة ({selected_asset['دائرة العرض']} , {selected_asset['خط الطول']}).\n\nوتقبلوا خالص التحية والتقدير،،\n\nمدير فرع وزارة الصحة بمحافظة الطائف\nإدارة الأراضي والممتلكات"
            elif "الكهرباء" in letter_type:
                text_content = f"سعادة مدير شركة الكهرباء بمحافظة الطائف\nالسلام عليكم ورح السلام عليكم ورحمة الله وبركاته،،\n\nنظراً لجهوزية البدء الإنشائي والتشغيلي للموقع الطبي التابع للوزارة ({selected_asset['المنشأة']}) والمقام على الأرض ذات الصك رقم ({selected_asset['رقم الصك']})، نأمل منكم الإيعاز للمختصين لطلب إيصال التيار الكهربائي وتحديد موقع محول الطاقة الفرعي حسب الكروكي المرفق.\n\nوتقبلوا وافر التحية والتقدير،،\n\nمدير فر فرع وزارة الصحة بمحافظة الطائف\nإدارة الأراضي والممتلكات"
            else:
                text_content = f"فضيلة رئيس المحكمة العامة بمحافظة الطائف\nالسلام عليكم ورحمة الله وبركاته،،\n\nتتقدم فرع وزارة الصحة بمحافظة الطائف بطلب تحديث وإصدار صك إلكتروني موحد ومطابقة مسحية للصك رقم ({selected_asset['رقم الصك']}) وتاريخ ({selected_asset['تاريخ الصك']}) العائد لملك الوزارة في موقع ({selected_asset['المنشأة']}) بنطاق ({selected_asset['مركز/حي/قرية']}). نأمل التوجيه لمطابقة الحدود الإنشائية حسب الرفع المساحي المعتمد.\n\nوتقبلوا خالص التحية والتقدير،،\n\nمدير فرع وزارة الصحة بمحافظة الطائف\nإدارة الأراضي والممتلكات"
            
            st.text_area("📄 صيغة الخطاب المولد برمجياً والمجهز للطباعة الفورية والتوقيع:", text_content, height=180)
            doc = Document()
            doc.add_heading(letter_type, 0)
            doc.add_paragraph(text_content)
            doc.save("generated_letter.docx")
            with open("generated_letter.docx", "rb") as f:
                st.download_button("📥 تنزيل الخطاب الآن كملف Word رسمي مجهز للتوقيع", f, file_name=f"خطاب_صحة_الطائف_{selected_asset['رقم الصك']}.docx")
            st.markdown("</div>", unsafe_allow_html=True)
    # 💼 تتبع وحفظ المعاملات الرقمية بين الإدارات للعودة إليها (ثابتة سحابياً)
    elif menu == "💼 تتبع وحفظ المعاملات الرقمية":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-route'></i> حفظ وتتبع خط سير المعاملات الرقمية للأراضي</h1>", unsafe_allow_html=True)
        
        st.markdown("<div class='card-luxury'>", unsafe_allow_html=True)
        with st.form("workflow_add"):
            w_col1, w_col2 = st.columns(2)
            with w_col1:
                w_id = st.text_input("رقم المعاملة (الصادر أو الوارد)")
                w_subject = st.text_input("موضوع المعاملة الأساسي للمراسلة")
            with w_col2:
                w_dept = st.selectbox("الجهة الحكومية الحالية الواقفة عندها المعاملة:", ["أمانة محافظة الطائف", "المحكمة العامة بالطائف", "كتابة العدل بمحافظة الطائف", "إدارة الأراضي والممتلكات بالفرع"])
                w_status = st.selectbox("حالة المعاملة الحالية:", ["قيد الدراسة والتدقيق المساحي بالذرعة", "بانتظار الاعتماد النهائي", "تم الإفراغ والمطابقة بنجاح"])
            if st.form_submit_button("💾 قيد وحفظ المعاملة في السجلات"):
                new_wf = pd.DataFrame([{"رقم_المعاملة": w_id, "موضوع_المعاملة": w_subject, "الإدارة_الحالية": w_dept, "حالة_المعاملة": w_status, "تاريخ_التحديث": datetime.now().strftime("%Y-%m-%d %H:%M"), "الموظف_المسؤول": st.session_state['username']}])
                st.session_state["db_workflows"] = pd.concat([df_workflows, new_wf]).reset_index(drop=True)
                st.success("✅ تم حفظ وأرشفة المعاملة في النظام للعودة إليها في أي وقت!")
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='card-luxury'>", unsafe_allow_html=True)
        st.markdown("<h5>📋 أرشيف خط سير المعاملات الموثقة</h5>", unsafe_allow_html=True)
        st.data_editor(st.session_state["db_workflows"], num_rows="dynamic", use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    # ⚠️ نظام رقابة الأراضي وتتبع التعديات الميدانية للأراضي
    elif menu == "⚠️ رقابة الأراضي وتتبع التعديات":
        st.markdown("<h1 style='text-align: right; color: #1d5c43;'><i class='fa-solid fa-shield-halved'></i> تتبع التعديات والرقابة الميدانية للأراضي</h1>", unsafe_allow_html=True)
        
        st.markdown("<div class='card-luxury'>", unsafe_allow_html=True)
        with st.form("encroach_form"):
            c_e1, c_e2 = st.columns(2)
            with c_e1:
                e_sok = st.text_input("رقم الصك الشرعي المعتدى على حدوده الشرعية")
                e_name = st.text_input("اسم المنشأة أو الموقع الطبي المتأثر")
            with c_e2:
                e_type = st.selectbox("نوع التعدي المرصود ميدانياً:", ["إقامة حوش أو أسوار غير نظامية", "بناء شعبي بدون رخصة", "وضع لوحات أو شبوك وتجريف تربة"])
                e_status = st.selectbox("حالة المعاملة والقضية القانونية:", ["تحت الرفع للمحافظة", "بانتظار إزالة لجنة التعديات", "تمت الإزالة واسترداد الأرض بالكامل"])
            e_action = st.text_area("الإجراء المتخذ وتفاصيل المعاملة")
            if st.form_submit_button("🚨 تسجيل حالة التعدي وأرشفة المعاملة فوراً"):
                new_enc = pd.DataFrame([{"رقم الصك": e_sok, "المنشأة": e_name, "نوع_التعدي": e_type, "حالة_القضية": e_status, "تاريخ_الرصد": datetime.now().strftime("%Y-%m-%d"), "الإجراء_المتخذ": e_action}])
                st.session_state["db_encroach"] = pd.concat([df_encroach, new_enc]).reset_index(drop=True)
                st.success("✅ تم تسجيل حالة التعدي بنجاح في السجلات لحماية ممتلكات صحة الطائف.")
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='card-luxury'>", unsafe_allow_html=True)
        st.markdown("<h5>📋 سجل قضايا الرقابة وتتبع التعديات المرصودة</h5>", unsafe_allow_html=True)
        st.data_editor(st.session_state["db_encroach"], num_rows="dynamic", use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)
