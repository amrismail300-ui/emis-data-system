import streamlit as st
import sqlite3
import pandas as pd
from datetime import date
import io

# ==========================================
# 1. إعداد قاعدة البيانات والربط العلائقي
# ==========================================
def init_db():
    conn = sqlite3.connect('emis_system.db')
    conn.execute("PRAGMA foreign_keys = ON") 
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS Schools (
            School_ID TEXT PRIMARY KEY,
            School_Name TEXT NOT NULL,
            School_Type TEXT NOT NULL,
            Governorate TEXT NOT NULL
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS Learners (
            Learner_ID TEXT PRIMARY KEY,
            Full_Name TEXT NOT NULL,
            Gender TEXT NOT NULL,
            Status TEXT NOT NULL
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS Enrollment (
            Enrollment_ID INTEGER PRIMARY KEY AUTOINCREMENT,
            Learner_ID TEXT NOT NULL,
            School_ID TEXT NOT NULL,
            Academic_Year TEXT NOT NULL,
            Enrollment_Date DATE NOT NULL,
            FOREIGN KEY (Learner_ID) REFERENCES Learners(Learner_ID),
            FOREIGN KEY (School_ID) REFERENCES Schools(School_ID)
        )
    ''')
    conn.commit()
    return conn

conn = init_db()

# ==========================================
# 2. واجهة المستخدم
# ==========================================
st.set_page_config(page_title="نظام إدارة معلومات التعليم (EMIS)", layout="wide")
st.title("📚 نظام إدارة معلومات التعليم (EMIS)")

menu = ["لوحة القيادة والتصدير", "إدارة المدارس", "إدارة الطلاب", "التسجيل الأكاديمي (الربط)"]
choice = st.sidebar.selectbox("القائمة الرئيسية", menu)

# --- شاشة لوحة القيادة والتصدير ---
if choice == "لوحة القيادة والتصدير":
    st.header("📊 المؤشرات العامة والتسجيلات")
    
    # جلب البيانات الأساسية
    schools_df = pd.read_sql_query("SELECT * FROM Schools", conn)
    learners_df = pd.read_sql_query("SELECT * FROM Learners", conn)
    
    col1, col2 = st.columns(2)
    col1.info(f"🏫 إجمالي المدارس: {len(schools_df)}")
    col2.success(f"🎓 إجمالي الطلاب: {len(learners_df)}")
    
    st.subheader("سجل الانضمام (الطلاب المسجلين في المدارس)")
    query = '''
        SELECT e.Enrollment_ID, l.Full_Name AS 'اسم الطالب', s.School_Name AS 'المدرسة', 
               e.Academic_Year AS 'العام الدراسي', e.Enrollment_Date AS 'تاريخ التسجيل'
        FROM Enrollment e
        JOIN Learners l ON e.Learner_ID = l.Learner_ID
        JOIN Schools s ON e.School_ID = s.School_ID
    '''
    enrollment_df = pd.read_sql_query(query, conn)
    st.dataframe(enrollment_df, use_container_width=True)

    # --- وحدة التصدير (Export Module) ---
    st.markdown("---")
    st.subheader("📥 تصدير التقارير (Data Export)")
    st.write("استخراج قاعدة البيانات بالكامل إلى ملف Excel منظم يحتوي على أوراق (Sheets) متعددة.")
    
    if not schools_df.empty or not learners_df.empty:
        # إنشاء ملف Excel في الذاكرة الوهمية (Buffer)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            if not enrollment_df.empty:
                enrollment_df.to_excel(writer, sheet_name='سجل التسجيل', index=False)
            schools_df.to_excel(writer, sheet_name='المدارس', index=False)
            learners_df.to_excel(writer, sheet_name='الطلاب', index=False)
            
        st.download_button(
            label="تحميل التقرير الشامل (Excel) 📊",
            data=buffer.getvalue(),
            file_name=f"EMIS_Database_Export_{date.today()}.xlsx",
            mime="application/vnd.ms-excel"
        )
    else:
        st.warning("لا توجد بيانات في النظام لتصديرها.")

# --- شاشة إدارة المدارس ---
elif choice == "إدارة المدارس":
    st.header("🏫 إضافة مدرسة جديدة")
    with st.form(key='school_form'):
        col1, col2 = st.columns(2)
        school_id = col1.text_input("رمز المدرسة (مثال: SCH-001)")
        school_name = col1.text_input("اسم المدرسة")
        school_type = col2.selectbox("نوع المدرسة", ["حكومية", "خاصة", "مجتمعية"])
        governorate = col2.selectbox("المحافظة", ["تعز", "عدن", "صنعاء", "حضرموت", "الحديدة", "إب"])
        
        if st.form_submit_button('حفظ'):
            try:
                conn.execute("INSERT INTO Schools VALUES (?, ?, ?, ?)", (school_id, school_name, school_type, governorate))
                conn.commit()
                st.success("تم الحفظ بنجاح!")
            except:
                st.error("تأكد من عدم تكرار رمز المدرسة.")

# --- شاشة إدارة الطلاب ---
elif choice == "إدارة الطلاب":
    st.header("🎓 تسجيل طالب جديد بالهوية")
    with st.form(key='learner_form'):
        col1, col2 = st.columns(2)
        learner_id = col1.text_input("رقم الطالب (مثال: STU-1001)")
        full_name = col1.text_input("الاسم الرباعي")
        gender = col2.radio("الجنس", ["ذكر", "أنثى"])
        status = col2.selectbox("حالة القيد", ["نشط", "منقول", "متخرج"])
        
        if st.form_submit_button('حفظ'):
            try:
                conn.execute("INSERT INTO Learners VALUES (?, ?, ?, ?)", (learner_id, full_name, gender, status))
                conn.commit()
                st.success("تم الحفظ بنجاح!")
            except:
                st.error("تأكد من عدم تكرار رقم الطالب.")

# --- شاشة التسجيل الأكاديمي ---
elif choice == "التسجيل الأكاديمي (الربط)":
    st.header("🔗 ربط الطالب بمدرسة (Enrollment)")
    schools_df = pd.read_sql_query("SELECT School_ID, School_Name FROM Schools", conn)
    learners_df = pd.read_sql_query("SELECT Learner_ID, Full_Name FROM Learners", conn)
    
    if schools_df.empty or learners_df.empty:
        st.warning("⚠️ يجب إضافة مدرسة واحدة وطالب واحد على الأقل قبل إجراء عملية التسجيل.")
    else:
        with st.form(key='enroll_form'):
            school_options = dict(zip(schools_df.School_Name, schools_df.School_ID))
            learner_options = dict(zip(learners_df.Full_Name, learners_df.Learner_ID))
            
            selected_learner_name = st.selectbox("اختر الطالب", list(learner_options.keys()))
            selected_school_name = st.selectbox("اختر المدرسة", list(school_options.keys()))
            academic_year = st.selectbox("العام الدراسي", ["2025-2026", "2026-2027"])
            enrollment_date = st.date_input("تاريخ التسجيل", date.today())
            
            if st.form_submit_button('تأكيد التسجيل'):
                l_id = learner_options[selected_learner_name]
                s_id = school_options[selected_school_name]
                
                conn.execute("INSERT INTO Enrollment (Learner_ID, School_ID, Academic_Year, Enrollment_Date) VALUES (?, ?, ?, ?)", 
                             (l_id, s_id, academic_year, enrollment_date))
                conn.commit()
                st.success(f"تم ربط الطالب {selected_learner_name} بمدرسة {selected_school_name} بنجاح!")
