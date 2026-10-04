# =========================================================
# AI CAREER INTELLIGENCE PLATFORM
# Complete Single-File Advanced Version
# =========================================================

import streamlit as st
import re
from io import BytesIO

# =========================================================
# PDF / DOCX SUPPORT
# =========================================================

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

try:
    from docx import Document
except ImportError:
    Document = None

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer
)


# =========================================================
# PAGE CONFIGURATION
# =========================================================

st.set_page_config(
    page_title="AI Career Intelligence Platform",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded"
)


# =========================================================
# SESSION STATE INITIALIZATION
# IMPORTANT: MUST COME BEFORE DARK MODE
# =========================================================

if "dark_mode" not in st.session_state:
    st.session_state["dark_mode"] = False

if "resume_score" not in st.session_state:
    st.session_state["resume_score"] = 0

if "skill_score" not in st.session_state:
    st.session_state["skill_score"] = 0

if "project_score" not in st.session_state:
    st.session_state["project_score"] = 0

if "interview_score" not in st.session_state:
    st.session_state["interview_score"] = 0

if "resume_text" not in st.session_state:
    st.session_state["resume_text"] = ""

if "resume_name" not in st.session_state:
    st.session_state["resume_name"] = ""

if "selected_career" not in st.session_state:
    st.session_state["selected_career"] = "Data Analyst"

if "project_count" not in st.session_state:
    st.session_state["project_count"] = 0

if "overall_score" not in st.session_state:
    st.session_state["overall_score"] = 0


# =========================================================
# SIDEBAR DARK MODE
# ONLY ONE TOGGLE IN THE ENTIRE APP
# =========================================================

dark_mode = st.sidebar.toggle(
    "🌙 Dark Mode",
    value=st.session_state["dark_mode"],
    key="dark_mode_toggle"
)

st.session_state["dark_mode"] = dark_mode


# =========================================================
# CUSTOM CSS
# =========================================================

if dark_mode:

    st.markdown(
        """
        <style>

        .stApp {
            background-color: #0f172a;
            color: #f8fafc;
        }

        .main .block-container {
            max-width: 1400px;
            padding-top: 2rem;
            padding-bottom: 3rem;
        }

        .main-title {
            color: #f8fafc;
            font-size: 42px;
            font-weight: 800;
            margin-bottom: 5px;
        }

        .subtitle {
            color: #cbd5e1;
            font-size: 17px;
            margin-bottom: 25px;
        }

        .section-title {
            color: #f8fafc;
            font-size: 25px;
            font-weight: 750;
            margin-top: 25px;
            margin-bottom: 15px;
        }

        .metric-card {
            background-color: #1e293b;
            color: #f8fafc;
            border: 1px solid #334155;
            border-radius: 18px;
            padding: 22px;
            box-shadow: 0 8px 25px rgba(0,0,0,0.20);
        }

        div[data-testid="stMetric"] {
            background-color: #1e293b;
            border: 1px solid #334155;
            border-radius: 16px;
            padding: 18px;
        }

        div[data-testid="stMetricValue"] {
            color: #f8fafc !important;
        }

        div[data-testid="stMetricLabel"] {
            color: #cbd5e1 !important;
        }

        .stTextInput input {
            background-color: #1e293b;
            color: #f8fafc;
            border: 1px solid #334155;
            border-radius: 10px;
        }

        .stTextArea textarea {
            background-color: #1e293b;
            color: #f8fafc;
            border: 1px solid #334155;
            border-radius: 10px;
        }

        section[data-testid="stSidebar"] {
            background: linear-gradient(
                180deg,
                #111827 0%,
                #1e1b4b 100%
            );
        }

        section[data-testid="stSidebar"] * {
            color: white !important;
        }

        .stButton > button {
            border-radius: 12px;
            font-weight: 600;
            padding: 0.65rem 1.2rem;
        }

        .stProgress > div > div {
            border-radius: 20px;
        }

        div[data-testid="stAlert"] {
            border-radius: 12px;
        }

        hr {
            border-color: #334155;
        }

        .footer {
            text-align: center;
            color: #cbd5e1;
            font-size: 13px;
            margin-top: 40px;
            padding: 20px;
        }

        </style>
        """,
        unsafe_allow_html=True
    )

else:

    st.markdown(
        """
        <style>

        .stApp {
            background-color: #f8fafc;
            color: #0f172a;
        }

        .main .block-container {
            max-width: 1400px;
            padding-top: 2rem;
            padding-bottom: 3rem;
        }

        .main-title {
            color: #0f172a;
            font-size: 42px;
            font-weight: 800;
            margin-bottom: 5px;
        }

        .subtitle {
            color: #64748b;
            font-size: 17px;
            margin-bottom: 25px;
        }

        .section-title {
            color: #0f172a;
            font-size: 25px;
            font-weight: 750;
            margin-top: 25px;
            margin-bottom: 15px;
        }

        .metric-card {
            background-color: #ffffff;
            color: #0f172a;
            border: 1px solid #e2e8f0;
            border-radius: 18px;
            padding: 22px;
            box-shadow: 0 8px 25px rgba(15,23,42,0.08);
        }

        div[data-testid="stMetric"] {
            background-color: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 16px;
            padding: 18px;
        }

        div[data-testid="stMetricValue"] {
            color: #0f172a !important;
        }

        div[data-testid="stMetricLabel"] {
            color: #64748b !important;
        }

        .stTextInput input {
            background-color: #ffffff;
            color: #0f172a;
            border: 1px solid #e2e8f0;
            border-radius: 10px;
        }

        .stTextArea textarea {
            background-color: #ffffff;
            color: #0f172a;
            border: 1px solid #e2e8f0;
            border-radius: 10px;
        }

        section[data-testid="stSidebar"] {
            background: linear-gradient(
                180deg,
                #111827 0%,
                #1e1b4b 100%
            );
        }

        section[data-testid="stSidebar"] * {
            color: white !important;
        }

        .stButton > button {
            border-radius: 12px;
            font-weight: 600;
            padding: 0.65rem 1.2rem;
        }

        .stProgress > div > div {
            border-radius: 20px;
        }

        div[data-testid="stAlert"] {
            border-radius: 12px;
        }

        .footer {
            text-align: center;
            color: #64748b;
            font-size: 13px;
            margin-top: 40px;
            padding: 20px;
        }

        </style>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# CAREER DATA
# =========================================================

CAREER_SKILLS = {

    "Data Analyst": [
        "Python",
        "SQL",
        "Excel",
        "Power BI",
        "Statistics",
        "Data Visualization",
        "Pandas",
        "NumPy"
    ],

    "Data Scientist": [
        "Python",
        "SQL",
        "Statistics",
        "Machine Learning",
        "Pandas",
        "NumPy",
        "Data Visualization",
        "Deep Learning"
    ],

    "AI Engineer": [
        "Python",
        "Machine Learning",
        "Deep Learning",
        "TensorFlow",
        "PyTorch",
        "SQL",
        "NLP",
        "Computer Vision"
    ],

    "Machine Learning Engineer": [
        "Python",
        "Machine Learning",
        "Deep Learning",
        "Scikit-learn",
        "TensorFlow",
        "PyTorch",
        "SQL",
        "MLOps"
    ],

    "Business Analyst": [
        "Excel",
        "SQL",
        "Power BI",
        "Statistics",
        "Business Analysis",
        "Communication",
        "Data Visualization",
        "Requirements Analysis"
    ]
}


CAREER_ROADMAPS = {

    "Data Analyst": [
        "Python Basics",
        "Excel",
        "SQL",
        "Statistics",
        "Pandas & NumPy",
        "Power BI",
        "Data Visualization",
        "Portfolio Projects",
        "Internship Preparation"
    ],

    "Data Scientist": [
        "Python",
        "Statistics",
        "SQL",
        "Pandas & NumPy",
        "Machine Learning",
        "Data Visualization",
        "Deep Learning",
        "Projects",
        "Internship Preparation"
    ],

    "AI Engineer": [
        "Python",
        "Data Structures",
        "Machine Learning",
        "Deep Learning",
        "TensorFlow / PyTorch",
        "NLP",
        "Computer Vision",
        "AI Projects",
        "Deployment"
    ],

    "Machine Learning Engineer": [
        "Python",
        "Statistics",
        "Machine Learning",
        "Scikit-learn",
        "Deep Learning",
        "TensorFlow / PyTorch",
        "MLOps",
        "Model Deployment",
        "Projects"
    ],

    "Business Analyst": [
        "Excel",
        "Business Fundamentals",
        "SQL",
        "Statistics",
        "Power BI",
        "Requirements Analysis",
        "Communication",
        "Business Projects",
        "Interview Preparation"
    ]
}


JOB_DATABASE = [

    {
        "role": "Data Analyst",
        "company": "Analytics Company",
        "skills": ["Python", "SQL", "Excel", "Power BI"],
        "location": "Chennai",
        "type": "Full Time"
    },

    {
        "role": "Data Scientist",
        "company": "AI Solutions",
        "skills": ["Python", "SQL", "Machine Learning", "Statistics"],
        "location": "Bangalore",
        "type": "Full Time"
    },

    {
        "role": "Machine Learning Engineer",
        "company": "Tech AI Labs",
        "skills": ["Python", "Machine Learning", "Deep Learning"],
        "location": "Chennai",
        "type": "Internship"
    },

    {
        "role": "Business Analyst",
        "company": "Business Insights",
        "skills": ["Excel", "SQL", "Power BI", "Communication"],
        "location": "Hyderabad",
        "type": "Full Time"
    },

    {
        "role": "AI Engineer",
        "company": "Future AI",
        "skills": ["Python", "Machine Learning", "NLP"],
        "location": "Remote",
        "type": "Internship"
    }
]


# =========================================================
# HELPER FUNCTIONS
# =========================================================

def extract_resume_text(uploaded_file):

    if uploaded_file is None:
        return ""

    file_name = uploaded_file.name.lower()

    try:

        if file_name.endswith(".txt"):

            return uploaded_file.read().decode(
                "utf-8",
                errors="ignore"
            )

        elif file_name.endswith(".pdf"):

            if PdfReader is None:
                return ""

            reader = PdfReader(uploaded_file)

            text = ""

            for page in reader.pages:
                text += page.extract_text() or ""

            return text

        elif file_name.endswith(".docx"):

            if Document is None:
                return ""

            document = Document(uploaded_file)

            return "\n".join(
                paragraph.text
                for paragraph in document.paragraphs
            )

    except Exception as error:

        st.error(
            f"Unable to read file: {error}"
        )

    return ""


def calculate_resume_score(text):

    if not text.strip():
        return 0

    text_lower = text.lower()

    score = 0

    sections = {

        "education": [
            "education",
            "degree",
            "university",
            "college"
        ],

        "skills": [
            "skills",
            "technical skills",
            "programming"
        ],

        "projects": [
            "projects",
            "project"
        ],

        "experience": [
            "experience",
            "internship",
            "work experience"
        ],

        "contact": [
            "@",
            "phone",
            "email",
            "linkedin"
        ],

        "certifications": [
            "certification",
            "certificate"
        ]
    }

    for keywords in sections.values():

        if any(
            keyword in text_lower
            for keyword in keywords
        ):
            score += 12

    words = len(text.split())

    if words >= 150:
        score += 8

    if words >= 300:
        score += 5

    return min(score, 100)


def extract_skills(text):

    if not text:
        return []

    text_lower = text.lower()

    all_skills = set()

    for skills in CAREER_SKILLS.values():

        for skill in skills:

            if skill.lower() in text_lower:
                all_skills.add(skill)

    return sorted(all_skills)


def calculate_skill_score(
    extracted_skills,
    career
):

    required = CAREER_SKILLS.get(
        career,
        []
    )

    if not required:
        return 0

    extracted_lower = [
        skill.lower()
        for skill in extracted_skills
    ]

    matched = sum(
        1
        for skill in required
        if skill.lower() in extracted_lower
    )

    return int(
        matched / len(required) * 100
    )


def calculate_overall_score():

    return int(
        (
            st.session_state.resume_score
            + st.session_state.skill_score
            + st.session_state.project_score
            + st.session_state.interview_score
        ) / 4
    )


# =========================================================
# PDF CREATOR
# =========================================================

def create_pdf(
    name,
    email,
    phone,
    location,
    objective,
    education,
    skills,
    projects,
    certifications,
    experience,
    title="Professional Resume"
):

    buffer = BytesIO()

    document = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=18 * mm,
        leftMargin=18 * mm,
        topMargin=18 * mm,
        bottomMargin=18 * mm
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "TitleCustom",
        parent=styles["Title"],
        fontSize=24,
        alignment=TA_CENTER,
        spaceAfter=12
    )

    heading_style = ParagraphStyle(
        "HeadingCustom",
        parent=styles["Heading2"],
        fontSize=14,
        spaceBefore=10,
        spaceAfter=6,
        textColor=colors.HexColor("#1F4E79")
    )

    normal_style = ParagraphStyle(
        "NormalCustom",
        parent=styles["BodyText"],
        fontSize=10,
        leading=15
    )

    story = []

    story.append(
        Paragraph(
            name if name else "Your Name",
            title_style
        )
    )

    contact = " | ".join(
        x for x in [
            email,
            phone,
            location
        ]
        if x
    )

    if contact:

        story.append(
            Paragraph(
                contact,
                ParagraphStyle(
                    "Contact",
                    parent=normal_style,
                    alignment=TA_CENTER
                )
            )
        )

    story.append(
        Spacer(1, 10)
    )

    sections = [
        ("Career Objective", objective),
        ("Education", education),
        ("Skills", skills),
        ("Projects", projects),
        ("Certifications", certifications),
        ("Experience", experience)
    ]

    for heading, content in sections:

        if content:

            story.append(
                Paragraph(
                    heading,
                    heading_style
                )
            )

            formatted_content = (
                str(content)
                .replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\n", "<br/>")
            )

            story.append(
                Paragraph(
                    formatted_content,
                    normal_style
                )
            )

            story.append(
                Spacer(1, 5)
            )

    document.build(story)

    buffer.seek(0)

    return buffer.getvalue()


# =========================================================
# SIDEBAR
# IMPORTANT: NO SECOND DARK MODE TOGGLE HERE
# =========================================================

st.sidebar.title(
    "🚀 AI Career Platform"
)

st.sidebar.caption(
    "Your AI-powered career preparation system"
)

st.sidebar.divider()

page = st.sidebar.selectbox(
    "Choose a Module",
    [
        "🏠 Home",
        "📊 Career Dashboard",
        "📤 Advanced Resume Analyzer",
        "✨ Resume Builder",
        "🎨 CV Maker",
        "🤖 AI Career Helper",
        "💼 Job Matcher",
        "🎯 Career Roadmap",
        "🧠 Skill Gap Analyzer",
        "🎤 Interview Preparation"
    ],
    key="main_navigation"
)

st.sidebar.divider()

st.sidebar.info(
    "💡 Complete each module to improve your career readiness."
)


# =========================================================
# HOME
# =========================================================

if page == "🏠 Home":

    st.markdown(
        '<div class="main-title">🚀 AI Career Intelligence Platform</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="subtitle">Your complete AI-powered career preparation assistant</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="metric-card">
            <h2>Build Your Career With Data 🚀</h2>
            <p>
            Analyze your resume, discover skill gaps,
            prepare for interviews and track your career progress.
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.divider()

    col1, col2, col3, col4 = st.columns(4)

    with col1:

        st.metric(
            "📄 Resume",
            f"{st.session_state.resume_score}%"
        )

    with col2:

        st.metric(
            "🧠 Skills",
            f"{st.session_state.skill_score}%"
        )

    with col3:

        st.metric(
            "💻 Projects",
            f"{st.session_state.project_score}%"
        )

    with col4:

        st.metric(
            "🎤 Interview",
            f"{st.session_state.interview_score}%"
        )

    st.divider()

    st.markdown(
        '<div class="section-title">✨ Platform Features</div>',
        unsafe_allow_html=True
    )

    col1, col2, col3 = st.columns(3)

    with col1:

        st.markdown(
            """
            ### 📄 Resume Intelligence

            • Resume Analyzer  
            • ATS readiness  
            • Skill detection  
            • Resume Builder  
            • CV Maker
            """
        )

    with col2:

        st.markdown(
            """
            ### 🧠 Career Intelligence

            • Career Dashboard  
            • Skill Gap Analyzer  
            • Career Roadmap  
            • AI Career Helper
            """
        )

    with col3:

        st.markdown(
            """
            ### 💼 Career Preparation

            • Job Matcher  
            • Interview Preparation  
            • Technical Questions  
            • HR Questions
            """
        )

    st.divider()

    st.markdown(
        '<div class="section-title">🔄 How It Works</div>',
        unsafe_allow_html=True
    )

    steps = [
        ("1️⃣", "Upload Resume", "Analyze your resume and detect your skills."),
        ("2️⃣", "Find Skill Gaps", "Compare your skills with your target career."),
        ("3️⃣", "Build Skills", "Follow your personalized roadmap."),
        ("4️⃣", "Match Jobs", "Compare your profile with job requirements."),
        ("5️⃣", "Practice Interviews", "Improve your interview preparation.")
    ]

    for icon, title, description in steps:

        st.markdown(
            f"""
            <div class="metric-card" style="margin-bottom:12px;">
                <h3>{icon} {title}</h3>
                <p>{description}</p>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.success(
        "Select a module from the sidebar to get started."
    )


# =========================================================
# CAREER DASHBOARD
# =========================================================

elif page == "📊 Career Dashboard":

    st.markdown(
        '<div class="main-title">📊 Career Intelligence Dashboard</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="subtitle">Track your career preparation and identify what to improve next.</div>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns([2, 1])

    with col1:

        st.subheader("👤 Career Profile")

        profile_name = st.text_input(
            "Your Name",
            value="Rithika",
            key="dashboard_name"
        )

        career_goal = st.selectbox(
            "🎯 Target Career",
            list(CAREER_SKILLS.keys()),
            key="dashboard_career"
        )

    with col2:

        st.metric(
            "🎯 Target Career",
            career_goal
        )

    resume_score = st.session_state.resume_score
    skill_score = st.session_state.skill_score
    interview_score = st.session_state.interview_score

    project_count = st.number_input(
        "💻 Number of completed projects",
        min_value=0,
        max_value=10,
        value=st.session_state.project_count,
        step=1,
        key="dashboard_project_count"
    )

    st.session_state.project_count = project_count

    project_score = min(
        project_count * 20,
        100
    )

    st.session_state.project_score = project_score

    overall_score = round(
        (
            resume_score
            + skill_score
            + project_score
            + interview_score
        ) / 4
    )

    st.session_state.overall_score = overall_score

    st.divider()

    st.subheader("🚀 Career Readiness")

    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:

        st.metric(
            "Overall",
            f"{overall_score}%"
        )

    with col2:

        st.metric(
            "Resume",
            f"{resume_score}%"
        )

    with col3:

        st.metric(
            "Skills",
            f"{skill_score}%"
        )

    with col4:

        st.metric(
            "Projects",
            f"{project_score}%"
        )

    with col5:

        st.metric(
            "Interview",
            f"{interview_score}%"
        )

    st.subheader("📈 Preparation Progress")

    progress_data = {
        "Resume": resume_score,
        "Technical Skills": skill_score,
        "Projects": project_score,
        "Interview": interview_score
    }

    for name, score in progress_data.items():

        st.write(
            f"**{name} — {score}%**"
        )

        st.progress(
            min(score / 100, 1.0)
        )

    st.divider()

    st.subheader("🧠 Skill Analysis")

    required_skills = CAREER_SKILLS.get(
        career_goal,
        []
    )

    resume_text = st.session_state.resume_text

    detected_skills = extract_skills(
        resume_text
    )

    detected_lower = [
        x.lower()
        for x in detected_skills
    ]

    if resume_text:

        matched_skills = [
            skill
            for skill in required_skills
            if skill.lower() in detected_lower
        ]

        missing_skills = [
            skill
            for skill in required_skills
            if skill.lower() not in detected_lower
        ]

        col1, col2 = st.columns(2)

        with col1:

            st.markdown("### ✅ Skills You Have")

            if matched_skills:

                for skill in matched_skills:

                    st.success(
                        f"✓ {skill}"
                    )

            else:

                st.info(
                    "No matching skills detected yet."
                )

        with col2:

            st.markdown("### ⚠️ Skills To Learn")

            if missing_skills:

                for skill in missing_skills:

                    st.warning(
                        f"→ {skill}"
                    )

            else:

                st.success(
                    "🎉 You have all required skills!"
                )

    else:

        st.info(
            "📄 Analyze your resume first to see your personalized skill gap."
        )

    st.divider()

    st.subheader("🎯 Career Readiness Level")

    if overall_score < 30:

        level = "🌱 Beginner"
        message = "Start building your foundation with Python, SQL and basic projects."

    elif overall_score < 50:

        level = "📚 Developing"
        message = "You are building your skills. Focus on projects and resume improvement."

    elif overall_score < 70:

        level = "🚀 Growing"
        message = "Good progress! Strengthen your technical skills and interview preparation."

    elif overall_score < 85:

        level = "💼 Job Ready"
        message = "Your preparation is progressing well. Focus on real-world projects and applications."

    else:

        level = "🏆 Highly Prepared"
        message = "Maintain your skills and continue improving your portfolio."

    st.markdown(
        f"## {level}"
    )

    st.info(message)

    st.subheader("💡 Recommended Next Steps")

    recommendations = []

    if resume_score < 70:

        recommendations.append(
            "📄 Improve your resume and ATS compatibility."
        )

    if skill_score < 70:

        recommendations.append(
            "🧠 Learn the missing technical skills."
        )

    if project_score < 60:

        recommendations.append(
            "💻 Build more practical projects."
        )

    if interview_score < 60:

        recommendations.append(
            "🎤 Practice technical and HR interview questions."
        )

    if not recommendations:

        recommendations.append(
            "🔥 Keep improving your portfolio and practical experience."
        )

    for recommendation in recommendations:

        st.write(
            f"• {recommendation}"
        )

    st.divider()

    st.subheader("📌 Career Summary")

    st.write(
        f"**Candidate:** {profile_name}"
    )

    st.write(
        f"**Target Career:** {career_goal}"
    )

    st.write(
        f"**Overall Career Readiness:** {overall_score}%"
    )

    st.write(
        f"**Projects Completed:** {project_count}"
    )


# =========================================================
# ADVANCED RESUME ANALYZER
# =========================================================

elif page == "📤 Advanced Resume Analyzer":

    st.markdown(
        '<div class="main-title">📤 Advanced AI Resume Analyzer</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="subtitle">Analyze your resume for ATS readiness, skills and career alignment.</div>',
        unsafe_allow_html=True
    )

    st.divider()

    uploaded_file = st.file_uploader(
        "📄 Upload your Resume",
        type=["pdf", "docx", "txt"],
        key="resume_upload"
    )

    pasted_resume = st.text_area(
        "Or paste your resume text here",
        height=250,
        key="resume_paste"
    )

    analyze = st.button(
        "🔍 Analyze Resume",
        use_container_width=True
    )

    if analyze:

        resume_text = ""

        if uploaded_file:

            resume_text = extract_resume_text(
                uploaded_file
            )

            st.session_state.resume_name = (
                uploaded_file.name
            )

        elif pasted_resume:

            resume_text = pasted_resume

        if not resume_text.strip():

            st.warning(
                "Please upload or paste a resume."
            )

        else:

            st.session_state.resume_text = resume_text

            score = calculate_resume_score(
                resume_text
            )

            st.session_state.resume_score = score

            detected_skills = extract_skills(
                resume_text
            )

            skill_score = calculate_skill_score(
                detected_skills,
                st.session_state.selected_career
            )

            st.session_state.skill_score = skill_score

            st.success(
                "Resume analysis completed!"
            )

            st.divider()

            col1, col2, col3 = st.columns(3)

            with col1:

                st.metric(
                    "ATS Score",
                    f"{score}%"
                )

            with col2:

                st.metric(
                    "Skills Detected",
                    len(detected_skills)
                )

            with col3:

                st.metric(
                    "Word Count",
                    len(resume_text.split())
                )

            st.divider()

            st.subheader("🧠 Detected Skills")

            if detected_skills:

                st.info(
                    " • ".join(detected_skills)
                )

            else:

                st.warning(
                    "No recognized technical skills detected."
                )

            st.divider()

            st.subheader(
                f"🎯 Skill Match for {st.session_state.selected_career}"
            )

            required_skills = CAREER_SKILLS[
                st.session_state.selected_career
            ]

            detected_lower = [
                x.lower()
                for x in detected_skills
            ]

            missing_skills = [
                skill
                for skill in required_skills
                if skill.lower() not in detected_lower
            ]

            if missing_skills:

                for skill in missing_skills:

                    st.warning(
                        f"⚠️ {skill}"
                    )

            else:

                st.success(
                    "🎉 All major career skills were detected!"
                )

            st.divider()

            st.subheader(
                "📋 Resume Section Check"
            )

            sections = [
                "Education",
                "Skills",
                "Projects",
                "Experience",
                "Certifications",
                "LinkedIn"
            ]

            lower_text = resume_text.lower()

            for section in sections:

                if section.lower() in lower_text:

                    st.success(
                        f"✅ {section}"
                    )

                else:

                    st.warning(
                        f"⚠️ {section} section not detected"
                    )

            st.divider()

            st.subheader(
                "💡 Improvement Suggestions"
            )

            if score < 50:

                st.warning(
                    "Add more resume sections, projects and technical skills."
                )

            elif score < 75:

                st.info(
                    "Improve project descriptions and add measurable achievements."
                )

            else:

                st.success(
                    "Your resume has a strong basic structure. Continue improving keywords and achievements."
                )


# =========================================================
# RESUME BUILDER
# =========================================================

elif page == "✨ Resume Builder":

    st.markdown(
        '<div class="main-title">✨ Professional Resume Builder</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="subtitle">Create and download your professional resume.</div>',
        unsafe_allow_html=True
    )

    st.divider()

    col1, col2 = st.columns(2)

    with col1:

        builder_name = st.text_input(
            "Full Name",
            key="builder_name"
        )

        builder_email = st.text_input(
            "Email",
            key="builder_email"
        )

        builder_phone = st.text_input(
            "Phone",
            key="builder_phone"
        )

        builder_location = st.text_input(
            "Location",
            key="builder_location"
        )

    with col2:

        builder_objective = st.text_area(
            "Career Objective",
            height=120,
            key="builder_objective"
        )

    st.divider()

    builder_education = st.text_area(
        "🎓 Education",
        placeholder="B.Tech AI & Data Science - College - Year",
        key="builder_education"
    )

    builder_skills = st.text_area(
        "🧠 Skills",
        placeholder="Python, SQL, Power BI, Excel...",
        key="builder_skills"
    )

    builder_projects = st.text_area(
        "💻 Projects",
        placeholder="Project Name - Description",
        key="builder_projects"
    )

    builder_certifications = st.text_area(
        "📜 Certifications",
        key="builder_certifications"
    )

    builder_experience = st.text_area(
        "💼 Experience / Internship",
        key="builder_experience"
    )

    st.divider()

    if st.button(
        "📄 Generate Professional Resume",
        use_container_width=True,
        key="generate_resume"
    ):

        pdf = create_pdf(
            builder_name,
            builder_email,
            builder_phone,
            builder_location,
            builder_objective,
            builder_education,
            builder_skills,
            builder_projects,
            builder_certifications,
            builder_experience,
            "Professional Resume"
        )

        st.success(
            "Resume generated successfully!"
        )

        st.download_button(
            "⬇️ Download Resume",
            data=pdf,
            file_name="Professional_Resume.pdf",
            mime="application/pdf",
            use_container_width=True,
            key="download_resume"
        )


# =========================================================
# CV MAKER
# =========================================================

elif page == "🎨 CV Maker":

    st.markdown(
        '<div class="main-title">🎨 Professional CV Maker</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="subtitle">Create a clean professional CV.</div>',
        unsafe_allow_html=True
    )

    st.divider()

    cv_template = st.selectbox(
        "🎨 Choose CV Template",
        [
            "Professional Blue",
            "Modern Purple",
            "Fresh Green",
            "Classic Black"
        ],
        key="cv_template"
    )

    st.divider()

    cv_name = st.text_input(
        "Full Name",
        key="cv_name"
    )

    cv_email = st.text_input(
        "Email",
        key="cv_email"
    )

    cv_phone = st.text_input(
        "Phone",
        key="cv_phone"
    )

    cv_location = st.text_input(
        "Location",
        key="cv_location"
    )

    cv_objective = st.text_area(
        "Career Objective",
        key="cv_objective"
    )

    cv_education = st.text_area(
        "Education",
        key="cv_education"
    )

    cv_skills = st.text_area(
        "Skills",
        key="cv_skills"
    )

    cv_projects = st.text_area(
        "Projects",
        key="cv_projects"
    )

    cv_certifications = st.text_area(
        "Certifications",
        key="cv_certifications"
    )

    cv_experience = st.text_area(
        "Experience",
        key="cv_experience"
    )

    if st.button(
        "🎨 Generate CV",
        use_container_width=True,
        key="generate_cv"
    ):

        pdf = create_pdf(
            cv_name,
            cv_email,
            cv_phone,
            cv_location,
            cv_objective,
            cv_education,
            cv_skills,
            cv_projects,
            cv_certifications,
            cv_experience,
            cv_template
        )

        st.success(
            f"{cv_template} CV generated!"
        )

        st.download_button(
            "⬇️ Download CV",
            data=pdf,
            file_name="Professional_CV.pdf",
            mime="application/pdf",
            use_container_width=True,
            key="download_cv"
        )


# =========================================================
# AI CAREER HELPER
# =========================================================

elif page == "🤖 AI Career Helper":

    st.markdown(
        '<div class="main-title">🤖 AI Career Helper</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="subtitle">Get personalized career guidance based on your profile.</div>',
        unsafe_allow_html=True
    )

    career = st.selectbox(
        "🎯 Your Target Career",
        list(CAREER_SKILLS.keys()),
        key="helper_career"
    )

    resume_text = st.session_state.resume_text

    resume_score = st.session_state.resume_score
    skill_score = st.session_state.skill_score
    project_score = st.session_state.project_score
    interview_score = st.session_state.interview_score

    detected_skills = extract_skills(
        resume_text
    )

    required_skills = CAREER_SKILLS[
        career
    ]

    detected_lower = [
        skill.lower()
        for skill in detected_skills
    ]

    missing_skills = [
        skill
        for skill in required_skills
        if skill.lower() not in detected_lower
    ]

    st.divider()

    st.subheader(
        "👤 Your Career Profile"
    )

    col1, col2, col3, col4 = st.columns(4)

    with col1:

        st.metric(
            "📄 Resume",
            f"{resume_score}%"
        )

    with col2:

        st.metric(
            "🧠 Skills",
            f"{skill_score}%"
        )

    with col3:

        st.metric(
            "💻 Projects",
            f"{project_score}%"
        )

    with col4:

        st.metric(
            "🎤 Interview",
            f"{interview_score}%"
        )

    st.divider()

    st.subheader(
        "💬 Ask Your Career Helper"
    )

    question = st.text_area(
        "What do you want help with?",
        placeholder="Example: What skills should I learn next?",
        height=120,
        key="career_question"
    )

    if st.button(
        "🤖 Get Career Guidance",
        use_container_width=True,
        key="career_guidance"
    ):

        if not question.strip():

            st.warning(
                "Please enter a question first."
            )

        else:

            q = question.lower()

            st.subheader(
                "🤖 Career Guidance"
            )

            if (
                "skill" in q
                or "learn" in q
                or "technology" in q
            ):

                if missing_skills:

                    st.write(
                        f"For a **{career}** career, focus on:"
                    )

                    for index, skill in enumerate(
                        missing_skills,
                        start=1
                    ):

                        st.write(
                            f"{index}. **{skill}**"
                        )

                    st.info(
                        "After learning each skill, build a small project using it."
                    )

                else:

                    st.success(
                        "Your resume currently covers the listed skills for this career. Focus on projects and interview preparation."
                    )

            elif (
                "resume" in q
                or "cv" in q
                or "ats" in q
            ):

                if resume_score < 50:

                    st.warning(
                        "Improve your resume sections, projects, technical skills and achievements."
                    )

                elif resume_score < 80:

                    st.info(
                        "Your resume has a good foundation. Improve project descriptions and measurable results."
                    )

                else:

                    st.success(
                        "Your resume score is strong. Keep tailoring it to the roles you target."
                    )

            elif (
                "project" in q
                or "portfolio" in q
                or "github" in q
            ):

                st.write(
                    f"For a **{career}**, build projects that demonstrate practical use of the required skills."
                )

                st.write(
                    "💻 Solve a real problem."
                )

                st.write(
                    "📊 Include data, analysis or model results where relevant."
                )

                st.write(
                    "📝 Explain your tools and results clearly."
                )

                st.write(
                    "🌐 Upload projects to GitHub with a README."
                )

            elif (
                "interview" in q
                or "question" in q
                or "hr" in q
            ):

                if interview_score < 50:

                    st.warning(
                        "Practice more interview questions and explain your projects clearly."
                    )

                elif interview_score < 80:

                    st.info(
                        "Your preparation is developing. Practice structured answers."
                    )

                else:

                    st.success(
                        "Your practice score is strong. Continue practicing unexpected questions."
                    )

            elif (
                "job" in q
                or "career" in q
                or "company" in q
            ):

                st.write(
                    f"Your selected career goal is **{career}**."
                )

                st.write(
                    "Focus on:"
                )

                st.write(
                    "1️⃣ Technical skills"
                )

                st.write(
                    "2️⃣ Practical projects"
                )

                st.write(
                    "3️⃣ Resume and portfolio"
                )

                st.write(
                    "4️⃣ Interview preparation"
                )

            elif (
                "roadmap" in q
                or "path" in q
                or "steps" in q
            ):

                roadmap = CAREER_ROADMAPS.get(
                    career,
                    []
                )

                st.write(
                    f"### 🗺️ {career} Roadmap"
                )

                for index, step in enumerate(
                    roadmap,
                    start=1
                ):

                    st.write(
                        f"**Step {index}:** {step}"
                    )

            else:

                st.write(
                    f"Your current target is **{career}**."
                )

                st.write(
                    "Based on your profile, focus on:"
                )

                if resume_score < 70:
                    st.write(
                        "📄 Improving your resume"
                    )

                if skill_score < 70:
                    st.write(
                        "🧠 Learning missing skills"
                    )

                if project_score < 70:
                    st.write(
                        "💻 Building projects"
                    )

                if interview_score < 70:
                    st.write(
                        "🎤 Practicing interviews"
                    )

    st.divider()

    st.subheader(
        "🚀 Your Personalized Next Steps"
    )

    if resume_score < 60:

        st.write(
            "1️⃣ Improve your resume"
        )

    elif skill_score < 60:

        st.write(
            "1️⃣ Focus on missing technical skills"
        )

    elif project_score < 60:

        st.write(
            "1️⃣ Build more practical projects"
        )

    elif interview_score < 60:

        st.write(
            "1️⃣ Practice interview questions"
        )

    else:

        st.write(
            "1️⃣ Continue building practical experience"
        )

    st.write(
        "2️⃣ Update your GitHub portfolio regularly"
    )

    st.write(
        "3️⃣ Practice explaining your projects"
    )

    st.write(
        "4️⃣ Keep your resume tailored to target roles"
    )


# =========================================================
# JOB MATCHER
# =========================================================

elif page == "💼 Job Matcher":

    st.markdown(
        '<div class="main-title">💼 AI Job Matcher</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="subtitle">Find jobs that match your resume skills and career goal.</div>',
        unsafe_allow_html=True
    )

    resume_text = st.session_state.resume_text

    if not resume_text:

        st.warning(
            "📄 Please analyze your resume first in Advanced Resume Analyzer."
        )

        st.stop()

    user_skills = extract_skills(
        resume_text
    )

    user_skills_lower = [
        skill.lower()
        for skill in user_skills
    ]

    selected_career = st.selectbox(
        "🎯 Select your target career",
        list(CAREER_SKILLS.keys()),
        key="job_match_career"
    )

    locations = sorted(
        list(
            set(
                job["location"]
                for job in JOB_DATABASE
            )
        )
    )

    selected_location = st.selectbox(
        "📍 Preferred Location",
        ["All"] + locations,
        key="job_location"
    )

    matched_jobs = []

    for job in JOB_DATABASE:

        if job["role"] != selected_career:
            continue

        if (
            selected_location != "All"
            and job["location"] != selected_location
        ):
            continue

        required_skills = job["skills"]

        matched = [
            skill
            for skill in required_skills
            if skill.lower() in user_skills_lower
        ]

        missing = [
            skill
            for skill in required_skills
            if skill.lower() not in user_skills_lower
        ]

        if required_skills:

            match_percentage = round(
                len(matched)
                / len(required_skills)
                * 100
            )

        else:

            match_percentage = 0

        matched_jobs.append(
            {
                "job": job,
                "matched": matched,
                "missing": missing,
                "match": match_percentage
            }
        )

    matched_jobs.sort(
        key=lambda x: x["match"],
        reverse=True
    )

    st.divider()

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "📄 Resume Skills",
            len(user_skills)
        )

    with col2:

        st.metric(
            "💼 Jobs Found",
            len(matched_jobs)
        )

    with col3:

        if matched_jobs:

            average_match = round(
                sum(
                    item["match"]
                    for item in matched_jobs
                )
                / len(matched_jobs)
            )

        else:

            average_match = 0

        st.metric(
            "📊 Average Match",
            f"{average_match}%"
        )

    st.subheader(
        "🧠 Your Detected Skills"
    )

    if user_skills:

        st.info(
            " • ".join(user_skills)
        )

    else:

        st.warning(
            "No skills were detected from your resume."
        )

    st.divider()

    st.subheader(
        "🔎 Matching Jobs"
    )

    if not matched_jobs:

        st.info(
            "No jobs match the selected career and location."
        )

    else:

        for index, item in enumerate(
            matched_jobs,
            start=1
        ):

            job = item["job"]

            match = item["match"]

            matched = item["matched"]

            missing = item["missing"]

            with st.container():

                st.markdown(
                    f"### {index}. {job['role']}"
                )

                col1, col2, col3 = st.columns(3)

                with col1:

                    st.write(
                        f"🏢 **Company:** {job['company']}"
                    )

                with col2:

                    st.write(
                        f"📍 **Location:** {job['location']}"
                    )

                with col3:

                    st.write(
                        f"📊 **Skill Match:** {match}%"
                    )

                st.progress(
                    match / 100
                )

                if matched:

                    st.write(
                        "**✅ Matching Skills:**"
                    )

                    st.success(
                        ", ".join(matched)
                    )

                if missing:

                    st.write(
                        "**⚠️ Skills To Improve:**"
                    )

                    st.warning(
                        ", ".join(missing)
                    )

                if match >= 80:

                    st.info(
                        "Your resume contains most of the skills listed for this role."
                    )

                elif match >= 50:

                    st.info(
                        "You have several relevant skills, but some requirements are missing."
                    )

                else:

                    st.info(
                        "Consider learning the missing skills before targeting this role."
                    )

                st.divider()

    st.subheader(
        f"🎯 Skills Commonly Used for {selected_career}"
    )

    career_skills = CAREER_SKILLS[
        selected_career
    ]

    for skill in career_skills:

        if skill.lower() in user_skills_lower:

            st.write(
                f"✅ {skill}"
            )

        else:

            st.write(
                f"⚠️ {skill}"
            )

    st.divider()

    st.subheader(
        "🚀 Next Step"
    )

    st.write(
        "Improve the missing skills shown above, update your resume, and run the Resume Analyzer again."
    )


# =========================================================
# CAREER ROADMAP
# =========================================================

elif page == "🎯 Career Roadmap":

    st.markdown(
        '<div class="main-title">🎯 Personalized Career Roadmap</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="subtitle">Follow a structured path from beginner to job-ready.</div>',
        unsafe_allow_html=True
    )

    career = st.selectbox(
        "🎯 Choose your target career",
        list(CAREER_ROADMAPS.keys()),
        key="roadmap_career"
    )

    ROADMAP_DETAILS = {

        "Data Analyst": {

            "Beginner": [
                "Python Basics",
                "Excel",
                "SQL Basics",
                "Statistics Fundamentals"
            ],

            "Intermediate": [
                "Pandas & NumPy",
                "Data Cleaning",
                "Power BI",
                "Data Visualization"
            ],

            "Advanced": [
                "Advanced SQL",
                "Business Analytics",
                "Advanced Power BI",
                "Portfolio Projects"
            ],

            "Projects": [
                "Sales Dashboard",
                "Customer Analysis",
                "Financial Data Analysis"
            ],

            "Career Skills": [
                "Python",
                "SQL",
                "Excel",
                "Power BI",
                "Statistics",
                "Pandas"
            ]
        },

        "Data Scientist": {

            "Beginner": [
                "Python",
                "Statistics",
                "SQL",
                "NumPy & Pandas"
            ],

            "Intermediate": [
                "Data Cleaning",
                "Machine Learning",
                "Feature Engineering",
                "Data Visualization"
            ],

            "Advanced": [
                "Deep Learning",
                "Model Deployment",
                "MLOps Basics",
                "Advanced Machine Learning"
            ],

            "Projects": [
                "Customer Churn Prediction",
                "House Price Prediction",
                "Healthcare Prediction System"
            ],

            "Career Skills": [
                "Python",
                "SQL",
                "Statistics",
                "Machine Learning",
                "Pandas",
                "NumPy",
                "Deep Learning"
            ]
        },

        "AI Engineer": {

            "Beginner": [
                "Python",
                "Programming Fundamentals",
                "Mathematics",
                "Data Structures"
            ],

            "Intermediate": [
                "Machine Learning",
                "Deep Learning",
                "NLP",
                "Computer Vision"
            ],

            "Advanced": [
                "Generative AI",
                "Large Language Models",
                "Model Deployment",
                "AI Applications"
            ],

            "Projects": [
                "AI Chatbot",
                "Image Classification System",
                "AI Recommendation System"
            ],

            "Career Skills": [
                "Python",
                "Machine Learning",
                "Deep Learning",
                "NLP",
                "Computer Vision"
            ]
        },

        "Machine Learning Engineer": {

            "Beginner": [
                "Python",
                "Statistics",
                "Linear Algebra",
                "SQL"
            ],

            "Intermediate": [
                "Machine Learning",
                "Feature Engineering",
                "Model Evaluation",
                "Scikit-learn"
            ],

            "Advanced": [
                "Deep Learning",
                "Model Deployment",
                "MLOps",
                "Cloud Platforms"
            ],

            "Projects": [
                "Spam Detection",
                "Customer Churn Prediction",
                "Recommendation System"
            ],

            "Career Skills": [
                "Python",
                "Machine Learning",
                "Statistics",
                "Scikit-learn",
                "Deep Learning"
            ]
        },

        "Business Analyst": {

            "Beginner": [
                "Excel",
                "Business Fundamentals",
                "Communication",
                "Basic Statistics"
            ],

            "Intermediate": [
                "SQL",
                "Power BI",
                "Requirements Analysis",
                "Data Visualization"
            ],

            "Advanced": [
                "Business Strategy",
                "Stakeholder Management",
                "Advanced Analytics",
                "Business Intelligence"
            ],

            "Projects": [
                "Sales Performance Dashboard",
                "Customer Segmentation",
                "Business KPI Analysis"
            ],

            "Career Skills": [
                "Excel",
                "SQL",
                "Power BI",
                "Statistics",
                "Communication"
            ]
        }
    }

    roadmap = ROADMAP_DETAILS.get(
        career,
        {}
    )

    state_key = f"roadmap_{career}"

    if state_key not in st.session_state:

        st.session_state[state_key] = {}

    progress_state = st.session_state[state_key]

    st.divider()

    st.subheader(
        "🌱 Beginner Level"
    )

    beginner_items = roadmap.get(
        "Beginner",
        []
    )

    beginner_done = 0

    for index, item in enumerate(
        beginner_items
    ):

        key = f"{career}_beginner_{index}"

        checked = st.checkbox(
            item,
            value=progress_state.get(
                key,
                False
            ),
            key=key
        )

        progress_state[key] = checked

        if checked:

            beginner_done += 1

    st.divider()

    st.subheader(
        "🚀 Intermediate Level"
    )

    intermediate_items = roadmap.get(
        "Intermediate",
        []
    )

    intermediate_done = 0

    for index, item in enumerate(
        intermediate_items
    ):

        key = f"{career}_intermediate_{index}"

        checked = st.checkbox(
            item,
            value=progress_state.get(
                key,
                False
            ),
            key=key
        )

        progress_state[key] = checked

        if checked:

            intermediate_done += 1

    st.divider()

    st.subheader(
        "🏆 Advanced Level"
    )

    advanced_items = roadmap.get(
        "Advanced",
        []
    )

    advanced_done = 0

    for index, item in enumerate(
        advanced_items
    ):

        key = f"{career}_advanced_{index}"

        checked = st.checkbox(
            item,
            value=progress_state.get(
                key,
                False
            ),
            key=key
        )

        progress_state[key] = checked

        if checked:

            advanced_done += 1

    st.divider()

    st.subheader(
        "💻 Portfolio Projects"
    )

    project_items = roadmap.get(
        "Projects",
        []
    )

    projects_done = 0

    for index, item in enumerate(
        project_items
    ):

        key = f"{career}_project_{index}"

        checked = st.checkbox(
            item,
            value=progress_state.get(
                key,
                False
            ),
            key=key
        )

        progress_state[key] = checked

        if checked:

            projects_done += 1

    total_items = (
        len(beginner_items)
        + len(intermediate_items)
        + len(advanced_items)
        + len(project_items)
    )

    completed_items = (
        beginner_done
        + intermediate_done
        + advanced_done
        + projects_done
    )

    if total_items > 0:

        roadmap_progress = round(
            completed_items
            / total_items
            * 100
        )

    else:

        roadmap_progress = 0

    st.divider()

    st.subheader(
        "📊 Roadmap Progress"
    )

    st.metric(
        "Career Roadmap Completion",
        f"{roadmap_progress}%"
    )

    st.progress(
        roadmap_progress / 100
    )

    st.write(
        f"Completed **{completed_items}** of **{total_items}** roadmap items."
    )

    st.divider()

    st.subheader(
        "🎯 Current Level"
    )

    if roadmap_progress < 25:

        st.info(
            "🌱 Beginner — Build your fundamentals first."
        )

    elif roadmap_progress < 50:

        st.info(
            "📚 Foundation Built — Start developing intermediate skills."
        )

    elif roadmap_progress < 75:

        st.info(
            "🚀 Intermediate — Focus on advanced skills and projects."
        )

    elif roadmap_progress < 100:

        st.info(
            "💼 Nearly Job Ready — Complete the remaining roadmap items."
        )

    else:

        st.success(
            "🏆 Roadmap Completed — Continue applying your skills through projects and opportunities."
        )

    st.divider()

    st.subheader(
        f"🧠 Important Skills for {career}"
    )

    career_skills = roadmap.get(
        "Career Skills",
        []
    )

    resume_text = st.session_state.resume_text

    detected_skills = extract_skills(
        resume_text
    )

    detected_lower = [
        skill.lower()
        for skill in detected_skills
    ]

    for skill in career_skills:

        if skill.lower() in detected_lower:

            st.success(
                f"✅ {skill}"
            )

        else:

            st.warning(
                f"⚠️ {skill}"
            )

    st.divider()

    st.subheader(
        "🚀 Recommended Next Action"
    )

    if beginner_done < len(beginner_items):

        st.write(
            "👉 Complete your remaining beginner-level fundamentals."
        )

    elif intermediate_done < len(intermediate_items):

        st.write(
            "👉 Move to intermediate technical skills."
        )

    elif advanced_done < len(advanced_items):

        st.write(
            "👉 Start advanced topics and practical implementation."
        )

    elif projects_done < len(project_items):

        st.write(
            "👉 Complete more portfolio projects."
        )

    else:

        st.success(
            "🔥 You completed this roadmap!"
        )


# =========================================================
# SKILL GAP ANALYZER
# =========================================================

elif page == "🧠 Skill Gap Analyzer":

    st.markdown(
        '<div class="main-title">🧠 AI Skill Gap Analyzer</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="subtitle">Compare your current skills with the skills required for your target career.</div>',
        unsafe_allow_html=True
    )

    career = st.selectbox(
        "🎯 Select your target career",
        list(CAREER_SKILLS.keys()),
        key="skill_gap_career"
    )

    required_skills = CAREER_SKILLS[
        career
    ]

    resume_text = st.session_state.resume_text

    if not resume_text:

        st.warning(
            "📄 Please analyze your resume first using the Advanced Resume Analyzer."
        )

        st.stop()

    detected_skills = extract_skills(
        resume_text
    )

    detected_lower = [
        skill.lower()
        for skill in detected_skills
    ]

    matched_skills = []

    missing_skills = []

    for skill in required_skills:

        if skill.lower() in detected_lower:

            matched_skills.append(skill)

        else:

            missing_skills.append(skill)

    total_skills = len(required_skills)

    matched_count = len(matched_skills)

    if total_skills > 0:

        match_percentage = round(
            matched_count
            / total_skills
            * 100
        )

    else:

        match_percentage = 0

    st.session_state.skill_score = match_percentage

    st.divider()

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "🎯 Skill Match",
            f"{match_percentage}%"
        )

    with col2:

        st.metric(
            "✅ Skills Matched",
            matched_count
        )

    with col3:

        st.metric(
            "⚠️ Skills Missing",
            len(missing_skills)
        )

    st.subheader(
        "📊 Skill Match Progress"
    )

    st.progress(
        match_percentage / 100
    )

    st.write(
        f"You currently match **{match_percentage}%** of the skills required for **{career}**."
    )

    st.divider()

    col1, col2 = st.columns(2)

    with col1:

        st.subheader(
            "✅ Skills You Have"
        )

        if matched_skills:

            for skill in matched_skills:

                st.success(
                    f"✓ {skill}"
                )

        else:

            st.info(
                "No required skills detected."
            )

    with col2:

        st.subheader(
            "⚠️ Skills To Learn"
        )

        if missing_skills:

            for skill in missing_skills:

                st.warning(
                    f"→ {skill}"
                )

        else:

            st.success(
                "🎉 You already have all required skills!"
            )

    st.divider()

    st.subheader(
        "📚 Recommended Learning Path"
    )

    if missing_skills:

        for index, skill in enumerate(
            missing_skills,
            start=1
        ):

            st.markdown(
                f"### {index}️⃣ {skill}"
            )

            st.progress(
                max(
                    0.1,
                    match_percentage / 100
                )
            )

    else:

        st.success(
            "🔥 Your current skill set covers the listed career requirements."
        )

    st.divider()

    st.subheader(
        f"🗺️ {career} Learning Roadmap"
    )

    roadmap = CAREER_ROADMAPS.get(
        career,
        []
    )

    for index, step in enumerate(
        roadmap,
        start=1
    ):

        st.write(
            f"**Step {index}:** {step}"
        )

    st.divider()

    st.subheader(
        "🚀 Your Action Plan"
    )

    if match_percentage < 40:

        st.error(
            "🌱 Start with the fundamentals before moving to advanced topics."
        )

    elif match_percentage < 70:

        st.warning(
            "📚 You have a foundation. Focus on your missing technical skills."
        )

    elif match_percentage < 90:

        st.info(
            "🚀 You're progressing well. Build projects using your missing skills."
        )

    else:

        st.success(
            "🏆 Your skill coverage is strong. Focus on projects and interview preparation."
        )

    st.write(
        "💡 Tip: After learning a missing skill, update your resume and run the Resume Analyzer again."
    )


# =========================================================
# INTERVIEW PREPARATION
# =========================================================

elif page == "🎤 Interview Preparation":

    st.markdown(
        '<div class="main-title">🎤 AI Interview Preparation</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="subtitle">Practice career-specific interview questions using your resume skills.</div>',
        unsafe_allow_html=True
    )

    career = st.selectbox(
        "🎯 Select your target career",
        list(CAREER_SKILLS.keys()),
        key="interview_career"
    )

    resume_text = st.session_state.resume_text

    detected_skills = extract_skills(
        resume_text
    )

    interview_type = st.selectbox(
        "📝 Interview Type",
        [
            "HR Interview",
            "Technical Interview",
            "Resume Based Interview",
            "Mixed Interview"
        ],
        key="interview_type"
    )

    HR_QUESTIONS = [

        "Tell me about yourself.",

        "Why should we hire you?",

        "What are your strengths?",

        "What is your biggest weakness?",

        "Where do you see yourself in five years?",

        "Why did you choose Artificial Intelligence and Data Science?",

        "Tell me about a challenge you faced and how you handled it.",

        "Why are you interested in this career?"
    ]

    TECHNICAL_QUESTIONS = {

        "Data Analyst": [

            "What is the difference between SQL and NoSQL?",

            "What is a primary key in SQL?",

            "What is data cleaning?",

            "What is the difference between mean, median and mode?",

            "What is data visualization?",

            "What is Power BI used for?",

            "What is the difference between INNER JOIN and LEFT JOIN?",

            "What is an Excel Pivot Table?"
        ],

        "Data Scientist": [

            "What is the difference between supervised and unsupervised learning?",

            "What is overfitting?",

            "What is cross-validation?",

            "What is feature engineering?",

            "What is the difference between classification and regression?",

            "What is precision and recall?",

            "What is a confusion matrix?",

            "What is the difference between NumPy and Pandas?"
        ],

        "AI Engineer": [

            "What is Artificial Intelligence?",

            "What is Machine Learning?",

            "What is Deep Learning?",

            "What is an artificial neural network?",

            "What is NLP?",

            "What is computer vision?",

            "What is the difference between AI and ML?",

            "What is an AI model?"
        ],

        "Machine Learning Engineer": [

            "What is supervised learning?",

            "What is unsupervised learning?",

            "What is overfitting?",

            "What is regularization?",

            "What is feature scaling?",

            "What is gradient descent?",

            "What is model deployment?",

            "What is the difference between training and testing data?"
        ],

        "Business Analyst": [

            "What does a Business Analyst do?",

            "What is requirements gathering?",

            "What is a stakeholder?",

            "What is SWOT analysis?",

            "What is a business requirement document?",

            "What is data-driven decision making?",

            "What is the difference between a functional and non-functional requirement?",

            "How do you handle conflicting stakeholder requirements?"
        ]
    }

    if interview_type == "HR Interview":

        questions = HR_QUESTIONS.copy()

    elif interview_type == "Technical Interview":

        questions = TECHNICAL_QUESTIONS.get(
            career,
            []
        ).copy()

    elif interview_type == "Resume Based Interview":

        questions = [

            "Explain one of the projects mentioned in your resume.",

            "What technologies have you used in your projects?",

            "What was your role in your project?",

            "What was the biggest challenge in your project?",

            "How did you test your project?",

            "What did you learn from your project?",

            "Which technical skill are you currently improving?",

            "How would you improve your project in the future?"
        ]

    else:

        questions = (
            HR_QUESTIONS[:4]
            + TECHNICAL_QUESTIONS.get(
                career,
                []
            )[:4]
        )

    if detected_skills:

        st.divider()

        st.subheader(
            "🧠 Questions Based on Your Resume Skills"
        )

        skill_questions = []

        for skill in detected_skills[:5]:

            skill_lower = skill.lower()

            if skill_lower == "python":

                skill_questions.append(
                    "Explain the difference between a list, tuple and dictionary in Python."
                )

            elif skill_lower == "sql":

                skill_questions.append(
                    "Explain the difference between WHERE and HAVING in SQL."
                )

            elif skill_lower == "machine learning":

                skill_questions.append(
                    "Explain the basic steps involved in building a Machine Learning model."
                )

            elif skill_lower == "pandas":

                skill_questions.append(
                    "What is Pandas and why is it used in Data Science?"
                )

            elif skill_lower == "numpy":

                skill_questions.append(
                    "What is NumPy and why is it useful for numerical computing?"
                )

            elif skill_lower == "power bi":

                skill_questions.append(
                    "What is Power BI and how is it used for business analytics?"
                )

            elif skill_lower == "excel":

                skill_questions.append(
                    "How would you use Excel to analyze a large dataset?"
                )

            elif skill_lower == "deep learning":

                skill_questions.append(
                    "What is the difference between Machine Learning and Deep Learning?"
                )

        questions.extend(
            skill_questions
        )

    questions = list(
        dict.fromkeys(
            questions
        )
    )

    if not questions:

        st.warning(
            "No interview questions available."
        )

    else:

        max_questions = min(
            10,
            len(questions)
        )

        number_of_questions = st.slider(
            "Number of questions",
            min_value=3,
            max_value=max_questions,
            value=min(5, max_questions),
            key="question_count"
        )

        selected_questions = questions[
            :number_of_questions
        ]

        st.divider()

        st.subheader(
            f"🎯 {career} Interview Questions"
        )

        answers = {}

        for index, question in enumerate(
            selected_questions,
            start=1
        ):

            st.markdown(
                f"### Question {index}"
            )

            st.write(
                question
            )

            answers[index] = st.text_area(
                "Your Answer",
                key=f"interview_answer_{index}",
                height=120
            )

        st.divider()

        if st.button(
            "📊 Evaluate My Interview",
            use_container_width=True,
            key="evaluate_interview"
        ):

            answered = 0

            strong_answers = 0

            for index, answer in answers.items():

                if answer.strip():

                    answered += 1

                    word_count = len(
                        answer.split()
                    )

                    if word_count >= 30:

                        strong_answers += 1

            interview_score = round(
                strong_answers
                / len(selected_questions)
                * 100
            )

            st.session_state.interview_score = (
                interview_score
            )

            st.subheader(
                "📊 Interview Performance"
            )

            col1, col2, col3 = st.columns(3)

            with col1:

                st.metric(
                    "Interview Score",
                    f"{interview_score}%"
                )

            with col2:

                st.metric(
                    "Questions Answered",
                    answered
                )

            with col3:

                st.metric(
                    "Strong Answers",
                    strong_answers
                )

            st.progress(
                interview_score / 100
            )

            if interview_score < 40:

                st.warning(
                    "🌱 Keep practicing. Try giving longer, structured answers."
                )

            elif interview_score < 70:

                st.info(
                    "📚 Good progress. Add examples and explain your reasoning clearly."
                )

            elif interview_score < 90:

                st.success(
                    "🚀 Strong preparation. Continue practicing technical depth."
                )

            else:

                st.success(
                    "🏆 Excellent practice performance. Keep refining your answers."
                )

            st.subheader(
                "💡 Answer Feedback"
            )

            for index, answer in answers.items():

                if not answer.strip():

                    st.warning(
                        f"Question {index}: No answer provided."
                    )

                elif len(answer.split()) < 15:

                    st.warning(
                        f"Question {index}: Try giving a more detailed answer with an example."
                    )

                elif len(answer.split()) < 30:

                    st.info(
                        f"Question {index}: Good start. Add more explanation and specific examples."
                    )

                else:

                    st.success(
                        f"Question {index}: Detailed answer provided."
                    )

            st.subheader(
                "📝 Recommended Answer Structure"
            )

            st.write(
                "1️⃣ Start with a direct answer."
            )

            st.write(
                "2️⃣ Explain your reasoning or technical concept."
            )

            st.write(
                "3️⃣ Give a real example from your project or experience."
            )

            st.write(
                "4️⃣ End with the result or what you learned."
            )

    st.divider()

    st.subheader(
        "🔥 Interview Preparation Tips"
    )

    tips = [

        "Know every project mentioned in your resume.",

        "Practice explaining technical concepts in simple language.",

        "Prepare examples for your strengths and achievements.",

        "Do not memorize answers word-for-word.",

        "Use real project experiences whenever possible.",

        "Practice speaking clearly and confidently."
    ]

    for tip in tips:

        st.write(
            f"• {tip}"
        )


# =========================================================
# FOOTER
# =========================================================

st.divider()

st.markdown(
    """
    <div class="footer">
        🚀 <b>AI Career Intelligence Platform</b><br>
        Resume • Skills • Jobs • Roadmap • Interviews<br><br>
        Built with Python + Streamlit
    </div>
    """,
    unsafe_allow_html=True
)