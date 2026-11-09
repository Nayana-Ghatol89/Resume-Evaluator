import io
import re
from datetime import date
import streamlit as st
from docx import Document
from pypdf import PdfReader

st.set_page_config(
    page_title="Resume Evaluator",
    page_icon="📄",
    layout="wide",
)

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Space+Grotesk:wght@500;600;700&display=swap');
    :root { --ink: #172a3a; --muted: #647789; --teal: #087f8c; --mint: #dff3ef; --line: #d9e5e8; --paper: #ffffff; }
    .stApp { background: #eef4f3; color: var(--ink); }
    [data-testid="stHeader"] { background: rgba(238, 244, 243, 0.88); }
    [data-testid="stAppViewContainer"] > .main { background: radial-gradient(circle at 85% 0%, #d9eeea 0, transparent 30rem); }
    .block-container { max-width: 1120px; padding: 2.5rem 3rem 4rem; }
    .hero { border-bottom: 1px solid var(--line); padding: 1rem 0 2rem; margin-bottom: 1.6rem; }
    .hero-kicker { color: var(--teal); font: 700 .75rem 'DM Sans', sans-serif; letter-spacing: .16em; text-transform: uppercase; }
    .hero h1 { color: var(--ink); font: 700 3.2rem/1 'Space Grotesk', sans-serif; letter-spacing: 0; margin: .55rem 0 .7rem; }
    .hero p { color: var(--muted); font: 400 1.05rem/1.6 'DM Sans', sans-serif; margin: 0; max-width: 38rem; }
    h2, h3 { color: var(--ink); font-family: 'Space Grotesk', sans-serif; letter-spacing: 0; }
    [data-testid="stFileUploader"] { background: var(--paper); border: 1px solid var(--line); border-radius: 10px; padding: .8rem 1rem; box-shadow: 0 8px 24px rgba(23, 42, 58, .05); }
    [data-testid="stFileUploaderDropzone"] { border: 1px dashed #9ac6c5; background: #f6fbfa; }
    [data-testid="stMetric"] { background: var(--paper); border: 1px solid var(--line); border-radius: 10px; padding: 1rem 1.1rem; box-shadow: 0 8px 24px rgba(23, 42, 58, .05); }
    [data-testid="stMetricLabel"] { color: var(--muted); font-family: 'DM Sans', sans-serif; }
    [data-testid="stMetricValue"] { color: var(--teal); font-family: 'Space Grotesk', sans-serif; }
    [data-testid="stAlert"] { border-radius: 8px; }
    .snapshot-intro { background: var(--paper); border-left: 4px solid var(--teal); border-radius: 0 10px 10px 0; padding: .8rem 1rem; margin: .5rem 0 1rem; box-shadow: 0 8px 24px rgba(23, 42, 58, .04); }
    .role-strip { background: #173b4b; color: #edf9f7; border-radius: 10px; padding: 1rem 1.15rem; margin: 0 0 1.25rem; }
    .role-strip strong { color: #83e0d3; font: 600 .72rem 'DM Sans', sans-serif; letter-spacing: .12em; text-transform: uppercase; }
    .role-strip span { display: inline-block; background: rgba(255,255,255,.1); border: 1px solid rgba(255,255,255,.16); border-radius: 999px; color: #fff; font: 500 .86rem 'DM Sans', sans-serif; margin: .55rem .35rem 0 0; padding: .35rem .65rem; }
    .stChatMessage { border: 1px solid var(--line); border-radius: 10px; background: var(--paper); }
    @media (max-width: 700px) { .block-container { padding: 1.5rem 1rem 3rem; } .hero h1 { font-size: 2.4rem; } }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    "<div class='hero'><div class='hero-kicker'>Resume intelligence</div><h1>Resume Evaluator</h1><p>Turn a resume upload into practical, evidence-based feedback.</p></div>",
    unsafe_allow_html=True,
)

SKILL_DATABASE = [
    "python", "java", "javascript", "typescript", "c++", "c#", "sql", "nosql",
    "mongodb", "postgresql", "machine learning", "ai", "deep learning", "nlp",
    "excel", "power bi", "tableau", "pandas", "numpy", "scikit-learn",
    "tensorflow", "pytorch", "react", "angular", "vue", "nodejs", "django",
    "flask", "aws", "azure", "gcp", "docker", "kubernetes", "git", "html",
    "css", "rest api", "graphql", "statistics", "r", "scala", "spark", "hadoop",
    "etl", "data warehousing", "data visualization", "business intelligence",
    "data analysis", "data mining", "data engineering", "big data", "cloud computing",
    "devops", "agile", "scrum", "pyspark", "selenium", "testing",
    "manual testing", "automation testing", "jira", "spring", "dax",
    "project management", "leadership", "team management", "communication",
    "risk management", "stakeholder management", "budget management", "roadmapping",
    "figma", "adobe xd", "sketch", "ui design", "ux design", "wireframing",
    "prototyping", "user research", "usability testing", "accessibility",
    "responsive design", "bootstrap", "tailwind css", "sass", "webpack",
    "jest", "cypress", "postman", "pytest", "junit", "opencv",
    "generative ai", "llm", "prompt engineering", "rag", "retrieval augmented generation",
    "vector database", "embedding models", "langchain", "langgraph", "llamaindex",
    "hugging face", "transformers", "fine-tuning", "reinforcement learning",
    "computer vision", "opencv", "keras", "xgboost", "model deployment", "mlops",
    "feature engineering", "model evaluation", "hyperparameter tuning", "model monitoring",
    "github actions", "ci/cd", "streamlit", "fastapi", "data cleaning", "data storytelling",
    "excel advanced", "look up functions", "pivot tables", "dashboarding", "ab testing",
    "experiment design", "data mining", "snowflake", "databricks", "airflow", "dbt",
    "kafka", "bigquery", "redshift", "terraform", "linux", "bash", "oracle", "mysql",
    "sql server", "tsql", "stored procedures", "query optimization", "indexing",
    "data modeling", "asyncio", "oop", "design patterns", "sqlalchemy", "redis", "celery",
    "spring boot", "microservices", "hibernate", "maven", "gradle", "multithreading",
    "power query", "playwright", "appium", "cucumber", "bdd", "regression testing",
    "smoke testing", "api testing", "performance testing", "test cases", "defect lifecycle",
    "functional testing", "user acceptance testing", "next.js", "storybook", "design systems",
    "wcag", "adobe illustrator", "persona", "information architecture", "user journey",
    "confluence", "trello", "asana", "sprint planning", "resource planning", "change management",
    "stakeholder communication", "risk analysis", "budget tracking", "program management",
]

ROLE_PROFILES = {
    "Data Analyst": {"required": {"sql", "excel", "statistics", "data analysis", "data visualization", "python", "pandas", "power bi", "tableau", "business intelligence", "dax", "etl", "data cleaning", "data storytelling", "excel advanced", "pivot tables", "dashboarding", "look up functions", "data modeling"}},
    "Data Scientist": {"required": {"python", "statistics", "machine learning", "pandas", "numpy", "scikit-learn", "tensorflow", "pytorch", "sql", "nlp", "deep learning", "data visualization", "feature engineering", "model evaluation", "hyperparameter tuning", "model monitoring", "ab testing", "experiment design", "reinforcement learning", "xgboost", "keras"}},
    "Data Engineer": {"required": {"python", "sql", "etl", "data engineering", "spark", "pyspark", "hadoop", "aws", "azure", "gcp", "docker", "kubernetes", "data warehousing", "big data", "cloud computing", "snowflake", "databricks", "airflow", "dbt", "kafka", "bigquery", "redshift", "terraform", "linux", "bash"}},
    "Python Developer": {"required": {"python", "django", "flask", "sql", "git", "rest api", "docker", "pytest", "postman", "github actions", "ci/cd", "fastapi", "asyncio", "oop", "design patterns", "sqlalchemy", "redis", "celery", "kubernetes", "linux"}},
    "Java Developer": {"required": {"java", "sql", "spring", "rest api", "git", "docker", "junit", "postman", "github actions", "ci/cd", "spring boot", "microservices", "hibernate", "maven", "gradle", "multithreading", "kubernetes", "linux"}},
    "SQL Developer": {"required": {"sql", "postgresql", "mongodb", "nosql", "etl", "data warehousing", "python", "data analysis", "mysql", "oracle", "sql server", "tsql", "stored procedures", "query optimization", "indexing", "data modeling"}},
    "Power BI Developer": {"required": {"power bi", "sql", "dax", "excel", "data visualization", "data analysis", "business intelligence", "etl", "power query", "data modeling", "dashboarding", "data storytelling"}},
    "Automation Tester": {"required": {"testing", "automation testing", "selenium", "java", "python", "sql", "git", "agile", "rest api", "jira", "postman", "cypress", "jest", "ci/cd", "playwright", "appium", "cucumber", "bdd", "api testing", "regression testing", "smoke testing", "performance testing"}},
    "Manual Tester": {"required": {"testing", "manual testing", "sql", "java", "python", "rest api", "jira", "agile", "postman", "test cases", "defect lifecycle", "functional testing", "user acceptance testing", "regression testing"}},
    "UI/UX Developer": {"required": {"html", "css", "javascript", "ui design", "ux design", "figma", "react", "responsive design", "wireframing", "prototyping", "user research", "usability testing", "accessibility", "adobe xd", "sketch", "typescript", "tailwind css", "bootstrap", "sass", "webpack", "jest", "git", "next.js", "storybook", "design systems", "wcag", "adobe illustrator", "persona", "information architecture", "user journey"}},
    "AI/ML Developer": {"required": {"python", "machine learning", "statistics", "deep learning", "ai", "nlp", "pandas", "numpy", "scikit-learn", "tensorflow", "pytorch", "sql", "data engineering", "opencv", "generative ai", "llm", "prompt engineering", "rag", "retrieval augmented generation", "vector database", "embedding models", "langchain", "langgraph", "llamaindex", "hugging face", "transformers", "fine-tuning", "reinforcement learning", "computer vision", "keras", "xgboost", "model deployment", "mlops", "feature engineering", "model evaluation", "hyperparameter tuning", "model monitoring", "docker", "fastapi"}},
    "Project Manager": {"required": {"project management", "leadership", "team management", "agile", "scrum", "jira", "communication", "risk management", "stakeholder management", "budget management", "roadmapping", "devops", "git", "confluence", "trello", "asana", "sprint planning", "resource planning", "change management", "stakeholder communication", "risk analysis", "budget tracking", "program management"}},
}

ANALYSIS_VERSION = "2026-11-04-target-role-1"

SKILL_DISPLAY_NAMES = {
    "sql": "SQL",
    "nosql": "NoSQL",
    "mongodb": "MongoDB",
    "postgresql": "PostgreSQL",
    "javascript": "JavaScript",
    "typescript": "TypeScript",
    "pandas": "Pandas",
    "numpy": "NumPy",
    "scikit-learn": "Scikit-learn",
    "tensorflow": "TensorFlow",
    "pytorch": "PyTorch",
    "power bi": "Power BI",
    "tableau": "Tableau",
    "nodejs": "Node.js",
    "rest api": "REST API",
    "graphql": "GraphQL",
    "html": "HTML",
    "css": "CSS",
    "aws": "AWS",
    "azure": "Azure",
    "gcp": "GCP",
    "etl": "ETL",
    "nlp": "NLP",
    "ai": "AI",
    "c++": "C++",
    "c#": "C#",
    "r": "R",
    "figma": "Figma",
    "next.js": "Next.js",
    "github actions": "GitHub Actions",
    "ci/cd": "CI/CD",
    "rag": "RAG",
    "llm": "LLM",
    "adobe xd": "Adobe XD",
    "ui design": "UI Design",
    "ux design": "UX Design",
}

def extract_text(uploaded_file):
    """Extract text without depending on the uploaded file pointer position."""
    file_bytes = uploaded_file.getvalue()
    filename = uploaded_file.name.lower()

    try:
        if filename.endswith(".pdf"):
            reader = PdfReader(io.BytesIO(file_bytes))
            return "\n".join(page.extract_text() or "" for page in reader.pages).strip()

        if filename.endswith(".docx"):
            document = Document(io.BytesIO(file_bytes))
            parts = [paragraph.text for paragraph in document.paragraphs if paragraph.text.strip()]

            for table in document.tables:
                for row in table.rows:
                    row_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                    if row_text:
                        parts.append("\t".join(row_text))

            for section in document.sections:
                for header_footer in (section.header, section.footer):
                    for paragraph in header_footer.paragraphs:
                        if paragraph.text.strip():
                            parts.append(paragraph.text.strip())
                    for table in header_footer.tables:
                        for row in table.rows:
                            row_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                            if row_text:
                                parts.append("\t".join(row_text))

            return "\n".join(parts).strip()
    except Exception as exc:
        raise ValueError(f"Could not read {uploaded_file.name}: {exc}") from exc

    return ""

def get_resume_analysis(resume_text):
    text = resume_text.lower()
    role_keywords = (
        r"(?:senior|sr\.?|junior|jr\.?|lead|principal|assistant|associate)?\s*"
        r"(?:data analyst|data scientist|data engineer|software engineer|"
        r"python developer|java developer|sql developer|power bi developer|"
        r"tableau developer|automation tester|manual tester|test engineer|"
        r"ui/ux developer|ui developer|ux developer|"
        r"ai/ml developer|ai developer|ml developer|machine learning developer|"
        r"project manager|product manager|program manager|supervisor|team lead|technical lead|"
        r"developer|engineer|analyst|tester|manager)"
    )
    heading_lines = {
        "professional summary", "resume summary", "work experience",
        "professional experience", "technical skills", "education",
        "certifications", "projects", "contact information", "objective",
        "profile", "skills", "experience", "summary", "data science",
        "data analytics", "software development", "professional profile",
    }
    role_match = re.search(
        rf"(?im)^[ \t]*(?:role|position|job title|title)[ \t]*[:\-][ \t]*(.+?)[ \t]*$",
        resume_text,
    )
    candidate_role = role_match.group(1).strip() if role_match else "Not Found"
    if candidate_role == "Not Found":
        role_match = re.search(rf"(?im)^[ \t]*({role_keywords})[ \t]*$", resume_text)
        if role_match:
            candidate_role = re.sub(r"\s+", " ", role_match.group(1)).strip().title()

    explicit_name = re.search(
        r"(?im)^[ \t]*(?:name|candidate name)[ \t]*[:\-][ \t]*(.+?)[ \t]*$",
        resume_text,
    )
    name_matches = re.findall(
        r"(?m)^[ \t]*([A-Za-z][A-Za-z.'-]+[ \t]+[A-Za-z][A-Za-z.'-]*(?:[ \t]+[A-Za-z][A-Za-z.'-]*){0,3})[ \t]*$",
        resume_text,
    )

    non_name_values = {
        re.sub(r"\s+", " ", value.strip()).lower()
        for value in heading_lines
    } | {"resume", "cv", "curriculum vitae"}
    role_line_pattern = re.compile(rf"^[ \t]*{role_keywords}[ \t]*$", re.IGNORECASE)

    early_lines = [line.strip() for line in resume_text.splitlines() if line.strip()][:5]
    def personal_info_name(line):
        value = re.sub(r"(?i)^\s*(?:name|candidate name)\s*[:\-]\s*", "", line).strip()
        value = re.split(r"\s*(?:\||•|·|;)\s*|\s+(?=[\w.+-]+@[\w.-]+)", value, maxsplit=1)[0].strip()
        if (
            not value
            or value.lower() in non_name_values
            or role_line_pattern.fullmatch(value)
            or re.search(r"https?://|linkedin|github|\d", value, re.IGNORECASE)
        ):
            return None
        return value

    early_name = next(
        (
            name for line in early_lines
            for name in [personal_info_name(line)]
            if name
        ),
        None,
    )
    candidate_name = (
        personal_info_name(explicit_name.group(1))
        if explicit_name
        else early_name or "Not Found"
    )
    candidate_name = candidate_name or early_name or "Not Found"
    normalized_candidate_name = re.sub(r"\s+", " ", candidate_name).lower()
    if normalized_candidate_name in non_name_values or candidate_name == "Not Found":
        candidate_name = next(
            (
                name.strip() for name in name_matches
                if re.sub(r"\s+", " ", name.strip()).lower() not in non_name_values
                and not role_line_pattern.fullmatch(name.strip())
                and not re.search(r"@|linkedin|github|https?://|\d", name, re.IGNORECASE)
            ),
            "Not Found",
        )
    skills = []
    for skill in SKILL_DATABASE:
        pattern = r"(?<![a-z0-9+#])" + re.escape(skill) + r"(?![a-z0-9+#])"
        if re.search(pattern, text):
            skills.append(SKILL_DISPLAY_NAMES.get(skill, skill.title()))

    skill_groups = {
        "Programming": {"python", "java", "javascript", "typescript", "c++", "c#", "r", "scala"},
        "Data and AI": {"sql", "nosql", "mongodb", "postgresql", "pandas", "numpy", "machine learning", "ai", "deep learning", "nlp", "scikit-learn", "tensorflow", "pytorch", "statistics", "spark", "hadoop", "pyspark"},
        "Web and UI/UX Development": {"html", "css", "javascript", "typescript", "react", "angular", "vue", "nodejs", "django", "flask", "rest api", "graphql", "figma", "adobe xd", "sketch", "ui design", "ux design", "wireframing", "prototyping", "user research", "usability testing", "accessibility", "responsive design"},
        "Cloud and Tools": {"aws", "azure", "gcp", "docker", "kubernetes", "git", "devops", "agile", "scrum"},
        "Business and Analytics": {"excel", "power bi", "tableau", "data visualization", "business intelligence", "data analysis", "data mining", "data engineering", "etl", "data warehousing", "big data", "cloud computing"},
    }
    grouped_skills = {
        group: [skill for skill in skills if skill.lower() in {SKILL_DISPLAY_NAMES.get(name, name.title()).lower() for name in skill_names}]
        for group, skill_names in skill_groups.items()
    }
    skill_details = []
    for skill in skills:
        skill_pattern = r"(?<![a-z0-9+#])" + re.escape(skill.lower()) + r"(?![a-z0-9+#])"
        occurrences = len(re.findall(skill_pattern, text))
        skill_context = re.search(
            rf"(?:{skill_pattern}).{{0,90}}(?:project|experience|developed|built|used|worked|implemented|proficient|expert)",
            text,
            re.IGNORECASE,
        )
        if occurrences >= 2 and skill_context:
            evidence = "Repeated use with experience context"
        elif occurrences >= 2:
            evidence = "Repeated mention"
        elif skill_context:
            evidence = "Experience or project context"
        else:
            evidence = "Single mention"
        skill_details.append({
            "skill": SKILL_DISPLAY_NAMES.get(skill.lower(), skill.title()),
            "occurrences": occurrences,
            "evidence": evidence,
        })
    skill_details.sort(key=lambda item: (item["occurrences"], item["evidence"] != "Single mention"), reverse=True)

    email_matches = re.findall(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+", text)
    phone_matches = re.findall(r"(?:\+?\d{1,3}[\s.-]?)?(?:\(?\d{3}\)?[\s.-]?)\d{3}[\s.-]?\d{4}", text)
    linkedin_matches = re.findall(
        r"(?:https?://)?(?:www\.)?linkedin\.com/in/[A-Za-z0-9%_-]+/?",
        text,
        re.IGNORECASE,
    )
    github_matches = re.findall(
        r"(?:https?://)?(?:www\.)?github\.com/[A-Za-z0-9_-]+/?",
        text,
        re.IGNORECASE,
    )
    degree_labels = []
    for label, pattern in {
        "Bachelor's degree": r"\b(bachelor|b\.s\.?|b\.a\.?|bs|ba)\b",
        "Master's degree": r"\b(master|m\.s\.?|m\.a\.?|ms|ma)\b",
        "PhD": r"\b(phd|ph\.d)\b",
    }.items():
        if re.search(pattern, text):
            degree_labels.append(label)

    education_status = "Found" if degree_labels else "Not Found"
    education_details = degree_labels

    experience = "Not Found"
    experience_period = ""
    experience_years = 0.0
    experience_evidence = "No years of experience detected"
    experience_source = "none"
    for pattern in [
        r"(\d+(?:\.\d+)?)\+?\s*(?:years?|yrs?)(?:\s+of)?\s+(?:professional\s+)?experience",
        r"(?:total|overall)\s+experience\s*[:\-]?\s*(\d+(?:\.\d+)?)\+?\s*(?:years?|yrs?)?",
        r"experience\s*[:\-]?\s*(\d+(?:\.\d+)?)\+?\s*(?:years?|yrs?)",
    ]:
        match = re.search(pattern, text)
        if match:
            experience_years = float(match.group(1))
            experience = f"{match.group(1)} Years"
            experience_evidence = f"Matched: {match.group(0)}"
            experience_source = "explicit"
            break

    if experience_years == 0:
        date_pattern = re.compile(
            r"\b(19\d{2}|20\d{2})\s*(?:-|–|—|to)\s*(present|current|19\d{2}|20\d{2})?(?=\s|$)",
            re.IGNORECASE,
        )
        employment_terms = r"experience|employment|work history|worked|developer|engineer|analyst|tester|manager|supervisor|lead"
        education_terms = r"education|university|college|school|bachelor|master|phd|degree"
        employment_ranges = []
        resume_lines = [line.strip() for line in text.splitlines() if line.strip()]
        for line_index, line in enumerate(resume_lines):
            context_lines = resume_lines[max(0, line_index - 1):line_index + 2]
            context = " ".join(context_lines)
            if re.search(education_terms, context, re.IGNORECASE) and not re.search(employment_terms, context, re.IGNORECASE):
                continue
            if not re.search(employment_terms, context, re.IGNORECASE):
                continue
            for date_match in date_pattern.finditer(line):
                start, end = date_match.groups()
                end_year = date.today().year if not end or end.lower() in {"present", "current"} else int(end)
                start_year = int(start)
                if start_year <= end_year <= date.today().year and end_year - start_year <= 60:
                    employment_ranges.append((start_year, end_year, date_match.group(0)))

        if employment_ranges:
            employment_ranges.sort()
            current_ranges = [
                date_range for date_range in employment_ranges
                if re.search(r"present|current", date_range[2], re.IGNORECASE)
            ]
            if current_ranges:
                start_year, end_year, matched_range = max(current_ranges, key=lambda item: item[0])
                selected_ranges = [(start_year, end_year, matched_range)]
            else:
                selected_ranges = employment_ranges

            merged_ranges = []
            for start_year, end_year, matched_range in selected_ranges:
                if merged_ranges and start_year <= merged_ranges[-1][1]:
                    merged_ranges[-1] = (
                        merged_ranges[-1][0],
                        max(merged_ranges[-1][1], end_year),
                        merged_ranges[-1][2],
                    )
                else:
                    merged_ranges.append((start_year, end_year, matched_range))
            total_years = sum(end_year - start_year for start_year, end_year, _ in merged_ranges)
            experience_years = float(total_years)
            experience = f"{experience_years:g} Years (estimated from dates)"
            experience_period = f"{start_year} - {end_year}"
            experience_evidence = "Matched employment date ranges: " + ", ".join(
                matched_range for _, _, matched_range in merged_ranges
            )
            experience_source = "date range"

    experience_flags = []
    if experience_years > 40:
        experience_flags.append("Unusually high value; verify the resume dates")
    if experience_source == "date range":
        experience_flags.append("Estimated from an employment date range")
    if experience_source == "explicit":
        experience_flags.append("Explicit duration found in resume text")
    if experience_source == "none":
        experience_flags.append("No explicit experience duration or employment date range found")
    if experience_years <= 0:
        experience_points = 0.0
    elif experience_years <= 1:
        experience_points = 0.5
    elif experience_years <= 3:
        experience_points = 1.0
    elif experience_years <= 5:
        experience_points = 1.5
    elif experience_years <= 10:
        experience_points = 2.0
    elif experience_years <= 15:
        experience_points = 2.5
    else:
        experience_points = 3.0

    normalized_skills = {skill.lower() for skill in skills}
    role_matches = []
    for role, profile in ROLE_PROFILES.items():
        required = profile["required"]
        required_matches = required & normalized_skills
        match_score = round(len(required_matches) / len(required) * 100, 1)
        missing = sorted(required - normalized_skills)
        role_matches.append({
            "role": role,
            "match_score": match_score,
            "matched_skills": sorted(required_matches),
            "missing_skills": missing,
        })
    role_matches.sort(key=lambda match: match["match_score"], reverse=True)

    measurable_evidence = len(re.findall(
        r"\b(?:increased|improved|reduced|saved|generated|managed|led|built|developed|automated)\b",
        text,
    ))
    score_breakdown = {
        "Skills coverage": round(min(len(skills) / 10 * 4, 4), 1),
        "Experience evidence": experience_points,
        "Education evidence": 1.5 if education_status == "Found" else 0.0,
        "Contact completeness": 0.5 if email_matches and phone_matches else 0.25 if email_matches or phone_matches else 0.0,
        "Achievement evidence": round(min(measurable_evidence / 3, 1), 1),
    }
    score = round(sum(score_breakdown.values()), 1)
    score_level = "Excellent" if score >= 8 else "Good" if score >= 6 else "Needs Improvement"
    recommendation = "Hire" if score >= 8 else "Consider" if score >= 5 else "Needs Improvement"

    top_role_match = role_matches[0] if role_matches else None
    missing_role_skills = top_role_match["missing_skills"] if top_role_match else []
    resume_suggestions = []
    if not re.search(r"(?im)^\s*(professional summary|resume summary|profile|objective|career objective)\s*:?$", text):
        resume_suggestions.append("Add a Professional Summary section that states your target role, strongest skills, experience, and career focus.")
    if not re.search(r"(?im)^\s*(projects?|key projects?|academic projects?)\s*:?$", text):
        resume_suggestions.append("Add a Projects section with the problem, technologies used, your contribution, and one measurable result for each project.")
    if measurable_evidence == 0:
        resume_suggestions.append("Add quantified achievements using percentages, time saved, revenue, users, accuracy, performance, or delivery improvements.")
    if not re.search(r"(?im)^\s*(certifications?|licenses? and certifications?)\s*:?$", text):
        resume_suggestions.append("Add a Certifications section if you have any relevant technical, cloud, Agile, or domain certifications.")
    if missing_role_skills:
        resume_suggestions.append(
            f"For the suggested {top_role_match['role']} role, add evidence of these missing role skills only if you genuinely have them: "
            + ", ".join(missing_role_skills[:8]) + "."
        )
    if email_matches and phone_matches and linkedin_matches and github_matches:
        pass
    else:
        missing_contact = []
        if not email_matches:
            missing_contact.append("email")
        if not phone_matches:
            missing_contact.append("phone")
        if not linkedin_matches:
            missing_contact.append("LinkedIn")
        if not github_matches:
            missing_contact.append("GitHub")
        resume_suggestions.append("Complete your Contact Information with " + ", ".join(missing_contact) + ".")
    if education_status == "Not Found":
        resume_suggestions.append("Add a clearly labeled Education section with your degree, institution, and graduation year.")

    return {
        "name": candidate_name,
        "role": candidate_role,
        "skills": skills,
        "grouped_skills": grouped_skills,
        "skill_details": skill_details,
        "experience": experience,
        "experience_period": experience_period,
        "experience_years": experience_years,
        "experience_evidence": experience_evidence,
        "experience_flags": experience_flags,
        "role_matches": role_matches,
        "education": education_details,
        "education_status": education_status,
        "score": score,
        "score_level": score_level,
        "score_breakdown": score_breakdown,
        "measurable_evidence": measurable_evidence,
        "recommendation": recommendation,
        "email": email_matches[0] if email_matches else "Not Found",
        "phone": phone_matches[0] if phone_matches else "Not Found",
        "linkedin": linkedin_matches[0] if linkedin_matches else "Not Found",
        "github": github_matches[0] if github_matches else "Not Found",
        "suggested_role": top_role_match["role"] if top_role_match else "Not Found",
        "missing_role_skills": missing_role_skills,
        "resume_suggestions": resume_suggestions,
    }

def analysis_markdown(analysis):
    resume_suggestions = "\n".join(
        f"{index}. {suggestion}"
        for index, suggestion in enumerate(analysis["resume_suggestions"], start=1)
    )
    grouped_skill_details = "\n".join(
        f"- **{group}:** {', '.join(group_skills) if group_skills else 'None detected'}"
        for group, group_skills in analysis["grouped_skills"].items()
    )
    role_details = "\n".join(
        f"- **{match['role']} ({match['match_score']}%)**: {', '.join(match['matched_skills']) or 'No matching skills'}"
        f" | Missing: {', '.join(match['missing_skills']) or 'None'}"
        for match in analysis["role_matches"][:5]
    )
    return f"""# Resume Analysis Report

## Candidate
**Name:** {analysis['name']}

**Role:** {analysis['role']}

## Contact
- **Email:** {analysis['email']}
- **Phone:** {analysis['phone']}
- **LinkedIn:** {analysis['linkedin']}
- **GitHub:** {analysis['github']}

## Score
**{analysis['score']}/10 ({analysis['score_level']})**

## Experience
**{f"{analysis['experience_years']:g} Years ({analysis['experience_period']})" if analysis['experience_years'] else 'Not Found'}**

## Skills
**{len(analysis['skills'])} skills detected**

### Skill Categories
{grouped_skill_details}

## Education
**Status:** {analysis['education_status']}

**Details:** {', '.join(analysis['education']) if analysis['education'] else 'No degree detected'}

**Measurable achievement phrases:** {analysis['measurable_evidence']}

## Recommended Roles
{role_details}

## Things to Add
{resume_suggestions}

## Recommendation
{analysis['recommendation']}
"""

def ai_reply(question, analysis):
    query = question.lower().strip()
    if query in {"hi", "hello", "hey", "hii"}:
        return "Hi! Ask me for the analysis report, candidate name, skills, experience, education, specifications, score, recommendation, or contact details."
    if any(word in query for word in ["report", "analysis", "full summary"]):
        return analysis_markdown(analysis)
    if any(word in query for word in ["email", "contact", "phone", "mobile", "linkedin", "github"]):
        return (
            f"**Email:** {analysis['email']}\n\n"
            f"**Phone:** {analysis['phone']}\n\n"
            f"**LinkedIn:** {analysis['linkedin']}\n\n"
            f"**GitHub:** {analysis['github']}"
        )
    if "name" in query or "candidate" in query:
        return f"**Candidate name:** {analysis['name']}\n\n**Role:** {analysis['role']}"
    if "specification" in query or "specifications" in query:
        return (
            f"**Specifications:**\n\n"
            f"- Skills: {len(analysis['skills'])} detected\n"
            f"- Experience: {analysis['experience']}\n"
            f"- Education: {', '.join(analysis['education']) or 'Not Found'}\n"
            f"- Score: {analysis['score']}/10 ({analysis['score_level']})"
        )
    if any(word in query for word in ["role", "job", "position", "career", "fit"]):
        role_details = "\n".join(
            f"- **{match['role']} ({match['match_score']}%)**: missing {', '.join(match['missing_skills']) or 'no required skills'}"
            for match in analysis["role_matches"][:5]
        )
        return f"**Recommended roles:**\n\n{role_details}"
    if "skill" in query:
        categories = "\n".join(
            f"- **{group}:** {', '.join(group_skills) if group_skills else 'None detected'}"
            for group, group_skills in analysis["grouped_skills"].items()
        )
        return f"**Skills found ({len(analysis['skills'])})**\n\n{categories}"
    if any(word in query for word in ["experience", "exp"]):
        if analysis["experience_years"]:
            period = f" ({analysis['experience_period']})" if analysis["experience_period"] else ""
            return f"**Experience:** {analysis['experience_years']:g} Years{period}"
        return "**Experience:** Not Found"
    if "education" in query or "degree" in query:
        return f"**Education ({analysis['education_status']}):** {', '.join(analysis['education']) or 'No degree detected.'}"
    if "score" in query:
        return f"**Score:** {analysis['score']}/10 ({analysis['score_level']})"
    if any(word in query for word in ["recommend", "hire"]):
        return f"**Recommendation:** {analysis['recommendation']}"
    if "strength" in query:
        return f"**Strengths:** {len(analysis['skills'])} relevant skills detected, with {analysis['experience']} of experience and {analysis['education_status'].lower()} education evidence."
    if any(word in query for word in ["improve", "weakness", "better"]):
        return "**Improvements:** Add quantified achievements, role-specific keywords, and clearly labeled experience and education details."
    if any(word in query for word in ["add", "missing", "improve", "suggest", "things to"]):
        suggestions = "\n".join(
            f"{index}. {suggestion}"
            for index, suggestion in enumerate(analysis["resume_suggestions"], start=1)
        )
        return f"**Suggested additions for your resume:**\n\n{suggestions}"
    if "summary" in query:
        return f"**Summary:** {analysis['experience']} experience, {len(analysis['skills'])} skills detected, scored {analysis['score']}/10. Recommendation: {analysis['recommendation']}."
    return "Try asking for the analysis report, name, specifications, skills, experience, education, score, recommendation, or contact details."

st.subheader("Upload Resume")
resume = st.file_uploader("Choose a PDF or DOCX", type=["pdf", "docx"])

if "messages" not in st.session_state:
    st.session_state.messages = []

if resume:
    file_key = f"{ANALYSIS_VERSION}:{resume.name}:{len(resume.getvalue())}"
    if st.session_state.get("file_key") != file_key:
        try:
            st.session_state.resume_text = extract_text(resume)
        except ValueError as exc:
            st.error(str(exc))
            st.stop()

        if not st.session_state.resume_text:
            st.error(
                "No selectable text was found in this resume. If it is a scanned PDF or an image, "
                "upload an OCR-enabled PDF or a DOCX file instead."
            )
            st.stop()

        st.session_state.file_key = file_key
        st.session_state.analysis = get_resume_analysis(st.session_state.resume_text)
        st.session_state.messages = []

    analysis = st.session_state.analysis
    st.success(f"Uploaded: {resume.name}")
    st.subheader("Candidate Snapshot")
    st.markdown(
        f"<div class='snapshot-intro'><strong>Candidate</strong><br>{analysis['name']}<br><small>{analysis['role']}</small></div>",
        unsafe_allow_html=True,
    )
    top_roles = "".join(
        f"<span>{match['role']} · {match['match_score']}%</span>"
        for match in analysis["role_matches"][:3]
    )
    st.markdown(
        f"<div class='role-strip'><strong>Suggested roles</strong><br>{top_roles}</div>",
        unsafe_allow_html=True,
    )
    st.subheader("Best Role for Your Current Resume")
    best_role = analysis["role_matches"][0]
    st.metric("Recommended Role", best_role["role"])
    st.progress(min(best_role["match_score"] / 100, 1.0), text=f"Current role match: {best_role['match_score']}%")
    st.markdown(
        "**Matched skills:** " + (", ".join(best_role["matched_skills"]) or "None detected") + "<br>"
        "**Skills to consider adding:** " + (", ".join(best_role["missing_skills"][:10]) or "No required skills missing") ,
        unsafe_allow_html=True,
    )
    metric1, metric2, metric3 = st.columns(3)
    metric1.metric("Resume Score", f"{analysis['score']}/10")
    metric2.metric("Skills Found", len(analysis["skills"]))
    experience_display = (
        f"{analysis['experience_years']:g} Years ({analysis['experience_period']})"
        if analysis["experience_years"]
        else "Not Found"
    )
    metric3.metric("Experience", experience_display)

    st.subheader("Customize Suggestions for a Role")
    role_options = ["Choose a role"] + list(ROLE_PROFILES.keys())
    selected_role = st.selectbox(
        "Choose your target role",
        options=role_options,
        index=role_options.index(analysis["suggested_role"]) if analysis["suggested_role"] in role_options else 0,
        key=f"selected_target_role:{analysis['name']}:{analysis['score']}",
    )
    custom_role = st.text_input(
        "Or enter another role",
        value="",
        placeholder="For example: AI or AI Engineer",
        key=f"custom_target_role:{analysis['name']}:{analysis['score']}",
    )
    target_role = custom_role.strip() or ("" if selected_role == "Choose a role" else selected_role)
    normalized_target_role = " ".join(target_role.strip().casefold().replace("/", " ").replace("-", " ").split())
    if normalized_target_role.endswith(" role"):
        normalized_target_role = normalized_target_role[:-5].strip()
    role_aliases = {
        "ai": "AI/ML Developer",
        "ai ml": "AI/ML Developer",
        "ai developer": "AI/ML Developer",
        "ml": "AI/ML Developer",
        "ml developer": "AI/ML Developer",
        "machine learning": "AI/ML Developer",
        "machine learning developer": "AI/ML Developer",
    }
    resolved_target_role = role_aliases.get(normalized_target_role, normalized_target_role)
    ai_terms = {
        "ai", "ai ml", "ai ml developer", "ai developer", "ai engineer", "artificial intelligence",
        "ml", "ml developer", "ml engineer", "machine learning", "machine learning developer",
        "machine learning engineer",
    }
    if normalized_target_role in ai_terms:
        target_profile = ("AI/ML Developer", ROLE_PROFILES["AI/ML Developer"])
    else:
        target_profile = next(
            (
                (role, profile) for role, profile in ROLE_PROFILES.items()
                if normalized_target_role == role.casefold().replace("/", " ").replace("-", " ")
                or normalized_target_role == role.casefold()
            ),
            None,
        )
        if resolved_target_role in ROLE_PROFILES:
            target_profile = (resolved_target_role, ROLE_PROFILES[resolved_target_role])
    target_matched = []
    target_missing = []
    target_match_percent = 0.0
    if target_role.strip():
        if target_profile is None:
            st.warning("This role is not in the built-in role list yet. Available roles: " + ", ".join(ROLE_PROFILES.keys()))
        else:
            target_role_name, profile = target_profile
            resume_skills = {skill.strip().casefold() for skill in analysis["skills"]}
            target_matched = sorted(profile["required"] & resume_skills, key=str.lower)
            target_missing = sorted(profile["required"] - resume_skills, key=str.lower)
            match_percent = round(len(target_matched) / len(profile["required"]) * 100, 1)
            target_match_percent = match_percent
            st.metric(f"Match for {target_role_name}", f"{match_percent}%")
            st.progress(min(match_percent / 100, 1.0), text=f"{len(target_matched)} of {len(profile['required'])} required skills found")
            st.markdown("**Skills already in your resume:** " + (", ".join(target_matched) or "None detected"))
            st.markdown("**Add or demonstrate these skills for this role:** " + (", ".join(target_missing) or "None — all required skills are detected"))
            if target_missing:
                st.caption("Only add skills you genuinely possess, and include project or work experience evidence for each skill.")

    st.subheader("Things to Add to Your Resume")
    role_focus = target_role.strip() if target_role.strip() else analysis["suggested_role"]
    st.caption(f"Based on gaps found for the selected role: {role_focus}")
    if target_profile and target_missing:
        st.markdown("### Missing skills for this role")
        for skill in target_missing:
            skill_name = SKILL_DISPLAY_NAMES.get(skill, skill.title())
            st.markdown(f"- **{skill_name}** — add genuine experience or a project/example that demonstrates this skill.")
    elif target_profile:
        st.success("Your resume contains all required skills for this role. Improve the evidence and quantified achievements next.")
    st.markdown("### Other resume improvements")
    for suggestion in analysis["resume_suggestions"]:
        st.markdown(f"- {suggestion}")

    st.subheader("Resume Assistant")
    for message in st.session_state.messages:
        with st.chat_message(message["role"]):
            st.markdown(message["content"])
    user_input = st.chat_input("Ask about this resume...")
    if user_input:
        st.session_state.messages.append({"role": "user", "content": user_input})
        st.session_state.messages.append({"role": "assistant", "content": ai_reply(user_input, analysis)})
        st.rerun()
else:
    st.info("Upload a resume above to begin analysis.")


# & ".\.venv\Scripts\streamlit.exe" run ".\Validator_code.py"