"""
MIU — AI Based Resume Screening & Recruitment Portal
Mohi-ud-Din Islamic University, Nerian Sharif, AJK  |  Final Year Project (CS & IT)

Portals : Admin · HR · Company · Job Seeker (role "recruiter")
Storage : SQLite (miu_recruitment.db)
Scoring : dataset-based resume scorer (datasets/*.csv) + optional Claude AI rating
HR view : top 5 % of applicants are shown automatically (no manual upload needed)
"""
import streamlit as st
import streamlit.components.v1 as components
import sqlite3
import pickle
import docx
import PyPDF2
import re
import csv
import html
import urllib.parse
import pandas as pd
import hashlib
import json
import time
import os
import math                      # math.ceil() is used for the "top 5 %" calculation
from collections import Counter
from datetime import datetime
from functools import lru_cache
import anthropic

# ══════════════════════════════════════════════════════════════════════════════
# UNIVERSITY LOGO
# ══════════════════════════════════════════════════════════════════════════════
# The logo is embedded below as a base64 string. If you ever want to replace it,
# either paste a new base64 string here, or leave it empty and put a file named
# "miu_logo.jpg" next to this script.
MIU_LOGO_B64 = "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"

# Fallback: if the string above is empty, look for a "miu_logo.jpg" next to this file.
if not MIU_LOGO_B64:
    import base64
    _logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "miu_logo.jpg")
    if os.path.exists(_logo_path):
        with open(_logo_path, "rb") as _f:
            MIU_LOGO_B64 = base64.b64encode(_f.read()).decode()

# ══════════════════════════════════════════════════════════════════════════════
# DATABASE  —  SQLite  (miu_recruitment.db)
# ══════════════════════════════════════════════════════════════════════════════
DB_PATH = "miu_recruitment.db"


def _hash(p):
    """SHA-256 hash of a password (passwords are never stored in plain text)."""
    return hashlib.sha256(p.encode()).hexdigest()


def get_conn():
    """Return a SQLite connection whose rows behave like dictionaries."""
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def init_db():
    """Create the tables and seed default data. Safe to run on every start."""
    conn = get_conn()
    c = conn.cursor()

    c.executescript("""
    CREATE TABLE IF NOT EXISTS users (
        id           INTEGER PRIMARY KEY AUTOINCREMENT,
        username     TEXT UNIQUE NOT NULL,
        password     TEXT NOT NULL,
        role         TEXT NOT NULL,
        name         TEXT NOT NULL,
        email        TEXT NOT NULL,
        first_name   TEXT DEFAULT '',
        last_name    TEXT DEFAULT '',
        age          INTEGER,
        gender       TEXT DEFAULT '',
        phone        TEXT DEFAULT '',
        company_info TEXT DEFAULT '{}',
        created      TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS jobs (
        id           INTEGER PRIMARY KEY AUTOINCREMENT,
        job_id       TEXT UNIQUE NOT NULL,
        title        TEXT NOT NULL,
        company      TEXT NOT NULL,
        posted_by    TEXT NOT NULL,
        department   TEXT,
        location     TEXT,
        type         TEXT,
        salary       TEXT,
        description  TEXT,
        requirements TEXT,
        deadline     TEXT,
        posted       TEXT,
        status       TEXT DEFAULT 'open'
    );

    CREATE TABLE IF NOT EXISTS applications (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        app_id      TEXT UNIQUE NOT NULL,
        job_id      TEXT NOT NULL,
        recruiter   TEXT NOT NULL,
        name        TEXT,
        email       TEXT,
        phone       TEXT,
        experience  TEXT,
        cover       TEXT,
        category    TEXT,
        resume_name TEXT,
        score       TEXT DEFAULT 'Pending',
        status      TEXT DEFAULT 'Under Review',
        date        TEXT,
        FOREIGN KEY (job_id) REFERENCES jobs(job_id)
    );

    CREATE TABLE IF NOT EXISTS ai_ratings (
        id             INTEGER PRIMARY KEY AUTOINCREMENT,
        session_key    TEXT NOT NULL,
        filename       TEXT NOT NULL,
        score          INTEGER,
        rating         TEXT,
        strengths      TEXT,
        weaknesses     TEXT,
        recommendation TEXT,
        years_exp      TEXT,
        top_skills     TEXT,
        created_at     TEXT,
        UNIQUE(session_key, filename)
    );

    -- Which companies an HR user may work with (assigned by the Admin).
    -- If an HR user has no rows here, that HR may choose ANY company.
    CREATE TABLE IF NOT EXISTS hr_company_access (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        hr_username TEXT NOT NULL,
        company     TEXT NOT NULL,
        UNIQUE(hr_username, company)
    );
    """)

    # Migration: add the newer user columns if the table was created by an older version
    existing = {r[1] for r in c.execute("PRAGMA table_info(users)").fetchall()}
    for col, ddl in [("first_name", "TEXT DEFAULT ''"), ("last_name", "TEXT DEFAULT ''"),
                     ("age", "INTEGER"), ("gender", "TEXT DEFAULT ''"),
                     ("phone", "TEXT DEFAULT ''")]:
        if col not in existing:
            c.execute(f"ALTER TABLE users ADD COLUMN {col} {ddl}")

    # Migration: store the full dataset-based score breakdown (JSON) for each application
    app_cols = {r[1] for r in c.execute("PRAGMA table_info(applications)").fetchall()}
    if "score_detail" not in app_cols:
        c.execute("ALTER TABLE applications ADD COLUMN score_detail TEXT DEFAULT ''")

    # Default accounts. NOTE: the HR display name is now simply the username.
    defaults = [
        ("admin", _hash("admin123"), "admin", "System Administrator",
         "admin@miu.edu.pk", "{}", "2026-01-01"),
        ("hr_manager", _hash("hr2026"), "hr", "hr_manager",
         "hr@miu.edu.pk", "{}", "2026-01-01"),
        ("techcorp", _hash("tech2026"), "company", "TechCorp Pakistan",
         "hr@techcorp.com.pk",
         json.dumps({"name": "TechCorp Pakistan", "industry": "Information Technology",
                     "size": "500-1000", "website": "https://techcorp.com.pk",
                     "about": "Leading IT company providing software solutions."}),
         "2026-01-01"),
        ("recruiter1", _hash("recruit123"), "recruiter", "Ahmed Raza",
         "ahmed@email.com", "{}", "2026-01-01"),
    ]
    for row in defaults:
        c.execute("""INSERT OR IGNORE INTO users
                     (username,password,role,name,email,company_info,created)
                     VALUES (?,?,?,?,?,?,?)""", row)

    # Migration: older databases still contain the old HR display name -> replace it
    # with the HR username so "Dr. Muhammad Ali" no longer appears anywhere.
    c.execute("UPDATE users SET name=username WHERE role='hr' AND name='Dr. Muhammad Ali'")

    # Two demo jobs so the job board is not empty on first run
    demo_jobs = [
        ("JOB001", "Senior Python Developer", "TechCorp Pakistan", "techcorp",
         "Engineering", "Islamabad / Remote", "Full-time", "PKR 150,000 - 250,000",
         "We are looking for a Senior Python Developer with 4+ years experience.",
         "4+ years Python, Django/FastAPI, PostgreSQL, Docker, AWS/GCP",
         "2026-06-30", "2026-05-01", "open"),
        ("JOB002", "Data Scientist", "TechCorp Pakistan", "techcorp",
         "Analytics", "Lahore", "Full-time", "PKR 120,000 - 200,000",
         "Join our data team to build ML models and analytics pipelines.",
         "3+ years ML, Python, TensorFlow/PyTorch, SQL, Power BI",
         "2026-06-15", "2026-05-05", "open"),
    ]
    for j in demo_jobs:
        c.execute("""INSERT OR IGNORE INTO jobs
                     (job_id,title,company,posted_by,department,location,type,
                      salary,description,requirements,deadline,posted,status)
                     VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""", j)

    conn.commit()
    conn.close()


# ── User helpers ─────────────────────────────────────────────────────────────
def db_authenticate(username, password):
    """Return the user row (as dict) if the credentials are valid, otherwise None.
    The login name can be the username OR the e-mail (companies sign in with their e-mail)."""
    conn = get_conn()
    ident = (username or "").strip().lower()
    row = conn.execute("SELECT * FROM users WHERE (username=? OR lower(email)=?) AND password=?",
                       (ident, ident, _hash(password))).fetchone()
    conn.close()
    return dict(row) if row else None


def db_get_all_users():
    conn = get_conn()
    rows = conn.execute("SELECT * FROM users ORDER BY created").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def db_get_user(username):
    conn = get_conn()
    row = conn.execute("SELECT * FROM users WHERE username=?", (username,)).fetchone()
    conn.close()
    return dict(row) if row else None


def db_create_user(username, password, role, name, email, company_info=None,
                   first_name="", last_name="", age=None, gender="", phone=""):
    """Insert a new user. Returns False if the username already exists."""
    conn = get_conn()
    try:
        conn.execute("""INSERT INTO users
                        (username,password,role,name,email,phone,first_name,last_name,age,gender,
                         company_info,created)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
                     (username, _hash(password), role, name, email, phone,
                      first_name, last_name, age, gender,
                      json.dumps(company_info or {}), datetime.now().strftime("%Y-%m-%d")))
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False
    finally:
        conn.close()


def db_delete_user(username):
    conn = get_conn()
    conn.execute("DELETE FROM users WHERE username=?", (username,))
    conn.commit()
    conn.close()


def db_update_password(username, new_password):
    conn = get_conn()
    conn.execute("UPDATE users SET password=? WHERE username=?",
                 (_hash(new_password), username))
    conn.commit()
    conn.close()


def db_update_company_info(username, info: dict):
    conn = get_conn()
    conn.execute("UPDATE users SET company_info=? WHERE username=?",
                 (json.dumps(info), username))
    conn.commit()
    conn.close()


def db_update_phone(username, phone):
    conn = get_conn()
    conn.execute("UPDATE users SET phone=? WHERE username=?", (phone, username))
    conn.commit()
    conn.close()


def db_company_name_exists(name):
    """True if a company with this name is already registered (case-insensitive)."""
    return name.strip().lower() in {c.lower() for c in db_get_companies()}


def db_email_exists(email):
    """True if an account with this e-mail already exists (case-insensitive)."""
    conn = get_conn()
    row = conn.execute("SELECT 1 FROM users WHERE lower(email)=lower(?)", (email,)).fetchone()
    conn.close()
    return row is not None


# ── Job helpers ───────────────────────────────────────────────────────────────
def db_get_all_jobs():
    conn = get_conn()
    rows = conn.execute("SELECT * FROM jobs ORDER BY posted DESC").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def db_get_jobs_by_user(username):
    conn = get_conn()
    rows = conn.execute("SELECT * FROM jobs WHERE posted_by=? ORDER BY posted DESC",
                        (username,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def db_get_open_jobs():
    conn = get_conn()
    rows = conn.execute("SELECT * FROM jobs WHERE status='open' ORDER BY posted DESC").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def db_create_job(job_id, title, company, posted_by, department, location,
                  jtype, salary, description, requirements, deadline, posted):
    conn = get_conn()
    conn.execute("""INSERT INTO jobs
                    (job_id,title,company,posted_by,department,location,type,
                     salary,description,requirements,deadline,posted,status)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,'open')""",
                 (job_id, title, company, posted_by, department, location, jtype,
                  salary, description, requirements, deadline, posted))
    conn.commit()
    conn.close()


def db_toggle_job_status(job_id, new_status):
    conn = get_conn()
    conn.execute("UPDATE jobs SET status=? WHERE job_id=?", (new_status, job_id))
    conn.commit()
    conn.close()


def db_count_apps_for_job(job_id):
    conn = get_conn()
    n = conn.execute("SELECT COUNT(*) FROM applications WHERE job_id=?",
                     (job_id,)).fetchone()[0]
    conn.close()
    return n


# ── Application helpers ───────────────────────────────────────────────────────
def db_get_all_applications():
    """All applications joined with the job title and company."""
    conn = get_conn()
    rows = conn.execute("""
        SELECT a.*, j.title as job_title, j.company
        FROM applications a
        JOIN jobs j ON a.job_id = j.job_id
        ORDER BY a.date DESC
    """).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def db_get_apps_for_job(job_id):
    conn = get_conn()
    rows = conn.execute("SELECT * FROM applications WHERE job_id=? ORDER BY date DESC",
                        (job_id,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def db_get_apps_by_recruiter(username):
    """Applications submitted by one job seeker."""
    conn = get_conn()
    rows = conn.execute("""
        SELECT a.*, j.title as job_title, j.company
        FROM applications a
        JOIN jobs j ON a.job_id = j.job_id
        WHERE a.recruiter=?
        ORDER BY a.date DESC
    """, (username,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def db_create_application(app_id, job_id, recruiter, name, email, phone,
                          experience, cover, category, resume_name,
                          score="Pending", score_detail=""):
    conn = get_conn()
    conn.execute("""INSERT INTO applications
                    (app_id,job_id,recruiter,name,email,phone,experience,cover,
                     category,resume_name,score,score_detail,status,date)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,'Under Review',?)""",
                 (app_id, job_id, recruiter, name, email, phone, experience, cover,
                  category, resume_name, score, score_detail,
                  datetime.now().strftime("%Y-%m-%d %H:%M")))
    conn.commit()
    conn.close()


def db_update_app_status(app_id, status):
    conn = get_conn()
    conn.execute("UPDATE applications SET status=? WHERE app_id=?", (status, app_id))
    conn.commit()
    conn.close()


# ── AI Rating helpers ─────────────────────────────────────────────────────────
def db_save_rating(session_key, filename, rating: dict):
    """Save (or overwrite) one AI rating so it survives page reloads."""
    conn = get_conn()
    conn.execute("""INSERT OR REPLACE INTO ai_ratings
                    (session_key,filename,score,rating,strengths,weaknesses,
                     recommendation,years_exp,top_skills,created_at)
                    VALUES (?,?,?,?,?,?,?,?,?,?)""",
                 (session_key, filename,
                  rating.get("score"), rating.get("rating"),
                  json.dumps(rating.get("strengths", [])),
                  json.dumps(rating.get("weaknesses", [])),
                  rating.get("recommendation", ""),
                  rating.get("years_exp", ""),
                  json.dumps(rating.get("top_skills", [])),
                  datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    conn.close()


def _decode_rating_row(row):
    """Convert the JSON text columns of a rating row back into Python lists."""
    r = dict(row)
    r["strengths"] = json.loads(r.get("strengths") or "[]")
    r["weaknesses"] = json.loads(r.get("weaknesses") or "[]")
    r["top_skills"] = json.loads(r.get("top_skills") or "[]")
    return r


def db_get_all_ratings_for_session(session_key):
    """Return {filename: rating_dict} for one user."""
    conn = get_conn()
    rows = conn.execute("SELECT * FROM ai_ratings WHERE session_key=?",
                        (session_key,)).fetchall()
    conn.close()
    return {r["filename"]: _decode_rating_row(r) for r in rows}

# ── Company / HR access helpers ───────────────────────────────────────────────
def db_get_companies():
    """All company names: companies that posted jobs + registered company accounts."""
    conn = get_conn()
    names = {r[0] for r in conn.execute("SELECT DISTINCT company FROM jobs WHERE company<>''")}
    for r in conn.execute("SELECT company_info FROM users WHERE role='company'"):
        try:
            n = (json.loads(r[0] or "{}").get("name") or "").strip()
        except Exception:
            n = ""
        if n:
            names.add(n)
    conn.close()
    return sorted(names)


def db_get_assigned_companies(hr_username):
    """Companies explicitly assigned to this HR user (may be empty)."""
    conn = get_conn()
    rows = conn.execute("SELECT company FROM hr_company_access WHERE hr_username=? ORDER BY company",
                        (hr_username,)).fetchall()
    conn.close()
    return [r[0] for r in rows]


def db_get_hr_companies(hr_username):
    """Companies shown in the HR company picker: the assigned ones, or all if none assigned."""
    return db_get_assigned_companies(hr_username) or db_get_companies()


def db_set_hr_companies(hr_username, companies):
    """Replace the company list of an HR user."""
    conn = get_conn()
    conn.execute("DELETE FROM hr_company_access WHERE hr_username=?", (hr_username,))
    for c in companies:
        conn.execute("INSERT OR IGNORE INTO hr_company_access (hr_username,company) VALUES (?,?)",
                     (hr_username, c))
    conn.commit()
    conn.close()


def db_get_apps_for_company(company):
    """Only the applications that were sent to jobs posted by this company."""
    conn = get_conn()
    rows = conn.execute("""
        SELECT a.*, j.title as job_title, j.company
        FROM applications a
        JOIN jobs j ON a.job_id = j.job_id
        WHERE j.company=?
        ORDER BY a.date DESC
    """, (company,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def db_get_jobs_for_company(company):
    conn = get_conn()
    rows = conn.execute("SELECT * FROM jobs WHERE company=? ORDER BY posted DESC", (company,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


# ══════════════════════════════════════════════════════════════════════════════
# SCORING DATASETS (CSV generator) — creates datasets/*.csv on first run
# ══════════════════════════════════════════════════════════════════════════════
BASE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE, "datasets")

# skill, category, domain, aliases (separated by |)
SKILLS = [
    # Programming
    ("Python", "Programming", "it", ""), ("Java", "Programming", "it", ""),
    ("JavaScript", "Programming", "it", "ecmascript|es6"), ("TypeScript", "Programming", "it", ""),
    ("C++", "Programming", "it", "cpp"), ("C#", "Programming", "it", "csharp|c sharp"),
    ("C Programming", "Programming", "it", "c language|c programming"), ("PHP", "Programming", "it", ""),
    ("Ruby", "Programming", "it", ""), ("Go", "Programming", "it", "golang|go language"),
    ("Kotlin", "Programming", "it", ""), ("Swift", "Programming", "it", ""),
    ("R Programming", "Programming", "data", "rstudio|r language"),
    ("MATLAB", "Programming", "engineering", ""), ("SQL", "Programming", "it", "t-sql|pl/sql"),
    ("HTML", "Web", "it", "html5"), ("CSS", "Web", "it", "css3"), ("Bash", "Programming", "it", "shell scripting"),
    # Web
    ("React", "Web", "it", "react.js|reactjs"), ("Angular", "Web", "it", "angularjs"),
    ("Vue.js", "Web", "it", "vue|vuejs"), ("Node.js", "Web", "it", "nodejs|node js"),
    ("Django", "Web", "it", ""), ("Flask", "Web", "it", ""), ("FastAPI", "Web", "it", ""),
    ("Spring Boot", "Web", "it", "spring framework"), ("Laravel", "Web", "it", ""),
    ("ASP.NET", "Web", "it", ".net|dotnet|asp.net core"),
    ("REST API", "Web", "it", "restful|rest apis|web services"), ("GraphQL", "Web", "it", ""),
    ("Bootstrap", "Web", "it", ""), ("Tailwind CSS", "Web", "it", "tailwind"),
    ("jQuery", "Web", "it", ""), ("WordPress", "Web", "it", ""),
    # Mobile
    ("Flutter", "Mobile", "it", "dart"), ("React Native", "Mobile", "it", ""),
    ("Android", "Mobile", "it", "android studio"), ("iOS", "Mobile", "it", "ios development"),
    # Data / AI
    ("Machine Learning", "Data/AI", "data", "ml"), ("Deep Learning", "Data/AI", "data", "neural networks"),
    ("NLP", "Data/AI", "data", "natural language processing"),
    ("Computer Vision", "Data/AI", "data", "image processing"),
    ("TensorFlow", "Data/AI", "data", ""), ("PyTorch", "Data/AI", "data", ""), ("Keras", "Data/AI", "data", ""),
    ("Scikit-learn", "Data/AI", "data", "sklearn|scikit learn"), ("Pandas", "Data/AI", "data", ""),
    ("NumPy", "Data/AI", "data", ""), ("Power BI", "Data/AI", "data", "powerbi"),
    ("Tableau", "Data/AI", "data", ""), ("Excel", "Data/AI", "data", "ms excel|microsoft excel|advanced excel"),
    ("Data Analysis", "Data/AI", "data", "data analytics"),
    ("Data Visualization", "Data/AI", "data", "data visualisation"),
    ("Statistics", "Data/AI", "data", "statistical analysis"), ("Big Data", "Data/AI", "data", ""),
    ("Apache Spark", "Data/AI", "data", "spark|pyspark"), ("Hadoop", "Data/AI", "data", ""),
    ("ETL", "Data/AI", "data", "data pipelines"), ("Data Mining", "Data/AI", "data", ""),
    ("OpenCV", "Data/AI", "data", ""),
    ("Generative AI", "Data/AI", "data", "llm|large language models|prompt engineering"),
    # Databases
    ("MySQL", "Database", "it", ""), ("PostgreSQL", "Database", "it", "postgres"),
    ("MongoDB", "Database", "it", "nosql"), ("Oracle", "Database", "it", "oracle database"),
    ("SQL Server", "Database", "it", "mssql|microsoft sql server"), ("SQLite", "Database", "it", ""),
    ("Redis", "Database", "it", ""), ("Firebase", "Database", "it", ""),
    # DevOps / Cloud
    ("AWS", "Cloud/DevOps", "it", "amazon web services"), ("Azure", "Cloud/DevOps", "it", "microsoft azure"),
    ("Google Cloud", "Cloud/DevOps", "it", "gcp"), ("Docker", "Cloud/DevOps", "it", "containers"),
    ("Kubernetes", "Cloud/DevOps", "it", "k8s"), ("Jenkins", "Cloud/DevOps", "it", ""),
    ("CI/CD", "Cloud/DevOps", "it", "ci cd|continuous integration"), ("Git", "Cloud/DevOps", "it", ""),
    ("GitHub", "Cloud/DevOps", "it", ""), ("GitLab", "Cloud/DevOps", "it", ""),
    ("Linux", "Cloud/DevOps", "it", "ubuntu|unix"), ("Terraform", "Cloud/DevOps", "it", ""),
    ("Ansible", "Cloud/DevOps", "it", ""), ("Nginx", "Cloud/DevOps", "it", ""),
    # Security / Network
    ("Cyber Security", "Security", "it", "cybersecurity|information security|infosec"),
    ("Penetration Testing", "Security", "it", "pentesting|pen testing"),
    ("Ethical Hacking", "Security", "it", ""),
    ("Networking", "Network", "it", "computer networks|tcp/ip"), ("Cisco", "Network", "it", ""),
    ("Firewalls", "Security", "it", "firewall"),
    # Software engineering practice
    ("Agile", "Practice", "it", ""), ("Scrum", "Practice", "it", ""), ("JIRA", "Practice", "it", ""),
    ("Microservices", "Practice", "it", ""),
    ("OOP", "Practice", "it", "object oriented programming|object-oriented"),
    ("Data Structures", "Practice", "it", "dsa"), ("Algorithms", "Practice", "it", ""),
    ("Unit Testing", "Testing", "it", "junit|pytest"), ("Selenium", "Testing", "it", "test automation"),
    ("Software Testing", "Testing", "it", "manual testing|qa testing|quality assurance"),
    ("System Design", "Practice", "it", ""), ("UML", "Practice", "it", ""),
    # Design
    ("Figma", "Design", "design", ""), ("Adobe Photoshop", "Design", "design", "photoshop"),
    ("Adobe Illustrator", "Design", "design", "illustrator"),
    ("UI/UX", "Design", "design", "ui ux|ux design|user experience|ui design"),
    ("Adobe XD", "Design", "design", ""), ("Canva", "Design", "design", ""),
    ("Video Editing", "Design", "design", "premiere pro|after effects"),
    ("Graphic Design", "Design", "design", "graphic designing"),
    # Soft skills / management
    ("Communication", "Soft Skill", "business", "communication skills"),
    ("Leadership", "Soft Skill", "business", ""), ("Teamwork", "Soft Skill", "business", "team work|team player"),
    ("Problem Solving", "Soft Skill", "business", "problem-solving"),
    ("Time Management", "Soft Skill", "business", ""),
    ("Presentation Skills", "Soft Skill", "business", "presentations"),
    ("Negotiation", "Soft Skill", "business", ""), ("Critical Thinking", "Soft Skill", "business", ""),
    ("Project Management", "Management", "business", ""),
    ("MS Office", "Office", "business", "microsoft office"),
    # Marketing
    ("SEO", "Marketing", "marketing", "search engine optimization"),
    ("Digital Marketing", "Marketing", "marketing", ""),
    ("Social Media Marketing", "Marketing", "marketing", "social media"),
    ("Content Writing", "Marketing", "marketing", "content creation"),
    ("Google Analytics", "Marketing", "marketing", ""), ("Email Marketing", "Marketing", "marketing", ""),
    ("Copywriting", "Marketing", "marketing", ""), ("Google Ads", "Marketing", "marketing", "ppc|adwords"),
    # Finance
    ("Accounting", "Finance", "finance", "bookkeeping"), ("QuickBooks", "Finance", "finance", ""),
    ("Financial Analysis", "Finance", "finance", "financial modeling"),
    ("Budgeting", "Finance", "finance", "forecasting"), ("Taxation", "Finance", "finance", "tax"),
    ("Auditing", "Finance", "finance", "audit"), ("Tally", "Finance", "finance", ""),
    ("SAP", "Finance", "finance", ""), ("Financial Reporting", "Finance", "finance", "ifrs"),
    # HR
    ("Recruitment", "HR", "hr", "talent acquisition|recruiting"), ("Payroll", "HR", "hr", ""),
    ("Employee Relations", "HR", "hr", ""), ("Onboarding", "HR", "hr", ""),
    ("Performance Management", "HR", "hr", ""), ("HRIS", "HR", "hr", "hr management"),
    # Engineering
    ("AutoCAD", "Engineering", "engineering", ""), ("SolidWorks", "Engineering", "engineering", ""),
    ("PLC", "Engineering", "engineering", "plc programming"),
    ("Embedded Systems", "Engineering", "engineering", "embedded c|microcontrollers"),
    ("Arduino", "Engineering", "engineering", ""), ("Circuit Design", "Engineering", "engineering", "pcb design"),
    ("Revit", "Engineering", "engineering", ""), ("Raspberry Pi", "Engineering", "engineering", ""),
    # Education
    ("Teaching", "Education", "education", "lecturing|instruction"),
    ("Curriculum Development", "Education", "education", "lesson planning"),
    ("Classroom Management", "Education", "education", ""),
    ("Research", "Education", "education", "research methodology"),
    ("Academic Writing", "Education", "education", "research papers"),
    # Sales
    ("Customer Service", "Sales", "sales", "customer support"),
    ("Sales", "Sales", "sales", "business development"), ("CRM", "Sales", "sales", ""),
    ("Salesforce", "Sales", "sales", ""), ("Lead Generation", "Sales", "sales", ""),
]

# role, keywords (matched against the job title / category), domain, default skills
ROLE_SKILLS = [
    ("Python Developer", "python developer|python|django|flask|fastapi", "it", "Python|Django|Flask|REST API|SQL|Git|OOP|PostgreSQL|Docker"),
    ("Java Developer", "java developer|java|spring", "it", "Java|Spring Boot|SQL|REST API|OOP|Git|Microservices|MySQL"),
    ("Web Developer", "web developer|web design|frontend|front-end|front end|full stack|fullstack|mern", "it", "HTML|CSS|JavaScript|React|Node.js|REST API|Git|Bootstrap"),
    ("Data Scientist", "data scientist|data science|machine learning|ml engineer|ai engineer", "data", "Python|Machine Learning|Pandas|NumPy|Scikit-learn|SQL|Statistics|Data Visualization|TensorFlow"),
    ("Data Analyst", "data analyst|business intelligence|bi analyst", "data", "SQL|Excel|Power BI|Tableau|Data Analysis|Python|Statistics|Data Visualization"),
    ("Mobile Developer", "mobile|android|flutter|ios", "it", "Flutter|Android|Kotlin|Firebase|REST API|Git|UI/UX"),
    ("DevOps Engineer", "devops|cloud engineer|sre|system administrator", "it", "Linux|Docker|Kubernetes|AWS|CI/CD|Git|Terraform|Jenkins"),
    ("Network/Security Engineer", "network|security|cyber", "it", "Networking|Cisco|Cyber Security|Linux|Firewalls|Penetration Testing|Ethical Hacking"),
    ("QA Engineer", "qa|tester|quality assurance|sqa|testing", "it", "Software Testing|Selenium|JIRA|Agile|SQL|Unit Testing"),
    ("UI/UX Designer", "ui/ux|ui ux|graphic|designer", "design", "Figma|Adobe Photoshop|Adobe Illustrator|UI/UX|Adobe XD|Graphic Design"),
    ("Digital Marketer", "marketing|seo|social media|content", "marketing", "SEO|Digital Marketing|Social Media Marketing|Google Analytics|Content Writing|Email Marketing|Copywriting"),
    ("Accountant", "accountant|accounts|finance|audit", "finance", "Accounting|QuickBooks|Financial Reporting|Taxation|Auditing|Excel|Budgeting"),
    ("HR Officer", "hr|human resource|recruiter", "hr", "Recruitment|Payroll|Employee Relations|Onboarding|Performance Management|Communication"),
    ("Teacher / Lecturer", "teacher|lecturer|professor|instructor", "education", "Teaching|Curriculum Development|Classroom Management|Communication|Research"),
    ("Sales Executive", "sales|business development|customer", "sales", "Sales|CRM|Customer Service|Negotiation|Communication|Lead Generation"),
    ("Electrical/Embedded Engineer", "electrical|electronics|embedded", "engineering", "Embedded Systems|Arduino|Circuit Design|PLC|AutoCAD|MATLAB"),
    ("Project Manager", "project manager|product manager|scrum master", "business", "Project Management|Agile|Scrum|JIRA|Leadership|Communication"),
    ("Database Administrator", "dba|database", "it", "SQL|MySQL|PostgreSQL|Oracle|MongoDB|SQL Server"),
]

# keyword, level, level_name  (higher level = higher qualification)
EDU_LEVELS = [
    ("phd", 6, "PhD"), ("ph.d", 6, "PhD"), ("doctorate", 6, "PhD"), ("doctoral", 6, "PhD"),
    ("master of", 5, "Master"), ("master's", 5, "Master"), ("masters", 5, "Master"),
    ("master degree", 5, "Master"), ("msc", 5, "Master"), ("m.sc", 5, "Master"), ("mba", 5, "Master"),
    ("mphil", 5, "Master"), ("m.phil", 5, "Master"), ("mcs", 5, "Master"), ("ms in", 5, "Master"),
    ("ms cs", 5, "Master"), ("m.s", 5, "Master"),
    ("bachelor", 4, "Bachelor"), ("bsc", 4, "Bachelor"), ("b.sc", 4, "Bachelor"), ("bscs", 4, "Bachelor"),
    ("bsit", 4, "Bachelor"), ("bsse", 4, "Bachelor"), ("bs in", 4, "Bachelor"), ("bs cs", 4, "Bachelor"),
    ("bs computer", 4, "Bachelor"), ("bs", 4, "Bachelor"), ("bba", 4, "Bachelor"), ("b.e", 4, "Bachelor"),
    ("b.tech", 4, "Bachelor"), ("btech", 4, "Bachelor"), ("bca", 4, "Bachelor"),
    ("diploma", 3, "Diploma"), ("associate degree", 3, "Diploma"), ("adp", 3, "Diploma"), ("dae", 3, "Diploma"),
    ("intermediate", 2, "Intermediate"), ("fsc", 2, "Intermediate"), ("f.sc", 2, "Intermediate"),
    ("hssc", 2, "Intermediate"), ("higher secondary", 2, "Intermediate"), ("a-level", 2, "Intermediate"),
    ("a level", 2, "Intermediate"),
    ("matric", 1, "Matric"), ("matriculation", 1, "Matric"), ("ssc", 1, "Matric"),
    ("o-level", 1, "Matric"), ("o level", 1, "Matric"), ("secondary school certificate", 1, "Matric"),
]

# field keyword, relevant job domains (separated by |)
EDU_FIELDS = [
    ("computer science", "it|data"), ("software engineering", "it|data"), ("information technology", "it|data"),
    ("data science", "data|it"), ("artificial intelligence", "data|it"), ("cyber security", "it"),
    ("bscs", "it|data"), ("bsit", "it|data"), ("bsse", "it|data"), ("mcs", "it|data"), ("cs", "it|data"),
    ("information systems", "it|data|business"), ("mathematics", "data|it|education"),
    ("statistics", "data|finance"), ("electrical", "engineering|it"), ("electronics", "engineering|it"),
    ("telecommunication", "engineering|it"), ("mechanical", "engineering"), ("civil", "engineering"),
    ("business administration", "business|marketing|hr|finance|sales"), ("bba", "business|marketing|hr|finance|sales"),
    ("mba", "business|marketing|hr|finance|sales"), ("commerce", "finance|business|sales"),
    ("accounting", "finance"), ("finance", "finance"), ("economics", "finance|business|data"),
    ("marketing", "marketing|sales|business"), ("human resource", "hr|business"),
    ("graphic design", "design"), ("fine arts", "design"), ("multimedia", "design|marketing"),
    ("education", "education"), ("islamic studies", "education"), ("english", "education|marketing"),
    ("mass communication", "marketing|design"),
]

# name, issuer, domain, aliases
CERTS = [
    ("AWS Certified Solutions Architect", "Amazon", "it", "aws solutions architect"),
    ("AWS Certified Developer", "Amazon", "it", "aws developer associate"),
    ("AWS Cloud Practitioner", "Amazon", "it", "aws certified cloud practitioner"),
    ("Microsoft Azure Fundamentals", "Microsoft", "it", "az-900"),
    ("Microsoft Azure Administrator", "Microsoft", "it", "az-104"),
    ("Google Associate Cloud Engineer", "Google", "it", "associate cloud engineer"),
    ("Google Data Analytics Certificate", "Google", "data", "google data analytics"),
    ("Google Project Management Certificate", "Google", "business", "google project management"),
    ("Google IT Support Certificate", "Google", "it", "google it support"),
    ("Google Analytics Certification", "Google", "marketing", "google analytics certified"),
    ("PMP", "PMI", "business", "project management professional"),
    ("CAPM", "PMI", "business", "certified associate in project management"),
    ("PRINCE2", "Axelos", "business", ""),
    ("Certified ScrumMaster", "Scrum Alliance", "it", "csm|professional scrum master|psm"),
    ("CompTIA A+", "CompTIA", "it", "comptia a plus"),
    ("CompTIA Security+", "CompTIA", "it", "security+"),
    ("CompTIA Network+", "CompTIA", "it", "network+"),
    ("CCNA", "Cisco", "it", "cisco certified network associate"),
    ("CCNP", "Cisco", "it", "cisco certified network professional"),
    ("CEH", "EC-Council", "it", "certified ethical hacker"),
    ("CISSP", "ISC2", "it", ""), ("CISA", "ISACA", "it", ""),
    ("ITIL", "Axelos", "it", ""),
    ("Oracle Certified Professional", "Oracle", "it", "ocp|oracle certified associate|oca"),
    ("TensorFlow Developer Certificate", "Google", "data", ""),
    ("Microsoft Power BI Data Analyst", "Microsoft", "data", "pl-300"),
    ("Tableau Desktop Specialist", "Tableau", "data", "tableau certified"),
    ("IBM Data Science Professional Certificate", "IBM", "data", "ibm data science"),
    ("Machine Learning Specialization", "Coursera", "data", "deeplearning.ai|andrew ng"),
    ("Meta Front-End Developer Certificate", "Meta", "it", "meta frontend"),
    ("Certified Kubernetes Administrator", "CNCF", "it", "cka|ckad"),
    ("Terraform Associate", "HashiCorp", "it", ""),
    ("Red Hat Certified System Administrator", "Red Hat", "it", "rhcsa"),
    ("Docker Certified Associate", "Docker", "it", ""),
    ("HubSpot Certification", "HubSpot", "marketing", "hubspot academy"),
    ("CFA", "CFA Institute", "finance", "chartered financial analyst"),
    ("ACCA", "ACCA", "finance", ""), ("CPA", "AICPA", "finance", "certified public accountant"),
    ("CMA", "IMA", "finance", "certified management accountant"),
    ("SHRM-CP", "SHRM", "hr", "shrm cp|phr"),
    ("TEFL", "TEFL", "education", "tesol|celta"),
    ("Adobe Certified Professional", "Adobe", "design", "adobe certified"),
]

# section, type (heading|regex), pattern, weight   (weights total 100)
SECTIONS = [
    ("Email", "regex", r"[\w.+-]+@[\w-]+\.[\w.-]+", 8),
    ("Phone", "regex", r"\+?\d[\d\s\-().]{8,16}\d", 6),
    ("Online profile", "regex", r"linkedin\.com|github\.com|behance\.net|portfolio", 6),
    ("Summary", "heading", "summary|objective|profile|about me|career objective|professional summary", 10),
    ("Education", "heading", "education|academic|qualification|qualifications|academic background", 15),
    ("Experience", "heading", "experience|employment|work history|internship|internships|professional experience", 20),
    ("Skills", "heading", "skills|technical skills|key skills|core competencies|competencies|expertise", 15),
    ("Projects", "heading", "projects|project|academic projects|portfolio", 10),
    ("Certifications", "heading", "certifications|certification|certificates|courses|training|trainings", 5),
    ("Achievements", "heading", "achievements|awards|honors|honours|accomplishments", 5),
]


def _write(name, header, rows):
    """Write one dataset as a CSV file inside the datasets/ folder."""
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(os.path.join(DATA_DIR, name), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def build_all(force=False):
    """Create all scoring datasets (skipped if they already exist, unless force=True)."""
    if not force and os.path.exists(os.path.join(DATA_DIR, "resume_sections_dataset.csv")):
        return
    _write("skills_dataset.csv", ["skill", "category", "domain", "aliases"], SKILLS)
    _write("role_skills_dataset.csv", ["role", "keywords", "domain", "skills"], ROLE_SKILLS)
    _write("education_dataset.csv", ["type", "keyword", "level", "level_name"],
           [("level", k, l, n) for k, l, n in EDU_LEVELS] + [("field", f, "", d) for f, d in EDU_FIELDS])
    _write("certifications_dataset.csv", ["certification", "issuer", "domain", "aliases"], CERTS)
    _write("resume_sections_dataset.csv", ["section", "type", "pattern", "weight"], SECTIONS)


# ══════════════════════════════════════════════════════════════════════════════
# DATASET-BASED RESUME SCORER
# (skills · experience · education · certifications · completeness  ->  score out of 100)
# ══════════════════════════════════════════════════════════════════════════════
WEIGHTS = {"skills": 35, "experience": 25, "education": 15,
           "certifications": 10, "completeness": 15}
LABELS = {"skills": "Skills Match", "experience": "Experience Match",
          "education": "Education Match", "certifications": "Certifications",
          "completeness": "Resume Completeness"}
LEVEL_NAMES = {0: "Not found", 1: "Matric", 2: "Intermediate", 3: "Diploma",
               4: "Bachelor", 5: "Master", 6: "PhD"}
MONTHS = {m: i + 1 for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"])}


# ───────────────────────── dataset loading ─────────────────────────
def _term(word):
    """Regex that matches a whole word/phrase (so 'c' does not match inside 'science')."""
    return re.compile(r"(?<![a-z0-9+#])" + re.escape(word.lower().strip()) + r"(?![a-z0-9+#])")


def _read(name):
    """Read a dataset CSV; regenerate the datasets if the file is missing."""
    path = os.path.join(DATA_DIR, name)
    if not os.path.exists(path):
        build_all(force=True)
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


@lru_cache(maxsize=1)
def _load():
    """Load every dataset once and pre-compile the regex patterns (cached)."""
    skills = []
    for r in _read("skills_dataset.csv"):
        names = [r["skill"]] + [a for a in r["aliases"].split("|") if a]
        skills.append({"skill": r["skill"], "category": r["category"], "domain": r["domain"],
                       "pats": [_term(n) for n in names]})
    roles = [{"role": r["role"], "keywords": [k for k in r["keywords"].split("|") if k],
              "domain": r["domain"], "skills": r["skills"].split("|")}
             for r in _read("role_skills_dataset.csv")]
    levels, fields = [], []
    for r in _read("education_dataset.csv"):
        if r["type"] == "level":
            levels.append((_term(r["keyword"]), int(r["level"])))
        else:
            fields.append((_term(r["keyword"]), set(r["level_name"].split("|"))))
    certs = [{"name": r["certification"], "domain": r["domain"],
              "pats": [_term(n) for n in [r["certification"]] + [a for a in r["aliases"].split("|") if a]]}
             for r in _read("certifications_dataset.csv")]
    sections = []
    for r in _read("resume_sections_dataset.csv"):
        if r["type"] == "regex":
            pats = [re.compile(r["pattern"], re.I)]
        else:
            pats = [_term(p) for p in r["pattern"].split("|")]
        sections.append({"name": r["section"], "type": r["type"], "pats": pats,
                         "weight": float(r["weight"]), "kw": r["pattern"].split("|")})
    return {"skills": skills, "roles": roles, "levels": levels, "fields": fields,
            "certs": certs, "sections": sections}


# ───────────────────────── text helpers ─────────────────────────
def _find_skills(text_l):
    """Return every dataset skill that appears in the (lower-case) text."""
    return [s for s in _load()["skills"] if any(p.search(text_l) for p in s["pats"])]


def _heading_of(line):
    """If the line looks like a resume section heading, return the section name."""
    s = line.strip()
    if not s or len(s) > 40 or re.search(r"[\d,.]", s.rstrip(":")):
        return None
    cleaned = re.sub(r"[^a-z& ]", " ", s.lower())
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    words = cleaned.split()
    if not words or len(words) > 4:
        return None
    if not (len(words) == 1 or s.isupper() or s.istitle() or s.endswith(":")):
        return None
    for sec in _load()["sections"]:
        if sec["type"] != "heading":
            continue
        for kw in sec["kw"]:
            if re.fullmatch(r"(?:[a-z&]+ ){0,2}" + re.escape(kw) + r"(?: [a-z&]+){0,2}", cleaned):
                return sec["name"]
    return None


def split_sections(text):
    """Split the resume into {section name: text} using the detected headings."""
    out, cur = {}, None
    for line in text.splitlines():
        if not line.strip():
            continue
        h = _heading_of(line)
        if h:
            cur = h
            out.setdefault(cur, [])
        elif cur:
            out[cur].append(line.strip())
    return {k: "\n".join(v) for k, v in out.items()}


def _month(s):
    return MONTHS.get((s or "").strip().lower()[:3], None)


def estimate_years(text, sections):
    """Estimate years of experience using both explicit 'X years experience'
    statements and the date ranges found in the Experience section."""
    tl = text.lower()
    explicit = 0.0
    for pat in (r"(\d{1,2}(?:\.\d)?)\s*\+?\s*(?:years?|yrs?)(?:\s+of)?\s+(?:[a-z]+\s+){0,3}experience",
                r"experience\s*(?:of|:|-)?\s*(\d{1,2}(?:\.\d)?)\s*\+?\s*(?:years?|yrs?)"):
        for m in re.finditer(pat, tl):
            explicit = max(explicit, float(m.group(1)))

    exp_text = sections.get("Experience", "")
    now = datetime.now()
    mon = r"(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?"
    rng = re.compile(r"(?:(" + mon + r")[\s,]*)?((?:19|20)\d{2})\s*(?:-|–|—|to)\s*"
                     r"(?:(?:(" + mon + r")[\s,]*)?((?:19|20)\d{2})|(present|current|now|ongoing|date))", re.I)
    spans = []
    for m in rng.finditer(exp_text):
        sm, sy, em, ey, pres = m.groups()
        start = int(sy) + ((_month(sm) or 1) - 1) / 12
        if pres:
            end = now.year + (now.month - 1) / 12
        else:
            end = int(ey) + ((_month(em) or 12) - 1) / 12
        if end > start:
            spans.append((start, end))
    # Merge overlapping date ranges so parallel jobs are not counted twice
    spans.sort()
    merged = []
    for s, e in spans:
        if merged and s <= merged[-1][1]:
            merged[-1][1] = max(merged[-1][1], e)
        else:
            merged.append([s, e])
    computed = sum(e - s for s, e in merged)
    return round(min(max(explicit, computed), 40), 1)


def required_years(job_text_l):
    """Years of experience requested in the job text (0 if none is mentioned)."""
    m = re.search(r"(\d{1,2})(?:\s*(?:-|to)\s*\d{1,2})?\s*\+?\s*(?:years?|yrs?)", job_text_l)
    return int(m.group(1)) if m else 0


def detect_level(text_l, mode="max"):
    """Highest (or lowest) education level mentioned in the text; 0 if none."""
    found = [lvl for pat, lvl in _load()["levels"] if pat.search(text_l)]
    if not found:
        return 0
    return max(found) if mode == "max" else min(found)


def _job_domain(job_skills, role):
    """Guess the job domain (it, data, finance, ...) from its skills or role."""
    votes = Counter(s["domain"] for s in job_skills if s["domain"] != "business")
    if votes:
        return votes.most_common(1)[0][0]
    return role["domain"] if role else None


def _match_role(title_cat_l):
    """Match the job title / predicted category to a role in the role dataset."""
    for r in _load()["roles"]:
        if any(k in title_cat_l for k in r["keywords"]):
            return r
    return None


def _rating(score):
    return "Excellent" if score >= 80 else "Good" if score >= 60 else "Average" if score >= 40 else "Poor"


# ───────────────────────── main function ─────────────────────────
def score_resume(resume_text, job, category="General"):
    """Score a resume against a job. Returns a dict with the total score (0-100),
    a per-criterion breakdown, strengths, weaknesses and a recommendation."""
    data = _load()
    text = resume_text or ""
    tl = text.lower()
    sections = split_sections(text)

    title = (job.get("title") or "")
    job_text_l = " ".join([title, job.get("description") or "", job.get("requirements") or ""]).lower()
    role = _match_role((title + " " + (category or "")).lower())

    # ---- 1) SKILLS (35 pts): share of the job's skills found in the resume ----
    job_skills = _find_skills(job_text_l)
    used_role_fallback = False
    if not job_skills and role:                      # job lists no known skills -> use role defaults
        names = set(role["skills"])
        job_skills = [s for s in data["skills"] if s["skill"] in names]
        used_role_fallback = True
    res_skills = _find_skills(tl)
    res_names = {s["skill"] for s in res_skills}
    job_names = [s["skill"] for s in job_skills]
    matched = [n for n in job_names if n in res_names]
    missing = [n for n in job_names if n not in res_names]
    if job_names:
        frac = len(matched) / len(job_names)
        sk_detail = f"{len(matched)}/{len(job_names)} required skills found" + \
                    (" (default skills for the role)" if used_role_fallback else "")
    else:
        frac = min(len(res_names) / 8, 1) * 0.8
        sk_detail = f"No skills detected in the job; resume lists {len(res_names)} skills"
    skills_pts = frac * WEIGHTS["skills"]

    # ---- 2) EXPERIENCE (25 pts): candidate years vs required years ----
    years = estimate_years(text, sections)
    req = required_years(job_text_l)
    if req > 0:
        efrac = min(years / req, 1)
        ex_detail = f"Candidate: {years:g} yrs | Required: {req}+ yrs"
    else:
        efrac = 0.6 + 0.4 * min(years / 3, 1)
        ex_detail = f"Candidate: {years:g} yrs | No experience requirement in the job"
    exp_pts = efrac * WEIGHTS["experience"]

    # ---- 3) EDUCATION (15 pts): level (75 %) + relevance of the field (25 %) ----
    edu_text_l = sections.get("Education", "").lower() or tl
    cand_level = detect_level(edu_text_l)
    req_level = detect_level(job_text_l, mode="min") or 4      # default requirement: Bachelor
    lfrac = min(cand_level / req_level, 1) if cand_level else 0
    domain = _job_domain(job_skills, role)
    cand_fields = [doms for pat, doms in data["fields"] if pat.search(edu_text_l)]
    if domain is None:
        ffrac = 1.0
    elif not cand_fields:
        ffrac = 0.5
    else:
        ffrac = 1.0 if any(domain in d for d in cand_fields) else 0.4
    edu_pts = (0.75 * lfrac + 0.25 * ffrac) * WEIGHTS["education"] if cand_level else 0
    ed_detail = (f"Candidate: {LEVEL_NAMES[cand_level]} | Required: {LEVEL_NAMES[req_level]}"
                 + (" | Field relevant" if cand_level and ffrac == 1.0 and domain else ""))

    # ---- 4) CERTIFICATIONS (10 pts): known certificates + other listed ones ----
    found_certs = [c for c in data["certs"] if any(p.search(tl) for p in c["pats"])]
    pts = 0.0
    for c in found_certs:
        pts += 1.0 if (domain is None or c["domain"] == domain) else 0.5
    cert_lines = [l for l in sections.get("Certifications", "").splitlines() if len(l.strip()) > 4]
    pts += 0.4 * min(max(len(cert_lines) - len(found_certs), 0), 4)
    cert_pts = min(pts / 3, 1) * WEIGHTS["certifications"]
    ce_detail = (f"{len(found_certs)} known + {max(len(cert_lines) - len(found_certs), 0)} other certificate(s)"
                 if (found_certs or cert_lines) else "No certifications found")

    # ---- 5) COMPLETENESS (15 pts): which standard resume sections are present ----
    present, absent, got, total = [], [], 0.0, 0.0
    for sec in data["sections"]:
        total += sec["weight"]
        ok = (sec["name"] in sections) if sec["type"] == "heading" else False
        ok = ok or any(p.search(text if sec["type"] == "regex" else tl) for p in sec["pats"])
        (present if ok else absent).append(sec["name"])
        got += sec["weight"] if ok else 0
    cfrac = got / total if total else 0
    words = len(text.split())
    if words < 80:                     # very short resumes are penalised
        cfrac *= 0.5
    elif words < 150:
        cfrac *= 0.75
    comp_pts = cfrac * WEIGHTS["completeness"]
    cm_detail = f"{len(present)}/{len(present) + len(absent)} sections | {words} words"

    parts = {
        "skills": (skills_pts, sk_detail), "experience": (exp_pts, ex_detail),
        "education": (edu_pts, ed_detail), "certifications": (cert_pts, ce_detail),
        "completeness": (comp_pts, cm_detail),
    }
    breakdown = {k: {"label": LABELS[k], "score": round(v[0], 1), "max": WEIGHTS[k], "detail": v[1]}
                 for k, v in parts.items()}
    breakdown["skills"].update(matched=matched, missing=missing)
    breakdown["certifications"]["found"] = [c["name"] for c in found_certs]
    breakdown["completeness"].update(present=present, missing=absent)
    total_score = int(round(sum(v[0] for v in parts.values())))
    total_score = max(0, min(100, total_score))

    # ---- strengths / weaknesses / recommendation ----
    strengths, weaknesses = [], []
    if frac >= 0.7 and job_names:
        strengths.append(f"Strong skills match ({len(matched)}/{len(job_names)} required skills)")
    if req and years >= req:
        strengths.append(f"Meets experience requirement ({years:g} yrs vs {req}+ yrs)")
    elif not req and years >= 2:
        strengths.append(f"{years:g} years of experience")
    if cand_level and cand_level >= req_level:
        strengths.append(f"Education level meets requirement ({LEVEL_NAMES[cand_level]})")
    if found_certs:
        strengths.append("Relevant certifications: " + ", ".join(c["name"] for c in found_certs[:3]))
    if cfrac >= 0.8:
        strengths.append("Well-structured, complete resume")
    if missing:
        weaknesses.append("Missing skills: " + ", ".join(missing[:5]))
    if req and years < req:
        weaknesses.append(f"Experience is below requirement ({years:g} yrs vs {req}+ yrs)")
    if not cand_level:
        weaknesses.append("No education details found in the resume")
    elif cand_level < req_level:
        weaknesses.append(f"Education level below requirement ({LEVEL_NAMES[cand_level]} < {LEVEL_NAMES[req_level]})")
    if not found_certs and not cert_lines:
        weaknesses.append("No certifications listed")
    if absent:
        weaknesses.append("Missing resume sections: " + ", ".join(absent[:4]))
    strengths = strengths or ["Resume reviewed"]
    weaknesses = weaknesses or ["No major gaps found"]

    if total_score >= 80:
        rec = "Strongly recommended — shortlist for an interview right away."
    elif total_score >= 60:
        rec = "Recommended — can be shortlisted for an interview."
    elif total_score >= 40:
        rec = "Average match — compare with other candidates before deciding."
    else:
        rec = "Weak match for this job."

    top = matched[:5] or [s["skill"] for s in res_skills[:5]]
    return {"score": total_score, "rating": _rating(total_score), "breakdown": breakdown,
            "strengths": strengths[:4], "weaknesses": weaknesses[:4], "recommendation": rec,
            "years_exp": f"{years:g} years", "top_skills": top, "method": "dataset"}

# ══════════════════════════════════════════════════════════════════════════════
# UI HELPERS (header, login panel, bordered tables, Enter-to-next, job search, score breakdown)
# ══════════════════════════════════════════════════════════════════════════════
# ═════════════════════════ 1) PORTAL HEADER ═════════════════════════
HEADER_CSS = """<style>
.miu-head{background:linear-gradient(120deg,#003366 0%,#006400 60%,#004d99 100%);border-radius:20px;
padding:1.2rem 1.6rem 0.9rem;margin-bottom:1.4rem;box-shadow:0 10px 40px rgba(0,51,102,.28);}
.miu-brand{display:flex;align-items:center;justify-content:center;gap:1.4rem;flex-wrap:wrap;}
.miu-brand img{width:96px;height:96px;border-radius:50%;border:3px solid rgba(255,255,255,.55);
object-fit:contain;background:#fff;padding:5px;flex-shrink:0;}
.miu-brand-text{display:flex;flex-direction:column;justify-content:center;min-height:96px;}
.miu-uni{font-family:'Playfair Display',serif;font-size:1.7rem;font-weight:900;color:#fff;line-height:1.15;}
.miu-loc{font-size:1rem;font-weight:700;color:#a8d5a2;letter-spacing:1.4px;text-transform:uppercase;margin:3px 0;}
.miu-sys{font-size:.95rem;font-weight:600;color:rgba(255,255,255,.9);}
.miu-user{margin-top:.9rem;padding-top:.7rem;border-top:1px solid rgba(255,255,255,.25);text-align:center;
color:rgba(255,255,255,.85);font-size:.8rem;}
.miu-user .fyp{background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.3);color:#fff;
border-radius:20px;padding:2px 12px;font-size:.72rem;font-weight:700;margin-left:6px;}
</style>"""


def portal_header_html(logo_b64, name, role, rc, role_label, subtitle=""):
    """HTML of the green header shown at the top of every portal."""
    sub = f" &nbsp;|&nbsp; {html.escape(subtitle)}" if subtitle else ""
    img = f'<img src="data:image/jpeg;base64,{logo_b64}" alt="MIU"/>' if logo_b64 else ""
    return (HEADER_CSS +
            '<div class="miu-head">'
            '<div class="miu-brand">'
            f'{img}'
            '<div class="miu-brand-text">'
            '<div class="miu-uni">Mohi-ud-Din Islamic University</div>'
            '<div class="miu-loc">Nerian Sharif, AJK</div>'
            '<div class="miu-sys">AI Resume Screening System</div>'
            '</div></div>'
            '<div class="miu-user">'
            f'Signed in as <b>{html.escape(name)}</b> &nbsp;<span class="{rc}">{role.upper()}</span>'
            f' &nbsp;|&nbsp; {html.escape(role_label)}{sub}'
            ' &nbsp;|&nbsp; <span class="db-badge">🗄️ SQLite DB</span>'
            '<span class="fyp">🎓 Final Year Project — CS &amp; IT</span>'
            '</div></div>')


# ═════════════════════════ 2) LOGIN PAGE ═════════════════════════
# Layout: full-width green hero on top. The university LOGO sits right BEFORE the
# university name (same line, left side), then the location and the project name.
# The Sign In / Sign Up form is below the hero.
AUTH_CSS = """<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');
.stApp{background:#eef3f0!important};
.block-container{max-width:100%!important;padding:1rem clamp(1rem,3vw,3rem) 1.5rem!important;margin:0 auto!important;}

/* ---------- GREEN HERO ---------- */
.auth-hero{width:100%;margin:0 auto 1.6rem;background:linear-gradient(120deg,#0b2a4a 0%,#0f5132 100%);
border-radius:24px;padding:1.8rem clamp(1.2rem,3vw,3rem);color:#fff;font-family:'Plus Jakarta Sans',sans-serif;
display:flex;flex-direction:column;gap:1.4rem;box-shadow:0 14px 40px rgba(11,42,74,.25);}
/* LOGO on top (big), then university name, then project name */
.auth-brand{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:.7rem;text-align:center;}
.auth-brand-uni{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:.9rem;
font-family:'Playfair Display',serif;font-size:clamp(1.4rem,3.4vw,2.5rem);font-weight:900;
line-height:1.2;color:#fff;}
/* BIG logo */
.uni-inline{width:clamp(110px,14vw,160px);height:clamp(110px,14vw,160px);border-radius:50%;
background:#fff;padding:6px;object-fit:contain;border:4px solid rgba(255,255,255,.65);
box-shadow:0 8px 28px rgba(0,0,0,.3);flex-shrink:0;}
.auth-brand-uni .loc{color:#a8d5a2;}
.auth-brand-proj{font-size:clamp(1.1rem,2.3vw,1.6rem);font-weight:800;color:#fff;}
.auth-mid{text-align:center;max-width:900px;margin:0 auto;}
.auth-hero h1{font-family:'Playfair Display',serif;font-size:1.45rem;line-height:1.25;margin:0 0 .6rem;color:#fff;}
.auth-hero p{color:rgba(255,255,255,.78);font-size:.98rem;line-height:1.65;margin:0;}
.auth-points{display:flex;flex-direction:row;flex-wrap:wrap;gap:16px;margin-top:.4rem;}
.auth-point{display:flex;gap:14px;align-items:flex-start;flex:1 1 260px;background:rgba(255,255,255,.08);
border-radius:16px;padding:14px 16px;}
.auth-point .ic{width:44px;height:44px;border-radius:12px;background:rgba(255,255,255,.14);
display:flex;align-items:center;justify-content:center;font-size:1.25rem;flex-shrink:0;}
.auth-point b{display:block;font-size:.95rem;color:#fff;}
.auth-point span{font-size:.82rem;color:rgba(255,255,255,.72);}
.auth-uni{padding-top:.9rem;border-top:1px solid rgba(255,255,255,.2);font-size:.78rem;
color:rgba(255,255,255,.65);text-align:center;}

/* ---------- FORM ---------- */
.auth-form-title{font-family:'Playfair Display',serif;font-size:1.6rem;font-weight:800;color:#0b2a4a;margin:.2rem 0;}
.auth-form-sub{color:#667;font-size:.88rem;margin-bottom:1rem;}
div[data-testid="stTabs"] button[role="tab"]{font-weight:600;font-size:1rem;}
div[data-testid="stTextInput"] input{border-radius:10px;}
@media (max-width:900px){.auth-hero h1{font-size:1.2rem;}}
</style>"""

def auth_hero_html(logo_b64):
    """Login hero: logo + university name + Nerian Sharif, AJK on the SAME line."""
    inline = (f'<img class="uni-inline" src="data:image/jpeg;base64,{logo_b64}" alt="MIU"/>'
              if logo_b64 else "")
    return (
        '<div class="auth-hero">'
        '<div class="auth-brand">'
        f'<div class="auth-brand-uni">{inline}'
        '<span>Mohi-ud-Din Islamic University, <span class="loc">Nerian Sharif, AJK</span></span></div>'
        '<div class="auth-brand-proj">AI Based Resume Screening System</div>'
        '</div>'
        '<div class="auth-mid">'
        '<h1>Connect the right candidate with the right job.</h1>'
        '<p>AI-powered resume screening, company job postings and candidate applications — all in one portal.</p>'
        '</div>'
        '<div class="auth-points">'
        '<div class="auth-point"><div class="ic">🤖</div><div><b>Smart Resume Scoring</b>'
        '<span>Skills, experience, education, certificates &amp; completeness — 1-100 score</span></div></div>'
        '<div class="auth-point"><div class="ic">🏢</div><div><b>For Companies</b>'
        '<span>Post jobs and view all applications in one place</span></div></div>'
        '<div class="auth-point"><div class="ic">📄</div><div><b>For Job Seekers</b>'
        '<span>Upload your resume and your details are filled in automatically</span></div></div>'
        '</div>'
        '<div class="auth-uni">Mohi-ud-Din Islamic University, Nerian Sharif AJK · '
        'Final Year Project — CS &amp; IT Department</div>'
        '</div>')
        # ═════════════════════════ 3) BOXED TABLES ═════════════════════════
TABLE_CSS = """<style>
.miu-table-box{background:#fff;border:2px solid #003366;border-radius:12px;overflow:auto;
box-shadow:0 4px 18px rgba(0,51,102,.12);margin:.4rem 0 1rem;}
.miu-table{width:100%;border-collapse:collapse;font-size:.84rem;}
.miu-table th{background:linear-gradient(90deg,#003366,#006400);color:#fff;text-align:left;padding:10px 12px;
border:1px solid #1d4e7a;position:sticky;top:0;white-space:nowrap;}
.miu-table td{padding:9px 12px;border:1px solid #c5d0dc;color:#1f2d3d;vertical-align:top;}
.miu-table td.idx{background:#eaf1f8;color:#003366;font-weight:700;text-align:center;width:44px;}
.miu-table tr:nth-child(even) td{background:#f4f8fb;}
.miu-table tr:hover td{background:#e8f5e9;}
</style>"""


def boxed_table(df, max_height=420, empty_text="No records found.", show_index=True, **_ignored):
    """Replacement for st.dataframe: a table with a visible box around every cell (grid lines).
    Old arguments such as use_container_width / hide_index are accepted and ignored."""
    if df is None or len(df) == 0:
        st.markdown(TABLE_CSS + f'<div class="miu-table-box" style="padding:1rem;color:#667;">'
                    f'{html.escape(empty_text)}</div>', unsafe_allow_html=True)
        return
    head = ("<th>#</th>" if show_index else "") + "".join(f"<th>{html.escape(str(c))}</th>" for c in df.columns)
    rows = []
    for i, row in enumerate(df.itertuples(index=False, name=None), start=1):
        cells = "".join(
            "<td>" + html.escape("" if (not isinstance(v, (list, dict)) and pd.isna(v)) else str(v)).replace("\n", " ") + "</td>"
            for v in row)
        rows.append("<tr>" + (f'<td class="idx">{i}</td>' if show_index else "") + cells + "</tr>")
    st.markdown(TABLE_CSS +
                f'<div class="miu-table-box" style="max-height:{max_height}px;"><table class="miu-table">'
                f'<thead><tr>{head}</tr></thead><tbody>{"".join(rows)}</tbody></table></div>',
                unsafe_allow_html=True)


def table_with_search(df, key, max_height=420):
    """Search bar + boxed table. Only rows matching what you type are shown (all columns are searched)."""
    q = st.text_input("🔍 Search table", key=f"tbl_{key}", placeholder="🔍 Type to search in table… (press Enter)",
                      label_visibility="collapsed")
    shown = df
    if q.strip() and len(df):
        as_text = df.astype(str)
        mask = pd.Series(True, index=df.index)
        for tok in q.lower().split():          # every typed word must match some column
            mask &= as_text.apply(lambda c: c.str.lower().str.contains(tok, regex=False)).any(axis=1)
        shown = df[mask]
        st.caption(f"{len(shown)} of {len(df)} rows match “{q}”")
    boxed_table(shown, max_height=max_height, empty_text="No matching records.")


# ═════════════════════════ 4) JOB SEARCH ═════════════════════════
_JOB_FIELDS = ("job_id", "title", "company", "department", "location", "type",
               "salary", "description", "requirements", "deadline", "posted")


def filter_jobs(jobs, query):
    """Every word of the query must appear in some field of the job (title, company, location,
    department, type, salary, description, requirements...)."""
    tokens = [t for t in re.split(r"\s+", (query or "").lower().strip()) if t]
    if not tokens:
        return jobs
    out = []
    for j in jobs:
        hay = " ".join(str(j.get(f) or "") for f in _JOB_FIELDS).lower()
        if all(t in hay for t in tokens):
            out.append(j)
    return out


# ═════════════════════════ 5) ENTER  ->  NEXT BOX ═════════════════════════
_ENTER_JS = """
<script>
(function () {
  var doc = window.parent.document;
  if (doc.__miuEnterHandler) doc.removeEventListener('keydown', doc.__miuEnterHandler, true);
  var OK = ['text', 'password', 'number', 'email', 'tel', 'search'];
  function visible(el) { return !!(el.offsetParent || el.getClientRects().length); }
  function usable(x) {
    if (x.disabled || !visible(x)) return false;
    if (x.getAttribute('role') === 'combobox' || x.getAttribute('aria-autocomplete')) return false;
    if (x.tagName === 'TEXTAREA') return true;
    return x.tagName === 'INPUT' && OK.indexOf((x.getAttribute('type') || 'text').toLowerCase()) !== -1;
  }
  function handler(e) {
    if (e.key !== 'Enter' || e.shiftKey || e.isComposing) return;
    var el = e.target;
    if (!el || el.tagName !== 'INPUT' || !usable(el)) return;
    if ((el.getAttribute('aria-label') || '').indexOf('\\uD83D\\uDD0D') === 0) return;  // search boxes: plain Enter only
    var all = Array.prototype.slice.call(doc.querySelectorAll('input, textarea')).filter(usable);
    var i = all.indexOf(el);
    if (i === -1) return;
    var next = all[i + 1];
    if (next) {
      setTimeout(function () { next.focus(); try { next.select(); } catch (_) {} }, 80);
      return;
    }
    // last box: commit the value, then press the next primary button
    var btns = Array.prototype.slice.call(
      doc.querySelectorAll('button[data-testid="stBaseButton-primary"], button[kind="primary"]')
    ).filter(function (b) {
      return visible(b) && !b.disabled && (el.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_FOLLOWING);
    });
    if (btns.length) {
      el.blur();
      setTimeout(function () { btns[0].click(); }, 350);
    }
  }
  doc.__miuEnterHandler = handler;
  doc.addEventListener('keydown', handler, true);
})();
</script>
"""


def enable_enter_to_next():
    """Call once per page. Enter moves the cursor to the next box; Enter in the last box
    presses the form's primary button (Sign in / Create / Publish)."""
    if hasattr(st, "iframe"):           # newer Streamlit (components.html is deprecated)
        st.iframe(_ENTER_JS, height=1)
    else:                              # older Streamlit
        components.html(_ENTER_JS, height=0)


# ═════════════════════════ 6) SCORE BREAKDOWN ═════════════════════════
def _bar_color(pct):
    """Colour of a progress bar depending on the percentage."""
    return "#28a745" if pct >= 80 else "#17a2b8" if pct >= 60 else "#fd7e14" if pct >= 40 else "#dc3545"


def render_score_breakdown(result):
    """Show the per-criterion score bars, matched/missing skills and the recommendation."""
    if not result or "breakdown" not in result:
        st.info("No score details available for this application.")
        return
    rows = ""
    for b in result["breakdown"].values():
        pct = int(100 * b["score"] / b["max"]) if b["max"] else 0
        rows += (
            '<div style="margin:10px 0;">'
            '<div style="display:flex;justify-content:space-between;font-size:.84rem;font-weight:700;color:#003366;">'
            f'<span>{html.escape(b["label"])}</span><span>{b["score"]:g} / {b["max"]}</span></div>'
            '<div style="background:#eceff1;border-radius:8px;height:9px;overflow:hidden;">'
            f'<div style="width:{pct}%;background:{_bar_color(pct)};height:100%;"></div></div>'
            f'<div style="font-size:.74rem;color:#667;margin-top:2px;">{html.escape(b.get("detail", ""))}</div></div>')
    sk = result["breakdown"].get("skills", {})
    chips = "".join(f'<span class="skill-chip">✅ {html.escape(s)}</span>' for s in sk.get("matched", []))
    chips += "".join(f'<span class="skill-chip" style="background:#fdecea;color:#b71c1c;">❌ {html.escape(s)}</span>'
                     for s in sk.get("missing", []))
    certs = result["breakdown"].get("certifications", {}).get("found", [])
    cert_html = ("<div style='margin-top:8px;font-size:.8rem;'><b>🎓 Certifications:</b> " +
                 html.escape(", ".join(certs)) + "</div>") if certs else ""
    st.markdown(
        '<div class="card">'
        f'<div style="font-weight:800;color:#003366;margin-bottom:4px;">Total: {result["score"]}/100 — '
        f'{html.escape(result["rating"])}</div>{rows}'
        f'<div style="margin-top:8px;">{chips}</div>{cert_html}'
        f'<div style="margin-top:10px;font-size:.82rem;"><b>💡</b> {html.escape(result.get("recommendation", ""))}</div>'
        '</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="MIU — AI Recruitment Portal",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ══════════════════════════════════════════════════════════════════════════════
# GLOBAL CSS (cards, tags, score ring, role badges, buttons ...)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700;900&family=Inter:wght@300;400;500;600;700&display=swap');
html,body,[class*="css"]{font-family:'Inter',sans-serif;background:#f0f4f8;}
#MainMenu,footer,header{visibility:hidden;}
.block-container{padding-top:1rem;}
.sec-title{font-family:'Playfair Display',serif;font-size:1.15rem;font-weight:700;
  color:#003366;border-left:4px solid #006400;padding-left:12px;margin:1.6rem 0 1rem;display:block;}
.card{background:white;border-radius:16px;padding:1.4rem 1.6rem;border:1px solid #e0e8f0;
  margin-bottom:1rem;box-shadow:0 4px 20px rgba(0,51,102,.07);transition:box-shadow .2s;}
.card:hover{box-shadow:0 8px 32px rgba(0,51,102,.14);}
.job-card{background:white;border-radius:16px;padding:1.4rem 1.6rem;border:1px solid #e0e8f0;
  margin-bottom:1rem;box-shadow:0 4px 20px rgba(0,51,102,.07);border-left:5px solid #006400;}
.job-title{font-size:1.05rem;font-weight:700;color:#003366;}
.job-company{font-size:.85rem;color:#555;margin-top:2px;}
.job-tag{display:inline-block;background:#e8f4fd;color:#0d6efd;border-radius:20px;
  padding:2px 10px;font-size:.72rem;font-weight:600;margin:2px;}
.job-tag-green{background:#e8f5e9;color:#1b5e20;}
.job-tag-orange{background:#fff3e0;color:#e65100;}
.stat-tile{background:white;border-radius:16px;padding:1.3rem 1rem;text-align:center;
  border:1px solid #e0e8f0;box-shadow:0 4px 16px rgba(0,51,102,.07);}
.stat-num{font-family:'Playfair Display',serif;font-size:2.5rem;font-weight:900;
  color:#003366;line-height:1;}
.stat-lbl{font-size:.73rem;font-weight:600;color:#666;text-transform:uppercase;
  letter-spacing:.8px;margin-top:6px;}
.score-ring{display:inline-flex;align-items:center;justify-content:center;width:78px;
  height:78px;border-radius:50%;flex-direction:column;box-shadow:0 6px 20px rgba(0,0,0,.2);flex-shrink:0;}
.score-ring .num{font-size:1.7rem;font-weight:900;color:white;line-height:1;}
.score-ring .sub{font-size:.6rem;color:rgba(255,255,255,.8);}
.rating-excellent{background:#28a745;color:white;border-radius:20px;padding:4px 14px;
  font-size:.8rem;font-weight:700;display:inline-block;}
.rating-good{background:#17a2b8;color:white;border-radius:20px;padding:4px 14px;
  font-size:.8rem;font-weight:700;display:inline-block;}
.rating-average{background:#fd7e14;color:white;border-radius:20px;padding:4px 14px;
  font-size:.8rem;font-weight:700;display:inline-block;}
.rating-poor{background:#dc3545;color:white;border-radius:20px;padding:4px 14px;
  font-size:.8rem;font-weight:700;display:inline-block;}
.btn-email{background:#0066cc;color:white!important;padding:7px 16px;border-radius:8px;
  text-decoration:none!important;font-size:.8rem;font-weight:600;display:inline-block;margin:2px;}
.btn-whatsapp{background:#25D366;color:white!important;padding:7px 16px;border-radius:8px;
  text-decoration:none!important;font-size:.8rem;font-weight:600;display:inline-block;margin:2px;}
.btn-zoom{background:#2D8CFF;color:white!important;padding:7px 16px;border-radius:8px;
  text-decoration:none!important;font-size:.8rem;font-weight:600;display:inline-block;margin:2px;}
.cat-badge{background:linear-gradient(90deg,#003366,#006400);color:white;border-radius:20px;
  padding:3px 14px;font-size:.75rem;font-weight:600;display:inline-block;}
.role-admin{background:#dc3545;color:white;padding:2px 10px;border-radius:10px;
  font-size:.72rem;font-weight:700;}
.role-hr{background:#fd7e14;color:white;padding:2px 10px;border-radius:10px;
  font-size:.72rem;font-weight:700;}
.role-company{background:#6610f2;color:white;padding:2px 10px;border-radius:10px;
  font-size:.72rem;font-weight:700;}
.role-recruiter{background:#0d6efd;color:white;padding:2px 10px;border-radius:10px;
  font-size:.72rem;font-weight:700;}
.skill-chip{display:inline-block;background:#e3f2fd;color:#0d47a1;border-radius:14px;
  padding:3px 11px;font-size:.75rem;font-weight:600;margin:2px;}
.footer{text-align:center;color:#888;font-size:.76rem;padding:1.5rem 0 1rem;
  border-top:1px solid #e0e8f0;margin-top:2rem;}
.db-badge{background:#e8f5e9;color:#1b5e20;border-radius:8px;padding:4px 10px;
  font-size:.75rem;font-weight:600;display:inline-block;border:1px solid #c8e6c9;}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# INIT
# ══════════════════════════════════════════════════════════════════════════════
init_db()            # create tables + default accounts
build_all()          # create the scoring datasets (datasets/*.csv) on first run

if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
if 'claude_api_key' not in st.session_state:
    st.session_state.claude_api_key = ""


# ══════════════════════════════════════════════════════════════════════════════
# ML MODELS (resume category classifier)
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_resource
def load_models():
    """Load the pickled classifier, TF-IDF vectoriser and label encoder (None if missing)."""
    try:
        with open("clf.pkl", "rb") as a, open("tfidf.pkl", "rb") as b, open("encoder.pkl", "rb") as c:
            return pickle.load(a), pickle.load(b), pickle.load(c)
    except Exception:
        return None, None, None


svc_model, tfidf, le = load_models()


def clean_resume(txt):
    """Remove URLs and non-ASCII characters before classification."""
    txt = re.sub(r'http\S+', ' ', txt)
    txt = re.sub(r'[^\x00-\x7f]', ' ', txt)
    return re.sub(r'\s+', ' ', txt).strip()


def extract_text(f):
    """Read the text of an uploaded PDF / DOCX / TXT resume."""
    f.seek(0)   # rewind the file before reading it again (important!)
    ext = f.name.rsplit('.', 1)[-1].lower()
    if ext == 'pdf':
        return ''.join(p.extract_text() or '' for p in PyPDF2.PdfReader(f).pages)
    if ext == 'docx':
        return '\n'.join(p.text for p in docx.Document(f).paragraphs)
    raw = f.read()
    try:
        return raw.decode('utf-8')
    except Exception:
        return raw.decode('latin-1')


def predict_category(text):
    """Predict the resume category (e.g. 'Data Science'); 'General' if no model is available."""
    if svc_model is None:
        return "General"
    try:
        vec = tfidf.transform([clean_resume(text)])
        try:
            vec = vec.toarray()
        except AttributeError:
            pass
        return le.inverse_transform(svc_model.predict(vec))[0]
    except Exception:
        return "General"


def extract_candidate_info(text: str) -> dict:
    """Pull the candidate's name, e-mail and phone out of the resume text (used for auto-fill)."""
    info = {"name": "", "email": "", "phone": ""}
    m = re.search(r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}', text)
    if m:
        info["email"] = m.group().strip()
    m = re.search(r'(\+?\d[\d\s\-().]{8,16}\d)', text)
    if m:
        raw = m.group().strip()
        if 7 <= len(re.sub(r'\D', '', raw)) <= 15:
            info["phone"] = raw
    # Words that are headings, not names
    SKIP = {'resume', 'cv', 'curriculum', 'vitae', 'objective', 'summary',
            'profile', 'skills', 'experience', 'education', 'email', 'phone',
            'address', 'contact', 'linkedin', 'github', 'portfolio',
            'certifications', 'languages', 'projects', 'awards'}
    # The name is usually one of the first short lines made only of letters
    for line in [l.strip() for l in text.split('\n') if l.strip()][:20]:
        if '@' in line or 'http' in line.lower():
            continue
        if sum(c.isdigit() for c in line) > 3:
            continue
        words = line.split()
        if not 2 <= len(words) <= 5:
            continue
        if not all(re.match(r"^[A-Za-z][A-Za-z'\-\.]*$", w) for w in words):
            continue
        if any(w.lower() in SKIP for w in words):
            continue
        info["name"] = " ".join(w.capitalize() for w in words)
        break
    return info


# ══════════════════════════════════════════════════════════════════════════════
# AI RATING (optional — uses the Claude API)
# ══════════════════════════════════════════════════════════════════════════════
def rate_resume_with_ai(resume_text, category, api_key):
    """Ask Claude to rate a resume. Always returns a dict (errors are returned, not raised)."""
    raw = ""
    try:
        client = anthropic.Anthropic(api_key=api_key.strip())
        system = ("You are an expert HR recruiter AI. "
                  "You ONLY respond with valid JSON — no markdown, no extra text.")
        prompt = f"""Analyze this resume for the role: {category}

RESUME (first 2500 chars):
{resume_text[:2500]}

Respond EXACTLY with this JSON (real values from the resume):
{{
  "score": 72,
  "rating": "Good",
  "strengths": ["point1","point2","point3"],
  "weaknesses": ["gap1","gap2"],
  "recommendation": "One sentence recommendation.",
  "years_exp": "3-5 years",
  "top_skills": ["Skill1","Skill2","Skill3"]
}}

Rules:
- score: integer 1-100
- rating: exactly one of Excellent(80-100) Good(60-79) Average(40-59) Poor(0-39)
- strengths: 3 specific real observations
- weaknesses: 2 specific gaps
- top_skills: 3 skills found in resume"""

        msg = client.messages.create(
            model="claude-haiku-4-5-20251001", max_tokens=600,
            system=system,
            messages=[{"role": "user", "content": prompt}]
        )
        raw = msg.content[0].text.strip()
        raw = re.sub(r'```(?:json)?\s*', '', raw).strip()      # strip markdown fences if any
        s, e = raw.find('{'), raw.rfind('}')
        if s != -1 and e > s:
            raw = raw[s:e + 1]
        result = json.loads(raw)
        score = max(1, min(100, int(result.get("score", 50))))
        result["score"] = score
        auto_r = ("Excellent" if score >= 80 else "Good" if score >= 60
                  else "Average" if score >= 40 else "Poor")
        if result.get("rating", "") not in ["Excellent", "Good", "Average", "Poor"]:
            result["rating"] = auto_r
        result.setdefault("strengths", ["Resume reviewed"])
        result.setdefault("weaknesses", ["Manual review recommended"])
        result.setdefault("top_skills", [])
        result.setdefault("years_exp", "Not specified")
        result.setdefault("recommendation", "Review manually.")
        return result
    except json.JSONDecodeError as e:
        # The reply was not valid JSON: fall back to the first number as a rough score
        nums = re.findall(r'\b([1-9][0-9]?|100)\b', raw)
        fs = int(nums[0]) if nums else 55
        return {"score": fs, "rating": "Good" if fs >= 60 else "Average",
                "strengths": ["AI response received", "Text analysed", "Category matched"],
                "weaknesses": ["Parse failed", "Manual review needed"],
                "recommendation": "Partial AI response — review manually.",
                "years_exp": "Unknown", "top_skills": [], "error": str(e)}
    except anthropic.AuthenticationError:
        return {"score": 0, "rating": "Auth Error", "strengths": [], "top_skills": [],
                "weaknesses": ["Invalid API key", "Get key at console.anthropic.com"],
                "recommendation": "Fix API key.", "years_exp": "N/A", "error": "bad_api_key"}
    except anthropic.RateLimitError:
        return {"score": 0, "rating": "Rate Limited", "strengths": [], "top_skills": [],
                "weaknesses": ["Too many requests", "Wait 30-60 sec and retry"],
                "recommendation": "Retry later.", "years_exp": "N/A", "error": "rate_limit"}
    except Exception as e:
        return {"score": 0, "rating": "Error", "strengths": [], "top_skills": [],
                "weaknesses": [str(e)[:100], "Check internet / API key"],
                "recommendation": "Manual review needed.", "years_exp": "N/A", "error": str(e)}


def score_color(s):
    """Colour for a score: green / blue / orange / red."""
    if s >= 80:
        return "#28a745"
    if s >= 60:
        return "#17a2b8"
    if s >= 40:
        return "#fd7e14"
    return "#dc3545"


def rating_class(r):
    """CSS class for a rating label."""
    r = (r or "").lower()
    if "excellent" in r:
        return "rating-excellent"
    if "good" in r:
        return "rating-good"
    if "average" in r:
        return "rating-average"
    return "rating-poor"


# ══════════════════════════════════════════════════════════════════════════════
# SHARED UI HELPERS (header + sidebar)
# ══════════════════════════════════════════════════════════════════════════════
def display_name(user):
    """HR users are shown by their USERNAME (e.g. 'hr_manager'); every other role
    is shown by its real name."""
    return user['username'] if user.get('role') == 'hr' else user['name']


ROLE_CLASS = {"admin": "role-admin", "hr": "role-hr",
              "company": "role-company", "recruiter": "role-recruiter"}


def render_header(subtitle=""):
    """Green portal header with the logo, university name and the signed-in user."""
    role = st.session_state.user_info['role']
    name = display_name(st.session_state.user_info)
    rc = ROLE_CLASS.get(role, "role-recruiter")
    labels = {"admin": "Admin Portal", "hr": "HR Portal",
              "company": "Company Portal", "recruiter": "Recruiter Portal"}
    st.markdown(portal_header_html(MIU_LOGO_B64, name, role, rc, labels.get(role, ""), subtitle),
                unsafe_allow_html=True)


def render_sidebar():
    """Sidebar: user card, API key + interview settings (HR/Admin only) and Logout."""
    user = st.session_state.user_info
    role = user['role']
    rc = ROLE_CLASS.get(role, "role-recruiter")
    shown_name = html.escape(display_name(user))
    with st.sidebar:
        icon = {"admin": "👑", "hr": "🏢", "company": "🏗️", "recruiter": "👤"}.get(role, "👤")
        st.markdown(f"""
<div style="background:linear-gradient(135deg,#003366,#006400);border-radius:14px;
            padding:1.1rem;text-align:center;margin-bottom:1rem;">
  <div style="font-size:2.2rem;">{icon}</div>
  <div style="font-weight:700;font-size:.95rem;color:white;">{shown_name}</div>
  <div style="font-size:.75rem;color:rgba(255,255,255,.7);">{html.escape(user['email'])}</div>
  <span class="{rc}">{role.upper()}</span>
</div>
<div style="background:#e8f5e9;border-radius:8px;padding:.6rem .9rem;
            font-size:.75rem;color:#1b5e20;font-weight:600;margin-bottom:.8rem;
            border:1px solid #c8e6c9;">
  🗄️ Data stored in <b>SQLite</b> — persists across sessions
</div>""", unsafe_allow_html=True)

        # HR: show the company that is currently open + a button to change it
        if role == 'hr' and st.session_state.get('hr_company'):
            st.markdown(
                '<div style="background:#fff3e0;border:1px solid #ffe0b2;border-radius:8px;'
                'padding:.6rem .9rem;font-size:.8rem;color:#bf360c;font-weight:600;margin-bottom:.6rem;">'
                f'🏢 Company: <b>{html.escape(st.session_state.hr_company)}</b></div>',
                unsafe_allow_html=True)
            if st.button("🔄 Change company", use_container_width=True, key="chg_co"):
                st.session_state.pop('hr_company', None)
                st.rerun()

        if role in ['admin', 'hr']:
            st.markdown("---")
            st.markdown("**🤖 AI Rating (Claude API)**")
            st.text_input("API Key (sk-ant-...)", type="password",
                          key="claude_api_key", placeholder="sk-ant-api03-...",
                          help="Only visible to HR & Admin.")
            if st.session_state.claude_api_key:
                st.success("✅ API key ready — AI rating enabled")
            else:
                st.info("Paste key above to enable AI scoring")

            st.markdown("---")
            st.markdown("**📅 Interview Settings**")
            st.session_state['int_interviewer'] = st.text_input("Interviewer", display_name(user), key="s_int")
            st.session_state['int_date'] = st.text_input("Date & Time", "Monday, May 20, 2026 at 3 PM", key="s_dt")
            st.session_state['int_zoom'] = st.text_input("Zoom Link", "https://zoom.us/j/your-id", key="s_zm")
            st.session_state['int_company'] = st.text_input("Organization", "MIU", key="s_co")

        st.markdown("---")
        st.caption(f"🕐 Login: {st.session_state.get('login_time', '')}")
        if st.button("🚪 Logout", use_container_width=True):
            for k in ['logged_in', 'username', 'user_info', 'login_time',
                      'claude_api_key', 'fresh_ratings', 'hr_company']:
                st.session_state.pop(k, None)
            st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# LOGIN / SIGN UP
# ══════════════════════════════════════════════════════════════════════════════
def _start_session(user):
    """Store the user in the session and reload the app."""
    st.session_state.logged_in = True
    st.session_state.username = user["username"]
    st.session_state.user_info = user
    st.session_state.pop('hr_company', None)     # HR must pick a company after every sign-in
    st.session_state.login_time = datetime.now().strftime("%Y-%m-%d %H:%M")
    time.sleep(0.4)
    st.rerun()


def show_login():
    # Inject the login-page CSS (logo before the university name, bigger/bolder names, smaller tagline)
    st.markdown(AUTH_CSS, unsafe_allow_html=True)

    # Green hero on TOP (full width)
    st.markdown(auth_hero_html(MIU_LOGO_B64), unsafe_allow_html=True)

    # Sign In / Sign Up form BELOW, centred with auto margins
    _, right, _ = st.columns([1, 1.6, 1])

    with right:
        tab_in, tab_up = st.tabs(["🔐 Sign In", "📝 Sign Up"])

        # ---------------- Sign In ----------------
        with tab_in:
            st.markdown('<div class="auth-form-title">Welcome back</div>'
                        '<div class="auth-form-sub">Sign in to your account.</div>',
                        unsafe_allow_html=True)
            username = st.text_input("Username", key="li_user", placeholder="username")
            password = st.text_input("Password", type="password", key="li_pass", placeholder="password")
            if st.button("Sign in", use_container_width=True, type="primary", key="li_btn"):
                if not (username and password):
                    st.warning("Please enter both username and password.")
                else:
                    user = db_authenticate(username.strip().lower(), password)
                    if user:
                        _start_session(user)
                    else:
                        st.error("Incorrect username or password.")

            with st.expander("Demo accounts"):
                st.markdown("""
| Portal | Username | Password |
|---|---|---|
| Admin | `admin` | `admin123` |
| HR | `hr_manager` | `hr2026` |
| Company | `techcorp` | `tech2026` |
| Job Seeker | `recruiter1` | `recruit123` |
""")

        # ---------------- Sign Up ----------------
        with tab_up:
            st.markdown('<div class="auth-form-title">Create a new account</div>'
                        '<div class="auth-form-sub">First choose what you want to do.</div>',
                        unsafe_allow_html=True)

            # 1st option = Post Job (company) · 2nd = Apply Job (job seeker) · 3rd = Hire Staff (HR)
            purpose = st.radio(
                "What do you want to do?",
                ["📝 Post Job (Company)",
                 "💼 Apply Job (Job Seeker)",
                 "👥 Hire Staff (HR)"],
                key="su_purpose",
            )
            if purpose.startswith("📝"):
                new_role = "company"
            elif purpose.startswith("💼"):
                new_role = "recruiter"
            else:
                new_role = "hr"

            email_re = r"^[^@\s]+@[^@\s]+\.[A-Za-z]{2,}$"

            if new_role == "company":
                # ---- Post Job: company name, phone, e-mail, create / confirm password ----
                company_name = st.text_input("Company name*", key="su_company")
                c_phone = st.text_input("Company phone number*", key="su_cphone", placeholder="+92 300 1234567")
                c_email = st.text_input("Company email*", key="su_cemail", placeholder="hr@company.com")
                c_pw1 = st.text_input("Create password* (at least 6 characters)", type="password", key="su_cpw1")
                c_pw2 = st.text_input("Confirm password*", type="password", key="su_cpw2")

                if st.button("Create company account", use_container_width=True, type="primary", key="su_cbtn"):
                    digits = re.sub(r"\D", "", c_phone)
                    if not all([company_name.strip(), c_phone.strip(), c_email.strip(), c_pw1, c_pw2]):
                        st.error("All fields marked with (*) are required.")
                    elif not re.match(email_re, c_email.strip()):
                        st.error("Please enter a valid email address.")
                    elif not 10 <= len(digits) <= 15:
                        st.error("Please enter a valid phone number (10-15 digits).")
                    elif len(c_pw1) < 6:
                        st.error("Password must be at least 6 characters.")
                    elif c_pw1 != c_pw2:
                        st.error("Passwords do not match.")
                    elif db_email_exists(c_email.strip()):
                        st.error("An account with this email already exists.")
                    elif db_company_name_exists(company_name):
                        st.error("A company with this name is already registered.")
                    else:
                        login_name = c_email.strip().lower()      # the company signs in with its e-mail
                        info = {"name": company_name.strip(), "industry": "", "size": "1-50",
                                "website": "", "about": "", "phone": c_phone.strip(),
                                "email": c_email.strip()}
                        ok = db_create_user(login_name, c_pw1, "company", company_name.strip(),
                                            c_email.strip(), info, phone=c_phone.strip())
                        if not ok:
                            st.error("An account with this email already exists.")
                        else:
                            st.success("Company account created! Signing you in…")
                            _start_session(db_authenticate(login_name, c_pw1))
                st.caption("You will sign in with your company email and password.")

            else:
                # ---- Apply Job (job seeker) and Hire Staff (HR): personal details ----
                n1, n2 = st.columns(2)
                first_name = n1.text_input("First name*", key="su_fn")
                last_name = n2.text_input("Last name*", key="su_ln")
                uname = st.text_input("Username*", key="su_user", placeholder="short, no spaces")
                email = st.text_input("Email*", key="su_email", placeholder="you@email.com")
                g1, g2 = st.columns(2)
                age = g1.number_input("Age*", min_value=16, max_value=100, value=18, step=1, key="su_age")
                gender = g2.selectbox("Gender*", ["Male", "Female", "Other"], key="su_gender")
                pw1 = st.text_input("Create password* (at least 6 characters)", type="password", key="su_pw1")
                pw2 = st.text_input("Confirm password*", type="password", key="su_pw2")

                if st.button("Create account", use_container_width=True, type="primary", key="su_btn"):
                    u = (uname or "").strip().lower()
                    required = [first_name.strip(), last_name.strip(), u, email.strip(), pw1, pw2]
                    # Validate every field before creating the account
                    if not all(required):
                        st.error("All fields marked with (*) are required.")
                    elif not re.match(email_re, email.strip()):
                        st.error("Please enter a valid email address.")
                    elif not re.match(r"^[a-z0-9_.]{3,20}$", u):
                        st.error("Username must be 3-20 characters: only a-z, 0-9, _ and .")
                    elif len(pw1) < 6:
                        st.error("Password must be at least 6 characters.")
                    elif pw1 != pw2:
                        st.error("Passwords do not match.")
                    elif db_email_exists(email.strip()):
                        st.error("An account with this email already exists.")
                    else:
                        full_name = f"{first_name.strip()} {last_name.strip()}"
                        ok = db_create_user(u, pw1, new_role, full_name, email.strip(), None,
                                            first_name=first_name.strip(), last_name=last_name.strip(),
                                            age=int(age), gender=gender)
                        if not ok:
                            st.error("This username is already taken.")
                        else:
                            st.success("Account created! Signing you in…")
                            _start_session(db_authenticate(u, pw1))
                if new_role == "hr":
                    st.caption("After signing in, HR chooses which company to work on.")

            st.caption("Admin accounts can only be created from the Admin portal.")


# ══════════════════════════════════════════════════════════════════════════════
# DATABASE VIEWER (used by the HR and Admin portals)
# ══════════════════════════════════════════════════════════════════════════════
def database_viewer(key_prefix="db", company=None):
    """Browse the main tables with search + CSV download.
    - company=None  (Admin): all tables, all data, plus a full .db backup.
    - company="X"   (HR)   : only the Applications and Jobs of company X.
    The password column is never selected, so it can never be displayed."""
    st.markdown('<span class="sec-title">🗄️ Database</span>', unsafe_allow_html=True)
    if company:
        queries = {
            "Applications": ("SELECT a.app_id,a.job_id,j.title AS job,a.name,a.email,a.phone,a.experience,"
                             "a.category,a.score,a.status,a.date FROM applications a "
                             "JOIN jobs j ON a.job_id=j.job_id WHERE j.company=? ORDER BY a.date DESC",
                             (company,)),
            "Jobs": ("SELECT job_id,title,department,location,type,salary,deadline,posted,status "
                     "FROM jobs WHERE company=? ORDER BY posted DESC", (company,)),
        }
    else:
        queries = {
            "Applications": ("SELECT app_id,job_id,name,email,phone,experience,category,score,status,date "
                             "FROM applications ORDER BY date DESC", ()),
            "Jobs": ("SELECT job_id,title,company,posted_by,department,location,type,salary,deadline,posted,status "
                     "FROM jobs ORDER BY posted DESC", ()),
            "Users": ("SELECT username,first_name,last_name,email,phone,age,gender,role,created "
                      "FROM users ORDER BY created", ()),
            "AI Ratings": ("SELECT session_key,filename,score,rating,recommendation,years_exp,created_at "
                           "FROM ai_ratings", ()),
            "HR Company Access": ("SELECT hr_username,company FROM hr_company_access "
                                  "ORDER BY hr_username", ()),
        }
    choice = st.radio("Table", list(queries), horizontal=True, key=f"{key_prefix}_tbl")
    sql, params = queries[choice]
    conn = get_conn()
    df = pd.read_sql_query(sql, conn, params=params)
    conn.close()
    st.caption(f"📁 {DB_PATH} · {choice}: {len(df)} rows" + (f" · company: {company}" if company else ""))
    table_with_search(df, key=f"{key_prefix}_{choice}")

    d1, d2 = st.columns(2)
    d1.download_button(f"📥 {choice} (CSV)", df.to_csv(index=False),
                       file_name=f"{choice.lower().replace(' ', '_')}.csv",
                       mime="text/csv", use_container_width=True, key=f"{key_prefix}_csv")
    # The full backup contains every company's data, so only the Admin gets it
    if not company and os.path.exists(DB_PATH):
        with open(DB_PATH, "rb") as f:
            d2.download_button("💾 Full database backup (.db)", f.read(),
                               file_name="miu_recruitment.db", use_container_width=True,
                               key=f"{key_prefix}_bak")


# ══════════════════════════════════════════════════════════════════════════════
# PORTAL 1 — ADMIN
# ══════════════════════════════════════════════════════════════════════════════
def admin_portal():
    render_header("System Dashboard")
    render_sidebar()

    tab_dash, tab_users, tab_jobs, tab_db = st.tabs(["📊 Dashboard", "👥 Users", "💼 All Jobs", "🗄️ Database"])

    # ---------------- Dashboard ----------------
    with tab_dash:
        users = db_get_all_users()
        jobs = db_get_all_jobs()
        apps = db_get_all_applications()
        rc = {}                                  # number of users per role
        for u in users:
            rc[u['role']] = rc.get(u['role'], 0) + 1

        c1, c2, c3, c4 = st.columns(4)
        c1.markdown(f'<div class="stat-tile"><div class="stat-num">{len(users)}</div><div class="stat-lbl">Users</div></div>', unsafe_allow_html=True)
        c2.markdown(f'<div class="stat-tile"><div class="stat-num">{len(jobs)}</div><div class="stat-lbl">Jobs</div></div>', unsafe_allow_html=True)
        c3.markdown(f'<div class="stat-tile"><div class="stat-num">{len(apps)}</div><div class="stat-lbl">Applications</div></div>', unsafe_allow_html=True)
        c4.markdown(f'<div class="stat-tile"><div class="stat-num">{rc.get("company", 0)}</div><div class="stat-lbl">Companies</div></div>', unsafe_allow_html=True)

        st.markdown('<span class="sec-title">📈 Users by Role</span>', unsafe_allow_html=True)
        if rc:
            st.bar_chart(pd.DataFrame(list(rc.items()), columns=["Role", "Count"]).set_index("Role"))

        st.markdown('<span class="sec-title">🗄️ Database Info</span>', unsafe_allow_html=True)
        db_size = os.path.getsize(DB_PATH) if os.path.exists(DB_PATH) else 0
        st.info(f"📁 **Database file:** `{DB_PATH}` · **Size:** {db_size / 1024:.1f} KB · "
                f"**Tables:** users, jobs, applications, ai_ratings")

    # ---------------- Users ----------------
    with tab_users:
        st.markdown('<span class="sec-title">👥 All Users</span>', unsafe_allow_html=True)
        users = db_get_all_users()
        rows = [{"Username": u["username"], "First Name": u.get("first_name") or "",
                 "Last Name": u.get("last_name") or "", "Email": u["email"],
                 "Phone": u.get("phone") or "", "Age": u.get("age") or "", "Gender": u.get("gender") or "",
                 "Role": u["role"].upper(), "Created": u["created"]} for u in users]
        # Bordered database table with a search box
        table_with_search(pd.DataFrame(rows), key="users")

        st.markdown("#### ➕ Add User")
        a1, a2 = st.columns(2)
        with a1:
            nu = st.text_input("Username*", key="au_un")
            nn = st.text_input("Full Name*", key="au_nm")
            ne = st.text_input("Email*", key="au_em")
        with a2:
            nr = st.selectbox("Role", ["recruiter", "hr", "company", "admin"], key="au_rl")
            np_ = st.text_input("Password*", type="password", key="au_pw")
            nc = st.text_input("Confirm*", type="password", key="au_cp")
        if st.button("✅ Create User", type="primary", key="au_btn"):
            nu_clean = nu.strip().lower()
            if not all([nu_clean, nn, ne, np_, nc]):
                st.error("All fields required")
            elif np_ != nc:
                st.error("Passwords don't match")
            elif not db_create_user(nu_clean, np_, nr, nn, ne):
                st.error("Username already taken")
            else:
                st.success(f"✅ User '{nu_clean}' created ({nr.upper()})")
                st.rerun()

        st.markdown("#### 🗑️ Delete User")
        del_opts = [u["username"] for u in users if u["username"] != st.session_state.username]
        if del_opts:
            du = st.selectbox("Select user", del_opts, key="du_sel")
            if st.button("Delete", type="primary", key="du_btn"):
                db_delete_user(du)
                st.success(f"Deleted '{du}'")
                st.rerun()

        st.markdown("#### 🔑 Reset Password")
        rp1, rp2, rp3 = st.columns(3)
        with rp1:
            ru = st.selectbox("User", [u["username"] for u in users], key="rp_u")
        with rp2:
            rp = st.text_input("New Password", type="password", key="rp_p")
        with rp3:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Reset", key="rp_btn"):
                if rp:
                    db_update_password(ru, rp)
                    st.success(f"Password updated for '{ru}'")

        st.markdown("#### 🔗 HR ↔ Company Access")
        hr_users = [u["username"] for u in users if u["role"] == "hr"]
        if hr_users:
            hsel = st.selectbox("HR user", hr_users, key="hca_u")
            all_c = db_get_companies()
            cur = [c for c in db_get_assigned_companies(hsel) if c in all_c]
            picked = st.multiselect("Companies this HR can open", all_c, default=cur, key=f"hca_c_{hsel}")
            st.caption("Leave empty = this HR can choose ANY company at sign-in.")
            if st.button("💾 Save access", key="hca_save"):
                db_set_hr_companies(hsel, picked)
                st.success(f"Saved company access for '{hsel}'")
        else:
            st.info("No HR users yet.")

    # ---------------- All jobs ----------------
    with tab_jobs:
        st.markdown('<span class="sec-title">💼 All Jobs</span>', unsafe_allow_html=True)
        jobs = db_get_all_jobs()
        if not jobs:
            st.info("No jobs posted yet.")
        for j in jobs:
            n_apps = db_count_apps_for_job(j["job_id"])
            dot = '🟢' if j['status'] == 'open' else '🔴'
            with st.expander(f"{dot} {j['title']} — {j['company']} | {n_apps} apps"):
                st.markdown(
                    f"**ID:** {j['job_id']} | **Posted by:** {j['posted_by']}  \n"
                    f"**Location:** {j['location']} | **Type:** {j['type']}  \n"
                    f"**Salary:** {j['salary']} | **Deadline:** {j['deadline']}"
                )
                st.write(j['description'])

    # ---------------- Database ----------------
    with tab_db:
        database_viewer("adm")

    st.markdown('<div class="footer">🎓 MIU | AI Recruitment Portal | SQLite Database</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# PORTAL 2 — HR
# ══════════════════════════════════════════════════════════════════════════════
# ── Top-5 % logic ────────────────────────────────────────────────────────────
TOP_PERCENT = 5          # share of applicants shown automatically to HR


def _score_num(a):
    """Numeric score of an application; -1 if it is still 'Pending' (not scored)."""
    try:
        return int(float(a.get("score")))
    except (TypeError, ValueError):
        return -1


def top_percent_apps(apps, pct=TOP_PERCENT):
    """Return the best `pct` % of the scored applications (highest score first).
    100 applicants -> top 5, 40 applicants -> top 2, and always at least 1."""
    scored = [a for a in apps if _score_num(a) >= 0]
    if not scored:
        return []
    scored.sort(key=_score_num, reverse=True)
    n = max(1, math.ceil(len(scored) * pct / 100))
    return scored[:n]


def hr_top_tab(company):
    """Automatic view: HR sees the top 5 % applicants of the SELECTED COMPANY only,
    without uploading anything. Scores were already calculated when each candidate applied."""
    st.markdown(f'<span class="sec-title">🏆 Top {TOP_PERCENT}% Candidates — {html.escape(company)}</span>',
                unsafe_allow_html=True)
    apps = db_get_apps_for_company(company)          # only this company's applications
    if not apps:
        st.info(f"No applications have been submitted to {company} yet.")
        return

    # Optional filter: the 5 % is then calculated per job instead of over all jobs
    jobs = {f"{a['job_title']} — {a['company']} ({a['job_id']})": a["job_id"] for a in apps}
    sel = st.selectbox("Select job", ["All jobs of this company"] + sorted(jobs), key=f"top_job_{company}")
    pool = apps if sel == "All jobs of this company" else [a for a in apps if a["job_id"] == jobs[sel]]
    top = top_percent_apps(pool)

    if not top:
        st.info("There is no scored application in this selection yet.")
        return

    # Summary tiles
    c1, c2, c3 = st.columns(3)
    c1.markdown(f'<div class="stat-tile"><div class="stat-num">{len(pool)}</div><div class="stat-lbl">Total Applicants</div></div>', unsafe_allow_html=True)
    c2.markdown(f'<div class="stat-tile"><div class="stat-num">{len(top)}</div><div class="stat-lbl">Top {TOP_PERCENT}% Shown</div></div>', unsafe_allow_html=True)
    c3.markdown(f'<div class="stat-tile"><div class="stat-num">{_score_num(top[-1])}</div><div class="stat-lbl">Cut-off Score</div></div>', unsafe_allow_html=True)

    # Interview settings come from the sidebar
    interviewer = st.session_state.get("int_interviewer", display_name(st.session_state.user_info))
    idate = st.session_state.get("int_date", "Monday, May 20, 2026 at 3 PM")
    izoom = st.session_state.get("int_zoom", "https://zoom.us/j/your-id")
    icompany = st.session_state.get("int_company", "MIU")

    for rank, a in enumerate(top, 1):
        sc_val = _score_num(a)
        col = score_color(sc_val)
        st.markdown(f"""
<div class="card" style="border-left:5px solid {col};">
  <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;">
    <div>
      <div style="font-size:1.05rem;font-weight:700;color:#003366;">#{rank} &nbsp;👤 {html.escape(a["name"] or "")}</div>
      <div style="font-size:.82rem;color:#555;">💼 {html.escape(a["job_title"])} — {html.escape(a["company"])}</div>
      <div style="font-size:.8rem;color:#555;">✉️ {html.escape(a["email"] or "")} &nbsp;|&nbsp; 📞 {html.escape(a["phone"] or "—")}
      &nbsp;|&nbsp; Status: <b>{html.escape(a["status"])}</b></div>
      <span class="cat-badge">🏷️ {html.escape(a["category"] or "General")}</span>
    </div>
    <div class="score-ring" style="background:{col};"><span class="num">{sc_val}</span><span class="sub">/100</span></div>
  </div>
</div>""", unsafe_allow_html=True)

        with st.expander(f"📊 Details & Interview — {a['name']}"):
            if a.get("score_detail"):
                render_score_breakdown(json.loads(a["score_detail"]))
            if a.get("cover"):
                st.markdown(f"**Cover letter:** {a['cover']}")

            # Interview invitation texts (e-mail + WhatsApp)
            cname = a["name"] or "Candidate"
            subj = f"Interview Invitation — {a['job_title']} | {icompany}"
            body = (f"Dear {cname},\n\nWe reviewed your application and are pleased to shortlist you "
                    f"for the {a['job_title']} position at {icompany}.\n\n"
                    f"Date & Time : {idate}\nZoom Link   : {izoom}\n\nPlease confirm.\n\n"
                    f"Best regards,\n{interviewer}\n{icompany}")
            wa = (f"Hello {cname}! 👋\n\nI am *{interviewer}* from *{icompany}*. "
                  f"You are shortlisted for *{a['job_title']}*.\n\n📅 {idate}\n🔗 {izoom}\n\nPlease confirm. 🙏")

            b1, b2, b3 = st.columns(3)
            if a["email"]:
                # mailto needs '?' and spaces must be %20 (not '+')
                params = urllib.parse.urlencode({"subject": subj, "body": body}, quote_via=urllib.parse.quote)
                b1.markdown(f'<a class="btn-email" href="mailto:{urllib.parse.quote(a["email"])}?{params}" target="_blank">✉️ Email</a>', unsafe_allow_html=True)
            if a["phone"]:
                phone_d = re.sub(r"\D", "", a["phone"])      # digits only for wa.me
                b2.markdown(f'<a class="btn-whatsapp" href="https://wa.me/{phone_d}?text={urllib.parse.quote(wa)}" target="_blank">💬 WhatsApp</a>', unsafe_allow_html=True)
            zh = izoom if izoom.startswith("https://") else "https://zoom.us"
            b3.markdown(f'<a class="btn-zoom" href="{zh}" target="_blank">📹 Zoom</a>', unsafe_allow_html=True)

            if a["status"] != "Shortlisted":
                if st.button("✅ Mark as Shortlisted", key=f"sl_{a['app_id']}"):
                    db_update_app_status(a["app_id"], "Shortlisted")
                    st.rerun()


# ── Manual screening tab (optional: HR uploads resumes by hand) ──────────────
def hr_screen_tab():
    """Resume screening tab (wrapped in a function so `return` can be used —
    st.stop() used to block the other tabs as well)."""
    st.markdown('<span class="sec-title">📂 Upload & Screen Resumes</span>', unsafe_allow_html=True)
    api_key = st.session_state.get("claude_api_key", "")
    enable_ai = bool(api_key)
    sess_key = st.session_state.get("username", "anon")

    uploaded = st.file_uploader("Drop resumes (PDF · DOCX · TXT)",
                                type=["pdf", "docx", "txt"], accept_multiple_files=True)

    if not uploaded:
        st.info("👆 Upload resumes to begin.")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("""
**Steps:**
1. Upload resumes → auto-classified
2. AI rates with Claude (API key in sidebar)
3. Ratings saved to **SQLite** — persist across sessions
4. Send interview invites
5. Export CSV
""")
        with c2:
            st.markdown("""
**Ratings:**
- 🟢 Excellent 80-100
- 🔵 Good 60-79
- 🟡 Average 40-59
- 🔴 Poor 0-39
""")
        return

    # Read every uploaded file and predict its category
    results, errors, raw_texts = [], [], {}
    prog = st.progress(0, "Analysing…")
    for i, f in enumerate(uploaded):
        prog.progress((i + 1) / len(uploaded), f"Reading {f.name}")
        try:
            text = extract_text(f)
            cat = predict_category(text)
            raw_texts[f.name] = text
            results.append({"filename": f.name, "category": cat, "preview": text[:500].strip()})
        except Exception as e:
            errors.append({"file": f.name, "error": str(e)})
    prog.empty()
    if not results:
        for e in errors:
            st.error(f"⚠️ {e['file']}: {e['error']}")
        return

    df = pd.DataFrame(results)
    cat_cnts = df["category"].value_counts()

    # Dashboard
    st.markdown('<span class="sec-title">📊 Dashboard</span>', unsafe_allow_html=True)
    cols = st.columns(min(len(cat_cnts) + 1, 5))
    cols[0].markdown(f'<div class="stat-tile"><div class="stat-num">{len(df)}</div><div class="stat-lbl">Resumes</div></div>', unsafe_allow_html=True)
    for i, (cat, cnt) in enumerate(cat_cnts.items()):
        if i + 1 < len(cols):
            cols[i + 1].markdown(f'<div class="stat-tile"><div class="stat-num">{cnt}</div><div class="stat-lbl">{cat[:13]}</div></div>', unsafe_allow_html=True)
    st.bar_chart(cat_cnts.rename_axis("Category").reset_index(name="Count").set_index("Category"))

    # Ratings: saved ones from the DB + ones created in this session
    ai_ratings = db_get_all_ratings_for_session(sess_key)
    if "fresh_ratings" not in st.session_state:
        st.session_state.fresh_ratings = {}
    ai_ratings.update(st.session_state.fresh_ratings)

    def _rate_and_store(fn, text, category):
        r = rate_resume_with_ai(text, category, api_key)
        if not r.get("error") or r.get("score", 0) > 0:
            db_save_rating(sess_key, fn, r)      # save only successful ratings to the DB
        st.session_state.fresh_ratings[fn] = r
        return r

    if enable_ai:
        if st.button("🤖 Rate ALL Resumes with AI", type="primary", key="rate_all"):
            with st.spinner("Claude evaluating all resumes…"):
                for row in results:
                    fn = row["filename"]
                    _rate_and_store(fn, raw_texts.get(fn, row["preview"]), row["category"])
            st.success("✅ All resumes rated!")
            st.rerun()
    else:
        st.warning("⚠️ Add Claude API key in sidebar to enable AI scoring.")

    # Sort the best-rated resumes first
    display_df = df.copy()
    if ai_ratings:
        display_df["_score"] = display_df["filename"].map(
            lambda fn: ai_ratings.get(fn, {}).get("score", 0) or 0)
        display_df = display_df.sort_values("_score", ascending=False)

    st.markdown('<span class="sec-title">📋 Candidate Profiles</span>', unsafe_allow_html=True)

    saved_in_db = set(db_get_all_ratings_for_session(sess_key).keys())

    for idx, row in display_df.iterrows():
        fn = row["filename"]
        category = row["category"]
        preview = row["preview"]
        rating = ai_ratings.get(fn, {})
        score = rating.get("score", None)
        sc = score_color(score) if score else "#888"
        rc_cls = rating_class(rating.get("rating", ""))

        score_html = ""
        if score:
            score_html = (f'<div class="score-ring" style="background:{sc};">'
                          f'<span class="num">{score}</span><span class="sub">/100</span></div>'
                          f'<span class="{rc_cls}" style="margin-left:8px;">{rating.get("rating", "")}</span>')
        db_tag = '<span class="db-badge" style="margin-left:8px;">🗄️ Saved</span>' if fn in saved_in_db else ""

        st.markdown(f"""
<div class="card">
  <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;">
    <div>
      <div style="font-size:1rem;font-weight:700;color:#003366;">📄 {fn} {db_tag}</div>
      <span class="cat-badge">🏷️ {category}</span>
    </div>
    <div style="display:flex;align-items:center;">{score_html}</div>
  </div>
</div>""", unsafe_allow_html=True)

        with st.expander(f"👤 Profile — {fn}"):
            if rating and score and score > 0:
                strengths = rating.get("strengths", [])
                weaknesses = rating.get("weaknesses", [])
                skills = rating.get("top_skills", [])
                years = rating.get("years_exp", "N/A")
                rec = rating.get("recommendation", "")
                skills_html = "".join(f'<span class="skill-chip">{s}</span>' for s in skills)

                st.markdown(f"""
<div class="card" style="border-left:4px solid {sc};">
  <div style="display:flex;align-items:center;margin-bottom:1rem;flex-wrap:wrap;gap:12px;">
    <div class="score-ring" style="background:{sc};">
      <span class="num">{score}</span><span class="sub">/100</span>
    </div>
    <div>
      <div style="font-size:1rem;font-weight:700;color:#003366;">AI Evaluation</div>
      <span class="{rc_cls}">{rating.get("rating", "")}</span>
      <div style="font-size:.8rem;color:#555;margin-top:4px;">🕐 Experience: <b>{years}</b></div>
    </div>
  </div>
  <div style="background:#f5f5f5;border-radius:8px;height:10px;overflow:hidden;margin-bottom:6px;">
    <div style="background:{sc};height:100%;width:{score}%;border-radius:8px;"></div>
  </div>
  <div style="text-align:right;font-size:.75rem;color:#888;">{score}/100</div>
</div>""", unsafe_allow_html=True)

                sc_col, wk_col = st.columns(2)
                with sc_col:
                    items = "".join(f'<div style="margin:4px 0;font-size:.83rem;">✅ {s}</div>' for s in strengths)
                    st.markdown(f'<div style="background:#f0fff4;border:1px solid #c8e6c9;border-radius:12px;padding:1rem;"><div style="font-weight:700;color:#1b5e20;margin-bottom:6px;font-size:.88rem;">💪 Strengths</div>{items}</div>', unsafe_allow_html=True)
                with wk_col:
                    items = "".join(f'<div style="margin:4px 0;font-size:.83rem;">⚠️ {w}</div>' for w in weaknesses)
                    st.markdown(f'<div style="background:#fff8f0;border:1px solid #ffe0b2;border-radius:12px;padding:1rem;"><div style="font-weight:700;color:#bf360c;margin-bottom:6px;font-size:.88rem;">📈 Areas to Improve</div>{items}</div>', unsafe_allow_html=True)

                st.markdown(f"""
<div style="margin-top:.7rem;padding:.7rem 1rem;background:#e3f2fd;border-radius:12px;">
  <b style="color:#0d47a1;font-size:.83rem;">🔧 Top Skills: </b>{skills_html}
</div>
<div style="margin-top:.6rem;padding:.8rem 1rem;background:#e8f5e9;border-radius:12px;">
  <b style="color:#1b5e20;font-size:.83rem;">💡 Recommendation: </b>
  <span style="color:#1b4332;font-size:.83rem;">{rec}</span>
</div><br/>""", unsafe_allow_html=True)

            elif rating and rating.get("error"):
                err = rating.get("error", "")
                if "bad_api_key" in err:
                    st.error("❌ Invalid API key — console.anthropic.com → API Keys")
                elif "rate_limit" in err:
                    st.warning("⏳ Rate limited — wait 30 sec")
                else:
                    st.error(f"❌ {err[:200]}")

            if enable_ai and not (rating and score and score > 0):
                if st.button("🤖 Rate This Resume", key=f"rate_one_{idx}"):
                    with st.spinner("Rating…"):
                        _rate_and_store(fn, raw_texts.get(fn, preview), category)
                    st.rerun()

            st.text_area("📃 Resume Preview", preview + "…", height=120, key=f"prev_{idx}")

            # Contact section — name/email/phone are auto-filled from the resume
            st.markdown("---")
            st.markdown("#### 📬 Send Interview Invitation")
            _ext = extract_candidate_info(raw_texts.get(fn, preview))
            for _k, _v in [(f"cn_{idx}", _ext["name"]),
                           (f"ce_{idx}", _ext["email"]),
                           (f"cp_{idx}", _ext["phone"])]:
                if _k not in st.session_state and _v:
                    st.session_state[_k] = _v

            if any([_ext["name"], _ext["email"], _ext["phone"]]):
                st.info(f"✨ Auto-filled — Name: **{_ext['name'] or '—'}** · "
                        f"Email: **{_ext['email'] or '—'}** · "
                        f"Phone: **{_ext['phone'] or '—'}**")

            cc1, cc2, cc3 = st.columns(3)
            cn = cc1.text_input("Candidate Name", key=f"cn_{idx}", placeholder="Ahmed Khan")
            cem = cc2.text_input("Email", key=f"ce_{idx}", placeholder="candidate@email.com")
            cph = cc3.text_input("WhatsApp Number", key=f"cp_{idx}", placeholder="+923001234567")

            interviewer = st.session_state.get("int_interviewer", display_name(st.session_state.user_info))
            idate = st.session_state.get("int_date", "Monday, May 20, 2026 at 3 PM")
            izoom = st.session_state.get("int_zoom", "https://zoom.us/j/your-id")
            icompany = st.session_state.get("int_company", "MIU")
            cname = cn or "Candidate"

            subj = f"Interview Invitation — {category} | {icompany}"
            body = (f"Dear {cname},\n\nWe reviewed your resume and are pleased to shortlist you "
                    f"for the {category} position at {icompany}.\n\n"
                    f"Date & Time : {idate}\nZoom Link   : {izoom}\n\n"
                    f"Please confirm.\n\nBest regards,\n{interviewer}\n{icompany}")
            wa = (f"Hello {cname}! 👋\n\nI am *{interviewer}* from *{icompany}*. "
                  f"You are shortlisted for *{category}*.\n\n📅 {idate}\n🔗 {izoom}\n\n"
                  f"Please confirm. Thank you! 🙏")

            b1, b2, b3 = st.columns(3)
            with b1:
                if cem:
                    # mailto needs '?' and spaces must be %20 (not '+')
                    params = urllib.parse.urlencode({"subject": subj, "body": body},
                                                    quote_via=urllib.parse.quote)
                    st.markdown(f'<a class="btn-email" href="mailto:{urllib.parse.quote(cem)}?{params}" target="_blank">✉️ Email</a>', unsafe_allow_html=True)
                else:
                    st.caption("Enter email")
            with b2:
                if cph:
                    phone_d = re.sub(r'\D', '', cph)
                    st.markdown(f'<a class="btn-whatsapp" href="https://wa.me/{phone_d}?text={urllib.parse.quote(wa)}" target="_blank">💬 WhatsApp</a>', unsafe_allow_html=True)
                else:
                    st.caption("Enter phone")
            with b3:
                zh = izoom if izoom.startswith("https://") else "https://zoom.us"
                st.markdown(f'<a class="btn-zoom" href="{zh}" target="_blank">📹 Zoom</a>', unsafe_allow_html=True)

    # Export the screening result as CSV
    st.markdown('<span class="sec-title">⬇️ Export</span>', unsafe_allow_html=True)
    exp_df = df[["filename", "category"]].copy()
    if ai_ratings:
        exp_df["score"] = exp_df["filename"].map(lambda f: ai_ratings.get(f, {}).get("score", ""))
        exp_df["rating"] = exp_df["filename"].map(lambda f: ai_ratings.get(f, {}).get("rating", ""))
        exp_df["rec"] = exp_df["filename"].map(lambda f: ai_ratings.get(f, {}).get("recommendation", ""))
        exp_df["skills"] = exp_df["filename"].map(lambda f: ", ".join(ai_ratings.get(f, {}).get("top_skills", [])))
    st.download_button("📥 Download CSV",
                       data=exp_df.to_csv(index=False),
                       file_name=f"screening_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                       mime="text/csv", use_container_width=True)
    for e in errors:
        st.caption(f"⚠️ {e['file']}: {e['error']}")


def show_company_picker():
    """Shown right after an HR user signs in: HR chooses which company to work on.
    Everything in the HR portal is then limited to that company."""
    render_header("Select Company")
    user = st.session_state.user_info
    companies = db_get_hr_companies(user["username"])      # assigned companies, or all
    _, mid, _ = st.columns([1, 1.6, 1])
    with mid:
        st.markdown('<div class="auth-form-title" style="margin-top:1rem;">Choose a company</div>'
                    '<div class="auth-form-sub">You will only see resumes of jobs posted by this company.</div>',
                    unsafe_allow_html=True)
        if not companies:
            st.warning("No company has posted a job yet. Please try again later.")
        else:
            choice = st.selectbox("Company", companies, key="hr_company_pick")
            if st.button("Continue →", type="primary", use_container_width=True, key="hr_company_go"):
                st.session_state.hr_company = choice
                st.rerun()
        if st.button("🚪 Logout", use_container_width=True, key="hr_pick_logout"):
            for k in ['logged_in', 'username', 'user_info', 'login_time',
                      'claude_api_key', 'fresh_ratings', 'hr_company']:
                st.session_state.pop(k, None)
            st.rerun()


def hr_portal():
    """HR portal for ONE company: Top 5 % · Applications · Jobs · Database · Manual screening."""
    company = st.session_state.hr_company            # chosen in show_company_picker()
    render_header(f"{company} — Top Candidates")
    render_sidebar()

    tab_top, tab_apps, tab_jobs, tab_db, tab_screen = st.tabs([
        f"🏆 Top {TOP_PERCENT}%", "📋 Applications", "💼 Jobs", "🗄️ Database", "📂 Manual Screening"])

    # Opens first: the best applicants appear automatically, nothing needs to be uploaded
    with tab_top:
        hr_top_tab(company)

    # All applications from the database, with the score breakdown of each
    with tab_apps:
        st.markdown('<span class="sec-title">📋 Applications — all (from DB)</span>', unsafe_allow_html=True)
        apps = db_get_apps_for_company(company)       # only this company's applications
        if not apps:
            st.info("No applications yet.")
        else:
            rows = [{"Job": a["job_title"], "Company": a["company"], "Applicant": a["name"],
                     "Email": a["email"], "Category": a["category"], "Score": a["score"],
                     "Status": a["status"], "Applied": a["date"]} for a in apps]
            table_with_search(pd.DataFrame(rows), key="hr_apps")
            st.caption(f"Total: {len(apps)} applications in database")

            st.markdown('<span class="sec-title">📊 Resume Score Breakdown</span>', unsafe_allow_html=True)
            for a in apps:
                if a.get("score_detail"):
                    with st.expander(f"📊 {a['name']} — {a['job_title']} — {a['score']}/100"):
                        render_score_breakdown(json.loads(a["score_detail"]))

    # Open jobs
    with tab_jobs:
        st.markdown('<span class="sec-title">💼 Active Jobs</span>', unsafe_allow_html=True)
        for j in [x for x in db_get_jobs_for_company(company) if x["status"] == "open"]:
            n = db_count_apps_for_job(j["job_id"])
            st.markdown(f"""
<div class="job-card">
  <div class="job-title">💼 {j["title"]}</div>
  <div class="job-company">🏢 {j["company"]} | 📍 {j["location"]}</div>
  <div style="margin-top:6px;">
    <span class="job-tag">{j["type"]}</span>
    <span class="job-tag job-tag-green">{j["department"]}</span>
    <span class="job-tag job-tag-orange">💰 {j["salary"]}</span>
  </div>
  <div style="margin-top:6px;font-size:.78rem;color:#888;">
    📅 Deadline: {j["deadline"]} | 📨 {n} applications
  </div>
</div>""", unsafe_allow_html=True)

    # Database tables (view, search, export)
    with tab_db:
        database_viewer("hrdb", company)

    # Optional: HR can still upload resumes manually
    with tab_screen:
        hr_screen_tab()

    st.markdown('<div class="footer">🎓 MIU | AI Recruitment Portal | SQLite DB</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PORTAL 3 — COMPANY
# ══════════════════════════════════════════════════════════════════════════════
def company_portal():
    """Company portal: profile · post a job · my jobs · applications to my jobs."""
    render_header("Company Dashboard")
    render_sidebar()

    username = st.session_state.username
    user = db_get_user(username) or st.session_state.user_info
    ci = json.loads(user.get("company_info", "{}") or "{}")      # company profile (JSON)

    tab_profile, tab_post, tab_myjobs, tab_apps = st.tabs([
        "🏢 Profile", "➕ Post Job", "📋 My Jobs", "👥 Applications"
    ])

    # ---------------- Profile ----------------
    with tab_profile:
        st.markdown('<span class="sec-title">🏢 Company Profile</span>', unsafe_allow_html=True)
        st.markdown(f"""
<div class="card" style="border-left:5px solid #6610f2;">
  <div style="font-size:1.3rem;font-weight:700;color:#003366;">{ci.get("name", "Your Company")}</div>
  <div style="color:#555;margin:6px 0;">
    🏭 {ci.get("industry", "—")} &nbsp;|&nbsp;
    👥 {ci.get("size", "—")} &nbsp;|&nbsp;
    🌐 {ci.get("website", "—")} &nbsp;|&nbsp;
    📞 {ci.get("phone") or user.get("phone") or "—"} &nbsp;|&nbsp;
    ✉️ {html.escape(user.get("email") or "—")}
  </div>
  <div style="color:#444;font-size:.87rem;">{ci.get("about", "No description yet.")}</div>
</div>""", unsafe_allow_html=True)

        st.markdown("#### ✏️ Update Profile")
        p1, p2 = st.columns(2)
        sizes = ["1-50", "51-200", "201-500", "500-1000", "1000+"]
        with p1:
            nm = st.text_input("Company Name", ci.get("name", ""), key="cp_nm")
            ind = st.text_input("Industry", ci.get("industry", ""), key="cp_in")
            csz = st.selectbox("Size", sizes,
                               index=sizes.index(ci.get("size", "1-50")) if ci.get("size") in sizes else 0,
                               key="cp_sz")
        with p2:
            web = st.text_input("Website", ci.get("website", ""), key="cp_wb")
            ph = st.text_input("Phone number", ci.get("phone") or user.get("phone") or "", key="cp_ph")
            abt = st.text_area("About", ci.get("about", ""), height=120, key="cp_ab")

        if st.button("💾 Save Profile", type="primary", key="cp_save"):
            db_update_company_info(username, {"name": nm, "industry": ind, "size": csz,
                                              "website": web, "about": abt,
                                              "phone": ph, "email": user.get("email", "")})
            db_update_phone(username, ph)
            st.success("✅ Profile saved to database!")
            st.session_state.user_info = db_get_user(username)
            st.rerun()

    # ---------------- Post a job ----------------
    with tab_post:
        st.markdown('<span class="sec-title">➕ Post a Job</span>', unsafe_allow_html=True)
        p1, p2 = st.columns(2)
        with p1:
            jt = st.text_input("Job Title*", placeholder="Senior Python Developer", key="jt")
            jd = st.text_input("Department*", placeholder="Engineering", key="jd_dept")
            jl = st.text_input("Location*", placeholder="Islamabad / Remote", key="jl")
            jty = st.selectbox("Type", ["Full-time", "Part-time", "Contract", "Internship"], key="jty")
        with p2:
            js = st.text_input("Salary", placeholder="PKR 100,000-200,000", key="js")
            jdl = st.date_input("Deadline", key="jdl")
            jrq = st.text_area("Requirements*", placeholder="3+ years Python, Django", height=100, key="jrq")
        jdes = st.text_area("Description*", placeholder="Detailed job description…", height=150, key="jdes")

        if st.button("🚀 Publish Job", type="primary", key="pub_job"):
            if not all([jt.strip(), jd.strip(), jl.strip(), jdes.strip(), jrq.strip()]):
                st.error("Fill all required fields (*)")
            else:
                jid = f"JOB{int(time.time())}"
                db_create_job(jid, jt.strip(), ci.get("name") or user["name"],
                              username, jd.strip(), jl.strip(), jty,
                              js.strip() or "Competitive", jdes.strip(), jrq.strip(),
                              str(jdl), datetime.now().strftime("%Y-%m-%d"))
                st.success(f"✅ Job '{jt}' published! ID: {jid}")
                st.balloons()

    # ---------------- My jobs ----------------
    with tab_myjobs:
        st.markdown('<span class="sec-title">📋 My Jobs</span>', unsafe_allow_html=True)
        my_jobs = db_get_jobs_by_user(username)
        if not my_jobs:
            st.info("No jobs posted yet. Use 'Post Job' tab.")
        for j in my_jobs:
            n = db_count_apps_for_job(j["job_id"])
            sc = "#28a745" if j["status"] == "open" else "#dc3545"
            st.markdown(f"""
<div class="job-card">
  <div style="display:flex;justify-content:space-between;flex-wrap:wrap;">
    <div>
      <div class="job-title">💼 {j["title"]}</div>
      <div class="job-company">📍 {j["location"]} | 🏢 {j["department"]}</div>
      <span class="job-tag">{j["type"]}</span>
      <span class="job-tag job-tag-orange">💰 {j["salary"]}</span>
    </div>
    <div style="text-align:right;">
      <div style="font-size:1.3rem;font-weight:900;color:#003366;">{n}</div>
      <div style="font-size:.72rem;color:#666;">Applications</div>
      <span style="background:{sc};color:white;border-radius:12px;padding:2px 10px;
                   font-size:.72rem;font-weight:700;">{j["status"].upper()}</span>
    </div>
  </div>
  <div style="font-size:.78rem;color:#888;margin-top:6px;">
    ID: {j["job_id"]} | Posted: {j["posted"]} | Deadline: {j["deadline"]}
  </div>
</div>""", unsafe_allow_html=True)
            # Close / reopen the job
            ns = "closed" if j["status"] == "open" else "open"
            btn = "🔴 Close" if j["status"] == "open" else "🟢 Reopen"
            if st.button(btn, key=f"tog_{j['job_id']}"):
                db_toggle_job_status(j["job_id"], ns)
                st.rerun()

    # ---------------- Applications to my jobs ----------------
    with tab_apps:
        st.markdown('<span class="sec-title">👥 Applications to My Jobs</span>', unsafe_allow_html=True)
        my_jobs = db_get_jobs_by_user(username)
        found = False
        for job in my_jobs:
            apps = db_get_apps_for_job(job["job_id"])
            if apps:
                found = True
                st.markdown(f"**💼 {job['title']}** ({len(apps)} applications)")
                rows = [{"Name": a["name"], "Email": a["email"], "Category": a["category"],
                         "Score": a["score"], "Status": a["status"], "Applied": a["date"]}
                        for a in apps]
                boxed_table(pd.DataFrame(rows))
                for a in apps:
                    if a.get("score_detail"):
                        with st.expander(f"📊 {a['name']} — {a['score']}/100 score breakdown"):
                            render_score_breakdown(json.loads(a["score_detail"]))
        if not found:
            st.info("No applications yet. Job seekers apply from the Job Seeker Portal.")

    st.markdown('<div class="footer">🎓 MIU | AI Recruitment Portal</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# PORTAL 4 — RECRUITER / JOB SEEKER
# ══════════════════════════════════════════════════════════════════════════════
def recruiter_portal():
    """Job seeker portal: job board (with search) · apply · my applications."""
    render_header("Job Board & Applications")
    render_sidebar()

    username = st.session_state.username

    tab_jobs, tab_apply, tab_myapps = st.tabs(["💼 Job Board", "📤 Apply", "📋 My Applications"])

    # ---------------- Job board ----------------
    with tab_jobs:
        st.markdown('<span class="sec-title">💼 Open Jobs</span>', unsafe_allow_html=True)
        jobs = db_get_open_jobs()
        if not jobs:
            st.info("No open jobs at the moment.")
        else:
            c1, c2, c3 = st.columns(3)
            c1.markdown(f'<div class="stat-tile"><div class="stat-num">{len(jobs)}</div><div class="stat-lbl">Open Positions</div></div>', unsafe_allow_html=True)
            c2.markdown(f'<div class="stat-tile"><div class="stat-num">{len(set(j["company"] for j in jobs))}</div><div class="stat-lbl">Companies</div></div>', unsafe_allow_html=True)
            total_a = sum(db_count_apps_for_job(j["job_id"]) for j in jobs)
            c3.markdown(f'<div class="stat-tile"><div class="stat-num">{total_a}</div><div class="stat-lbl">Total Applicants</div></div>', unsafe_allow_html=True)

            # Search across title, company, location, department, type, salary,
            # description and requirements — only matching jobs are shown.
            search = st.text_input("🔍 Search jobs", "", key="job_search",
                                   placeholder="title, company, city, skill, salary… (press Enter)")
            if search:
                jobs = filter_jobs(jobs, search)
                st.caption(f"{len(jobs)} job(s) match “{search}”")
                if not jobs:
                    st.warning("No jobs match your search.")

            for j in jobs:
                n = db_count_apps_for_job(j["job_id"])
                st.markdown(f"""
<div class="job-card">
  <div style="display:flex;justify-content:space-between;flex-wrap:wrap;gap:10px;">
    <div style="flex:1;">
      <div class="job-title">💼 {j["title"]}</div>
      <div class="job-company">🏢 {j["company"]} | 📍 {j["location"]}</div>
      <div style="margin-top:6px;">
        <span class="job-tag">{j["type"]}</span>
        <span class="job-tag job-tag-green">{j["department"]}</span>
        <span class="job-tag job-tag-orange">💰 {j["salary"]}</span>
      </div>
      <div style="font-size:.83rem;color:#444;margin-top:8px;">{(j["description"] or "")[:150]}…</div>
      <div style="font-size:.78rem;color:#888;margin-top:6px;">
        📅 Deadline: <b>{j["deadline"]}</b> | 🔧 {(j["requirements"] or "")[:80]}…
      </div>
    </div>
    <div style="text-align:right;flex-shrink:0;">
      <div style="font-size:1.4rem;font-weight:900;color:#003366;">{n}</div>
      <div style="font-size:.72rem;color:#666;">Applicants</div>
    </div>
  </div>
</div>""", unsafe_allow_html=True)

    # ---------------- Apply ----------------
    with tab_apply:
        st.markdown('<span class="sec-title">📤 Apply for a Job</span>', unsafe_allow_html=True)
        jobs = db_get_open_jobs()
        if not jobs:
            st.info("No open jobs currently.")
        else:
            opts = {f"{j['title']} — {j['company']} ({j['job_id']})": j for j in jobs}
            sel_label = st.selectbox("Choose job:", list(opts.keys()), key="sel_job")
            sel_job = opts[sel_label]

            st.markdown(f"""
<div class="job-card" style="border-left:5px solid #0d6efd;margin-bottom:1rem;">
  <b>{sel_job["title"]}</b> at {sel_job["company"]}<br/>
  <span style="font-size:.8rem;color:#555;">
    📍 {sel_job["location"]} | 💰 {sel_job["salary"]} | 📅 {sel_job["deadline"]}
  </span><br/>
  <span style="font-size:.8rem;color:#333;">{sel_job["requirements"]}</span>
</div>""", unsafe_allow_html=True)

            st.markdown("#### 📎 Upload Your Resume First")
            resume_file = st.file_uploader("Resume (PDF · DOCX · TXT)",
                                           type=["pdf", "docx", "txt"], key="ap_res")

            # Auto-fill name / e-mail / phone when a resume is uploaded
            if resume_file:
                try:
                    _rf_text = extract_text(resume_file)
                    _rf_info = extract_candidate_info(_rf_text)
                    for _k, _v in [("ap_nm", _rf_info["name"]),
                                   ("ap_em", _rf_info["email"]),
                                   ("ap_ph", _rf_info["phone"])]:
                        if not st.session_state.get(_k) and _v:
                            st.session_state[_k] = _v
                    if any(_rf_info.values()):
                        st.info(f"✨ Auto-filled — Name: **{_rf_info['name'] or '—'}** · "
                                f"Email: **{_rf_info['email'] or '—'}** · "
                                f"Phone: **{_rf_info['phone'] or '—'}**")
                except Exception:
                    pass

            st.markdown("#### 📋 Your Details")
            a1, a2 = st.columns(2)
            with a1:
                ap_nm = st.text_input("Full Name*", key="ap_nm", placeholder="Your Name")
                ap_em = st.text_input("Email*", key="ap_em", placeholder="you@email.com")
            with a2:
                ap_ph = st.text_input("Phone/WhatsApp", key="ap_ph", placeholder="+92 300 1234567")
                ap_exp = st.text_input("Years Experience", key="ap_exp", placeholder="e.g. 3")
            ap_cv = st.text_area("Cover Letter (optional)", key="ap_cv",
                                 placeholder="Brief intro and why you're a great fit…", height=90)

            if st.button("🚀 Submit Application", type="primary", key="submit_app"):
                if not all([ap_nm.strip(), ap_em.strip(), resume_file]):
                    st.error("Name, Email and Resume are required.")
                else:
                    # Score the resume against this job using the datasets
                    # (skills, experience, education, certifications, completeness).
                    # The score is saved with the application, so HR sees the
                    # top 5 % automatically without uploading anything.
                    with st.spinner("Scoring your resume…"):
                        result = None
                        try:
                            text = extract_text(resume_file)
                            cat = predict_category(text)
                            result = score_resume(text, sel_job, cat)
                        except Exception:
                            cat = "General"
                    db_create_application(
                        f"APP{int(time.time())}", sel_job["job_id"], username,
                        ap_nm.strip(), ap_em.strip(), ap_ph, ap_exp, ap_cv,
                        cat, resume_file.name,
                        score=str(result["score"]) if result else "Pending",
                        score_detail=json.dumps(result) if result else "")
                    st.success(f"✅ Application submitted for **{sel_job['title']}** at **{sel_job['company']}**!")
                    st.balloons()
                    st.info(f"📂 Resume classified as: **{cat}** · Saved to database. "
                            "You will be contacted if shortlisted. Thank you! 🙏")
                    if result:
                        st.metric("Your resume match score", f"{result['score']}/100", result["rating"])
                        render_score_breakdown(result)

    # ---------------- My applications ----------------
    with tab_myapps:
        st.markdown('<span class="sec-title">📋 My Applications</span>', unsafe_allow_html=True)
        my_apps = db_get_apps_by_recruiter(username)
        if not my_apps:
            st.info("No applications yet.")
        else:
            st.markdown(f"**Total: {len(my_apps)} applications**")
            rows = [{"Job": a["job_title"], "Company": a["company"], "Applied": a["date"],
                     "Category": a["category"], "Score": a["score"], "Status": a["status"]}
                    for a in my_apps]
            boxed_table(pd.DataFrame(rows))

    st.markdown('<div class="footer">🎓 MIU | AI Recruitment Portal</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════
# NEW SCORING SYSTEM + HR AUTO-SCORING  (paste this whole block ONCE)
# Place: just ABOVE the line  "# ENTRY POINT"  at the very bottom of the file
#        (above  enable_enter_to_next() ).
# These definitions replace the older WEIGHTS, LABELS, score_resume() and
# hr_screen_tab() because Python uses the LAST definition of a name.
# ══════════════════════════════════════════════════════════════════════════

# ---------- 1) weights (total = 100) ----------
WEIGHTS = {"skills": 35, "education": 10, "experience": 10, "certifications": 10,
           "formatting": 10, "achievements": 10, "projects": 5,
           "contact": 5, "title_match": 5}                      # total = 100
LABELS = {"skills": "Skills Match", "education": "Education",
          "experience": "Experience", "certifications": "Certifications",
          "formatting": "Resume Formatting", "achievements": "Achievements",
          "projects": "Projects", "contact": "Contact & Online Profile",
          "title_match": "Job Title Relevance"}


# ---------- 2) new resume scorer ----------
_TITLE_STOP = {"senior", "junior", "lead", "sr", "jr", "the", "and", "for", "with", "of",
               "in", "at", "a", "an", "intern", "trainee", "remote", "full", "time", "part"}
_EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")
_PHONE_RE = re.compile(r"\+?\d[\d\s\-().]{8,16}\d")
_PROFILE_RE = re.compile(r"linkedin\.com|github\.com|behance\.net|portfolio", re.I)
_ACH_KW_RE = re.compile(r"\b(award|awarded|winner|won|rank|ranked|scholarship|dean'?s list|"
                        r"gold medal|silver medal|merit|appreciation|first position|"
                        r"second position|third position|honou?rs?|recognized|distinction|hackathon)\b")
_NUM_RE = re.compile(r"\d+(?:\.\d+)?\s*%|\b\d+\s*\+?\s*(?:users|clients|projects|customers|students|"
                     r"members|employees)\b|\$\s*\d|pkr\s*\d|\b(?:increased|reduced|improved|boosted)\b")


def score_resume(resume_text, job, category="General"):
    """Score a resume against a job out of 100 using 9 criteria:
    Skills 35 · Education 10 · Experience 10 · Certifications 10 · Formatting 10 ·
    Achievements 10 · Projects 5 · Contact & Online Profile 5 · Job Title Relevance 5."""
    data = _load()
    text = resume_text or ""
    tl = text.lower()
    sections = split_sections(text)

    title = (job.get("title") or "")
    job_text_l = " ".join([title, job.get("description") or "", job.get("requirements") or ""]).lower()
    role = _match_role((title + " " + (category or "")).lower())

    # ---- 1) SKILLS (35): share of the job's skills found in the resume ----
    job_skills = _find_skills(job_text_l)
    used_role_fallback = False
    if not job_skills and role:
        names = set(role["skills"])
        job_skills = [s for s in data["skills"] if s["skill"] in names]
        used_role_fallback = True
    res_skills = _find_skills(tl)
    res_names = {s["skill"] for s in res_skills}
    job_names = [s["skill"] for s in job_skills]
    matched = [n for n in job_names if n in res_names]
    missing = [n for n in job_names if n not in res_names]
    if job_names:
        frac = len(matched) / len(job_names)
        sk_detail = f"{len(matched)}/{len(job_names)} required skills found" + \
                    (" (default skills for the role)" if used_role_fallback else "")
    else:
        frac = min(len(res_names) / 8, 1) * 0.8
        sk_detail = f"No skills detected in the job; resume lists {len(res_names)} skills"
    skills_pts = frac * WEIGHTS["skills"]

    # ---- 2) EDUCATION (10): level (75 %) + field relevance (25 %) ----
    edu_text_l = sections.get("Education", "").lower() or tl
    cand_level = detect_level(edu_text_l)
    req_level = detect_level(job_text_l, mode="min") or 4
    lfrac = min(cand_level / req_level, 1) if cand_level else 0
    domain = _job_domain(job_skills, role)
    cand_fields = [doms for pat, doms in data["fields"] if pat.search(edu_text_l)]
    if domain is None:
        ffrac = 1.0
    elif not cand_fields:
        ffrac = 0.5
    else:
        ffrac = 1.0 if any(domain in d for d in cand_fields) else 0.4
    edu_pts = (0.75 * lfrac + 0.25 * ffrac) * WEIGHTS["education"] if cand_level else 0
    ed_detail = (f"Candidate: {LEVEL_NAMES[cand_level]} | Required: {LEVEL_NAMES[req_level]}"
                 + (" | Field relevant" if cand_level and ffrac == 1.0 and domain else ""))

    # ---- 3) EXPERIENCE (10): candidate years vs required years ----
    years = estimate_years(text, sections)
    req = required_years(job_text_l)
    if req > 0:
        efrac = min(years / req, 1)
        ex_detail = f"Candidate: {years:g} yrs | Required: {req}+ yrs"
    else:
        efrac = 0.6 + 0.4 * min(years / 3, 1)
        ex_detail = f"Candidate: {years:g} yrs | No experience requirement in the job"
    exp_pts = efrac * WEIGHTS["experience"]

    # ---- 4) CERTIFICATIONS (10) ----
    found_certs = [c for c in data["certs"] if any(p.search(tl) for p in c["pats"])]
    pts = 0.0
    for c in found_certs:
        pts += 1.0 if (domain is None or c["domain"] == domain) else 0.5
    cert_lines = [l for l in sections.get("Certifications", "").splitlines() if len(l.strip()) > 4]
    other_certs = max(len(cert_lines) - len(found_certs), 0)
    pts += 0.4 * min(other_certs, 4)
    cert_pts = min(pts / 3, 1) * WEIGHTS["certifications"]
    ce_detail = (f"{len(found_certs)} known + {other_certs} other certificate(s)"
                 if (found_certs or cert_lines) else "No certifications found")

    # ---- 5) RESUME FORMATTING (10): core sections 50 % · length 25 % · structure 25 % ----
    core = ["Summary", "Education", "Experience", "Skills"]
    core_present = [c for c in core if c in sections]
    core_frac = len(core_present) / len(core)
    words = len(text.split())
    if 200 <= words <= 1200:
        len_frac = 1.0
    elif 120 <= words < 200 or 1200 < words <= 1800:
        len_frac = 0.6
    elif words >= 60:
        len_frac = 0.3
    else:
        len_frac = 0.0
    lines = [l for l in text.splitlines() if l.strip()]
    bullets = sum(1 for l in lines if re.match(r"^\s*[•\-\*▪●◦‣–·]\s*\S", l))
    struct_frac = min(bullets / 6, 1)
    if struct_frac < 0.5 and len(lines) >= 15:       # PDF text often loses bullet symbols
        struct_frac = 0.5
    fmt_frac = 0.5 * core_frac + 0.25 * len_frac + 0.25 * struct_frac
    fmt_pts = fmt_frac * WEIGHTS["formatting"]
    fm_detail = f"{len(core_present)}/{len(core)} core sections | {words} words | {bullets} bullet lines"
    fm_missing = [c for c in core if c not in sections]

    # ---- 6) ACHIEVEMENTS (10): section 50 % · awards/keywords 30 % · measurable results 20 % ----
    ach_lines = [l for l in sections.get("Achievements", "").splitlines() if len(l.strip()) > 4]
    ach_kw = set(m for m in _ACH_KW_RE.findall(tl))
    nums = _NUM_RE.findall(tl)
    afrac = (0.5 * min(len(ach_lines) / 3, 1) + 0.3 * min(len(ach_kw) / 2, 1)
             + 0.2 * min(len(nums) / 3, 1))
    ach_pts = min(afrac, 1) * WEIGHTS["achievements"]
    ac_detail = (f"{len(ach_lines)} achievement line(s) | {len(ach_kw)} award keyword(s) | "
                 f"{len(nums)} measurable result(s)")

    # ---- 7) PROJECTS (5): section present · number of projects · uses the job's skills ----
    proj_text = sections.get("Projects", "")
    proj_lines = [l for l in proj_text.splitlines() if len(l.strip()) > 12]
    pfrac = 0.0
    if proj_lines:
        pfrac = 0.4 + 0.3 * min(len(proj_lines) / 4, 1)
        proj_names = {s["skill"] for s in _find_skills(proj_text.lower())}
        if job_names:
            pm = [n for n in job_names if n in proj_names]
            pfrac += 0.3 * min(len(pm) / max(1, min(len(job_names), 3)), 1)
        else:
            pfrac += 0.15
    proj_pts = min(pfrac, 1) * WEIGHTS["projects"]
    pj_detail = f"{len(proj_lines)} project line(s) found" if proj_lines else "No projects section found"

    # ---- 8) CONTACT & ONLINE PROFILE (5): e-mail 40 % · phone 30 % · LinkedIn/GitHub 30 % ----
    has_email = bool(_EMAIL_RE.search(text))
    has_phone = bool(_PHONE_RE.search(text))
    has_prof = bool(_PROFILE_RE.search(text))
    contact_pts = (0.4 * has_email + 0.3 * has_phone + 0.3 * has_prof) * WEIGHTS["contact"]
    ct_detail = " | ".join([f"Email {'✔' if has_email else '✘'}", f"Phone {'✔' if has_phone else '✘'}",
                            f"LinkedIn/GitHub {'✔' if has_prof else '✘'}"])

    # ---- 9) JOB TITLE RELEVANCE (5): words of the job title found in the resume ----
    twords = [w for w in re.findall(r"[a-z]+", title.lower()) if w not in _TITLE_STOP and len(w) > 2]
    if twords:
        hit = [w for w in twords if re.search(r"(?<![a-z0-9])" + re.escape(w[:6]), tl)]
        tfrac = len(hit) / len(twords)
        tt_detail = f"{len(hit)}/{len(twords)} job-title words found in resume"
    else:
        tfrac = 0.6
        tt_detail = "No specific job selected"
    title_pts = tfrac * WEIGHTS["title_match"]

    parts = {
        "skills": (skills_pts, sk_detail), "education": (edu_pts, ed_detail),
        "experience": (exp_pts, ex_detail), "certifications": (cert_pts, ce_detail),
        "formatting": (fmt_pts, fm_detail), "achievements": (ach_pts, ac_detail),
        "projects": (proj_pts, pj_detail), "contact": (contact_pts, ct_detail),
        "title_match": (title_pts, tt_detail),
    }
    breakdown = {k: {"label": LABELS[k], "score": round(v[0], 1), "max": WEIGHTS[k], "detail": v[1]}
                 for k, v in parts.items()}
    breakdown["skills"].update(matched=matched, missing=missing)
    breakdown["certifications"]["found"] = [c["name"] for c in found_certs]
    breakdown["formatting"].update(present=core_present, missing=fm_missing)
    total_score = int(round(sum(v[0] for v in parts.values())))
    total_score = max(0, min(100, total_score))

    # ---- strengths / weaknesses / recommendation ----
    strengths, weaknesses = [], []
    if frac >= 0.7 and job_names:
        strengths.append(f"Strong skills match ({len(matched)}/{len(job_names)} required skills)")
    if req and years >= req:
        strengths.append(f"Meets experience requirement ({years:g} yrs vs {req}+ yrs)")
    elif not req and years >= 2:
        strengths.append(f"{years:g} years of experience")
    if cand_level and cand_level >= req_level:
        strengths.append(f"Education level meets requirement ({LEVEL_NAMES[cand_level]})")
    if found_certs:
        strengths.append("Relevant certifications: " + ", ".join(c["name"] for c in found_certs[:3]))
    if afrac >= 0.6:
        strengths.append("Good achievements / measurable results")
    if len(proj_lines) >= 2:
        strengths.append(f"{len(proj_lines)} project entries listed")
    if fmt_frac >= 0.8:
        strengths.append("Well-structured, properly formatted resume")
    if missing:
        weaknesses.append("Missing skills: " + ", ".join(missing[:5]))
    if req and years < req:
        weaknesses.append(f"Experience is below requirement ({years:g} yrs vs {req}+ yrs)")
    if not cand_level:
        weaknesses.append("No education details found in the resume")
    elif cand_level < req_level:
        weaknesses.append(f"Education level below requirement ({LEVEL_NAMES[cand_level]} < {LEVEL_NAMES[req_level]})")
    if not found_certs and not cert_lines:
        weaknesses.append("No certifications listed")
    if not proj_lines:
        weaknesses.append("No projects section")
    if afrac < 0.2:
        weaknesses.append("No achievements or measurable results")
    if fm_missing:
        weaknesses.append("Missing resume sections: " + ", ".join(fm_missing))
    if not (has_email and has_phone):
        weaknesses.append("Contact details incomplete (email / phone)")
    strengths = strengths or ["Resume reviewed"]
    weaknesses = weaknesses or ["No major gaps found"]

    if total_score >= 80:
        rec = "Strongly recommended — shortlist for an interview right away."
    elif total_score >= 60:
        rec = "Recommended — can be shortlisted for an interview."
    elif total_score >= 40:
        rec = "Average match — compare with other candidates before deciding."
    else:
        rec = "Weak match for this job."

    top = matched[:5] or [s["skill"] for s in res_skills[:5]]
    return {"score": total_score, "rating": _rating(total_score), "breakdown": breakdown,
            "strengths": strengths[:4], "weaknesses": weaknesses[:4], "recommendation": rec,
            "years_exp": f"{years:g} years", "top_skills": top, "method": "dataset"}


# ---------- 3) HR manual screening with auto score ----------
def hr_screen_tab():
    """Manual resume screening for HR.
    Every uploaded resume is scored AUTOMATICALLY with the dataset-based scorer
    (score_resume) - no API key needed. Claude AI rating is an optional extra."""
    st.markdown('<span class="sec-title">📂 Upload & Screen Resumes</span>', unsafe_allow_html=True)
    api_key = st.session_state.get("claude_api_key", "")
    enable_ai = bool(api_key)
    sess_key = st.session_state.get("username", "anon")

    # ---- choose the job the resumes are scored against ----
    company = st.session_state.get("hr_company")
    jobs = db_get_jobs_for_company(company) if company else db_get_all_jobs()
    job_opts = {"General (no specific job)": {"title": "", "description": "", "requirements": ""}}
    for j in jobs:
        job_opts[f"{j['title']} ({j['job_id']})"] = j
    job_label = st.selectbox("🎯 Score resumes against job", list(job_opts), key="scr_job")
    sel_job = job_opts[job_label]

    uploaded = st.file_uploader("Drop resumes (PDF · DOCX · TXT)",
                                type=["pdf", "docx", "txt"], accept_multiple_files=True)

    if not uploaded:
        st.info("👆 Upload resumes to begin.")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("""
**Steps:**
1. Select the job
2. Upload resumes → auto-classified and **auto-scored (1-100)**
3. Open a profile to see the score breakdown
4. Send interview invites
5. Export CSV
""")
        with c2:
            st.markdown("""
**Ratings:**
- 🟢 Excellent 80-100
- 🔵 Good 60-79
- 🟡 Average 40-59
- 🔴 Poor 0-39
""")
        return

    # ---- read every file, predict its category and score it ----
    results, errors, raw_texts, ds_scores = [], [], {}, {}
    prog = st.progress(0, "Analysing…")
    for i, f in enumerate(uploaded):
        prog.progress((i + 1) / len(uploaded), f"Reading {f.name}")
        try:
            text = extract_text(f)
            cat = predict_category(text)
            raw_texts[f.name] = text
            ds = score_resume(text, sel_job, cat)          # dataset-based score
            ds_scores[f.name] = ds
            results.append({"filename": f.name, "category": cat,
                            "score": ds["score"], "rating": ds["rating"],
                            "preview": text[:500].strip()})
        except Exception as e:
            errors.append({"file": f.name, "error": str(e)})
    prog.empty()
    if not results:
        for e in errors:
            st.error(f"⚠️ {e['file']}: {e['error']}")
        return

    # best score first
    df = pd.DataFrame(results).sort_values("score", ascending=False).reset_index(drop=True)
    cat_cnts = df["category"].value_counts()

    # ---- dashboard ----
    st.markdown('<span class="sec-title">📊 Dashboard</span>', unsafe_allow_html=True)
    t1, t2, t3, t4 = st.columns(4)
    t1.markdown(f'<div class="stat-tile"><div class="stat-num">{len(df)}</div><div class="stat-lbl">Resumes</div></div>', unsafe_allow_html=True)
    t2.markdown(f'<div class="stat-tile"><div class="stat-num">{int(df["score"].mean())}</div><div class="stat-lbl">Average Score</div></div>', unsafe_allow_html=True)
    t3.markdown(f'<div class="stat-tile"><div class="stat-num">{int(df["score"].max())}</div><div class="stat-lbl">Highest Score</div></div>', unsafe_allow_html=True)
    t4.markdown(f'<div class="stat-tile"><div class="stat-num">{int((df["score"] >= 60).sum())}</div><div class="stat-lbl">Score 60+</div></div>', unsafe_allow_html=True)
    st.bar_chart(cat_cnts.rename_axis("Category").reset_index(name="Count").set_index("Category"))

    # ---- optional Claude AI ratings (saved in DB + this session) ----
    ai_ratings = db_get_all_ratings_for_session(sess_key)
    if "fresh_ratings" not in st.session_state:
        st.session_state.fresh_ratings = {}
    ai_ratings.update(st.session_state.fresh_ratings)

    def _rate_and_store(fn, text, category):
        r = rate_resume_with_ai(text, category, api_key)
        if not r.get("error") or r.get("score", 0) > 0:
            db_save_rating(sess_key, fn, r)
        st.session_state.fresh_ratings[fn] = r
        return r

    if enable_ai:
        if st.button("🤖 Also rate ALL resumes with Claude AI (optional)", key="rate_all"):
            with st.spinner("Claude evaluating all resumes…"):
                for row in results:
                    fn = row["filename"]
                    _rate_and_store(fn, raw_texts.get(fn, row["preview"]), row["category"])
            st.success("✅ AI ratings added!")
            st.rerun()
    else:
        st.caption("ℹ️ Scores below come from the dataset-based scorer. "
                   "Add a Claude API key in the sidebar only if you also want an AI second opinion.")

    st.markdown('<span class="sec-title">📋 Candidate Profiles (best score first)</span>', unsafe_allow_html=True)

    for idx, row in df.iterrows():
        fn = row["filename"]
        category = row["category"]
        preview = row["preview"]
        ds = ds_scores[fn]
        score = ds["score"]
        sc = score_color(score)
        rc_cls = rating_class(ds["rating"])

        st.markdown(f"""
<div class="card" style="border-left:5px solid {sc};">
  <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;">
    <div>
      <div style="font-size:1rem;font-weight:700;color:#003366;">#{idx + 1} &nbsp;📄 {html.escape(fn)}</div>
      <span class="cat-badge">🏷️ {html.escape(category)}</span>
      <span style="font-size:.78rem;color:#555;margin-left:6px;">🕐 {html.escape(ds["years_exp"])}</span>
    </div>
    <div style="display:flex;align-items:center;">
      <div class="score-ring" style="background:{sc};"><span class="num">{score}</span><span class="sub">/100</span></div>
      <span class="{rc_cls}" style="margin-left:8px;">{html.escape(ds["rating"])}</span>
    </div>
  </div>
</div>""", unsafe_allow_html=True)

        with st.expander(f"👤 Profile & Score — {fn}"):
            # full dataset score breakdown (bars, matched / missing skills, recommendation)
            render_score_breakdown(ds)

            sc_col, wk_col = st.columns(2)
            with sc_col:
                items = "".join(f'<div style="margin:4px 0;font-size:.83rem;">✅ {html.escape(s)}</div>' for s in ds["strengths"])
                st.markdown(f'<div style="background:#f0fff4;border:1px solid #c8e6c9;border-radius:12px;padding:1rem;"><div style="font-weight:700;color:#1b5e20;margin-bottom:6px;font-size:.88rem;">💪 Strengths</div>{items}</div>', unsafe_allow_html=True)
            with wk_col:
                items = "".join(f'<div style="margin:4px 0;font-size:.83rem;">⚠️ {html.escape(w)}</div>' for w in ds["weaknesses"])
                st.markdown(f'<div style="background:#fff8f0;border:1px solid #ffe0b2;border-radius:12px;padding:1rem;"><div style="font-weight:700;color:#bf360c;margin-bottom:6px;font-size:.88rem;">📈 Areas to Improve</div>{items}</div>', unsafe_allow_html=True)

            # optional Claude AI second opinion
            ai = ai_ratings.get(fn)
            if ai and ai.get("score", 0) > 0:
                ai_col = score_color(ai["score"])
                st.markdown(
                    f'<div style="margin-top:.8rem;padding:.8rem 1rem;background:#f3e8ff;border-radius:12px;">'
                    f'<b style="color:#4a148c;font-size:.85rem;">🤖 Claude AI second opinion: </b>'
                    f'<span style="color:{ai_col};font-weight:800;">{ai["score"]}/100</span> '
                    f'<span class="{rating_class(ai.get("rating", ""))}">{html.escape(str(ai.get("rating", "")))}</span>'
                    f'<div style="font-size:.82rem;color:#333;margin-top:4px;">{html.escape(ai.get("recommendation", ""))}</div></div>',
                    unsafe_allow_html=True)
            elif ai and ai.get("error"):
                err = ai.get("error", "")
                if "bad_api_key" in err:
                    st.error("❌ Invalid API key — console.anthropic.com → API Keys")
                elif "rate_limit" in err:
                    st.warning("⏳ Rate limited — wait 30 sec")
                else:
                    st.error(f"❌ {err[:200]}")
            if enable_ai and not (ai and ai.get("score", 0) > 0):
                if st.button("🤖 Get AI second opinion", key=f"rate_one_{idx}"):
                    with st.spinner("Rating…"):
                        _rate_and_store(fn, raw_texts.get(fn, preview), category)
                    st.rerun()

            st.text_area("📃 Resume Preview", preview + "…", height=120, key=f"prev_{idx}")

            # ---- contact section: name / email / phone auto-filled from the resume ----
            st.markdown("---")
            st.markdown("#### 📬 Send Interview Invitation")
            _ext = extract_candidate_info(raw_texts.get(fn, preview))
            for _k, _v in [(f"cn_{idx}", _ext["name"]),
                           (f"ce_{idx}", _ext["email"]),
                           (f"cp_{idx}", _ext["phone"])]:
                if _k not in st.session_state and _v:
                    st.session_state[_k] = _v

            if any([_ext["name"], _ext["email"], _ext["phone"]]):
                st.info(f"✨ Auto-filled — Name: **{_ext['name'] or '—'}** · "
                        f"Email: **{_ext['email'] or '—'}** · "
                        f"Phone: **{_ext['phone'] or '—'}**")

            cc1, cc2, cc3 = st.columns(3)
            cn = cc1.text_input("Candidate Name", key=f"cn_{idx}", placeholder="Ahmed Khan")
            cem = cc2.text_input("Email", key=f"ce_{idx}", placeholder="candidate@email.com")
            cph = cc3.text_input("WhatsApp Number", key=f"cp_{idx}", placeholder="+923001234567")

            interviewer = st.session_state.get("int_interviewer", display_name(st.session_state.user_info))
            idate = st.session_state.get("int_date", "Monday, May 20, 2026 at 3 PM")
            izoom = st.session_state.get("int_zoom", "https://zoom.us/j/your-id")
            icompany = st.session_state.get("int_company", "MIU")
            cname = cn or "Candidate"
            position = sel_job.get("title") or category

            subj = f"Interview Invitation — {position} | {icompany}"
            body = (f"Dear {cname},\n\nWe reviewed your resume and are pleased to shortlist you "
                    f"for the {position} position at {icompany}.\n\n"
                    f"Date & Time : {idate}\nZoom Link   : {izoom}\n\n"
                    f"Please confirm.\n\nBest regards,\n{interviewer}\n{icompany}")
            wa = (f"Hello {cname}! 👋\n\nI am *{interviewer}* from *{icompany}*. "
                  f"You are shortlisted for *{position}*.\n\n📅 {idate}\n🔗 {izoom}\n\n"
                  f"Please confirm. Thank you! 🙏")

            b1, b2, b3 = st.columns(3)
            with b1:
                if cem:
                    params = urllib.parse.urlencode({"subject": subj, "body": body},
                                                    quote_via=urllib.parse.quote)
                    st.markdown(f'<a class="btn-email" href="mailto:{urllib.parse.quote(cem)}?{params}" target="_blank">✉️ Email</a>', unsafe_allow_html=True)
                else:
                    st.caption("Enter email")
            with b2:
                if cph:
                    phone_d = re.sub(r'\D', '', cph)
                    st.markdown(f'<a class="btn-whatsapp" href="https://wa.me/{phone_d}?text={urllib.parse.quote(wa)}" target="_blank">💬 WhatsApp</a>', unsafe_allow_html=True)
                else:
                    st.caption("Enter phone")
            with b3:
                zh = izoom if izoom.startswith("https://") else "https://zoom.us"
                st.markdown(f'<a class="btn-zoom" href="{zh}" target="_blank">📹 Zoom</a>', unsafe_allow_html=True)

    # ---- export ----
    st.markdown('<span class="sec-title">⬇️ Export</span>', unsafe_allow_html=True)
    exp_df = df[["filename", "category", "score", "rating"]].copy()
    exp_df["recommendation"] = exp_df["filename"].map(lambda f: ds_scores[f]["recommendation"])
    exp_df["top_skills"] = exp_df["filename"].map(lambda f: ", ".join(ds_scores[f]["top_skills"]))
    exp_df["experience"] = exp_df["filename"].map(lambda f: ds_scores[f]["years_exp"])
    if ai_ratings:
        exp_df["ai_score"] = exp_df["filename"].map(lambda f: ai_ratings.get(f, {}).get("score", ""))
    st.download_button("📥 Download CSV",
                       data=exp_df.to_csv(index=False),
                       file_name=f"screening_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                       mime="text/csv", use_container_width=True)
    for e in errors:
        st.caption(f"⚠️ {e['file']}: {e['error']}")
# ══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════
enable_enter_to_next()   # Enter moves to the next box; Enter in the last box presses the form's main button

if not st.session_state.logged_in:
    show_login()
else:
    # Open the portal that matches the signed-in user's role
    role = st.session_state.user_info["role"]
    if role == "admin":
        admin_portal()
    elif role == "hr":
        if st.session_state.get("hr_company"):
            hr_portal()
        else:
            show_company_picker()      # HR must choose a company first
    elif role == "company":
        company_portal()
    elif role == "recruiter":
        recruiter_portal()
    else:
        st.error("Unknown role.")
        if st.button("Logout"):
            st.session_state.logged_in = False
            st.rerun()