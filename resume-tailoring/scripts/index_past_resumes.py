#!/usr/bin/env python3
"""
index_past_resumes.py - Enriches, HAMS-scores, and indexes past resume bullets
with Target Company, Core Tech Tags, and Purpose/Thesis metadata for semantic matching.
"""

import os
import sys
import re
import glob
import pypdf
from docx import Document

sys.stdout.reconfigure(encoding='utf-8')

WORKSPACE_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", ".."))
PAST_RESUMES_DIR = os.path.join(WORKSPACE_ROOT, "past_resumes")
VAULT_FILE = os.path.join(WORKSPACE_ROOT, ".agents", "skills", "resume-tailoring", "resources", "bullet-vault.md")
SYNC_VAULT_FILE = os.path.join(WORKSPACE_ROOT, "resume-tailoring", "resources", "bullet-vault.md")

# Known verified hard skills in Yash's stack
HARD_SKILLS_MAP = {
    "java": "Java", "spring boot": "Spring Boot", "spring security": "Spring Security",
    "hibernate": "Hibernate", "jpa": "JPA", "nestjs": "NestJS", "node.js": "Node.js",
    "nodejs": "Node.js", "express": "Express", "typescript": "TypeScript",
    "javascript": "JavaScript", "python": "Python", "go": "Go (Golang)", "golang": "Go",
    "c++": "C++", "c#": "C#", ".net": ".NET", "sql": "SQL", "postgresql": "PostgreSQL",
    "postgres": "PostgreSQL", "mysql": "MySQL", "mongodb": "MongoDB (NoSQL)", "redis": "Redis",
    "kafka": "Kafka", "websocket": "WebSockets", "websockets": "WebSockets", "vapi.ai": "Vapi.ai",
    "vapi": "Vapi.ai", "genesys": "Genesys AudioHook", "audiohook": "AudioHook",
    "langfuse": "Langfuse", "gemini": "Google Gemini", "next.js": "Next.js",
    "nextjs": "Next.js", "react": "React", "angular 14": "Angular 14", "angular": "Angular 14",
    "redux": "Redux", "tailwind": "Tailwind CSS", "docker": "Docker", "kubernetes": "Kubernetes",
    "aws": "AWS", "gcp": "GCP", "google cloud": "GCP", "azure": "Microsoft Azure",
    "jenkins": "Jenkins CI/CD", "ci/cd": "CI/CD", "sentry": "Sentry", "artillery": "Artillery Load Testing",
    "jest": "Jest", "junit": "JUnit", "mockito": "Mockito", "sonarqube": "SonarQube",
    "prisma": "Prisma ORM", "slate crm": "Slate CRM", "slate": "Slate CRM",
    "salesforce": "Salesforce CRM", "dynamodb": "DynamoDB", "terraform": "Terraform",
    "ecs": "AWS ECS", "lambda": "AWS Lambda", "sqlite": "SQLite", "flask": "Flask",
    "sqlalchemy": "SQLAlchemy",
    "ferpa": "FERPA Compliance", "oauth": "OAuth 2.0", "jwt": "JWT Auth", "stripe": "Stripe Payments",
    "hmac": "HMAC Signatures", "pub/sub": "Google Cloud Pub/Sub", "cloud run": "Google Cloud Run",
    "django": "Django", "fastapi": "FastAPI",
    "react native": "React Native",
    "uipath": "UiPath RPA",
    "rest": "RESTful APIs", "restful": "RESTful APIs", "grpc": "gRPC", "bff": "BFF Architecture",
    "microservices": "Microservices", "opentelemetry": "OpenTelemetry"
}

STRONG_ACTION_VERBS = [
    "architected", "engineered", "developed", "reduced", "automated",
    "launched", "modeled", "designed", "optimized", "streamlined",
    "spearheaded", "implemented", "delivered", "deployed", "structured",
    "authored", "scaled", "profiled", "accelerated", "orchestrated",
    "eliminated", "refactored", "integrated", "cut", "expanded", "slashed",
    "built", "shipped", "containerized", "troubleshot", "secured", "created",
    "won", "solved", "formulated", "instrumented", "established", "evaluated",
    "standardized"
]

WEAK_ACTION_VERBS = [
    "worked", "assisted", "helped", "responsible", "involved", "participated",
    "supported", "contributed", "maintained", "utilized", "used"
]

METRIC_PATTERNS = [
    r"\b\d+[\.,]?\d*\s*%",               # 50%, 99.9%
    r"\b\d+[\.,]?\d*\s*(?:[xX]\b|[×*])", # 3x, 3X, 3×
    r"\b\d+[\.,]?\d*\s*tps\b",           # 1,200 TPS
    r"\b\d+[\.,]?\d*\s*ms\b",            # 100ms, 120ms
    r"[\$₹]\s*\d+[\.,]?\d*[kKmMbB]?\+?", # $150K+, ₹50,000
    r"\brs\.?\s*\d+[\.,]?\d*[kKmM]?\b",  # Rs. 50,000
    r"\b\d+\s*k\+\b",                    # 1k+, 50k+, 1,000+
    r"\b\d+\+",                          # 200+, 450+, 4+
    r"\b\d{2,}\b"                         # 50,000, 15s to 5s
]

COMPANY_DETECTION_PATTERNS = {
    "Akamai": [r"akamai"],
    "American Express": [r"americanexpress", r"amex"],
    "Apple": [r"apple"],
    "Atlassian": [r"atlassian"],
    "Baazi Games": [r"baazi"],
    "Bajaj Finance": [r"bajaj"],
    "Bank of America": [r"bank_of_america", r"bofa"],
    "Barclays": [r"barclays"],
    "BlackRock": [r"blackrock", r"black_rock"],
    "Brillio": [r"brillio"],
    "Carelon": [r"carelon"],
    "Cisco": [r"cisco"],
    "Dentsu": [r"dentsu"],
    "ElevenLabs": [r"elevenlabs", r"eleven_labs"],
    "Ethos": [r"ethos"],
    "Expedia": [r"expedia"],
    "Fidelity": [r"fidelity"],
    "Flex": [r"\bflex\b"],
    "Franklin Templeton": [r"franklin", r"templeton"],
    "Gokiwi": [r"gokiwi", r"kiwi"],
    "Google": [r"google"],
    "Groww": [r"groww"],
    "Instahyre": [r"instahyre"],
    "IQVIA": [r"iqvia"],
    "JioHotstar": [r"jiohotstar", r"hotstar"],
    "Juspay": [r"juspay"],
    "KPMG": [r"kpmg"],
    "MakeMyTrip": [r"makemytrip", r"mmt"],
    "Mastercard": [r"mastercard"],
    "Microsoft": [r"microsoft", r"msft"],
    "NatWest": [r"natwest"],
    "Nike": [r"nike"],
    "ParkPlus": [r"parkplus", r"park_plus", r"park\+"],
    "PolicyBazaar": [r"policybazaar", r"polic-bazaar"],
    "Purplle": [r"purplle"],
    "RDash": [r"rdash"],
    "Rippling": [r"rippling"],
    "Rubrik": [r"rubrik"],
    "Salesforce": [r"salesforce"],
    "SOTI": [r"soti"],
    "Swiggy": [r"swiggy", r"instamart"],
    "Swish": [r"swish"],
    "Target": [r"(?:^|_)target(?:_|$)"],
    "Temple": [r"(?:^|_)temple(?:_|$)"],
    "Uber": [r"uber"],
    "Visa": [r"visa", r"visanet"],
    "Weekday": [r"weekday"],
    "Worldline": [r"worldline"],
    "Zebra": [r"zebra"],
    "Zomato": [r"zomato"],
    "ZS Associates": [r"(?:^|_)zs(?:_|$)"],
    "Airtel": [r"airtel"],
    "Cashfree": [r"cashfree"],
    "Citi": [r"(?:^|_)citi(?:_|$)", r"citigroup"],
    "FTSE Russell": [r"ftse"],
    "Fynd": [r"fynd"],
    "GalaxyAI": [r"galaxyai", r"galaxy_ai"],
    "JLL": [r"(?:^|_)jll(?:_|$)"],
    "LSEG": [r"lseg"],
    "Nasdaq": [r"nasdaq"],
    "OLX India": [r"(?:^|_)olx(?:_|$)"],
    "Optum": [r"optum"],
    "Oracle": [r"oracle"],
    "Stripe": [r"stripe"],
    "The Hartford": [r"the_hartford", r"hartford"],
    "TruAudience": [r"truaudience"],
    "UiPath": [r"uipath"],
    "Abstrabit": [r"abstrabit"]
}

SKILL_AND_HEADER_PREFIXES = (
    "languages:", "programming languages:", "databases:", "database &", "databases &",
    "frontend:", "frontend &", "frontend technologies:", "backend:", "backend &",
    "cloud:", "cloud &", "cloud /", "cloud & infrastructure:", "testing:", "testing &", "coursework:",
    "ai &", "ai, ", "payments &", "security &", "operations &",
    "engineering practices", "web &", "devops", "distributed systems &", "distributed systems:",
    "tools &", "ci/cd &", "ci/cd:", "data &", "reliability,", "reliability &",
    "software engineering &", "web development &", "backend & architecture:",
    "backend & microservices:", "backend & distributed systems:", "testing & qa:",
    "architecture & messaging:", "messaging & architecture:", "sdlc, testing & devops:",
    "devops, cloud & quality:", "methodology & security:", "distributed & messaging:",
    "devops & virtualization:", "frameworks & ai:", "platform & streaming systems:",
    "sdlc & devops:", "core engineering & sdlc:", "distributed & messaging systems:",
    "testing, ci/cd & quality:"
)

def is_skill_header_line(text):
    t_lower = text.lower().strip()
    if any(t_lower.startswith(pfx) for pfx in SKILL_AND_HEADER_PREFIXES) and ":" in text[:45]:
        return True
    if ":" in text[:35]:
        first_part, rest = text.split(":", 1)
        if len(first_part.split()) <= 5 and "," in rest and not any(v in rest.lower()[:30] for v in STRONG_ACTION_VERBS):
            return True
    return False

def detect_company_and_role(filename):
    fname_lower = filename.lower()
    detected_company = "General / Multi-Target"
    for comp, patterns in COMPANY_DETECTION_PATTERNS.items():
        if any(re.search(pat, fname_lower) for pat in patterns):
            detected_company = comp
            break
            
    detected_role = "Software Engineer"
    if "react_native" in fname_lower:
        detected_role = "React Native SDE"
    elif "applications_development_analyst" in fname_lower:
        detected_role = "Applications Development Analyst"
    elif "backoffice" in fname_lower:
        detected_role = "Analyst Backoffice Development"
    elif "product_analyst" in fname_lower:
        detected_role = "Associate Product Analyst"
    elif "application_software_engineer" in fname_lower:
        detected_role = "Application Software Engineer"
    elif "senior" in fname_lower:
        detected_role = "Senior Software Engineer"
    elif "mts" in fname_lower:
        detected_role = "Software Engineering MTS"
    elif "junior" in fname_lower:
        detected_role = "Junior Java Full-Stack Developer"
    elif "integrations" in fname_lower:
        detected_role = "Integrations Engineer"
    elif "business_technology" in fname_lower or "bts" in fname_lower:
        detected_role = "Business Technology Solutions Associate"
    elif "platform" in fname_lower:
        detected_role = "Platform Engineer"
    elif "central_people_team" in fname_lower:
        detected_role = "Software Engineer (Central People Team)"
    elif "java_backend" in fname_lower:
        detected_role = "Java Backend SDE"
    elif "backend" in fname_lower:
        detected_role = "Backend SDE"
    elif "frontend" in fname_lower:
        detected_role = "Frontend SDE"
    elif "fullstack" in fname_lower or "full_stack" in fname_lower:
        detected_role = "Full-Stack Engineer"
    elif "genai" in fname_lower or "ai_engineer" in fname_lower or "fde" in fname_lower:
        detected_role = "GenAI / AI Engineer"
    elif "python" in fname_lower:
        detected_role = "Python Backend Developer"
    elif "model_risk" in fname_lower:
        detected_role = "Model Risk & Quantitative SWE"
    elif "swe_ii" in fname_lower or "sde2" in fname_lower:
        detected_role = "SDE-2 / Systems"
    elif "cloud_security" in fname_lower:
        detected_role = "Cloud Security SWE"
    elif "managed_services" in fname_lower:
        detected_role = "Managed Services Engineer"
    elif "new_grad" in fname_lower:
        detected_role = "New Grad Software Engineer"
    elif "software_developer" in fname_lower:
        detected_role = "Software Developer"
    elif "ase" in fname_lower or "associate" in fname_lower:
        detected_role = "Associate Software Engineer"
        
    return f"{detected_company} ({detected_role})"

def extract_tech_tags(text):
    text_lower = text.lower()
    tags = set()
    for kw, canonical in HARD_SKILLS_MAP.items():
        if re.search(r'\b' + re.escape(kw) + r'\b', text_lower):
            tags.add(canonical)
    return sorted(list(tags))

def derive_core_purpose(text, category):
    text_lower = text.lower()
    if any(k in text_lower for k in ["won first place", "hackbpit", "codegeeks", "450+"]):
        return "Demonstrates high-signal competitive achievements, 1st-place hackathon engineering leadership (Rs. 50,000 award), and deep algorithmic problem-solving foundations (450+ DSA problems)."
    if "stripe" in text_lower or "subscription billing" in text_lower or "hmac" in text_lower:
        return "Demonstrates fintech subscription billing microservice engineering, HMAC-signed Stripe webhook verification, and 100% reconciliation idempotency."
    if "cloud run" in text_lower or "pub/sub" in text_lower or ("golang" in text_lower and "delta" in text_lower):
        return "Demonstrates cloud-native Go (Golang) data pipeline engineering on Cloud Run, Pub/Sub messaging, and lease-fenced idempotent writes in MongoDB."
    if "p0 incident" in text_lower or "schema drift" in text_lower or "post-mortem" in text_lower:
        return "Demonstrates mission-critical SRE incident response, on-call triaging, 2-hour P0 production schema drift resolution, and structured RCA post-mortem documentation."
    if "langfuse" in text_lower or "reasoning chains" in text_lower or "prompt drift" in text_lower or "hallucinations" in text_lower:
        return "Demonstrates LLM application observability, reasoning chain tracing with Langfuse, automated quality gates, and agentic prompt drift mitigation."
    if "1,200 tps" in text_lower or "double-charges" in text_lower or "idempotency" in text_lower or "fluxpay" in text_lower:
        return "Demonstrates high-throughput distributed payment processing (1,200 TPS), Saga pattern, Redis distributed locks, and financial ACID idempotency for Fintech / PBC roles."
    if "gpa" in text_lower or "150k" in text_lower or "transcript" in text_lower or "ferpa" in text_lower:
        return "Demonstrates end-to-end full-stack commercial product ownership ($150K+ contract value, 4+ university clients), FERPA student data privacy compliance, and Next.js BFF architecture."
    if "audiohook" in text_lower or "websocket" in text_lower or "vapi" in text_lower or "codec" in text_lower or "1k+" in text_lower or "telephony" in text_lower:
        return "Demonstrates low-latency real-time telephony streaming, bidirectional WebSocket architecture, µ-law/PCM audio transcoding, and voice AI orchestration under production traffic."
    if "15s to 5s" in text_lower or "analytics api" in text_lower or "3×" in text_lower or "3x" in text_lower or "query execution" in text_lower:
        return "Demonstrates relational database optimization, composite indexing strategies, SQL execution plan profiling, and 3× API latency reduction across microservice layers."
    if "spring boot" in text_lower or "etl" in text_lower or "slate" in text_lower or "salesforce" in text_lower or "dlq" in text_lower:
        return "Demonstrates enterprise CRM integration (Slate/Salesforce), object-oriented strategy design patterns, multi-tenant normalization, and Kafka DLQ fault-tolerance with Sentry observability."
    if "50k+" in text_lower or "python pipeline" in text_lower:
        return "Demonstrates automated data engineering, scheduled Python ETL batch workflows, and high-volume student record ingestion from PostgreSQL, cutting manual operational effort by 60%."
    if "interview" in text_lower or "200+" in text_lower or "turnaround" in text_lower:
        return "Demonstrates 1st-place hackathon-winning full-stack AI platform, automated voice screening pipelines, structured rubric evaluation, and 50% turnaround reduction for hiring teams."
    if "sentry" in text_lower or "mean-time-to-detect" in text_lower or "mttd" in text_lower:
        return "Demonstrates enterprise-wide reliability engineering, structured error capture telemetry, custom alert thresholds, and DLQ redrive monitoring to minimize incident MTTD."
    if "cdn" in text_lower or "lcp" in text_lower or "jenkins" in text_lower:
        return "Demonstrates Core Web Vitals optimization (25% LCP cut), zero-downtime CDN edge migrations, and automated Jenkins CI/CD release deployment in Agile/Scrum environments."
    return "Demonstrates scalable microservice design, robust unit testing gates, secure OAuth/JWT access control, and production API reliability."

def score_bullet_hams(bullet):
    text = bullet.strip()
    text_lower = text.lower()
    
    # 1. Hard Skill (H) - 25 pts
    h_count = sum(1 for kw in HARD_SKILLS_MAP.keys() if re.search(r'\b' + re.escape(kw) + r'\b', text_lower))
    if h_count >= 3:
        h_score = 25
    elif h_count == 2:
        h_score = 20
    elif h_count == 1:
        h_score = 15
    else:
        h_score = 5

    # 2. Action Verb (A) - 25 pts
    first_word = re.sub(r'[^a-zA-Z]', '', text.split()[0].lower()) if text.split() else ""
    if first_word in STRONG_ACTION_VERBS:
        a_score = 25
    elif any(verb in text_lower[:30] for verb in STRONG_ACTION_VERBS):
        a_score = 20
    elif first_word in WEAK_ACTION_VERBS:
        a_score = 5
    else:
        a_score = 12

    # 3. Metric (M) - 25 pts
    m_matches = []
    for pat in METRIC_PATTERNS:
        matches = re.findall(pat, text, re.IGNORECASE)
        m_matches.extend(matches)
    
    if len(m_matches) >= 2:
        m_score = 25
    elif len(m_matches) == 1:
        m_score = 20
    else:
        m_score = 0

    # 4. Structure & Line-Fitting (S) - 25 pts
    length = len(text)
    if 130 <= length <= 260:
        s_score = 25
    elif 90 <= length < 130:
        s_score = 20
    elif 260 < length <= 310:
        s_score = 18
    elif length > 310:
        s_score = 12
    else:
        s_score = 10

    penalties = 0
    penalty_reasons = []
    
    if "gpa" in text_lower and "kafka" in text_lower:
        penalties += 20
        penalty_reasons.append("Claims Kafka in GPA evaluation (disproven)")

    if "gemini" in text_lower and ("document intelligence" in text_lower or "articulation" in text_lower or "fallback" in text_lower):
        penalties += 50
        penalty_reasons.append("Claims Gemini API in Document Intelligence (disproven; LLM fallback only)")

    if "transactional outbox" in text_lower:
        penalties += 50
        penalty_reasons.append("Claims transactional outbox (disproven; Fluxpay uses Saga pattern)")

    if "maven" in text_lower:
        penalties += 50
        penalty_reasons.append("Claims Maven (disproven; candidate never used Maven)")

    if re.search(r'\b(spearheaded|pioneered|established)\b', text_lower[:25]):
        penalties += 20
        penalty_reasons.append("Uses inflated early-career power verb")
        
    if re.search(r'\b(senior|lead|principal)\s+(software|engineer|developer)\b', text_lower):
        penalties += 25
        penalty_reasons.append("Claims inflated Senior/Lead title")

    total_score = max(0, min(100, (h_score + a_score + m_score + s_score) - penalties))
    
    details = {
        "total": total_score,
        "H": h_score,
        "A": a_score,
        "M": m_score,
        "S": s_score,
        "penalties": penalty_reasons,
        "metrics_found": m_matches
    }
    return total_score, details

def sanitize_extracted_bullet(text):
    # 1. Purge Gemini API in Document Intelligence
    text = re.sub(r'\bGemini\s+API\s+fallback\b', 'LLM fallback', text, flags=re.IGNORECASE)
    text = re.sub(r'\bwith\s+Gemini\s+API\b', 'with LLM fallback', text, flags=re.IGNORECASE)
    text = re.sub(r'\bGemini\s+API\b', 'LLM fallback', text, flags=re.IGNORECASE) if 'document intelligence' in text.lower() or 'articulation' in text.lower() else text

    # 2. Purge Spearheaded/Pioneered on Document Intelligence
    text = re.sub(r'\b(?:Spearheaded|Pioneered)\s+(?:the|a)\s+Document\s+Intelligence', 'Improved the Document Intelligence', text, flags=re.IGNORECASE)
    text = re.sub(r'\b(?:Spearheaded|Pioneered)\s+an?\s+', 'Implemented an ', text, flags=re.IGNORECASE)
    text = re.sub(r'\b(?:Spearheaded|Pioneered)\s+', 'Implemented ', text, flags=re.IGNORECASE)

    # 3. Replace transactional outbox with Saga pattern
    text = re.sub(r'\btransactional\s+outbox(?:\s+pattern)?\b', 'Saga pattern', text, flags=re.IGNORECASE)

    # 4. Replace Established Jenkins with Maintained
    text = re.sub(r'\bEstablished\s+Jenkins\s+CI/CD\b', 'Maintained service reliability under Jenkins CI/CD', text, flags=re.IGNORECASE)
    text = re.sub(r'\bEstablished\b', 'Maintained', text) if 'jenkins' in text.lower() or 'ci/cd' in text.lower() else text

    # 5. Purge Maven completely
    text = re.sub(r'\b(?:built\s+)?with\s+Maven,?\s*', '', text, flags=re.IGNORECASE)
    text = re.sub(r',\s*Maven\b', '', text, flags=re.IGNORECASE)
    text = re.sub(r'\bMaven,\s*', '', text, flags=re.IGNORECASE)
    text = re.sub(r'\band\s+Maven\b', '', text, flags=re.IGNORECASE)
    text = re.sub(r'\bMaven\b', '', text)

    # Cleanup spacing
    text = " ".join(text.split()).strip()
    return text

def extract_from_docx(file_path):
    bullets = []
    try:
        doc = Document(file_path)
        for p in doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue
            style_name = getattr(p.style, "name", "") if p.style else ""
            if text.startswith(("•", "-", "*", "·", "–", "—")) or (style_name and style_name.startswith("List")):
                clean = re.sub(r"^[•\-\*·–—\t\s]+", "", text).strip()
                if is_skill_header_line(clean):
                    continue
                if len(clean) > 30:
                    bullets.append(sanitize_extracted_bullet(clean))
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
    return bullets

def extract_from_pdf(file_path):
    bullets = []
    try:
        reader = pypdf.PdfReader(file_path)
        full_text = ""
        for page in reader.pages:
            full_text += (page.extract_text() or "") + "\n"
        
        # Truncate if appended Job Description or About section exists
        full_text = re.split(r"\n\s*(?:About\s+[A-Za-z]+|Job Description|Role Overview|What you'll do|Qualifications|Requirements)\b", full_text, flags=re.IGNORECASE)[0]
        
        raw_items = re.split(r"\n\s*[•·\-\*–—]\s*", full_text)
        for item in raw_items[1:]:
            clean = " ".join(item.split()).strip()
            if is_skill_header_line(clean):
                continue
            if re.match(r"^(Education|Projects|Technical Skills|Achievements|Experience|Languages|Coursework)", clean):
                continue
            first_block = re.split(r"\n\s*(Education|Projects|Technical Skills|Achievements)\b", clean)[0].strip()
            first_block = re.sub(r"(\w+)-\s+(\w+)", r"\1\2", first_block)
            if is_skill_header_line(first_block):
                continue
            if len(first_block) > 35:
                bullets.append(sanitize_extracted_bullet(first_block))
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
    return bullets

def categorize_bullet(b):
    text_lower = b.lower()
    # 11. Achievements & Competitive Programming
    if any(k in text_lower for k in [
        "won first place at the ischoolconnect", "secured first position in hackbpit",
        "solved 450+ data structures", "hackbpit", "codegeeks hackathon", "leetcode, geeksforgeeks"
    ]):
        return "11. Achievements & Competitive Programming"
    # 1. Fluxpay
    if "fluxpay" in text_lower or "1,200 tps" in text_lower or "double-charges" in text_lower or ("idempotency" in text_lower and "payment" in text_lower):
        return "1. Projects: Fluxpay (Distributed Payments Microservice)"
    # 2. Subscription & Billing
    if "stripe" in text_lower or "subscription billing" in text_lower or ("reconciliation" in text_lower and "payment" in text_lower):
        return "2. Microservices & Payments: Subscription Billing (Stripe / HMAC)"
    # 3. AI Interviewer
    if "interview" in text_lower or "screening" in text_lower or "200+ monthly" in text_lower:
        return "3. Projects & Experience: AI Interviewer & Voice Evaluation"
    # 4. GPA Platform
    if "gpa" in text_lower or "transcript" in text_lower or "150k" in text_lower or "ferpa" in text_lower:
        return "4. Enterprise Flagship: GPA Evaluation Platform ($150K+ ACV)"
    # 5. Real-Time Telephony
    if "audiohook" in text_lower or "websocket" in text_lower or "vapi" in text_lower or "codec" in text_lower or "1k+" in text_lower or "telephony" in text_lower:
        return "5. Real-Time Telephony: Voice AI & WebSockets"
    # 6. Java / Spring Boot / CRM ETL
    if "spring boot" in text_lower or "etl" in text_lower or "slate" in text_lower or "salesforce" in text_lower or "dlq" in text_lower:
        return "6. Distributed Backend: Java / Spring Boot / CRM ETL"
    # 7. Database Optimization
    if "analytics api" in text_lower or "15s to 5s" in text_lower or "query execution" in text_lower or "composite index" in text_lower or "3×" in text_lower or "3x" in text_lower:
        return "7. Database & Performance: SQL Optimization (3× Speedup)"
    # 8. Python & Go Data Pipelines
    if "50k+" in text_lower or "reporting effort" in text_lower or "python pipeline" in text_lower or "cloud run" in text_lower:
        return "8. Data Engineering: Python & Go Pipelines"
    # 9. Reliability & Incident Triage
    if "sentry" in text_lower or "observability" in text_lower or "mean-time-to-detect" in text_lower or "mttd" in text_lower or "p0 incident" in text_lower or "post-mortem" in text_lower:
        return "9. Reliability & Observability: Sentry & Incident Triage"
    # 10. Web Performance & CI/CD
    if "cdn" in text_lower or "lcp" in text_lower or "jenkins" in text_lower or "largest contentful paint" in text_lower:
        return "10. Web Performance & DevOps: CDN Migration & Jenkins CI/CD"
    # 12. Other Core SWE & Microservices
    return "12. Other Core SWE & Microservices"

def main():
    print(f"Scanning for resumes in: {PAST_RESUMES_DIR} and {WORKSPACE_ROOT}")
    
    files = []
    for ext in ("*.docx", "*.pdf"):
        files.extend(glob.glob(os.path.join(PAST_RESUMES_DIR, ext)))
    for ext in ("*.docx", "*.pdf"):
        files.extend(glob.glob(os.path.join(WORKSPACE_ROOT, ext)))
        
    valid_files = [
        f for f in set(files)
        if not os.path.basename(f).startswith("~$")
        and not any(ign in os.path.basename(f) for ign in ("tune_test", "temp_test", "test_charter"))
    ]
    
    print(f"Found {len(valid_files)} resume files to index.")
    
    all_bullets = {}
    
    for fpath in valid_files:
        fname = os.path.basename(fpath)
        ext = os.path.splitext(fname)[1].lower()
        company_role_tag = detect_company_and_role(fname)
        
        extracted = []
        if ext == ".docx":
            extracted = extract_from_docx(fpath)
        elif ext == ".pdf":
            extracted = extract_from_pdf(fpath)
            
        for b in extracted:
            norm = re.sub(r"[^a-zA-Z0-9]", "", b.lower())
            if not norm or len(norm) < 25:
                continue
                
            score, details = score_bullet_hams(b)
            cat = categorize_bullet(b)
            tech_tags = extract_tech_tags(b)
            purpose = derive_core_purpose(b, cat)
            
            if norm not in all_bullets:
                all_bullets[norm] = {
                    "text": b,
                    "category": cat,
                    "score": score,
                    "details": details,
                    "sources": [fname],
                    "target_contexts": [company_role_tag],
                    "tech_tags": tech_tags,
                    "purpose": purpose
                }
            else:
                if fname not in all_bullets[norm]["sources"]:
                    all_bullets[norm]["sources"].append(fname)
                if company_role_tag not in all_bullets[norm]["target_contexts"]:
                    all_bullets[norm]["target_contexts"].append(company_role_tag)
                if score > all_bullets[norm]["score"]:
                    all_bullets[norm]["text"] = b
                    all_bullets[norm]["score"] = score
                    all_bullets[norm]["details"] = details
                    all_bullets[norm]["tech_tags"] = tech_tags
                    all_bullets[norm]["purpose"] = purpose

    s_tier = [b for b in all_bullets.values() if b["score"] >= 90]
    a_tier = [b for b in all_bullets.values() if 75 <= b["score"] < 90]
    b_tier = [b for b in all_bullets.values() if 60 <= b["score"] < 75]
    c_tier = [b for b in all_bullets.values() if b["score"] < 60]

    print("\n=== ENRICHED HAMS SCORING SUMMARY ===")
    print(f"Total Unique Bullets Cataloged: {len(all_bullets)}")
    print(f"👑 S-Tier (Score 90-100, Elite Production Bullets): {len(s_tier)}")
    print(f"⭐ A-Tier (Score 75-89, Strong Technical Bullets):   {len(a_tier)}")
    print(f"👍 B-Tier (Score 60-74, Good, Lacks Metric/Length):   {len(b_tier)}")
    print(f"⚠️ C-Tier (Score <60, Flagged or Weak):               {len(c_tier)}")
    
    categories = {}
    for b in all_bullets.values():
        categories.setdefault(b["category"], []).append(b)
        
    for cat in categories:
        categories[cat].sort(key=lambda x: -x["score"])

    # Generate Bullet Vault Markdown with Rich Metadata
    lines = [
        "# Candidate Master Bullet Vault & Semantic Query Catalog",
        "",
        "This file is the canonical library of all past-generated resume bullets, enriched with **Target Companies**, **Tech Tags**, and **Core Architectural Purpose** for instant semantic and keyword retrieval.",
        "",
        "> **Agent Semantic Matching Protocol:**",
        "> 1. **Analyze Job Description (JD)**: Identify target company style (GCC vs PBC vs Big Tech), required tech stack, and experience band.",
        "> 2. **Filter by Tech Tags & Target Context**: Match JD keywords against the tagged metadata below.",
        "> 3. **Select Highest-Scoring Bullets**: Always prioritize **👑 S-Tier (90–100)** or **⭐ A-Tier (75–89)** bullets.",
        "> 4. **Avoid Deprecated Phrasing**: Never use bullets flagged with penalties (e.g. Kafka in GPA).",
        "",
        "---",
        ""
    ]

    for cat in sorted(categories.keys()):
        lines.append(f"## {cat}")
        lines.append("")
        for item in categories[cat]:
            d = item["details"]
            badge = "👑 S-Tier" if item["score"] >= 90 else ("⭐ A-Tier" if item["score"] >= 75 else ("👍 B-Tier" if item["score"] >= 60 else "⚠️ C-Tier"))
            lines.append(f"### {badge} — Score: **{item['score']}/100** (H:{d['H']} | A:{d['A']} | M:{d['M']} | S:{d['S']})")
            lines.append(f"> {item['text']}")
            lines.append("")
            
            # Formatted Metadata Blocks
            contexts = ", ".join([f"`{c}`" for c in item["target_contexts"][:4]])
            if len(item["target_contexts"]) > 4:
                contexts += f" *(and {len(item['target_contexts'])-4} more)*"
            lines.append(f"- **🎯 Target Company / Context**: {contexts}")
            
            tags_str = ", ".join([f"`{t}`" for t in item["tech_tags"]]) if item["tech_tags"] else "`General SWE`"
            lines.append(f"- **💻 Tech Tags**: {tags_str}")
            
            lines.append(f"- **📌 Core Purpose / Thesis**: {item['purpose']}")
            
            if d["metrics_found"]:
                lines.append(f"- **📈 Key Metrics**: `{', '.join(d['metrics_found'][:4])}`")
                
            srcs = ", ".join([f"`{s}`" for s in item["sources"][:2]])
            if len(item["sources"]) > 2:
                srcs += f" *(and {len(item['sources'])-2} more)*"
            lines.append(f"- **📄 Proven Source Resumes**: {srcs}")
            
            if d["penalties"]:
                lines.append(f"- **⚠️ Penalties / Flagged**: {', '.join(d['penalties'])}")
                
            lines.append("")
        lines.append("---")
        lines.append("")

    os.makedirs(os.path.dirname(VAULT_FILE), exist_ok=True)
    with open(VAULT_FILE, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"\nSaved Enriched Bullet Vault to: {VAULT_FILE}")

    os.makedirs(os.path.dirname(SYNC_VAULT_FILE), exist_ok=True)
    with open(SYNC_VAULT_FILE, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
        
    print("Enriched Indexing completed successfully!")

if __name__ == "__main__":
    main()
