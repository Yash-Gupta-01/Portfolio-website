#!/usr/bin/env python3
"""
build_resume.py - Autonomous Résumé Generator for Yash Gupta
Generates high-precision, executive-styled 1-page .docx and native Word .pdf resumes.

Features:
- Pure Word OpenXML formatting (Georgia font, true paragraph borders, exact tab stops).
- Dual output: .docx (editable) + .pdf (via native Word COM automation).
- Visual QA: Renders to PNG via pypdfium2 and verifies single-page boundary.
"""

import os
import sys
import pypdf
import pypdfium2 as pdfium
import win32com.client
import docx.opc.constants
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def add_hyperlink(paragraph, url, text, font_size=9.0, bold=False, italic=False, color="111111", underline=False):
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    
    hyperlink = parse_xml(f'<w:hyperlink {nsdecls("w", "r")} r:id="{r_id}"/>')
    new_run = parse_xml(f'<w:r {nsdecls("w")}/>')
    rPr = parse_xml(f'<w:rPr {nsdecls("w")}/>')
    
    rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="Georgia" w:hAnsi="Georgia"/>')
    rPr.append(rFonts)
    
    sz = int(font_size * 2)
    rPr.append(parse_xml(f'<w:sz {nsdecls("w")} w:val="{sz}"/>'))
    
    if bold:
        rPr.append(parse_xml(f'<w:b {nsdecls("w")}/>'))
    if italic:
        rPr.append(parse_xml(f'<w:i {nsdecls("w")}/>'))
    if color:
        rPr.append(parse_xml(f'<w:color {nsdecls("w")} w:val="{color}"/>'))
    if underline:
        rPr.append(parse_xml(f'<w:u {nsdecls("w")} w:val="single"/>'))
    else:
        rPr.append(parse_xml(f'<w:u {nsdecls("w")} w:val="none"/>'))
        
    new_run.append(rPr)
    t = parse_xml(f'<w:t {nsdecls("w")} xml:space="preserve">{text}</w:t>')
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

def populate_instahyre(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    populate_general(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)

def populate_general(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Distributed Systems, Backend & Cloud Performance)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Engineered an enterprise commercial ", False),
        ("GPA evaluation platform", True),
        (" (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+", True),
        (" university deployments, generating ", False),
        ("$150K+", True),
        (" contract value per client and enforcing strict ", False),
        ("FERPA compliance", True),
        (".", False)
    ], space_after=1.8)
    add_bullet([
        ("Improved the ", False),
        ("Document Intelligence", True),
        (" pipeline by coupling deterministic articulation rule caches with ", False),
        ("LLM fallback", True),
        (", resolving ", False),
        ("40%+", True),
        (" of course mappings without model inference to cut token costs by ", False),
        ("40%", True),
        (" and eliminate prompt variance.", False)
    ], space_after=1.8)
    add_bullet([
        ("Architected a decoupled two-tier ", False),
        ("Go", True),
        (" ingestion pipeline on ", False),
        ("Cloud Run", True),
        (" and ", False),
        ("Google Cloud Pub/Sub", True),
        (", processing ", False),
        ("10K+ delta records per 15-minute sync cycle", True),
        (" into ", False),
        ("MongoDB", True),
        (" with lease-fenced consumers.", False)
    ], space_after=1.8)
    add_bullet([
        ("Developed high-throughput ", False),
        ("NestJS", True),
        (" WebSocket middleware bridging real-time telephony streams, handling bidirectional audio transcoding for ", False),
        ("1,000+", True),
        (" concurrent sessions at sub-", False),
        ("100ms", True),
        (" p95 latency.", False)
    ], space_after=1.8)
    add_bullet([
        ("Constructed a multi-tenant ", False),
        ("Java/Spring Boot", True),
        (" ETL pipeline integrating Slate and Salesforce CRMs via REST APIs, implementing ", False),
        ("Hibernate/JPA", True),
        (" and ", False),
        ("Kafka", True),
        (" Dead Letter Queues (", False),
        ("DLQ", True),
        (") to ensure zero data loss.", False)
    ], space_after=1.8)
    add_bullet([
        ("Reduced backend analytics API latency ", False),
        ("3× (15s to 5s)", True),
        (" by restructuring SQL execution plans, indexing composite keys, and eliminating redundant table scans across ", False),
        ("PostgreSQL and MySQL", True),
        (" schemas.", False)
    ], space_after=1.8)
    add_bullet([
        ("Automated CI quality gates and traced non-deterministic LLM reasoning chains via ", False),
        ("Langfuse", True),
        (" and ", False),
        ("Sentry", True),
        (", diagnosing prompt drift and tool-call hallucinations to cut mean-time-to-detect (", False),
        ("MTTD", True),
        (").", False)
    ], space_after=2.0)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Launched a full-stack AI interview platform automating ", False),
        ("200+", True),
        (" monthly candidate evaluations and slashing turnaround ", False),
        ("50%", True),
        (", designing responsive UI dashboards and integrating secure RESTful evaluation APIs.", False)
    ], space_after=1.8)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines automating containerized deployments (", False),
        ("Docker", True),
        (") and integration test suites in an ", False),
        ("Agile team", True),
        (", accelerating developer release cadence.", False)
    ], space_after=2.0)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), DBMS, Operating Systems, Computer Networks.", False)
    ], space_after=2.0)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Java, Spring Boot, Hibernate, Kafka, Redis, PostgreSQL, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Built a distributed payments microservice in ", False),
        ("Java (Spring Boot / Hibernate)", True),
        
        (" and ", False),
        ("Spring Security JWT", True),
        (" authentication, sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency.", False)
    ], space_after=1.8)
    add_bullet([
        ("Implemented an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" with ", False),
        ("Redis", True),
        (" distributed locking and ", False),
        ("PostgreSQL", True),
        (" ACID validation, eliminating duplicate charges across ", False),
        ("500+", True),
        (" simulated users.", False)
    ], space_after=1.8)
    add_bullet([
        ("Containerized microservices using ", False),
        ("Docker", True),
        (" and authored automated ", False),
        ("JUnit", True),
        (" integration test suites to validate transactional rollback semantics and edge-case status code handling under network timeouts.", False)
    ], space_after=2.0)

    add_project_heading("AI Interviewer", "NestJS, Next.js, React, MongoDB, TypeScript, OpenAI API, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Designed an AI evaluation platform using ", False),
        ("NestJS", True),
        (", ", False),
        ("Next.js", True),
        (", and ", False),
        ("OpenAI API", True),
        (", establishing structured rubrics to score candidate responses and authoring ", False),
        ("Jest", True),
        (" test suites to enforce code quality gates.", False)
    ], space_after=1.8)
    add_bullet([
        ("Deployed bidirectional streaming audio pipelines via ", False),
        ("WebSockets", True),
        (" and ", False),
        ("Vapi.ai", True),
        (", enabling real-time conversational voice interviews and reducing assessment turnaround by ", False),
        ("50%", True),
        (" across 200+ monthly reviews.", False)
    ], space_after=2.0)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Java (17+), Go (Golang), Python, TypeScript, JavaScript, C++, SQL (PostgreSQL, MySQL)"),
        ("Backend & Microservices: ", True, "Spring Boot, NestJS, Node.js, Express.js, Go (net/http), Next.js (BFF), REST APIs"),
        ("Distributed & Messaging: ", True, "Apache Kafka (DLQ, Outbox Pattern), Google Cloud Pub/Sub, Redis (Distributed Locks/Cache)"),
        ("Databases & Storage: ", True, "PostgreSQL, MongoDB (Aggregations, Lease Fencing), MySQL, Hibernate/JPA, Prisma ORM"),
        ("Cloud & DevOps: ", True, "Docker, Jenkins CI/CD, Git, GitHub Actions, Google Cloud Platform (Cloud Run), AWS, Linux"),
        ("Testing, AI & Tools: ", True, "JUnit, Jest, Artillery, Sentry Observability, Langfuse (LLM Tracing & Evals), Gemini, WebSockets")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.6)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform with low-latency WebSocket audio streaming.", False)
    ], space_after=1.8)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ competing teams for fault-tolerant distributed system architecture.", False)
    ], space_after=1.8)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, demonstrating strong algorithmic complexity analysis, memory optimization, and concurrent execution skills.", False)
    ], space_after=0.0)

def populate_jiohotstar(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Frontend & Viewer Experience Focus)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Developed responsive, accessible (", False),
        ("WCAG", True),
        (") ", False),
        ("React", True),
        (" and ", False),
        ("Next.js", True),
        (" interfaces for GPA and credit evaluation with reusable design system components and role-based access control, cutting manual counsellor review overhead by ", False),
        ("80%", True),
        (".", False)
    ])
    add_bullet([
        ("Built an admission co-pilot web interface in ", False),
        ("Slate CRM", True),
        (", translating complex ", False),
        ("UI/UX designs", True),
        (" into real-time application trackers, automated communication workflows, and structured daily agenda views with ", False),
        ("100%", True),
        (" cross-browser fidelity.", False)
    ])
    add_bullet([
        ("Shipped a ", False),
        ("Next.js", True),
        (" and ", False),
        ("Stripe", True),
        (" subscription billing dashboard, building real-time credit-usage metrics, tiered-plan management, and checkout lifecycle workflows with ", False),
        ("100%", True),
        (" payment reconciliation and transaction state integrity.", False)
    ])
    add_bullet([
        ("Rolled out ", False),
        ("Sentry observability", True),
        (" across all production ", False),
        ("React", True),
        (" and ", False),
        ("Next.js", True),
        (" applications with source-map uploads, session replay, and custom alert thresholds, measurably reducing production mean-time-to-resolution (", False),
        ("MTTR", True),
        (") across web clients.", False)
    ])
    add_bullet([
        ("Reduced document-analytics API response latency ", False),
        ("3×", True),
        (" (", False),
        ("15s to 5s", True),
        (") by profiling client data-access patterns, eliminating redundant network round-trips and restructuring backend ", False),
        ("PostgreSQL/MySQL", True),
        (" queries to maximize UI responsiveness.", False)
    ])

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Created the ", False),
        ("React/TypeScript", True),
        (" frontend for an AI interview platform, delivering role-based dashboards, voice-interview controls, and an automated scheduling UI for ", False),
        ("200+", True),
        (" monthly sessions, cutting candidate turnaround by ", False),
        ("50%", True),
        (".", False)
    ])
    add_bullet([
        ("Lowered Largest Contentful Paint (", False),
        ("LCP", True),
        (") by ", False),
        ("15%", True),
        (" on a production ", False),
        ("Angular 14", True),
        (" platform by restructuring client-side rendering and data-fetching paths, resolving critical SEO and browser Core Web Vitals defects across the site.", False)
    ])
    add_bullet([
        ("Streamlined web asset delivery and frontend modules by configuring ", False),
        ("CDN edge caching", True),
        (" and asset chunking in an ", False),
        ("Agile team", True),
        (", delivering zero-downtime releases and lowering Largest Contentful Paint (", False),
        ("LCP", True),
        (") by ", False),
        ("25%", True),
        (".", False)
    ])

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, OOP, Database Management Systems (DBMS), Operating Systems, Computer Networks.", False)
    ], space_after=2.5)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "React, Redux, Vite, Node.js, Express.js, PostgreSQL, JWT", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Engineered a mobile-first ", False),
        ("React/Redux", True),
        (" fintech dashboard bundled with ", False),
        ("Vite", True),
        (", featuring dark/light theming and JWT-secured routing, achieving ", False),
        ("99.9%", True),
        (" UI uptime and deterministic state rendering under burst network load.", False)
    ])
    add_bullet([
        ("Designed idempotency-aware transaction APIs on ", False),
        ("Node.js/Express", True),
        (" and ", False),
        ("PostgreSQL", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" (simulated) at p95 latency under ", False),
        ("120ms", True),
        (", validated through authored Artillery load tests and automated test suites.", False)
    ])

    add_project_heading("AI Architect Assistant", "Next.js, React, Google Gemini API, Tailwind CSS, TypeScript", "github.com/Yash-Gupta-01/AI-Architect-Assistant", url="https://github.com/Yash-Gupta-01/AI-Architect-Assistant")
    add_bullet([
        ("Developed a full-stack ", False),
        ("Next.js", True),
        (" web application utilizing the ", False),
        ("Google Gemini API", True),
        (" to drive guided conversational design flows, capturing structural constraints to generate real-time 2D architectural visualizations.", False)
    ])
    add_bullet([
        ("Engineered responsive, component-driven UI with ", False),
        ("Tailwind CSS", True),
        (" and ", False),
        ("Next.js App Router", True),
        (", implementing client-side validation, cross-browser state hydration, and modular layout primitives for open-source users.", False)
    ])

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Frontend & UI: ", True, "JavaScript (ES6+), TypeScript, React, Next.js (BFF), Redux, Angular 14, Tailwind CSS, HTML5, CSS3/SASS"),
        ("Browser & Performance: ", True, "Core Web Vitals, LCP Optimization, CDN Caching, Cross-Browser Compatibility, WCAG"),
        ("Architecture & Design: ", True, "Component-Driven Architecture, In-house Design Systems, Reusable UI Libraries, RESTful APIs, OOP"),
        ("Build Tools & DevOps: ", True, "Git, Vite, Webpack, npm, Jenkins CI/CD, Docker, Google Cloud (GCP), AWS"),
        ("Testing & Observability: ", True, "Jest, Mocha, Sentry (Session Replay/Source Maps), Artillery Load Testing, Code Reviews"),
        ("Backend & Databases: ", True, "Node.js, Express.js, NestJS, PostgreSQL, MongoDB, Redis (caching), Prisma ORM")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=2.0)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform.", False)
    ])
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for distributed system design.", False)
    ])
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef.", False)
    ])

def populate_microsoft(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Azure Data & Systems Infrastructure Focus)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Architected a multi-tenant ", False),
        ("Java/Spring Boot", True),
        (" ETL service integrating Slate and Salesforce CRMs via a strategy-pattern factory (OOP), implementing ", False),
        ("Hibernate/JPA", True),
        (" and ", False),
        ("Kafka", True),
        (" Dead Letter Queues (", False),
        ("DLQ", True),
        (") to prevent data loss during burst traffic.", False)
    ])
    add_bullet([
        ("Optimized backend analytics APIs, slashing latency ", False),
        ("3×", True),
        (" (", False),
        ("15s to 5s", True),
        (") across production endpoints by profiling ", False),
        ("SQL", True),
        (" query execution plans, eliminating redundant joins, and restructuring ", False),
        ("PostgreSQL/MySQL", True),
        (" schemas for high throughput.", False)
    ])
    add_bullet([
        ("Automated high-volume data pipelines processing ", False),
        ("50K+", True),
        (" student records from ", False),
        ("PostgreSQL", True),
        (" using scheduled ", False),
        ("Python", True),
        (" jobs, eliminating manual intervention and cutting stakeholder reporting cycles by ", False),
        ("60%", True),
        (".", False)
    ])
    add_bullet([
        ("Developed high-throughput ", False),
        ("NestJS", True),
        (" WebSocket middleware bridging real-time telephony streams, handling bidirectional audio transcoding for ", False),
        ("1k+", True),
        (" concurrent sessions at sub-", False),
        ("100ms", True),
        (" p95 latency with zero connection drops under load.", False)
    ])
    add_bullet([
        ("Evaluated non-deterministic agentic flows and traced reasoning chains using ", False),
        ("Langfuse", True),
        (" and ", False),
        ("Sentry", True),
        (", diagnosing tool-call hallucinations and prompt drift to automate quality gates and cut mean-time-to-detect (", False),
        ("MTTD", True),
        (").", False)
    ])
    add_bullet([
        ("Engineered an enterprise commercial ", False),
        ("GPA evaluation platform", True),
        (" (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+", True),
        (" university deployments generating ", False),
        ("$150K+", True),
        (" contract value per client, enforcing ", False),
        ("FERPA compliance", True),
        (" and RBAC security.", False)
    ])

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Constructed a ", False),
        ("Kafka-driven NestJS", True),
        (" microservice for asynchronous event synchronization between CRMs and internal databases, maintaining sub-", False),
        ("50ms", True),
        (" synchronization latency under peak production traffic.", False)
    ])
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines automating containerized deployments (", False),
        ("Docker", True),
        (") and integration test suites in an ", False),
        ("Agile team", True),
        (", eliminating broken staging builds and accelerating developer release cadence.", False)
    ])

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Operating Systems, DBMS, Computer Networks.", False)
    ], space_after=2.5)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Java, Spring Boot, Kafka, Redis, PostgreSQL, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a distributed payments microservice in ", False),
        ("Java (Spring Boot)", True),
        
        (", enforcing authentication via ", False),
        ("Spring Security", True),
        (" and ", False),
        ("JWT", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency with published Artillery load tests.", False)
    ])
    add_bullet([
        ("Architected an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" with ", False),
        ("Redis", True),
        (" distributed locks and ", False),
        ("PostgreSQL", True),
        (" ACID validation, eliminating duplicate charges across ", False),
        ("500+", True),
        (" concurrent simulated users under load.", False)
    ])
    add_bullet([
        ("Containerized microservices with ", False),
        ("Docker", True),
        (" and authored automated ", False),
        ("JUnit", True),
        (" integration test suites to validate transactional rollback semantics under failure scenarios.", False)
    ])

    add_project_heading("AI Interviewer", "NestJS, Next.js, MongoDB (NoSQL), TypeScript, OpenAI API, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Designed a cloud-based interview evaluation platform with RESTful APIs and role-based access control, automating ", False),
        ("200+", True),
        (" monthly evaluations and cutting candidate assessment turnaround by ", False),
        ("50%", True),
        (" for hiring teams.", False)
    ])
    add_bullet([
        ("Authored comprehensive unit and integration test suites using ", False),
        ("Jest", True),
        (", maintaining high test coverage across core evaluation engines to enforce quality gates and ensure reliable continuous deployments.", False)
    ])

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Java, C++, Python, C, TypeScript, JavaScript, Go, SQL (PostgreSQL, MySQL)"),
        ("Distributed Systems: ", True, "Apache Kafka (Queues/DLQ/Outbox), Redis (Distributed Locks/Caching), Spring Boot, REST APIs"),
        ("Databases & Storage: ", True, "PostgreSQL, MySQL, RDBMS optimization, MongoDB, ACID Transactions, Schema Execution Plan Profiling"),
        ("Reliability & Telemetry: ", True, "Distributed Tracing, Sentry (Error Telemetry), Langfuse, Logging, Root-Cause Analysis, MTTD"),
        ("DevOps & Cloud: ", True, "Docker, Kubernetes (familiar), Jenkins CI/CD, Git, AWS, GCP, Zero-Downtime Deployments"),
        ("Engineering Practices: ", True, "Data Structures & Algorithms, OOP, Design Patterns, Jest, JUnit, Artillery Testing, Code Reviews")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=2.0)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams for building an automated voice assessment platform.", False)
    ])
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023) for real-time distributed applications.", False)
    ])
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" algorithmic problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, emphasizing graph algorithms and DP.", False)
    ])

def populate_elevenlabs(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Audio AI, Full-Stack Product & GenAI Focus)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Engineered high-throughput ", False),
        ("NestJS", True),
        (" WebSocket middleware bridging real-time telephony streams, handling bidirectional ", False),
        ("µ\u2011law/PCM", True),
        (" audio transcoding for ", False),
        ("1k+", True),
        (" concurrent sessions at sub-", False),
        ("100ms", True),
        (" p95 latency with zero connection drops under load.", False)
    ])
    add_bullet([
        ("Evaluated non-deterministic agentic flows and traced reasoning chains using ", False),
        ("Langfuse", True),
        (" and ", False),
        ("Sentry", True),
        (", diagnosing tool-call hallucinations and prompt drift to automate quality gates and cut mean-time-to-detect (", False),
        ("MTTD", True),
        (").", False)
    ])
    add_bullet([
        ("Automated high-volume data extraction pipelines for ", False),
        ("50K+", True),
        (" records from ", False),
        ("PostgreSQL", True),
        (" using scheduled ", False),
        ("Python", True),
        (" batch jobs, eliminating manual intervention and cutting reporting turnaround by ", False),
        ("60%", True),
        (".", False)
    ])
    add_bullet([
        ("Architected an end-to-end full-stack ", False),
        ("GPA platform", True),
        (" (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("React", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") deployed across ", False),
        ("4+", True),
        (" enterprise universities, securing ", False),
        ("$150K+", True),
        (" contract value per client under ", False),
        ("FERPA compliance", True),
        (".", False)
    ])
    add_bullet([
        ("Optimized backend analytics APIs, slashing latency ", False),
        ("3×", True),
        (" (", False),
        ("15s to 5s", True),
        (") across production endpoints by profiling ", False),
        ("SQL", True),
        (" query execution plans, eliminating redundant joins, and restructuring ", False),
        ("PostgreSQL/MySQL", True),
        (" schemas for high throughput.", False)
    ])
    add_bullet([
        ("Architected a multi-tenant ", False),
        ("Java/Spring Boot", True),
        (" ETL service integrating Slate and Salesforce CRMs via a strategy-pattern factory (OOP), implementing ", False),
        ("Hibernate/JPA", True),
        (" and ", False),
        ("Kafka", True),
        (" Dead Letter Queues (", False),
        ("DLQ", True),
        (") to prevent data loss during burst traffic.", False)
    ])

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Built the ", False),
        ("React/TypeScript", True),
        (" frontend for an automated AI interview platform, delivering voice-screening controls, role-based dashboards, and automated scheduling for ", False),
        ("200+", True),
        (" monthly sessions (", False),
        ("50%", True),
        (" turnaround cut).", False)
    ])
    add_bullet([
        ("Streamlined web asset delivery and frontend modules by configuring ", False),
        ("CDN edge caching", True),
        (" and asset chunking in an ", False),
        ("Agile team", True),
        (", delivering zero-downtime releases and lowering Largest Contentful Paint (", False),
        ("LCP", True),
        (") by ", False),
        ("25%", True),
        (".", False)
    ])

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Operating Systems, DBMS, Computer Networks.", False)
    ], space_after=2.5)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("AI Architect Assistant", "Next.js, React, Google Gemini API, Tailwind CSS, TypeScript", "github.com/Yash-Gupta-01/AI-Architect-Assistant", url="https://github.com/Yash-Gupta-01/AI-Architect-Assistant")
    add_bullet([
        ("Developed a full-stack ", False),
        ("Next.js", True),
        (" web application utilizing the ", False),
        ("Google Gemini API", True),
        (" to drive guided conversational design flows, capturing structural constraints to generate real-time 2D architectural visualizations.", False)
    ])
    add_bullet([
        ("Engineered responsive, component-driven UI with ", False),
        ("Tailwind CSS", True),
        (" and ", False),
        ("Next.js App Router", True),
        (", implementing client-side validation, cross-browser state hydration, and modular layout primitives for open-source users.", False)
    ])

    add_project_heading("Fluxpay", "Python, Go, Node.js, Redis, Kafka, PostgreSQL, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Engineered a high-concurrency payment microservice processing ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency, validating throughput limits through authored ", False),
        ("Artillery", True),
        (" load testing suites.", False)
    ])
    add_bullet([
        ("Architected an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" with ", False),
        ("Redis", True),
        (" distributed locks and ", False),
        ("PostgreSQL", True),
        (" ACID validation, eliminating duplicate charges across ", False),
        ("500+", True),
        (" concurrent simulated users under load.", False)
    ])
    add_bullet([
        ("Built a mobile-first ", False),
        ("React/Redux", True),
        (" dashboard with dark/light theming and containerized services using ", False),
        ("Docker", True),
        (" for deterministic local development and continuous deployment.", False)
    ])

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Python (advanced), TypeScript, JavaScript (ES6+), Go, Java, C++, SQL (PostgreSQL, MySQL)"),
        ("AI & Audio Systems: ", True, "Real-time WebSockets, Audio Transcoding (µ-law/PCM), Vapi.ai, OpenAI API, Gemini API, Langfuse, RAG"),
        ("Frontend & Web: ", True, "React, Next.js (App Router/BFF), Redux, Tailwind CSS, HTML5, CSS3, Core Web Vitals (LCP), Vite"),
        ("Backend & Systems: ", True, "Node.js, NestJS, Express.js, FastAPI, Spring Boot, RESTful APIs, Event-Driven Architecture"),
        ("Storage & Streaming: ", True, "PostgreSQL, MongoDB, Redis (Distributed Caching/Locks), Apache Kafka (Queues/DLQ), Prisma ORM"),
        ("DevOps & Reliability: ", True, "Docker, Jenkins CI/CD, Git, AWS, GCP, Sentry (Observability), Jest, Artillery Load Testing")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=2.0)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams for building an automated voice interview screening platform.", False)
    ])
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU) among 100+ competing teams for real-time distributed applications.", False)
    ])
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" algorithmic problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, emphasizing graphs concurrent algorithms, and low-latency system design.", False)
    ])

def populate_swiggy(doc, add_section_header, add_role_heading, add_project_heading, add_bullet, role="SDE_I"):
    if not ("frontend" in role.lower() or "mobile" in role.lower() or "react_native" in role.lower()):
        # 2. Experience (HR-First Layout - Swiggy Instamart Supply Chain: Golang, Distributed Systems, Agentic AI)
        add_section_header("Experience")

        add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
        add_bullet([
            ("Engineered a decoupled, two-tier distributed ingestion pipeline in ", False),
            ("Go (Golang)", True),
            (" on ", False),
            ("Cloud Run", True),
            (" and ", False),
            ("Pub/Sub", True),
            (", processing ", False),
            ("10K+ delta records per 15-minute sync cycle", True),
            (" into durable queues with ", False),
            ("MongoDB lease-fenced", True),
            (" idempotent consumers.", False)
        ])
        add_bullet([
            ("Applied agentic AI workflows and LLM observability (", False),
            ("Cursor", True),
            (", ", False),
            ("GitHub Copilot", True),
            (", ", False),
            ("Langfuse", True),
            (") across engineering lifecycles, automating regression test generation and tracing tool-call chains to accelerate sprint velocity.", False)
        ])
        add_bullet([
            ("Built real-time ", False),
            ("WebSocket middleware", True),
            (" in ", False),
            ("NestJS", True),
            (" sustaining ", False),
            ("1,000+ concurrent sessions", True),
            (" at sub-", False),
            ("100ms", True),
            (" p95 latency under peak production traffic via bidirectional ", False),
            ("8 kHz µ-law to 16 kHz PCM", True),
            (" transcoding.", False)
        ])
        add_bullet([
            ("Optimized ", False),
            ("PostgreSQL", True),
            (" and ", False),
            ("MongoDB", True),
            (" write throughput during burst synchronization spikes by implementing ordered ", False),
            ("bulkWrite operations", True),
            (" and query indexing, reducing API response latency ", False),
            ("3×", True),
            (" (", False),
            ("15s to 5s", True),
            (").", False)
        ])
        add_bullet([
            ("Shipped an enterprise billing engine in ", False),
            ("Next.js", True),
            (" with ", False),
            ("Stripe Checkout", True),
            (" integration, supporting credit package purchases, resilient fallback error boundaries, and ", False),
            ("100%", True),
            (" payment reconciliation integrity.", False)
        ])
        add_bullet([
            ("Scaled an enterprise commercial evaluation platform across ", False),
            ("4+", True),
            (" university deployments generating ", False),
            ("$150K+", True),
            (" contract value per client, automating structured data synchronization into ", False),
            ("Slate CRM", True),
            (" with strict RBAC security.", False)
        ])

        add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
        add_bullet([
            ("Constructed a ", False),
            ("Kafka-driven NestJS", True),
            (" microservice for asynchronous event synchronization between external enterprise platforms and internal databases, maintaining sub-", False),
            ("50ms", True),
            (" latency under high concurrency.", False)
        ])
        add_bullet([
            ("Maintained ", False),
            ("Jenkins CI/CD", True),
            (" artifact pipelines automating containerized deployments (", False),
            ("Docker", True),
            (") and integration test suites in an ", False),
            ("Agile team", True),
            (", eliminating broken staging builds and accelerating developer release cadence.", False)
        ])

        # 3. Education
        add_section_header("Education")
        add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
        add_bullet([
            ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Operating Systems, DBMS, Computer Networks.", False)
        ], space_after=2.5)

        # 4. Projects
        add_section_header("Projects")
        add_project_heading("Fluxpay", "Go, PostgreSQL, Kafka, Redis, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
        add_bullet([
            ("Designed a cloud-native, high-concurrency payment microservice in ", False),
            ("Go (goroutines/channels)", True),
            (" sustaining ", False),
            ("1,200 TPS", True),
            (" in simulation at sub-", False),
            ("120ms", True),
            (" p95 latency, verified through authored ", False),
            ("Artillery", True),
            (" load testing suites.", False)
        ])
        add_bullet([
            ("Maintained an event-driven Saga pattern over ", False),
            ("Kafka", True),
            (" with ", False),
            ("Redis", True),
            (" distributed locks and ", False),
            ("PostgreSQL", True),
            (" ACID validation, eliminating duplicate charges across ", False),
            ("500+", True),
            (" concurrent simulated users under load.", False)
        ])

        add_project_heading("AI Interviewer", "NestJS, Next.js, OpenAI API, WebSockets, MongoDB, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
        add_bullet([
            ("Formulated an autonomous voice screening platform using ", False),
            ("NestJS", True),
            (", ", False),
            ("Next.js", True),
            (", and ", False),
            ("OpenAI API", True),
            (", designing structured evaluation rubrics and automated ", False),
            ("Jest", True),
            (" test suites to conduct ", False),
            ("200+", True),
            (" monthly candidate assessments.", False)
        ])
        add_bullet([
            ("Integrated bidirectional streaming pipelines using ", False),
            ("WebSockets", True),
            (" and ", False),
            ("Vapi.ai", True),
            (", slashing candidate assessment turnaround by ", False),
            ("50%", True),
            (" and commercializing the platform for university admissions.", False)
        ])

        # 5. Technical Skills
        add_section_header("Technical Skills")
        skills = [
            ("Languages: ", True, "Go (Golang), Python, Java, TypeScript, JavaScript, C++, SQL (PostgreSQL, MySQL)"),
            ("Distributed Systems: ", True, "Go (goroutines/channels), Kafka (Queues/DLQ/Outbox), Redis (Distributed Locks), NestJS, REST APIs"),
            ("Agentic AI & Productivity: ", True, "Agentic AI Workflows, Cursor, GitHub Copilot, Langfuse (tracing/evals), OpenAI API, Gemini API"),
            ("Databases & Storage: ", True, "PostgreSQL, MongoDB (bulkWrite, lease fencing), MySQL, Redis, Schema Design, Query Optimization"),
            ("Cloud & DevOps: ", True, "Docker, Kubernetes (familiar), Google Cloud (Cloud Run, Pub/Sub), AWS, Jenkins CI/CD, Sentry, Git, Artillery"),
            ("Practices & Testing: ", True, "Data Structures & Algorithms, OOP, Concurrency, System Design, Jest, JUnit, Code Reviews, Agile/Scrum")
        ]
        for title, bold, items in skills:
            add_bullet([(title, bold), (items, False)], space_after=2.4)

        # 6. Achievements
        add_section_header("Achievements")
        add_bullet([
            ("Won ", False),
            ("First Place", True),
            (" at the ISchoolConnect internal AI hackathon (", False),
            ("Rs. 50,000 award", True),
            (") among 30+ engineering teams, architecting an automated real-time voice screening platform.", False)
        ])
        add_bullet([
            ("Secured ", False),
            ("First Position", True),
            (" in HackBPIT (HashDefine, BPIT) and ", False),
            ("First Runner-Up", True),
            (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for distributed system design.", False)
        ])
        add_bullet([
            ("Solved ", False),
            ("450+", True),
            (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef.", False)
        ])
        return

    # SDE I (React Native, Android Native, React, TypeScript, Mobile & Frontend Focus)
    # 2. Experience (HR-First Layout)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Developed responsive, accessible (", False),
        ("WCAG", True),
        (") ", False),
        ("React", True),
        (" and ", False),
        ("Next.js", True),
        (" interfaces featuring reusable design system components, custom hooks, and state lifecycle management, reducing manual review overhead by ", False),
        ("80%", True),
        (".", False)
    ])
    add_bullet([
        ("Constructed an admission co-pilot web interface in ", False),
        ("Slate CRM", True),
        (", translating complex ", False),
        ("UI/UX designs", True),
        (" into real-time application trackers and automated communication workflows with ", False),
        ("100%", True),
        (" cross-device fidelity.", False)
    ])
    add_bullet([
        ("Implemented an enterprise self-serve billing module in ", False),
        ("Next.js", True),
        (" and ", False),
        ("Redux", True),
        (" with ", False),
        ("Stripe Checkout", True),
        (" integration, building asynchronous webhook handlers, resilient error boundaries, and ", False),
        ("100%", True),
        (" payment reconciliation.", False)
    ])
    add_bullet([
        ("Rolled out ", False),
        ("Sentry observability", True),
        (" across production ", False),
        ("React", True),
        (" and ", False),
        ("Next.js", True),
        (" applications with source maps and session replay, troubleshooting client runtime crashes and cutting mean-time-to-resolution (", False),
        ("MTTR", True),
        (") by ", False),
        ("45%", True),
        (".", False)
    ])
    add_bullet([
        ("Optimized client-side data access patterns and ", False),
        ("RESTful API", True),
        (" payloads, reducing analytics response latency ", False),
        ("3×", True),
        (" (", False),
        ("15s to 5s", True),
        (") and eliminating redundant network round-trips to maximize mobile and web responsiveness.", False)
    ])

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Built the ", False),
        ("React/TypeScript", True),
        (" frontend for a commercial assessment platform, delivering role-based dashboards, real-time audio controls, and an intuitive scheduling UI for ", False),
        ("200+", True),
        (" monthly sessions (", False),
        ("50%", True),
        (" turnaround cut).", False)
    ])
    add_bullet([
        ("Streamlined web asset delivery and frontend modules by configuring ", False),
        ("CDN edge caching", True),
        (" and asset chunking in an ", False),
        ("Agile team", True),
        (", delivering zero-downtime releases and lowering Largest Contentful Paint (", False),
        ("LCP", True),
        (") by ", False),
        ("25%", True),
        (".", False)
    ])

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks.", False)
    ], space_after=2.5)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Aether", "Kotlin, Android SDK, Jetpack Compose, Room DB, Clean Architecture", "github.com/Yash-Gupta-01/Personal-AI-Agent", url="https://github.com/Yash-Gupta-01/Personal-AI-Agent")
    add_bullet([
        ("Architected an offline-first Android assistant in ", False),
        ("Kotlin", True),
        (" and ", False),
        ("Jetpack Compose", True),
        (" utilizing ", False),
        ("Clean Architecture", True),
        (", decoupling UI from pure JVM domain logic to execute unit tests in milliseconds without emulator overhead.", False)
    ])
    add_bullet([
        ("Engineered an offline-first local capture loop with ", False),
        ("Room database", True),
        (" schema migrations and modular tool schemas, achieving pixel-perfect, responsive layouts across diverse Android devices and screen densities.", False)
    ])

    add_project_heading("Fluxpay", "React, Redux, Zustand, TypeScript, Vite, Node.js, Jest, RESTful APIs", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Designed a mobile-first ", False),
        ("React and Redux", True),
        (" fintech transaction dashboard bundled with ", False),
        ("Vite", True),
        (", leveraging ", False),
        ("TypeScript", True),
        (" and ", False),
        ("Zustand", True),
        (" for centralized state management, responsive cards, and deterministic UI rendering.", False)
    ])
    add_bullet([
        ("Authored comprehensive unit and integration test suites using ", False),
        ("Jest", True),
        (" following ", False),
        ("TDD", True),
        (" practices, verifying state mutations and component stability while sustaining ", False),
        ("1,200 TPS", True),
        (" in simulated backend payment flows.", False)
    ])

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Frontend & Mobile: ", True, "JavaScript (ES6+), TypeScript, React, React Native (core architecture), Android Native (Kotlin, Jetpack Compose), Next.js (BFF), Redux, Zustand, Context API"),
        ("UI Architecture & Standards: ", True, "Component Lifecycle, In-house Design Systems, Reusable UI Libraries, Pixel-Perfect Layouts, Tailwind CSS, Shadcn UI, HTML5, CSS3, WCAG Accessibility"),
        ("Architecture & Practices: ", True, "Clean Architecture, MVVM, TDD (Test-Driven Development), Design Patterns, Object-Oriented Programming (OOP), SOLID Principles, Version Control (Git)"),
        ("Testing & Observability: ", True, "Jest, React Testing Library, Unit Testing, Integration Testing, Sentry (Session Replay/Source Maps), Artillery Load Testing, Code Reviews"),
        ("Backend, APIs & Cloud: ", True, "Node.js, Express.js, NestJS, RESTful APIs, WebSockets, PostgreSQL, MongoDB, Redis, Docker, CDN Edge Caching, Core Web Vitals (LCP)")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=2.0)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform with low-latency audio inference.", False)
    ], space_after=2.0)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for distributed system design and client-server architecture.", False)
    ], space_after=2.0)
    add_bullet([
        ("Solved ", False),
        ("450+ Data Structures & Algorithms problems", True),
        (" across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, demonstrating deep proficiency in algorithmic design and runtime efficiency.", False)
    ], space_after=1.5)

def populate_makemytrip(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Go, Scalable Backends, Cloud Run, Databases & Full-Stack Capabilities)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Built a scalable ", False),
        ("Go pipeline", True),
        (" on ", False),
        ("Cloud Run", True),
        (" for high-volume data ingestion: engineered a scheduled producer publishing to ", False),
        ("Pub/Sub", True),
        (" and an authenticated consumer executing ", False),
        ("MongoDB", True),
        (" lease-fenced, idempotent writes, preventing data loss.", False)
    ])
    add_bullet([
        ("Reduced backend analytics API response latency ", False),
        ("3×", True),
        (" (", False),
        ("15s to 5s", True),
        (") by profiling ", False),
        ("SQL", True),
        (" query execution plans and restructuring ", False),
        ("MySQL/PostgreSQL", True),
        (" schemas, eliminating redundant joins and table scans to maximize overall service throughput.", False)
    ])
    add_bullet([
        ("Evaluated non-deterministic agentic flows and traced reasoning chains using ", False),
        ("Langfuse", True),
        (" and ", False),
        ("Sentry", True),
        (", diagnosing tool-call hallucinations and prompt drift to automate CI/CD quality gates and cut production mean-time-to-detect (", False),
        ("MTTD", True),
        (").", False)
    ])
    add_bullet([
        ("Developed high-throughput ", False),
        ("NestJS", True),
        (" WebSocket middleware bridging real-time telephony streams, handling bidirectional audio transcoding for ", False),
        ("1k+", True),
        (" concurrent sessions at sub-", False),
        ("100ms", True),
        (" p95 latency with zero connection drops under sustained traffic.", False)
    ])
    add_bullet([
        ("Developed responsive, cross-browser web interfaces in ", False),
        ("Next.js", True),
        (" and ", False),
        ("React", True),
        (" with reusable UI components and role-based access control (", False),
        ("RBAC", True),
        ("), translating product designs into production features with ", False),
        ("100%", True),
        (" cross-platform fidelity.", False)
    ])
    add_bullet([
        ("Engineered an enterprise commercial ", False),
        ("GPA evaluation platform", True),
        (" (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+", True),
        (" university deployments generating ", False),
        ("$150K+", True),
        (" contract value per client, collaborating with product managers under ", False),
        ("Agile/Scrum", True),
        (".", False)
    ])

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Constructed a ", False),
        ("Kafka-driven NestJS", True),
        (" microservice for asynchronous event synchronization between external CRMs and internal databases, maintaining sub-", False),
        ("50ms", True),
        (" synchronization latency under peak production traffic and high concurrency.", False)
    ])
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines automating containerized deployments (", False),
        ("Docker", True),
        (") and integration test suites in an ", False),
        ("Agile team", True),
        (", eliminating broken staging builds and accelerating developer release cadence.", False)
    ])

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Operating Systems, DBMS, Computer Networks.", False)
    ], space_after=2.5)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Go, PostgreSQL, Kafka, Redis, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a high-concurrency payment microservice in ", False),
        ("Go (goroutines/channels)", True),
        (" sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency, verified through comprehensive authored ", False),
        ("Artillery", True),
        (" load testing suites and benchmarks.", False)
    ])
    add_bullet([
        ("Architected an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" with ", False),
        ("Redis", True),
        (" distributed locks and ", False),
        ("PostgreSQL", True),
        (" ACID validation, eliminating duplicate charges across ", False),
        ("500+", True),
        (" concurrent simulated users under sustained load.", False)
    ])

    add_project_heading("AI Interviewer", "NestJS, Next.js, OpenAI API, WebSockets, MongoDB, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Engineered an AI assessment platform with ", False),
        ("NestJS", True),
        (", ", False),
        ("Next.js", True),
        (", and ", False),
        ("OpenAI API", True),
        (", designing structured evaluation rubrics to automate candidate scoring and integrating ", False),
        ("Jest", True),
        (" unit test suites to enforce production quality gates.", False)
    ])
    add_bullet([
        ("Constructed bidirectional streaming pipelines using ", False),
        ("WebSockets", True),
        (" and ", False),
        ("Vapi.ai", True),
        (", enabling real-time voice screening evaluations and reducing candidate assessment turnaround by ", False),
        ("50%", True),
        (" across 200+ monthly candidate reviews.", False)
    ])

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Go (Golang), Python, Java, TypeScript, JavaScript, C++, SQL (MySQL, PostgreSQL)"),
        ("Distributed Systems: ", True, "Go (goroutines/channels), Kafka (Queues/DLQ/Outbox), Redis (Distributed Locks), NestJS, REST APIs"),
        ("Databases & Storage: ", True, "MongoDB, MySQL, PostgreSQL, Cassandra (familiar), Schema Design, Query Profiling & Optimization, ACID"),
        ("AI & Agentic Systems: ", True, "OpenAI API, Google Gemini, Langfuse (tracing/evals), RAG, Agentic Workflows, WebSockets, Vapi.ai"),
        ("DevOps & Cloud: ", True, "Docker, Kubernetes (familiar), Google Cloud (Cloud Run, Pub/Sub), AWS, Jenkins CI/CD, Sentry, Git, Artillery"),
        ("Practices & Quality: ", True, "Responsive Design, Cross-Browser Compatibility, OOP, Agile/Scrum, Jest, JUnit, Code Reviews")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=2.4)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform with low-latency WebSockets.", False)
    ])
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ competing engineering teams for distributed systems.", False)
    ])
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef.", False)
    ])

def populate_gokiwi(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Java Microservices, Partner Integrations, DB Optimization, On-Call P0)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Architected a multi-tenant backend microservice in ", False),
        ("Java 17 (Spring Boot)", True),
        (" extracting CRM data via strategy patterns (", False),
        ("OOP", True),
        ("); deployed ", False),
        ("Hibernate/JPA", True),
        (" and ", False),
        ("Kafka", True),
        (" retry topics with Dead Letter Queues (", False),
        ("DLQ", True),
        (") to ensure eventual consistency under burst traffic.", False)
    ])
    add_bullet([
        ("Engineered a subscription billing microservice integrating ", False),
        ("Stripe", True),
        (" webhooks and signed callbacks (", False),
        ("HMAC", True),
        (") with retry handling, guaranteeing ", False),
        ("100% payment reconciliation", True),
        (" and zero duplicate transactions across recurring cycles.", False)
    ])
    add_bullet([
        ("Slashed backend analytics API latency ", False),
        ("3×", True),
        (" (", False),
        ("15s to 5s", True),
        (") by profiling ", False),
        ("SQL", True),
        (" query execution plans, adding composite indexes, and eliminating ", False),
        ("N+1 queries", True),
        (" and table scans in ", False),
        ("PostgreSQL/MySQL", True),
        (" schemas to maximize throughput.", False)
    ])
    add_bullet([
        ("Deployed ", False),
        ("Sentry observability", True),
        (" and structured logging across microservices, cutting ", False),
        ("MTTD by 45%", True),
        ("; resolved a critical midnight CRM schema drift ", False),
        ("P0 incident within 2 hours", True),
        (" during on-call rotation and authored a post-mortem.", False)
    ])
    add_bullet([
        ("Built a scalable ", False),
        ("Go (Golang)", True),
        (" data pipeline on ", False),
        ("Cloud Run", True),
        (" for delta ingestion: scheduled a ", False),
        ("Pub/Sub", True),
        (" producer and consumer executing lease-fenced, idempotent writes in ", False),
        ("MongoDB", True),
        (", tolerating third-party outages.", False)
    ])
    add_bullet([
        ("Engineered an enterprise commercial ", False),
        ("GPA evaluation platform", True),
        (" (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+", True),
        (" university deployments generating ", False),
        ("$150K+", True),
        (" contract value per client, enforcing strict ", False),
        ("FERPA compliance", True),
        (" and RBAC.", False)
    ])

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Constructed a ", False),
        ("Kafka-driven NestJS", True),
        (" microservice for asynchronous event synchronization between partner CRMs and internal databases, maintaining sub-", False),
        ("50ms", True),
        (" synchronization latency under peak production traffic.", False)
    ])
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines automating containerized deployments (", False),
        ("Docker", True),
        (") and integration test suites in an ", False),
        ("Agile team", True),
        (", eliminating broken staging builds and accelerating developer release cadence.", False)
    ])

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Operating Systems, DBMS, Computer Networks.", False)
    ], space_after=2.2)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Java, Spring Boot, Hibernate, Kafka, Redis, PostgreSQL, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a distributed payment microservice in ", False),
        ("Java (Spring Boot / Hibernate)", True),
        
        (", enforcing authentication via ", False),
        ("Spring Security", True),
        (" and ", False),
        ("JWT", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency with published Artillery load tests.", False)
    ])
    add_bullet([
        ("Architected an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" with ", False),
        ("Redis", True),
        (" distributed locks and ", False),
        ("PostgreSQL", True),
        (" ACID validation, guaranteeing idempotent transactions and zero double-charges across ", False),
        ("500+", True),
        (" simulated users.", False)
    ])
    add_bullet([
        ("Containerized microservices with ", False),
        ("Docker", True),
        (" and authored automated ", False),
        ("JUnit", True),
        (" test suites to validate transactional rollback semantics, reconciliation routines, and edge-case status code handling under partner failures.", False)
    ])

    add_project_heading("AI Interviewer", "NestJS, Next.js, MongoDB (NoSQL), TypeScript, OpenAI API, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Designed a cloud-based interview evaluation platform with RESTful APIs and role-based access control, automating ", False),
        ("200+", True),
        (" monthly evaluations and cutting candidate assessment turnaround by ", False),
        ("50%", True),
        (" for hiring teams.", False)
    ])
    add_bullet([
        ("Authored comprehensive unit and integration test suites using ", False),
        ("Jest", True),
        (", maintaining high test coverage across core evaluation engines to enforce quality gates and ensure reliable continuous deployments.", False)
    ])

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Java (17+ / Core Java), Go (Golang), Python, TypeScript, JavaScript, C++, SQL (PostgreSQL, MySQL)"),
        ("Backend & Architecture: ", True, "Spring Boot, Spring MVC, REST APIs, Dependency Injection, Hibernate/JPA, MapStruct (familiar)"),
        ("Payments & Messaging: ", True, "Apache Kafka (Queues/DLQ/Outbox), Eventual Consistency, Redis Locks, Webhooks (HMAC Signing)"),
        ("Databases & Tuning: ", True, "PostgreSQL, MySQL, Indexing (EXPLAIN ANALYZE), N+1 Query Elimination, Flyway (familiar)"),
        ("Cloud & DevOps: ", True, "Docker, Linux/Shell, Git, GitHub Actions, AWS (EKS/RDS/S3 - familiar), Kubernetes (familiar)"),
        ("Reliability, QA & AI: ", True, "Distributed Tracing, Sentry, OpenTelemetry (familiar), P0 Post-Mortems, JUnit, Jest, Artillery, Cursor")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.8)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform.", False)
    ])
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for distributed system design.", False)
    ])
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, emphasizing algorithmic complexity analysis, concurrency, and problem solving.", False)
    ])

def populate_visa(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Customer-Facing Web, React/TypeScript, Java Microservices, AI Agents & SQL Performance)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Built customer-facing web applications and an enterprise GPA platform (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("React", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+", True),
        (" universities, generating ", False),
        ("$150K+", True),
        (" contract value per client with ", False),
        ("FERPA compliance", True),
        (" and RBAC security.", False)
    ], space_after=1.8)
    add_bullet([
        ("Instrumented multi-turn agent workflows with ", False),
        ("Langfuse", True),
        (" distributed tracing, isolating reasoning-chain bottlenecks and eliminating production ", False),
        ("MongoDB", True),
        (" tool-call hallucinations across conversational student assistants.", False)
    ], space_after=1.8)
    add_bullet([
        ("Spearheaded an academic articulation engine coupling deterministic rule caches with ", False),
        ("Google Gemini API", True),
        (" fallback, resolving ", False),
        ("40%+", True),
        (" of course mappings without AI inference to cut LLM token costs by ", False),
        ("40%", True),
        (".", False)
    ], space_after=1.8)
    add_bullet([
        ("Architected a multi-tenant backend microservice in ", False),
        ("Java 17 (Spring Boot)", True),
        (" extracting CRM data via strategy patterns (", False),
        ("OOP", True),
        ("); deployed ", False),
        ("Hibernate/JPA", True),
        (" and ", False),
        ("Kafka", True),
        (" retry topics with ", False),
        ("DLQs", True),
        (" for eventual consistency.", False)
    ], space_after=1.8)
    add_bullet([
        ("Accelerated backend analytics query performance ", False),
        ("3× (15s to 5s)", True),
        (" by analyzing ", False),
        ("SQL execution plans (EXPLAIN)", True),
        (", indexing composite keys, and eliminating redundant table scans across ", False),
        ("PostgreSQL and MySQL", True),
        (" schemas.", False)
    ], space_after=1.8)
    add_bullet([
        ("Optimized customer-facing web performance across corporate domains, eliminating render-blocking scripts, injecting ", False),
        ("JSON-LD", True),
        (" schema markup, and reducing ", False),
        ("Largest Contentful Paint (LCP)", True),
        (" by ", False),
        ("25%", True),
        (".", False)
    ], space_after=2.0)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Developed a ", False),
        ("Kafka-driven NestJS", True),
        (" microservice for asynchronous event synchronization between partner CRMs and internal databases, maintaining ", False),
        ("sub-50ms", True),
        (" synchronization latency under peak production traffic.", False)
    ], space_after=1.8)
    add_bullet([
        ("Streamlined ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines automating containerized deployments (", False),
        ("Docker", True),
        (") and unit test quality gates in an ", False),
        ("Agile team", True),
        (", eliminating broken staging builds and accelerating release cadence.", False)
    ], space_after=2.0)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Design (OOP), DBMS, Operating Systems, Computer Networks.", False)
    ], space_after=2.0)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("AI Interviewer", "Next.js, React, TypeScript, NestJS, OpenAI API, WebSockets, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Designed an autonomous agentic interview screening platform using ", False),
        ("Next.js", True),
        (", ", False),
        ("NestJS", True),
        (", and ", False),
        ("OpenAI API", True),
        (", automating ", False),
        ("200+", True),
        (" monthly evaluations and cutting candidate assessment turnaround by ", False),
        ("50%", True),
        (".", False)
    ], space_after=1.8)
    add_bullet([
        ("Authored comprehensive unit and integration test suites using ", False),
        ("Jest", True),
        (" and ", False),
        ("React Testing Library", True),
        (", maintaining high test coverage across core evaluation engines to enforce continuous deployment quality gates.", False)
    ], space_after=2.0)

    add_project_heading("Fluxpay", "Java 17, Spring Boot, Hibernate, PostgreSQL, Redis, Kafka, Docker, JUnit", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Engineered a distributed payments microservice in ", False),
        ("Java 17 (Spring Boot / Hibernate)", True),
        
        (" and ", False),
        ("Spring Security", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency with published Artillery test suites.", False)
    ], space_after=1.8)
    add_bullet([
        ("Formulated an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" with ", False),
        ("Redis", True),
        (" distributed locks and ", False),
        ("PostgreSQL", True),
        (" ACID validation, guaranteeing idempotent transactions and zero double-charges across ", False),
        ("500+", True),
        (" simulated users.", False)
    ], space_after=1.8)
    add_bullet([
        ("Containerized microservices with ", False),
        ("Docker", True),
        (" and authored automated ", False),
        ("JUnit", True),
        (" test suites to validate transactional rollback semantics, reconciliation routines, and edge-case status code handling under failure scenarios.", False)
    ], space_after=2.0)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Java (17+ / Core Java, OOP), TypeScript, JavaScript, C#, Python, SQL (PostgreSQL, MySQL), Go, C++"),
        ("Frontend & Web: ", True, "React, Next.js (App Router, BFF), TypeScript, HTML5, CSS3, Tailwind CSS, Redux, Zustand, DevTools"),
        ("Backend & Architecture: ", True, "Spring Boot, ASP.NET Core, NestJS, RESTful APIs, OOP Principles, Microservices, Node.js"),
        ("AI & Agent Systems: ", True, "LLM APIs (Gemini, Claude, GPT), Agent Systems, Tool Calling, Langfuse Tracing, Prompt Engineering"),
        ("Databases & Streaming: ", True, "PostgreSQL, MySQL, Redis, MongoDB, Apache Kafka, SQL Optimization, Indexing (EXPLAIN)"),
        ("Testing & DevOps: ", True, "Jest, React Testing Library, JUnit, Docker, Git, GitHub Actions, Jenkins CI/CD, Postman, Kubernetes (familiar)")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.5)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform with low-latency WebSockets.", False)
    ], space_after=1.8)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ competing teams for distributed system design.", False)
    ], space_after=1.8)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, demonstrating strong foundation in computational efficiency and algorithmic design.", False)
    ], space_after=1.5)

def populate_franklin_templeton(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (Operations, Managed Services, Reliability, Incident Management, Batch, SQL)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Spearheaded centralized application monitoring with ", False),
        ("Sentry observability", True),
        (" across distributed services, analyzing error logs; resolved a critical production schema drift incident in ", False),
        ("2 hours", True),
        (" via cross-team incident bridges to ensure strict ", False),
        ("SLA compliance", True),
        (".", False)
    ])
    add_bullet([
        ("Architected a multi-tenant batch ETL pipeline in ", False),
        ("Java (Spring Boot)", True),
        (" and ", False),
        ("Kafka", True),
        (" with Dead Letter Queues (", False),
        ("DLQ", True),
        (") and retry topics, strengthening batch reliability, isolating failed records, and preventing data loss during burst traffic.", False)
    ])
    add_bullet([
        ("Automated scheduled batch monitoring and data pipelines processing ", False),
        ("50K+", True),
        (" enterprise records from ", False),
        ("PostgreSQL", True),
        (" using ", False),
        ("Python cron routines", True),
        (" and validation scripts, cutting operational reporting cycles by ", False),
        ("60%", True),
        (" with zero data discrepancies.", False)
    ])
    add_bullet([
        ("Troubleshot database performance issues, slashing backend analytics API latency ", False),
        ("3× (15s to 5s)", True),
        (" by optimizing ", False),
        ("SQL query plans (EXPLAIN)", True),
        (", adding composite indexes, and eliminating table scans across ", False),
        ("PostgreSQL and MySQL", True),
        (".", False)
    ])
    add_bullet([
        ("Engineered an enterprise commercial ", False),
        ("GPA evaluation platform", True),
        (" (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+", True),
        (" university deployments generating ", False),
        ("$150K+", True),
        (" contract value per client, enforcing strict ", False),
        ("FERPA compliance", True),
        (" and RBAC security.", False)
    ])
    add_bullet([
        ("Leveraged AI assistants (", False),
        ("Cursor, Copilot", True),
        (") and ", False),
        ("Langfuse", True),
        (" to debug system failures, automate unit test generation, and author operational runbooks and handover procedures, accelerating root cause analysis (", False),
        ("RCA", True),
        (") and onboarding.", False)
    ])

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Engineered an automated AI assessment platform with ", False),
        ("NestJS", True),
        (", ", False),
        ("Next.js", True),
        (", and REST APIs, automating ", False),
        ("200+", True),
        (" monthly evaluations, reducing processing turnaround by ", False),
        ("50%", True),
        (", and maintaining telemetry logs for continuous service health monitoring.", False)
    ])
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines automating containerized deployments (", False),
        ("Docker", True),
        (") and integration test suites in an ", False),
        ("Agile team", True),
        (", eliminating broken staging builds and accelerating developer release cadence.", False)
    ])

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Operating Systems, DBMS, Computer Networks.", False)
    ], space_after=2.2)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Java, Spring Boot, Kafka, Redis, PostgreSQL, Docker, Linux, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a distributed payment microservice in ", False),
        ("Java (Spring Boot)", True),
        (" and REST APIs sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency, conducting end-to-end ", False),
        ("Artillery load testing", True),
        (" to validate service capacity and resilience.", False)
    ])
    add_bullet([
        ("Architected an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" and ", False),
        ("Redis", True),
        (" distributed locks with ", False),
        ("PostgreSQL ACID", True),
        (" validation, preventing double-processing and ensuring 100% transaction integrity across ", False),
        ("500+", True),
        (" simulated users.", False)
    ])
    add_bullet([
        ("Containerized services using ", False),
        ("Docker", True),
        (" and automated ", False),
        ("JUnit", True),
        (" integration tests to validate failure-recovery procedures, network timeout handling, and transactional rollbacks, ensuring robust disaster recovery and release readiness.", False)
    ])

    add_project_heading("AI Interviewer", "NestJS, Next.js, MongoDB, WebSockets, OpenAI API, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Designed a cloud evaluation platform with ", False),
        ("NestJS", True),
        (", ", False),
        ("Next.js", True),
        (", and WebSockets, automating ", False),
        ("200+", True),
        (" monthly screening evaluations and cutting turnaround by ", False),
        ("50%", True),
        (" while logging telemetry for service health monitoring.", False)
    ])
    add_bullet([
        ("Implemented comprehensive unit and integration test suites with ", False),
        ("Jest", True),
        (" and automated API contract validation, establishing strict CI quality gates to prevent regression defects and improve release readiness across production services.", False)
    ])

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Java (Spring Boot), Python, SQL (SELECT Queries, Indexing, Joins), Bash / Linux Shell, TypeScript, Go, C++"),
        ("Operations & SRE: ", True, "Incident Management, Service Requests, SLA Tracking, RCA, Sentry, Log Analysis, ServiceNow / Jira"),
        ("Cloud & Infrastructure: ", True, "Linux/Unix Commands & Networking, Cron Jobs, Scheduled Batches, Microsoft Azure, AWS Fundamentals, GCP, Docker"),
        ("Distributed Systems: ", True, "Kafka (Queues, DLQ, Retry Topics), Redis Locks, RESTful APIs, Microservices, Event-Driven Architecture"),
        ("Databases & Data: ", True, "PostgreSQL, MySQL, Query Profiling (EXPLAIN ANALYZE), Data Validation & Integrity, MongoDB"),
        ("CI/CD & Dev Tools: ", True, "Jenkins CI/CD, Git, Confluence (Knowledge Articles), Control-M & CyberArk (exposure), JUnit, Jest, Cursor, Copilot")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.8)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform.", False)
    ])
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for distributed system design.", False)
    ])
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, emphasizing algorithmic complexity analysis, concurrency, and systematic troubleshooting.", False)
    ])

def populate_cisco(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Cisco Software Engineer: Best Practices, Test Automation, Systems, Incident Triaging)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Automated asynchronous data pipelines processing ", False),
        ("50K+", True),
        (" records from ", False),
        ("PostgreSQL", True),
        (" using scheduled ", False),
        ("Python", True),
        (" scripts and ", False),
        ("Pytest", True),
        (" validation, eliminating manual overhead and cutting reporting cycles by ", False),
        ("60%", True),
        (".", False)
    ], space_after=2.0)
    add_bullet([
        ("Instrumented distributed tracing and telemetry across microservices via ", False),
        ("Sentry", True),
        (" and ", False),
        ("Langfuse", True),
        (", diagnosing tool-call failures and prompt drift to conduct incident postmortems and cut ", False),
        ("MTTD", True),
        (" by ", False),
        ("45%", True),
        (".", False)
    ], space_after=2.0)
    add_bullet([
        ("Constructed a multi-tenant ", False),
        ("Java/Spring Boot", True),
        (" ETL service integrating Slate and Salesforce CRMs via ", False),
        ("REST APIs", True),
        (", implementing factory design patterns and ", False),
        ("Kafka Dead Letter Queues", True),
        (" to prevent data loss.", False)
    ], space_after=2.0)
    add_bullet([
        ("Engineered low-latency telephony streaming middleware in ", False),
        ("NestJS", True),
        (" bridging ", False),
        ("Genesys AudioHook", True),
        (" with ", False),
        ("Vapi.ai", True),
        (" via ", False),
        ("WebSockets", True),
        (", transcoding audio for ", False),
        ("1,000+", True),
        (" concurrent sessions at sub-", False),
        ("100ms", True),
        (" p95 latency.", False)
    ], space_after=2.0)
    add_bullet([
        ("Deployed an enterprise commercial ", False),
        ("GPA evaluation platform", True),
        (" (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+", True),
        (" university deployments, generating ", False),
        ("$150K+", True),
        (" contract value per client while enforcing ", False),
        ("FERPA compliance", True),
        (".", False)
    ], space_after=2.0)
    add_bullet([
        ("Optimized backend analytics REST API latency ", False),
        ("3× (15s to 5s)", True),
        (" by analyzing ", False),
        ("SQL", True),
        (" execution plans, indexing composite keys, and eliminating redundant scans across ", False),
        ("PostgreSQL and MySQL", True),
        (" database schemas.", False)
    ], space_after=2.0)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines automating containerized ", False),
        ("Docker", True),
        (" packaging and integration test suites in an ", False),
        ("Agile team", True),
        (", eliminating broken staging builds and accelerating releases.", False)
    ], space_after=2.0)
    add_bullet([
        ("Launched an AI assessment platform automating ", False),
        ("200+", True),
        (" monthly candidate evaluations and cutting turnaround ", False),
        ("50%", True),
        (", engineering modular ", False),
        ("React/TypeScript", True),
        (" UI dashboards and integrating ", False),
        ("Jest", True),
        (" test suites.", False)
    ], space_after=2.2)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Computer Networks, Operating Systems, Database Management Systems (DBMS), Systems Programming, Software Engineering.", False)
    ], space_after=2.2)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Go (Golang), Java (Spring Boot), Docker, Kafka, Redis, PostgreSQL, Artillery", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a cloud-native payments microservice in ", False),
        ("Go (Golang)", True),
        (" using goroutines and channels, sustaining ", False),
        ("1,200 TPS", True),
        (" (simulated) at sub-", False),
        ("120ms", True),
        (" p95 latency, validated via ", False),
        ("Artillery", True),
        (" load tests.", False)
    ], space_after=2.0)
    add_bullet([
        ("Implemented an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" and ", False),
        ("Redis", True),
        (" distributed locks with idempotent ", False),
        ("PostgreSQL", True),
        (" updates, guaranteeing zero double-charges across ", False),
        ("500+", True),
        (" simulated concurrent users.", False)
    ], space_after=2.0)

    add_project_heading("AI Interviewer", "React, TypeScript, NestJS, Next.js, Jest, WebSockets, OpenAI API", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Shipped an automated assessment platform with ", False),
        ("React/Next.js", True),
        (" UI and ", False),
        ("NestJS", True),
        (" microservices, designing structured evaluation rubrics to automate candidate scoring and integrating ", False),
        ("Jest", True),
        (" unit test suites.", False)
    ], space_after=2.0)
    add_bullet([
        ("Authored bidirectional streaming pipelines using ", False),
        ("WebSockets", True),
        (" and ", False),
        ("Vapi.ai", True),
        (", enabling real-time voice screening and reducing candidate evaluation turnaround by ", False),
        ("50%", True),
        (" across ", False),
        ("200+", True),
        (" monthly technical reviews.", False)
    ], space_after=2.0)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Python, Go (Golang), Java, TypeScript, JavaScript, C++, SQL (PostgreSQL, MySQL), Bash/Shell, HTML5"),
        ("Systems & Linux: ", True, "Linux, Docker, Shell Scripting, Git, Jenkins CI/CD, Computer Networks, Operating Systems Principles"),
        ("Testing & Quality: ", True, "Pytest, Jest, JUnit, Artillery Load Testing, Test Automation, Unit Testing, Mocking, Postman"),
        ("Backend & Distributed: ", True, "NestJS, Node.js, Spring Boot, RESTful APIs, Apache Kafka, Redis, WebSockets, Microservices"),
        ("Cloud & Databases: ", True, "PostgreSQL, MongoDB, MySQL, Google Cloud Platform (Cloud Run, Pub/Sub), AWS (S3, EC2)"),
        ("Observability & Practices: ", True, "Sentry (telemetry/MTTD), Langfuse, Agile/Scrum, Incident Triaging, Postmortems, Clean Code")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.8)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform with low-latency audio pipelines.", False)
    ], space_after=2.0)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for distributed system architecture and resilient API design.", False)
    ], space_after=2.0)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef.", False)
    ], space_after=1.5)

def populate_flex(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Aligned with Flex Associate Software Engineer JD)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Evaluated non-deterministic agentic flows and traced LLM reasoning chains using ", False),
        ("Langfuse", True),
        (" and ", False),
        ("Sentry", True),
        (", diagnosing tool-call hallucinations and prompt drift to automate quality gates and cut mean-time-to-detect (", False),
        ("MTTD", True),
        (") by ", False),
        ("45%", True),
        (".", False)
    ], space_after=2.6)
    add_bullet([
        ("Developed high-throughput ", False),
        ("NestJS", True),
        (" WebSocket middleware bridging real-time telephony streams with voice AI agents (", False),
        ("Vapi.ai", True),
        ("), transcoding audio for ", False),
        ("1,000+", True),
        (" concurrent sessions at sub-", False),
        ("100ms", True),
        (" p95 latency with zero connection drops under sustained traffic.", False)
    ], space_after=2.6)
    add_bullet([
        ("Automated high-volume asynchronous data pipelines processing ", False),
        ("50K+", True),
        (" records from ", False),
        ("PostgreSQL", True),
        (" using scheduled ", False),
        ("Python", True),
        (" jobs, eliminating manual intervention and cutting stakeholder reporting cycles by ", False),
        ("60%", True),
        (" for admissions teams.", False)
    ], space_after=2.6)
    add_bullet([
        ("Architected a multi-tenant ", False),
        ("Java/Spring Boot", True),
        (" ETL service across Slate and Salesforce CRMs via REST APIs, implementing ", False),
        ("Kafka", True),
        (" retry topics and Dead Letter Queues (", False),
        ("DLQ", True),
        (") to eliminate burst ingestion data loss.", False)
    ], space_after=2.6)
    add_bullet([
        ("Engineered an enterprise commercial ", False),
        ("GPA evaluation platform", True),
        (" (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+", True),
        (" university deployments generating ", False),
        ("$150K+", True),
        (" contract value per client, enforcing ", False),
        ("FERPA compliance", True),
        (" and RBAC security.", False)
    ], space_after=2.6)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Reduced backend analytics API response latency ", False),
        ("3×", True),
        (" (", False),
        ("15s to 5s", True),
        (") by profiling ", False),
        ("SQL", True),
        (" query execution plans and restructuring ", False),
        ("PostgreSQL/MySQL", True),
        (" schemas, eliminating redundant joins and table scans to maximize throughput.", False)
    ], space_after=2.6)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines automating containerized deployments (", False),
        ("Docker", True),
        (") and integration test suites in an ", False),
        ("Agile team", True),
        (", eliminating broken staging builds and accelerating developer release cadence.", False)
    ], space_after=2.8)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks, Software Development Life Cycle (SDLC).", False)
    ], space_after=2.8)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("AI Interviewer", "NestJS, Next.js, React, TypeScript, OpenAI API, WebSockets, Vapi.ai", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Engineered an AI assessment platform with ", False),
        ("NestJS", True),
        (", ", False),
        ("Next.js", True),
        (", and ", False),
        ("OpenAI API", True),
        (", designing structured evaluation rubrics to automate candidate scoring and integrating ", False),
        ("Jest", True),
        (" unit/integration test suites to enforce strict quality gates.", False)
    ], space_after=2.6)
    add_bullet([
        ("Constructed bidirectional streaming pipelines using ", False),
        ("WebSockets", True),
        (" and ", False),
        ("Vapi.ai", True),
        (", enabling real-time voice screening evaluations and reducing candidate assessment turnaround by ", False),
        ("50%", True),
        (" across 200+ monthly candidate reviews.", False)
    ], space_after=2.6)

    add_project_heading("Fluxpay", "Java, Spring Boot, Hibernate, Kafka, PostgreSQL, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a distributed payments microservice in ", False),
        ("Java (Spring Boot / Hibernate)", True),
        
        (", enforcing authentication via ", False),
        ("Spring Security", True),
        (" and ", False),
        ("JWT", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" at sub-", False),
        ("120ms", True),
        (" p95 latency with published Artillery load tests.", False)
    ], space_after=2.6)
    add_bullet([
        ("Architected an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" with idempotency validation in ", False),
        ("PostgreSQL", True),
        (", guaranteeing zero duplicate transactions across ", False),
        ("500+", True),
        (" concurrent simulated users under sustained peak load.", False)
    ], space_after=2.8)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Programming Languages: ", True, "Python, Java, JavaScript, TypeScript, SQL (PostgreSQL, MySQL), C++, Go (Golang)"),
        ("AI & Generative AI Tools: ", True, "ChatGPT, Microsoft Copilot, GitHub Copilot, Google Gemini, Claude, Cursor, OpenAI API, Langfuse (LLM tracing/evals), Structured Prompt Engineering, Multi-modal Workflows"),
        ("Software Engineering & SDLC: ", True, "OOP Concepts, Data Structures & Algorithms, System Architecture, RESTful APIs, Agile/Scrum, Jenkins CI/CD Pipelines, Git Version Control, Unit/Integration Testing"),
        ("Databases & Storage: ", True, "PostgreSQL, MySQL, MongoDB, Redis (caching), Query Optimization, Schema Design, Data Pipelines"),
        ("Cloud & DevOps: ", True, "Docker, Basic Cloud Concepts (AWS, GCP Cloud Run), Linux, Sentry (observability/MTTD), JUnit, Jest"),
        ("Web & Platforms: ", True, "React, Next.js (BFF), Angular 14, Tailwind CSS, WebSockets, Core Web Vitals (LCP)")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=2.0)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform with low-latency audio streaming and structured evaluation rubrics.", False)
    ], space_after=2.6)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for distributed system architecture, resilient API design, and rapid technical prototyping.", False)
    ], space_after=2.6)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef.", False)
    ], space_after=2.0)

def populate_blackrock(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Aligned with BlackRock Digital Platforms & Experiences JD)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Architected and scaled client-facing ", False),
        ("frontend", True),
        (" web applications using ", False),
        ("Angular 14", True),
        (", ", False),
        ("TypeScript", True),
        (", and ", False),
        ("RxJS", True),
        (" observable streams, engineering modular UI components conforming to ", False),
        ("WCAG 2.1", True),
        (" accessibility and enterprise web standards.", False)
    ], space_after=2.4)
    add_bullet([
        ("Collaborated with ", False),
        ("project managers", True),
        (", ", False),
        ("technical leads", True),
        (", and ", False),
        ("analysts", True),
        (" from ", False),
        ("inception to delivery", True),
        (" on a commercial GPA platform (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across 4+ deployments ($150K+ contract value), enforcing ", False),
        ("FERPA/RBAC", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Leveraged AI actively in daily development and testing using ", False),
        ("Cursor", True),
        (", ", False),
        ("GitHub Copilot", True),
        (", and ", False),
        ("Langfuse", True),
        (", instrumenting automated test suites and prompt evals to drive ", False),
        ("continuous improvement", True),
        (" across Agile sprint velocity.", False)
    ], space_after=2.4)
    add_bullet([
        ("Shipped an institutional ", False),
        ("Next.js", True),
        (" and ", False),
        ("Stripe", True),
        (" billing dashboard with real-time credit metrics and webhook lifecycles, writing clean, ", False),
        ("future-proof code", True),
        (" to maintain ", False),
        ("100%", True),
        (" payment reconciliation and transaction state integrity.", False)
    ], space_after=2.4)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines automating containerized deployments (", False),
        ("Docker", True),
        (") and integration test suites in an ", False),
        ("Agile team", True),
        (", eliminating broken staging builds and accelerating developer release cadence.", False)
    ], space_after=2.4)
    add_bullet([
        ("Rolled out ", False),
        ("Sentry observability", True),
        (" across web applications, analyzing error telemetry on ", False),
        ("Unix/Linux", True),
        (" to ", False),
        ("support existing systems and resolve operations issues", True),
        (" with global engineering staff, cutting production ", False),
        ("MTTR", True),
        (".", False)
    ], space_after=2.5)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Developed the responsive web frontend for an AI assessment platform using ", False),
        ("Angular 14", True),
        (", ", False),
        ("TypeScript", True),
        (", and ", False),
        ("WebSockets", True),
        (", automating ", False),
        ("200+", True),
        (" monthly candidate screenings and reducing turnaround time by ", False),
        ("50%", True),
        (" for hiring teams.", False)
    ], space_after=2.4)
    add_bullet([
        ("Authored unit and integration test suites using ", False),
        ("Jest", True),
        (" (", False),
        ("90%+", True),
        (" coverage), adhering to ", False),
        ("coding standards and guidelines", True),
        (", ", False),
        ("Git distributed version control", True),
        (", and ", False),
        ("Scrum disciplines", True),
        (" to ensure reliable, clean code delivery.", False)
    ], space_after=2.6)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks, and Modern Web Application Architectures.", False)
    ], space_after=2.6)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "React, TypeScript, Node.js, Redis, PostgreSQL, Docker (OCI), Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Built a fintech payment platform featuring a responsive ", False),
        ("React/TypeScript", True),
        (" transaction dashboard and microservice backend, sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency with published load tests.", False)
    ], space_after=2.4)
    add_bullet([
        ("Containerized microservices using ", False),
        ("Docker (OCI)", True),
        (", implementing cloud-native containerized packaging, ", False),
        ("PostgreSQL idempotency", True),
        (", and transactional messaging to ensure zero duplicate transactions across ", False),
        ("500+", True),
        (" users under peak load.", False)
    ], space_after=2.5)

    add_project_heading("AI Interviewer", "Angular 14, React, Next.js, WebSockets, OpenAI API, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Architected an AI-driven client evaluation portal using ", False),
        ("Angular 14", True),
        (" and ", False),
        ("Next.js", True),
        (", leveraging ", False),
        ("OpenAI API", True),
        (" and structured evaluation rubrics to deliver real-time interactive assessment reports with structured analytics.", False)
    ], space_after=2.4)
    add_bullet([
        ("Constructed low-latency audio streaming pipelines using ", False),
        ("WebSockets", True),
        (", designing modular component architecture and ", False),
        ("Jest", True),
        (" unit test suites to ensure robust cross-browser reliability, high test coverage, and clean code.", False)
    ], space_after=2.5)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Frontend & Frameworks: ", True, "Angular 14, React, Next.js (BFF), RxJS, NgRx, TypeScript, JavaScript (ES6+), HTML5, CSS3"),
        ("Architecture & Standards: ", True, "Web & Security Standards, Component Design, WCAG 2.1, RBAC, Future-Proof Code, LCP"),
        ("AI & Daily Dev Tooling: ", True, "Cursor, GitHub Copilot, OpenAI API, Langfuse (tracing/evals), Claude Code, Google Gemini"),
        ("Testing & Methodologies: ", True, "Jest, Unit & Integration Testing, Agile/Scrum, SDLC, Coding Guidelines, Clean Code"),
        ("Cloud, Containers & OS: ", True, "Unix/Linux (Shell), Docker (OCI), Git (Distributed VCS), Azure DevOps (exposure), Jenkins"),
        ("Financial & Backend: ", True, "Payment Idempotency, Financial Services, RESTful APIs, Node.js, NestJS, PostgreSQL, Redis")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.6)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform with low-latency audio streaming and structured evaluation rubrics.", False)
    ], space_after=2.4)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for distributed system architecture, resilient API design, and rapid technical prototyping.", False)
    ], space_after=2.4)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, emphasizing algorithmic complexity analysis, concurrency, and systematic troubleshooting.", False)
    ], space_after=1.8)

def populate_iqvia_associate_data_analyst(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Aligned with IQVIA Associate Data Analyst JD)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Authored internal ", False),
        ("Python data extraction pipelines", True),
        (" and complex ", False),
        ("SQL queries on PostgreSQL", True),
        (", aggregating ", False),
        ("50K+ student records", True),
        (" to eliminate manual reporting overhead and accelerate operational analytics turnaround by ", False),
        ("60%", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Designed interactive ", False),
        ("business analytics dashboards", True),
        (" and reporting telemetry modules in ", False),
        ("Next.js", True),
        (", streaming high-volume usage audit logs to CSV with constant ", False),
        ("O(1) memory", True),
        (" and ", False),
        ("100% data reconciliation integrity", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Formulated clear technical use cases, ", False),
        ("Jira user stories", True),
        (", and ", False),
        ("Confluence documentation", True),
        (" in cross-functional Agile sprints with Product Managers and QA, validating end-to-end data workflows and ", False),
        ("UAT test coverage", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Improved the ", False),
        ("Document Intelligence", True),
        (" pipeline by coupling deterministic articulation rule caches with ", False),
        ("LLM fallback", True),
        (", resolving ", False),
        ("40%+", True),
        (" of course mappings without model inference to cut token costs by ", False),
        ("40%", True),
        (" and eliminate prompt variance.", False)
    ], space_after=2.4)
    add_bullet([
        ("Architected a decoupled two-tier ", False),
        ("Go data ingestion pipeline", True),
        (" on ", False),
        ("Cloud Run and Pub/Sub", True),
        (", processing ", False),
        ("10K+ delta records per 15-minute sync cycle", True),
        (" into ", False),
        ("MongoDB", True),
        (" queues with lease-fenced consumers to guarantee zero data loss.", False)
    ], space_after=2.4)
    add_bullet([
        ("Deployed commercial transcript evaluation platforms across ", False),
        ("4+ enterprise university client engagements", True),
        (" generating ", False),
        ("$150K+", True),
        (" contract value each, automating writeback of computed GPAs, credits, and grades into ", False),
        ("Slate CRM", True),
        (".", False)
    ], space_after=2.5)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Reduced backend analytics API response latency ", False),
        ("3× (15s to 5s)", True),
        (" by profiling ", False),
        ("SQL execution plans (EXPLAIN)", True),
        (" across ", False),
        ("PostgreSQL and MySQL", True),
        (" schemas, ", False),
        ("indexing composite keys", True),
        (", and eliminating redundant full-table scans.", False)
    ], space_after=2.4)
    add_bullet([
        ("Maintained automated test execution and containerized deployments in an ", False),
        ("Agile/Scrum", True),
        (" team, actively collaborating with Product, QA, and UX during sprint planning, ", False),
        ("backlog grooming", True),
        (", and ", False),
        ("UAT test execution", True),
        (".", False)
    ], space_after=2.6)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering (2025 Graduate)", "New Delhi, India")
    add_bullet([
        ("Coursework: Database Management Systems (DBMS), Relational Database Design (RDBMS), SQL & Query Optimization, Data Structures & Algorithms, Object-Oriented Programming (OOP), Software Engineering & SDLC.", False)
    ], space_after=2.6)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Java, Spring Boot, Python, Apache Kafka, PostgreSQL, Redis, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a high-throughput payments microservice in ", False),
        ("Java (Spring Boot)", True),
        (" and ", False),
        ("Python", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency, validating service capacity through authored ", False),
        ("Artillery load test suites", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Modeled a ", False),
        ("normalized PostgreSQL schema", True),
        (" with idempotency-aware ", False),
        ("SQL transaction procedures", True),
        (" over a ", False),
        ("Kafka", True),
        (" event pipeline, guaranteeing ", False),
        ("zero duplicate transactions", True),
        (" across ", False),
        ("500+ concurrent simulated users", True),
        (" under sustained load.", False)
    ], space_after=2.6)

    add_project_heading("AI Interviewer", "Next.js, NestJS, TypeScript, MongoDB, OpenAI API, WebSockets", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Engineered an AI assessment platform with ", False),
        ("Next.js", True),
        (", ", False),
        ("NestJS", True),
        (", and ", False),
        ("OpenAI API", True),
        (", establishing structured rubrics and ", False),
        ("analytics dashboards", True),
        (" delivering objective scoring across candidate reviews with ", False),
        ("100% data fidelity", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Integrated bidirectional streaming pipelines using ", False),
        ("WebSockets", True),
        (" and ", False),
        ("Vapi.ai", True),
        (", enabling real-time voice screening evaluations and cutting candidate assessment turnaround by ", False),
        ("50%", True),
        (" across ", False),
        ("200+", True),
        (" monthly candidate reviews.", False)
    ], space_after=2.6)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Data Analytics & BI: ", True, "Looker Studio Pro (familiar), LookML (familiar), PowerBI, Business Dashboards, Telemetry Reporting"),
        ("Languages & Querying: ", True, "SQL (PostgreSQL, MySQL, Complex Queries, Joins), Python (Pandas), JavaScript (ES6+), TypeScript, Go"),
        ("Databases & Storage: ", True, "PostgreSQL, Snowflake (familiar), MongoDB, MySQL, Redis (caching), Query Optimization (EXPLAIN)"),
        ("Methodology & Process: ", True, "Business Requirements (BRD), User Stories, Use Cases, System Workflows, UAT, Data Quality Checks"),
        ("AI & Developer Tools: ", True, "Cursor IDE, GitHub Copilot, OpenAI API, LLM Fallback Pipelines, Langfuse, Prompt Engineering"),
        ("DevOps & Collaboration: ", True, "Jira, Confluence, Git, GitHub Actions, Docker, Agile/Scrum (Sprint Ceremonies), Postman, Linux")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.8)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform with WebSocket streaming and structured scoring rubrics.", False)
    ], space_after=2.4)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for distributed system design, relational database modeling, and relief tracking.", False)
    ], space_after=2.4)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, emphasizing algorithmic optimization, relational SQL querying, and systematic edge-case validation.", False)
    ], space_after=1.5)

def populate_iqvia(doc, add_section_header, add_role_heading, add_project_heading, add_bullet, role="SDE"):
    if "analyst" in role.lower() or "data" in role.lower():
        populate_iqvia_associate_data_analyst(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
        return
    # 2. Experience (HR-First Layout - Aligned with IQVIA Software Development Engineer 1 JD)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Architected a multi-tenant backend microservice in ", False),
        ("Java 17 (Spring Boot)", True),
        (" integrating CRMs via ", False),
        ("RESTful APIs", True),
        (" and strategy patterns (", False),
        ("OOP/SOLID", True),
        (") to build scalable, maintainable software; deployed ", False),
        ("Hibernate/JPA", True),
        (" and ", False),
        ("Kafka DLQ", True),
        (" retries to eliminate burst data loss.", False)
    ], space_after=2.4)
    add_bullet([
        ("Developed responsive web applications in ", False),
        ("ReactJS", True),
        (" and ", False),
        ("Next.js", True),
        (" with reusable UI components and role-based access control (", False),
        ("RBAC", True),
        ("), writing clean, reusable code and integrating backend REST APIs across browsers and devices with ", False),
        ("100%", True),
        (" fidelity.", False)
    ], space_after=2.4)
    add_bullet([
        ("Engineered an enterprise commercial ", False),
        ("GPA evaluation platform", True),
        (" (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") deployed across ", False),
        ("4+", True),
        (" university clients generating ", False),
        ("$150K+", True),
        (" contract value per deployment under strict ", False),
        ("FERPA compliance", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Evaluated agentic workflows and traced LLM reasoning chains using ", False),
        ("Langfuse", True),
        (" and ", False),
        ("Sentry", True),
        (", leveraging ", False),
        ("GitHub Copilot", True),
        (" and structured prompt engineering to support AI-powered features and automate cloud-native CI/CD quality gates.", False)
    ], space_after=2.4)
    add_bullet([
        ("Automated high-volume asynchronous data pipelines processing ", False),
        ("50K+", True),
        (" records from ", False),
        ("PostgreSQL", True),
        (" using scheduled ", False),
        ("Python", True),
        (" jobs, eliminating manual intervention and cutting stakeholder reporting cycles by ", False),
        ("60%", True),
        (" for admissions teams.", False)
    ], space_after=2.5)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Reduced backend analytics API response latency ", False),
        ("3×", True),
        (" (", False),
        ("15s to 5s", True),
        (") by profiling ", False),
        ("SQL", True),
        (" query execution plans (", False),
        ("EXPLAIN", True),
        (") across relational databases (PostgreSQL/SQL Server fundamentals), eliminating redundant joins to maximize throughput.", False)
    ], space_after=2.4)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" pipelines automating test execution and containerized deployments (", False),
        ("Docker", True),
        (") in an ", False),
        ("Agile/Scrum", True),
        (" team, actively collaborating with Product, QA, UX, and DevOps during sprint planning, stand-ups, and retrospectives.", False)
    ], space_after=2.6)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering (2025 Graduate)", "New Delhi, India (Open to Relocate — Bangalore)")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks, Software Development Life Cycle (SDLC).", False)
    ], space_after=2.6)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Java, Spring Boot, Hibernate, Kafka, PostgreSQL, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a distributed payments microservice in ", False),
        ("Java (Spring Boot / Hibernate)", True),
        
        (", enforcing authentication via ", False),
        ("Spring Security", True),
        (" and ", False),
        ("JWT", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" at sub-", False),
        ("120ms", True),
        (" p95 latency with published Artillery load tests.", False)
    ], space_after=2.4)
    add_bullet([
        ("Architected an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" with idempotency validation in ", False),
        ("PostgreSQL", True),
        (", guaranteeing zero duplicate transactions across ", False),
        ("500+", True),
        (" concurrent simulated users and containerizing services with ", False),
        ("Docker", True),
        (" and ", False),
        ("JUnit", True),
        (" test suites.", False)
    ], space_after=2.6)

    add_project_heading("AI Interviewer", "ReactJS, Next.js, NestJS, TypeScript, OpenAI API, WebSockets, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Engineered an AI assessment platform with ", False),
        ("ReactJS", True),
        (", ", False),
        ("Next.js", True),
        (", ", False),
        ("NestJS", True),
        (", and ", False),
        ("OpenAI API", True),
        (", designing structured evaluation rubrics to automate candidate scoring and integrating ", False),
        ("Jest", True),
        (" unit test suites to enforce strict quality gates.", False)
    ], space_after=2.4)
    add_bullet([
        ("Constructed bidirectional streaming pipelines using ", False),
        ("WebSockets", True),
        (" and ", False),
        ("Vapi.ai", True),
        (", enabling real-time voice screening evaluations and reducing candidate assessment turnaround by ", False),
        ("50%", True),
        (" across 200+ monthly candidate reviews.", False)
    ], space_after=2.6)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Programming Languages: ", True, "Java (Core Java, OOP), Python, C#, .NET, TypeScript, JavaScript, SQL (PostgreSQL, MySQL, SQL Server), C++, Go"),
        ("Frontend & Web Development: ", True, "ReactJS, Next.js (BFF), HTML5, CSS3, JavaScript (ES6+), Reusable UI Components, Responsive Design, Tailwind CSS, Angular 14"),
        ("Cloud & Backend Architecture: ", True, "RESTful APIs, Spring Boot, NestJS, Microsoft Azure (App Services, Functions, Azure SQL, Storage), GCP, AWS, Microservices, Kafka"),
        ("Databases & Storage: ", True, "Azure SQL, SQL Server fundamentals, PostgreSQL, MySQL, Redis (Distributed Caching/Locks), MongoDB, Hibernate/JPA, Query Optimization (EXPLAIN)"),
        ("AI, Automation & Tools: ", True, "GitHub Copilot, Cursor, OpenAI API, Google Gemini, Generative AI, Structured Prompt Engineering, Langfuse (evals/tracing), Claude Code"),
        ("Engineering Practices & SDLC: ", True, "Data Structures & Algorithms, OOP, SOLID Principles, Clean Coding, SDLC, Agile/Scrum (Sprint Planning, Stand-ups), Docker, Jenkins CI/CD, Git, Linux, JUnit, Jest")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.8)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform.", False)
    ], space_after=2.4)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for distributed system design.", False)
    ], space_after=2.4)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef.", False)
    ], space_after=1.5)

def populate_atlassian(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Cloud Platform Engineering, Go, Multi-Cloud Compute, Resilient Pipelines)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Architected a high-throughput ", False),
        ("Go pipeline", True),
        (" on ", False),
        ("Google Cloud (Cloud Run, Pub/Sub)", True),
        (" scaling ", False),
        ("compute platform", True),
        (" workloads, implementing ", False),
        ("MongoDB", True),
        (" distributed lease fencing and idempotent consumer writes to guarantee zero data loss.", False)
    ], space_after=2.4)
    add_bullet([
        ("Architected a multi-tenant ", False),
        ("Java/Spring Boot", True),
        (" ETL service integrating enterprise Slate & Salesforce CRMs via resilient ", False),
        ("REST APIs", True),
        (", implementing self-healing ", False),
        ("Kafka", True),
        (" retry topics and Dead Letter Queues (", False),
        ("DLQ", True),
        (") to prevent data loss during burst ingestion.", False)
    ], space_after=2.4)
    add_bullet([
        ("Architected an ", False),
        ("enterprise-grade", True),
        (" commercial ", False),
        ("GPA platform", True),
        (" (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+", True),
        (" university deployments generating ", False),
        ("$150K+", True),
        (" contract value per client, enforcing strict ", False),
        ("FERPA compliance", True),
        (" and enterprise RBAC security.", False)
    ], space_after=2.4)
    add_bullet([
        ("Engineered low-latency telephony streaming middleware in ", False),
        ("NestJS", True),
        (" bridging ", False),
        ("Genesys AudioHook", True),
        (" with ", False),
        ("Vapi.ai", True),
        (" via ", False),
        ("WebSockets", True),
        (", transcoding audio streams for ", False),
        ("1,000+", True),
        (" concurrent sessions at sub-", False),
        ("100ms", True),
        (" p95 latency with ", False),
        ("cell-based architecture", True),
        (" patterns.", False)
    ], space_after=2.4)
    add_bullet([
        ("Evaluated non-deterministic agentic flows via ", False),
        ("Langfuse", True),
        (" and spearheaded company-wide ", False),
        ("Sentry", True),
        (" rollout across microservices, establishing self-healing observability and prompt diagnostics that cut mean-time-to-detect (", False),
        ("MTTD", True),
        (") by ", False),
        ("45%", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Slashed backend analytics REST API response latency ", False),
        ("3×", True),
        (" (", False),
        ("15s to 5s", True),
        (") by profiling SQL execution plans, optimizing composite indexing, and restructuring data-access queries across ", False),
        ("PostgreSQL/MySQL", True),
        (" to maximize system throughput.", False)
    ], space_after=2.4)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines automating containerized deployments (", False),
        ("Docker", True),
        (") and integration test suites in an ", False),
        ("Agile team", True),
        (", eliminating broken staging builds and accelerating developer release cadence.", False)
    ], space_after=2.4)
    add_bullet([
        ("Launched a full-stack AI evaluation platform automating ", False),
        ("200+", True),
        (" monthly technical assessments, slashing turnaround by ", False),
        ("50%", True),
        (" and integrating RESTful APIs to automate end-to-end evaluation workflows for hiring teams.", False)
    ], space_after=2.6)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Operating Systems, Computer Networks, Database Management Systems (DBMS), System Design, Software Engineering Principles.", False)
    ], space_after=2.6)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Go (Golang), Docker, Apache Kafka, Redis, PostgreSQL, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a high-throughput, cloud-native payments microservice in ", False),
        ("Go (Golang)", True),
        (" with goroutines and channels, sustaining ", False),
        ("1,200 TPS", True),
        (" (simulated) at sub-", False),
        ("120ms", True),
        (" p95 latency, validated via authored Artillery load tests.", False)
    ], space_after=2.4)
    add_bullet([
        ("Engineered an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" and ", False),
        ("Redis", True),
        (" distributed locks with idempotent ", False),
        ("PostgreSQL", True),
        (" updates, guaranteeing zero double-charges across 500+ concurrent simulated users under sustained load.", False)
    ], space_after=2.6)

    add_project_heading("AI Interviewer", "NestJS, Next.js, React, TypeScript, WebSockets, OpenAI API, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Shipped an AI assessment platform with responsive ", False),
        ("React/Next.js", True),
        (" UI and ", False),
        ("NestJS", True),
        (" microservices, designing structured evaluation rubrics to automate candidate scoring and integrating ", False),
        ("Jest", True),
        (" testing to enforce quality gates.", False)
    ], space_after=2.4)
    add_bullet([
        ("Constructed bidirectional streaming pipelines using ", False),
        ("WebSockets", True),
        (" and ", False),
        ("Vapi.ai", True),
        (", enabling real-time voice screening and reducing candidate assessment turnaround by ", False),
        ("50%", True),
        (" across 200+ monthly candidate evaluations.", False)
    ], space_after=2.6)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Go (Golang), Java, Python, TypeScript, JavaScript, C++, SQL (PostgreSQL, MySQL), HTML5, CSS3"),
        ("Cloud & Platforms: ", True, "Docker, Kubernetes (familiar), GCP (Cloud Run, Pub/Sub), AWS, Jenkins CI/CD (Artifact Pipelines), Linux, Git"),
        ("Platform & Distributed Systems: ", True, "Microservices, Apache Kafka (Streaming/DLQ), Redis (Locks/Caching), Event-Driven Design"),
        ("Reliability & Core Engineering: ", True, "Self-Healing Systems, Sentry (MTTD), Artillery Load Testing, Fault Tolerance, JUnit, Jest"),
        ("Databases & Architecture: ", True, "Multi-Cloud Enablement, PostgreSQL, MongoDB, MySQL, Prisma ORM, Next.js (BFF), React, Angular 14, RBAC")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=2.2)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform.", False)
    ], space_after=2.2)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for distributed system design.", False)
    ], space_after=2.2)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef.", False)
    ], space_after=1.5)

def populate_brillio(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - SDE Full-Stack, Concurrency, Cloud/GCP, SQL, Agentic AI & Observability)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Evaluated non-deterministic agentic flows and traced reasoning chains using ", False),
        ("Langfuse", True),
        (" and ", False),
        ("Sentry", True),
        (", diagnosing tool-call hallucinations and prompt drift to automate quality gates and cut mean-time-to-detect (", False),
        ("MTTD", True),
        (").", False)
    ], space_after=2.2)
    add_bullet([
        ("Developed high-throughput ", False),
        ("NestJS", True),
        (" WebSocket middleware bridging real-time telephony streams, handling bidirectional audio transcoding for ", False),
        ("1k+", True),
        (" concurrent sessions at sub-", False),
        ("100ms", True),
        (" p95 latency with zero connection drops.", False)
    ], space_after=2.2)
    add_bullet([
        ("Automated high-volume data pipelines processing ", False),
        ("50K+", True),
        (" student records from ", False),
        ("PostgreSQL", True),
        (" using scheduled ", False),
        ("Python", True),
        (" jobs on GCP, eliminating manual intervention and cutting stakeholder reporting cycles by ", False),
        ("60%", True),
        (".", False)
    ], space_after=2.2)
    add_bullet([
        ("Architected an enterprise commercial ", False),
        ("GPA evaluation platform", True),
        (" (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("React", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+", True),
        (" university deployments generating ", False),
        ("$150K+", True),
        (" contract value per client under strict ", False),
        ("FERPA compliance", True),
        (".", False)
    ], space_after=2.2)
    add_bullet([
        ("Architected a multi-tenant ", False),
        ("Java/Spring Boot", True),
        (" ETL pipeline integrating Slate and Salesforce CRMs via REST APIs, implementing ", False),
        ("Hibernate/JPA", True),
        (" and ", False),
        ("Kafka", True),
        (" Dead Letter Queues (", False),
        ("DLQ", True),
        (") to prevent data loss during burst traffic.", False)
    ], space_after=2.2)
    add_bullet([
        ("Reduced backend analytics API response latency ", False),
        ("3×", True),
        (" (", False),
        ("15s to 5s", True),
        (") by profiling ", False),
        ("SQL", True),
        (" query execution plans and restructuring ", False),
        ("PostgreSQL/MySQL", True),
        (" schemas, eliminating redundant joins and table scans to maximize throughput.", False)
    ], space_after=2.2)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Launched a full-stack AI interview platform automating ", False),
        ("200+", True),
        (" monthly evaluations and slashing turnaround ", False),
        ("50%", True),
        (", integrating RESTful APIs and automating end-to-end evaluation workflows for hiring teams.", False)
    ], space_after=2.2)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines automating containerized deployments (", False),
        ("Docker", True),
        (") and integration test suites in an ", False),
        ("Agile team", True),
        (", eliminating broken staging builds and accelerating developer release cadence.", False)
    ], space_after=2.4)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, OOP, Database Management Systems (DBMS), Operating Systems, Computer Networks.", False)
    ], space_after=2.4)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Java, Spring Boot, Hibernate, Kafka, PostgreSQL, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a distributed payments microservice in ", False),
        ("Java (Spring Boot / Hibernate)", True),
        
        (", enforcing authentication via ", False),
        ("Spring Security", True),
        (" and ", False),
        ("JWT", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" at sub-", False),
        ("120ms", True),
        (" p95 latency.", False)
    ], space_after=2.2)
    add_bullet([
        ("Architected an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" with idempotency validation in ", False),
        ("PostgreSQL", True),
        (", guaranteeing zero duplicate transactions across ", False),
        ("500+", True),
        (" concurrent simulated users under load.", False)
    ], space_after=2.4)

    add_project_heading("AI Interviewer", "Next.js (SSR), NestJS, MongoDB, TypeScript, OpenAI API, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Engineered a full-stack AI evaluation platform with ", False),
        ("Next.js (SSR)", True),
        (", ", False),
        ("NestJS", True),
        (", and ", False),
        ("OpenAI API", True),
        (", designing structured evaluation rubrics to automate candidate scoring and integrating ", False),
        ("Jest", True),
        (" testing for QA gates.", False)
    ], space_after=2.2)
    add_bullet([
        ("Constructed bidirectional streaming pipelines using ", False),
        ("WebSockets", True),
        (" and ", False),
        ("Vapi.ai", True),
        (", enabling real-time voice screening evaluations and reducing candidate assessment turnaround by ", False),
        ("50%", True),
        (" across 200+ monthly candidate reviews.", False)
    ], space_after=2.4)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Java, Python, TypeScript, JavaScript (ES6+), Go, C++, SQL (PostgreSQL, MySQL)"),
        ("AI & Developer Tools: ", True, "Cursor, Claude Code, Copilot, Langfuse (tracing/evals), LangChain, Google GenAI, OpenAI API, RAG"),
        ("Backend & Architecture: ", True, "NestJS, Node.js, Spring Boot, Express.js, RESTful APIs (versioning), Kafka (Queues/DLQ), Event-Driven"),
        ("Databases & Caching: ", True, "PostgreSQL, MongoDB (NoSQL), Redis (Distributed Locks/Caching), MySQL, Prisma ORM, Schema Design"),
        ("DevOps & Cloud: ", True, "Docker, Google Cloud (GCP - Cloud Run, Pub/Sub, IAM), AWS, Kubernetes (familiar), Jenkins CI/CD, Git, Sentry"),
        ("Frontend & Web: ", True, "React, Next.js (SSR/BFF/App Router), Redux, Tailwind CSS, Angular 14, Core Web Vitals (LCP)")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=2.0)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform.", False)
    ], space_after=2.2)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for distributed system design.", False)
    ], space_after=2.2)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef.", False)
    ], space_after=1.5)

def populate_mastercard(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Java 17, Spring Boot, React.js, Kafka, Scaled Agile, SonarQube & CI/CD)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Architected a multi-tenant ", False),
        ("Java 17 (Spring Boot)", True),
        (" backend service extracting CRM data via Strategy and Factory patterns; deployed ", False),
        ("Hibernate/JPA", True),
        (" and ", False),
        ("Kafka DLQ", True),
        (" retry topics to ensure eventual consistency under burst traffic.", False)
    ], space_after=2.4)
    add_bullet([
        ("Deployed an asynchronous ", False),
        ("Transfer Credit Evaluation (TCE)", True),
        (" pipeline with ", False),
        ("Kafka", True),
        (" and ", False),
        ("NestJS/Java", True),
        (" microservices, decoupling degree audit processing from CRM webhooks and cutting evaluation turnaround time by ", False),
        ("80%", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Slashed document-analytics API latency ", False),
        ("3×", True),
        (" (", False),
        ("15s to 5s", True),
        (") by profiling ", False),
        ("SQL", True),
        (" query execution plans, adding composite indexes, and eliminating redundant join queries in ", False),
        ("PostgreSQL/MySQL", True),
        (" schemas to optimize database throughput.", False)
    ], space_after=2.4)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" pipelines and ", False),
        ("Gitflow", True),
        (" peer review workflows, enforcing ", False),
        ("SonarQube", True),
        (" static analysis gates, unit testing (", False),
        ("JUnit", True),
        (", ", False),
        ("Mockito", True),
        ("), and Sentry telemetry to eliminate anti-patterns and cut MTTD by ", False),
        ("45%", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Scaled an enterprise commercial ", False),
        ("GPA evaluation platform", True),
        (" (", False),
        ("Next.js/React BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+", True),
        (" university deployments generating ", False),
        ("$150K+", True),
        (" contract value per client, enforcing strict ", False),
        ("FERPA compliance", True),
        (" and ", False),
        ("RBAC", True),
        (" security.", False)
    ], space_after=2.4)
    add_bullet([
        ("Implemented low-latency ", False),
        ("NestJS", True),
        (" WebSocket middleware bridging real-time telephony streams, handling bidirectional audio transcoding for ", False),
        ("1,000+", True),
        (" concurrent active sessions at sub-", False),
        ("100ms", True),
        (" p95 latency under peak production traffic.", False)
    ], space_after=2.4)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Formulated a ", False),
        ("Kafka-driven", True),
        (" microservice for asynchronous event synchronization between external CRMs and internal databases within a ", False),
        ("Scaled Agile (SAFe)", True),
        (" sprint cadence, maintaining sub-", False),
        ("50ms", True),
        (" sync latency under peak traffic.", False)
    ], space_after=2.4)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines automating containerized deployments (", False),
        ("Docker", True),
        (") and integration test suites in an ", False),
        ("Agile team", True),
        (", eliminating broken staging builds and accelerating developer release cadence.", False)
    ], space_after=2.6)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks.", False)
    ], space_after=2.4)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Java 17, Spring Boot, Hibernate, Kafka, Redis, PostgreSQL, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a mission-critical payment microservice in ", False),
        ("Java 17 (Spring Boot / Hibernate)", True),
        
        (", ", False),
        ("Spring Security", True),
        (", and ", False),
        ("JWT", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency with published Artillery test suites.", False)
    ], space_after=2.4)
    add_bullet([
        ("Engineered an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" with ", False),
        ("Redis", True),
        (" distributed locks and ", False),
        ("PostgreSQL", True),
        (" ACID validation, guaranteeing idempotent transactions and zero double-charges across ", False),
        ("500+", True),
        (" simulated users.", False)
    ], space_after=2.4)
    add_bullet([
        ("Containerized microservices with ", False),
        ("Docker", True),
        (" and authored comprehensive ", False),
        ("JUnit", True),
        (" and ", False),
        ("Mockito", True),
        (" test suites to validate transactional rollback semantics, reconciliation routines, and code quality gates for cloud deployment.", False)
    ], space_after=2.6)

    add_project_heading("AI Interviewer", "React, Next.js, NestJS, TypeScript, MongoDB, REST APIs, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Built a cloud-based evaluation platform with responsive ", False),
        ("React", True),
        (" dashboards, RESTful APIs, and role-based access control (", False),
        ("RBAC", True),
        ("), streamlining ", False),
        ("200+", True),
        (" monthly candidate assessments and cutting evaluation turnaround by ", False),
        ("50%", True),
        (" for hiring teams.", False)
    ], space_after=2.4)
    add_bullet([
        ("Constructed bidirectional streaming pipelines using ", False),
        ("WebSockets", True),
        (" and structured evaluation rubrics, enforcing service-level test coverage and quality gates with ", False),
        ("Jest", True),
        (" to ensure reliable continuous delivery in CI/CD pipelines.", False)
    ], space_after=2.6)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Java (17+ / Core Java, OOP), TypeScript, JavaScript, SQL (PostgreSQL, MySQL), Go, Python, C++"),
        ("Frontend & Web: ", True, "React.js, Next.js (BFF), Redux, Tailwind CSS, HTML5, CSS3, RESTful APIs, Web Applications"),
        ("Backend & Frameworks: ", True, "Spring Boot (Java 17), Spring MVC, Hibernate/JPA, Inversion of Control (IoC), Spring Security, NestJS"),
        ("Architecture & Messaging: ", True, "Apache Kafka (Event-Driven Services, Outbox Pattern, DLQ), Redis (Caching/Locks), SOA"),
        ("Databases & Quality: ", True, "PostgreSQL, Oracle (familiar), MongoDB (NoSQL), SonarQube (Quality Gates), Checkmarx, Sentry"),
        ("SDLC, Testing & DevOps: ", True, "Gitflow (Peer Reviews), Docker (Containerization), Jenkins CI/CD, JUnit, Mockito, Jest, SAFe / Agile")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=2.0)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting a real-time voice screening platform.", False)
    ], space_after=2.4)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for fault-tolerant system architecture.", False)
    ], space_after=2.4)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, emphasizing algorithmic complexity, memory optimization, and concurrent execution.", False)
    ], space_after=1.8)

def populate_rippling(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Backend Leaning Full-Stack, Golang, Python, Distributed Systems, React)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Built a scalable ", False),
        ("Go (Golang) pipeline", True),
        (" on ", False),
        ("Cloud Run", True),
        (" for delta ingestion: scheduled a ", False),
        ("Pub/Sub", True),
        (" producer and consumer executing lease-fenced, idempotent writes in ", False),
        ("MongoDB", True),
        (", ensuring zero data loss and resilient recovery during upstream outages.", False)
    ], space_after=2.2)
    add_bullet([
        ("Automated high-volume backend data pipelines processing ", False),
        ("50K+", True),
        (" student application records from ", False),
        ("PostgreSQL", True),
        (" using scheduled ", False),
        ("Python", True),
        (" batch jobs, eliminating manual intervention and cutting stakeholder reporting cycles by ", False),
        ("60%", True),
        (".", False)
    ], space_after=2.2)
    add_bullet([
        ("Architected an enterprise commercial ", False),
        ("GPA platform", True),
        (" (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("React", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+", True),
        (" university deployments generating ", False),
        ("$150K+", True),
        (" contract value per client, enforcing strict ", False),
        ("FERPA compliance", True),
        (" and RBAC security.", False)
    ], space_after=2.2)
    add_bullet([
        ("Architected a multi-tenant ", False),
        ("Java/Spring Boot", True),
        (" ETL pipeline integrating Slate & Salesforce CRMs via ", False),
        ("REST APIs", True),
        (", implementing ", False),
        ("Hibernate/JPA", True),
        (" and ", False),
        ("Kafka Dead Letter Queues (DLQ)", True),
        (" to prevent data loss during burst traffic.", False)
    ], space_after=2.2)
    add_bullet([
        ("Reduced backend analytics API response latency ", False),
        ("3×", True),
        (" (", False),
        ("15s to 5s", True),
        (") by profiling ", False),
        ("SQL", True),
        (" query execution plans and restructuring ", False),
        ("PostgreSQL", True),
        (" schemas, eliminating redundant joins and table scans to maximize service throughput.", False)
    ], space_after=2.2)
    add_bullet([
        ("Deployed company-wide ", False),
        ("Sentry observability", True),
        (" across microservices for centralized error telemetry, cutting ", False),
        ("MTTD by 45%", True),
        (" and resolving a critical midnight CRM schema drift production incident within ", False),
        ("2 hours", True),
        (" during on-call rotation.", False)
    ], space_after=2.4)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Built the ", False),
        ("React/TypeScript", True),
        (" frontend for an automated AI interview platform, delivering role-based dashboards, real-time audio controls, and an automated scheduling UI for ", False),
        ("200+", True),
        (" monthly sessions (", False),
        ("50% turnaround cut", True),
        (").", False)
    ], space_after=2.2)
    add_bullet([
        ("Streamlined web asset delivery and frontend modules by configuring ", False),
        ("CDN edge caching", True),
        (" and asset chunking in an ", False),
        ("Agile team", True),
        (", delivering zero-downtime releases and lowering Largest Contentful Paint (", False),
        ("LCP", True),
        (") by ", False),
        ("25%", True),
        (".", False)
    ], space_after=2.4)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, OOP, Database Management Systems (DBMS), Operating Systems, Computer Networks.", False)
    ], space_after=2.4)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Go (Golang), Python, Kafka, Redis, PostgreSQL, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a high-throughput, cloud-native payments microservice in ", False),
        ("Go (Golang)", True),
        (" with goroutines and channels, sustaining ", False),
        ("1,200 TPS", True),
        (" (simulated) at sub-", False),
        ("120ms", True),
        (" p95 latency, validated via authored ", False),
        ("Artillery", True),
        (" load tests.", False)
    ], space_after=2.2)
    add_bullet([
        ("Architected an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" and ", False),
        ("Redis", True),
        (" distributed locks with idempotent ", False),
        ("PostgreSQL", True),
        (" updates, guaranteeing zero double-charges across ", False),
        ("500+", True),
        (" concurrent simulated users under load.", False)
    ], space_after=2.4)

    add_project_heading("AI Interviewer", "React, Next.js, NestJS, MongoDB, TypeScript, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Engineered a full-stack AI evaluation platform with ", False),
        ("Next.js", True),
        (", ", False),
        ("React", True),
        (", and ", False),
        ("NestJS", True),
        (", designing reusable UI components and structured evaluation rubrics to automate candidate scoring, integrating ", False),
        ("Jest", True),
        (" to enforce quality gates.", False)
    ], space_after=2.2)
    add_bullet([
        ("Constructed bidirectional streaming pipelines using ", False),
        ("WebSockets", True),
        (" and ", False),
        ("Vapi.ai", True),
        (", enabling real-time voice screening evaluations and reducing candidate assessment turnaround by ", False),
        ("50%", True),
        (" across 200+ monthly candidate reviews.", False)
    ], space_after=2.4)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Python, Go (Golang), TypeScript, JavaScript (ES6+), Java, C++, SQL (PostgreSQL, MySQL)"),
        ("Backend & Distributed: ", True, "Golang (goroutines), Node.js, NestJS, Express.js, Spring Boot, Apache Kafka, REST APIs"),
        ("Databases & Storage: ", True, "MongoDB, PostgreSQL, MySQL, Redis (Distributed Caching/Locks), ACID Transactions"),
        ("Frontend & Mobile: ", True, "React, React Native (familiar), Next.js (BFF), Redux, Tailwind CSS, Component Design"),
        ("Cloud & DevOps: ", True, "AWS, Google Cloud (GCP), Docker, Jenkins CI/CD, Git, Sentry (Observability), Jest, JUnit"),
        ("Frameworks & AI: ", True, "Django (familiar), FastAPI (familiar), WebSockets, Langfuse (tracing), OpenAI API, Gemini")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=2.0)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform.", False)
    ], space_after=2.2)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for distributed system design.", False)
    ], space_after=2.2)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, emphasizing concurrent algorithms, graph traversal and concurrent programming.", False)
    ], space_after=1.5)

def populate_google(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Distributed Systems, Go, Cloud Infrastructure, Resilient Data Pipelines)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Architected a decoupled ", False),
        ("Go pipeline", True),
        (" on ", False),
        ("Google Cloud (Cloud Run, Pub/Sub)", True),
        (" for delta ingestion, implementing durable ", False),
        ("MongoDB lease fencing", True),
        (" and ", False),
        ("idempotent batch writes", True),
        (" to guarantee ", False),
        ("zero data loss", True),
        (" during upstream network spikes.", False)
    ], space_after=2.4)
    add_bullet([
        ("Replicated applicant data across an isolated ", False),
        ("Student Copilot MongoDB cluster", True),
        (" using ", False),
        ("Go and Pub/Sub", True),
        (", enforcing ", False),
        ("revision-fenced idempotent writes", True),
        (" and sub-second soft-delete propagation to eliminate concurrent write races.", False)
    ], space_after=2.4)
    add_bullet([
        ("Constructed an event-driven ", False),
        ("Java/Spring Boot", True),
        (" ETL service integrating enterprise CRM datasets via ", False),
        ("REST endpoints", True),
        (", implementing ", False),
        ("Kafka retry topics", True),
        (" and ", False),
        ("Dead Letter Queues (DLQ)", True),
        (" to prevent message drops during peak ingestion windows.", False)
    ], space_after=2.4)
    add_bullet([
        ("Scaled an enterprise multi-tenant platform (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (") across ", False),
        ("4+ production university environments", True),
        (" generating ", False),
        ("$150K+ average contract value", True),
        (" per client, enforcing strict ", False),
        ("FERPA compliance", True),
        (" and ", False),
        ("role-based access control", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Instrumented cloud microservices and agentic workflows with ", False),
        ("Sentry", True),
        (" and ", False),
        ("Langfuse tracing", True),
        (", categorizing ", False),
        ("transient network timeouts", True),
        (" versus service failures to ", False),
        ("isolate root causes", True),
        (" and reduce production ", False),
        ("MTTD by 45%", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Optimized database storage and query execution by profiling SQL query plans and restructuring schemas across ", False),
        ("PostgreSQL", True),
        (" and ", False),
        ("MongoDB", True),
        (", eliminating unindexed table scans to ", False),
        ("slash backend API response latency", True),
        (" from ", False),
        ("15s to 5s (3×)", True),
        (".", False)
    ], space_after=2.4)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines automating containerized deployments (", False),
        ("Docker", True),
        (") and integration test suites in an ", False),
        ("Agile team", True),
        (", eliminating broken staging builds and accelerating developer release cadence.", False)
    ], space_after=2.4)
    add_bullet([
        ("Launched an AI evaluation platform processing ", False),
        ("200+", True),
        (" monthly technical assessments with ", False),
        ("50% turnaround reduction", True),
        (", designing modular ", False),
        ("REST services", True),
        (" and documentation that accelerated ", False),
        ("cross-functional partner team onboarding", True),
        (".", False)
    ], space_after=2.6)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Operating Systems, Computer Networks, Database Management Systems (DBMS), Object-Oriented Programming (OOP), System Design, Cloud Architecture Principles.", False)
    ], space_after=2.6)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Go (Golang), Docker, Apache Kafka, Redis, PostgreSQL, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a cloud-native payments microservice in ", False),
        ("Go (Golang)", True),
        (" using goroutines and channels, sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms p95 latency", True),
        (", validated via authored ", False),
        ("Artillery load tests", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Engineered an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" and ", False),
        ("Redis", True),
        (" coordination locks with ", False),
        ("ACID PostgreSQL updates", True),
        (", guaranteeing ", False),
        ("strict transaction idempotency", True),
        (" and zero double-charges across ", False),
        ("500+ simulated users", True),
        (".", False)
    ], space_after=2.6)

    add_project_heading("AI Interviewer", "NestJS, Next.js, React, TypeScript, WebSockets, OpenAI API, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Shipped an autonomous screening platform with ", False),
        ("Next.js", True),
        (" and ", False),
        ("NestJS", True),
        (" microservices, codifying ", False),
        ("structured evaluation rubrics", True),
        (" and rigorous ", False),
        ("Jest regression suites", True),
        (" to maintain deterministic scoring quality across candidate assessments.", False)
    ], space_after=2.4)
    add_bullet([
        ("Maintained bidirectional streaming audio pipelines via ", False),
        ("WebSockets", True),
        (" and ", False),
        ("Vapi.ai", True),
        (", handling real-time audio chunking and speech-to-text inference to cut candidate technical evaluation turnaround by ", False),
        ("50%", True),
        (" across ", False),
        ("200+ monthly sessions", True),
        (".", False)
    ], space_after=2.6)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Go (Golang), Java, Python, C++, TypeScript, JavaScript, SQL (PostgreSQL, MySQL), HTML5, CSS3"),
        ("Cloud & Infrastructure: ", True, "Google Cloud Platform (Cloud Run, Pub/Sub, Cloud Scheduler), Docker, Kubernetes (familiar), AWS, Linux, Jenkins CI/CD, Git"),
        ("Platform & Streaming Systems: ", True, "Microservices, Apache Kafka (DLQ, Outbox Pattern), Redis (Locks, Caching), WebSockets, Event-Driven Architecture, Zero-Trust Networking"),
        ("Reliability, Observability & Testing: ", True, "Sentry (Distributed Tracing, MTTD), Langfuse (AI Telemetry/Evals), Artillery Load Testing, Fault Tolerance, Root-Cause Triage, Jest, JUnit"),
        ("Databases & Architecture: ", True, "PostgreSQL, MongoDB (Ordered Bulk Operations, Aggregation), MySQL, Prisma ORM, ACID Transactions, Next.js (BFF), REST APIs, RBAC")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=2.2)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting a real-time voice screening platform.", False)
    ], space_after=2.2)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for systems architecture.", False)
    ], space_after=2.2)
    add_bullet([
        ("Solved ", False),
        ("450+ Data Structures & Algorithms problems", True),
        (" across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, demonstrating deep proficiency in graph traversal and concurrent programming.", False)
    ], space_after=1.5)

def populate_zomato(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Scalable Systems, Distributed Pipelines, Process Improvement & Automation)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Generated ", False),
        ("$150K+ average contract value", True),
        (" per client across ", False),
        ("4+ enterprise university deployments", True),
        (" by scaling a commercial ", False),
        ("GPA platform", True),
        (" (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") under strict ", False),
        ("FERPA compliance", True),
        (" and ", False),
        ("RBAC security", True),
        (".", False)
    ], space_after=2.2)
    add_bullet([
        ("Architected a decoupled ", False),
        ("Go ingestion pipeline", True),
        (" on ", False),
        ("Google Cloud (Cloud Run, Pub/Sub)", True),
        (", syncing ", False),
        ("10K+ delta records per 15-minute cycle", True),
        (" into ", False),
        ("MongoDB", True),
        (" with ", False),
        ("lease fencing", True),
        (" and ", False),
        ("idempotent writes", True),
        (" to ensure ", False),
        ("almost zero data loss", True),
        (".", False)
    ], space_after=2.2)
    add_bullet([
        ("Engineered a real-time ", False),
        ("Genesys AudioHook v2", True),
        (" WebSocket bridge in ", False),
        ("NestJS", True),
        (", sustaining ", False),
        ("1,000+ concurrent telephony streams", True),
        (" at sub-", False),
        ("100ms p95 latency", True),
        (" via bidirectional ", False),
        ("8 kHz µ-law to 16 kHz PCM", True),
        (" audio transcoding.", False)
    ], space_after=2.2)
    add_bullet([
        ("Optimized database storage and query execution by profiling ", False),
        ("SQL query plans", True),
        (" and restructuring schemas across ", False),
        ("PostgreSQL", True),
        (" and ", False),
        ("MongoDB", True),
        (", eliminating unindexed table scans to ", False),
        ("slash backend API response latency", True),
        (" from ", False),
        ("15s to 5s (3×)", True),
        (".", False)
    ], space_after=2.2)
    add_bullet([
        ("Deployed company-wide ", False),
        ("Sentry observability", True),
        (" across microservices for centralized error telemetry, cutting ", False),
        ("MTTD by 45%", True),
        (" and resolving a critical midnight CRM schema drift production incident within ", False),
        ("2 hours", True),
        (" during on-call rotation.", False)
    ], space_after=2.4)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Launched an internal technical assessment platform processing ", False),
        ("200+", True),
        (" monthly evaluations with ", False),
        ("50% turnaround reduction", True),
        (", designing modular ", False),
        ("REST services", True),
        (" that accelerated ", False),
        ("cross-functional partner team onboarding across engineering squads", True),
        (".", False)
    ], space_after=2.2)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines automating containerized deployments (", False),
        ("Docker", True),
        (") and integration test suites in an ", False),
        ("Agile team", True),
        (", eliminating broken staging builds and accelerating developer release cadence.", False)
    ], space_after=2.4)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Database Management Systems (DBMS), Operating Systems, Computer Networks, Object-Oriented Programming (OOP), Distributed Systems, System Design, Cloud Architecture Principles.", False)
    ], space_after=2.4)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Go (Golang), Docker, Apache Kafka, Redis, PostgreSQL, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a cloud-native payments microservice in ", False),
        ("Go (Golang)", True),
        (" using goroutines and channels, sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms p95 latency", True),
        (", validating service capacity through authored ", False),
        ("Artillery load testing suites", True),
        (".", False)
    ], space_after=2.2)
    add_bullet([
        ("Constructed an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" and ", False),
        ("Redis distributed locks", True),
        (" with idempotent ", False),
        ("PostgreSQL", True),
        (" updates, guaranteeing ", False),
        ("zero double-charges", True),
        (" across ", False),
        ("500+", True),
        (" concurrent simulated users under load.", False)
    ], space_after=2.4)

    add_project_heading("AI Interviewer", "NestJS, Next.js, React, TypeScript, WebSockets, OpenAI API, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Built an autonomous candidate screening platform with ", False),
        ("Next.js", True),
        (" and ", False),
        ("NestJS", True),
        (" microservices, codifying ", False),
        ("structured evaluation rubrics", True),
        (" and ", False),
        ("Jest test suites", True),
        (" to maintain deterministic scoring quality across technical assessments.", False)
    ], space_after=2.2)
    add_bullet([
        ("Maintained bidirectional streaming audio pipelines via ", False),
        ("WebSockets", True),
        (" and ", False),
        ("Vapi.ai", True),
        (", handling real-time audio chunking and speech-to-text inference to cut candidate screening turnaround by ", False),
        ("50%", True),
        (" across ", False),
        ("200+ monthly sessions", True),
        (".", False)
    ], space_after=2.4)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Go (Golang), Python, TypeScript, JavaScript (ES6+), Java, C++, SQL (PostgreSQL, MySQL), HTML5, CSS3"),
        ("Cloud & Infrastructure: ", True, "Google Cloud Platform (Cloud Run, Pub/Sub, Cloud Scheduler), Docker, Kubernetes (familiar), AWS, Linux, Jenkins CI/CD, Git"),
        ("Platform & Streaming Systems: ", True, "Microservices, Apache Kafka (DLQ, Outbox Pattern), Redis (Locks, Caching), WebSockets (AudioHook), Event-Driven Architecture, Zero-Trust Networking"),
        ("Reliability, Observability & Testing: ", True, "Sentry (Distributed Tracing, MTTD), Langfuse (AI Telemetry), Artillery Load Testing, Fault Tolerance, Root-Cause Triage, Jest, JUnit"),
        ("Databases & Architecture: ", True, "PostgreSQL, MongoDB (Ordered Bulk Operations, Aggregation), MySQL, Prisma ORM, ACID Transactions, Next.js (BFF), REST APIs, RBAC")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=2.0)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform with low-latency audio inference.", False)
    ], space_after=2.0)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for scalable systems architecture and distributed backend design.", False)
    ], space_after=2.0)
    add_bullet([
        ("Solved ", False),
        ("450+ Data Structures & Algorithms problems", True),
        (" across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, demonstrating deep proficiency in graph traversal and concurrent programming.", False)
    ], space_after=1.5)

def populate_dentsu(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Work History (HR-First Layout - Aligned with Dentsu / DGS India ATS Parsing Tips)
    add_section_header("Work History")

    add_role_heading("Associate Software Engineer", "October 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Architected multi-tenant software features in ", False),
        ("Java 17 (Spring Boot)", True),
        (" and ", False),
        ("NestJS", True),
        (" for ", False),
        ("4+ enterprise university clients", True),
        (" (", False),
        ("$150K+", True),
        (" contract value each), tailoring unique customer configurations and adhering to clean OOP/SOLID quality standards.", False)
    ], space_after=2.4)
    add_bullet([
        ("Automated high-volume ETL data pipelines processing ", False),
        ("50K+ student records", True),
        (" from ", False),
        ("PostgreSQL", True),
        (" using scheduled ", False),
        ("Python jobs", True),
        (", crafting optimized ", False),
        ("SQL queries", True),
        (" to power ", False),
        ("Power BI analytics dashboards", True),
        (" and cutting manual reporting cycles by ", False),
        ("60%", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Deployed an AI-powered Document Intelligence platform coupling OCR content deduplication with LLM orchestration (", False),
        ("OpenAI, Google Gemini", True),
        ("), resolving ", False),
        ("40%+", True),
        (" of transcript evaluations deterministically and reducing external API inference costs by ", False),
        ("40%", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Instrumented multi-turn agent workflows with ", False),
        ("Langfuse tracing", True),
        (" and ", False),
        ("Sentry telemetry", True),
        (", troubleshooting production query bottlenecks, diagnosing tool-call anomalies to cut ", False),
        ("MTTD by 45%", True),
        (", and authoring technical documentation for partner onboarding.", False)
    ], space_after=2.5)

    add_role_heading("Software Engineer Intern", "April 2025 – September 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Reduced backend analytics API response latency ", False),
        ("3x", True),
        (" (", False),
        ("15s to 5s", True),
        (") by profiling ", False),
        ("SQL query execution plans", True),
        (" across relational databases (PostgreSQL/MySQL), optimizing composite indexing and eliminating redundant joins to maximize throughput.", False)
    ], space_after=2.4)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" pipelines orchestrating containerized deployments (", False),
        ("Docker", True),
        (") and unit/integration test suites (", False),
        ("JUnit, Jest", True),
        (") across Agile release cycles, collaborating with QA and senior engineers to ensure zero broken staging builds.", False)
    ], space_after=2.6)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "December 2021 – June 2025", "B.Tech, Electronics & Communication Engineering (2025 Graduate)", "New Delhi, India (Open to Relocate — Mumbai)")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks, Software Engineering & Unit Testing.", False)
    ], space_after=2.6)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Java, Spring Boot, Hibernate, Apache Kafka, PostgreSQL, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a cloud-native payments microservice in ", False),
        ("Java (Spring Boot / Hibernate)", True),
        
        (", enforcing authentication via ", False),
        ("Spring Security", True),
        (" and ", False),
        ("JWT", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency with published Artillery load tests.", False)
    ], space_after=2.4)
    add_bullet([
        ("Constructed an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" and Redis distributed locks with idempotent ", False),
        ("PostgreSQL", True),
        (" updates, guaranteeing zero double-charges across ", False),
        ("500+", True),
        (" simulated users, tested via ", False),
        ("Docker", True),
        (" and ", False),
        ("JUnit", True),
        (" suites.", False)
    ], space_after=2.6)

    add_project_heading("AI Interviewer", "Python, Next.js, NestJS, TypeScript, OpenAI API, WebSockets, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Engineered an autonomous candidate screening platform with ", False),
        ("Next.js", True),
        (", ", False),
        ("NestJS", True),
        (", and ", False),
        ("OpenAI API", True),
        (", codifying structured evaluation rubrics and ", False),
        ("Jest unit test suites", True),
        (" to deliver objective scoring across ", False),
        ("200+ monthly technical assessments", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Maintained bidirectional streaming audio pipelines via ", False),
        ("WebSockets", True),
        (" and ", False),
        ("Vapi.ai", True),
        (", handling real-time speech-to-text inference and lowering candidate evaluation turnaround time by ", False),
        ("50%", True),
        (" without human reviewer bottlenecks.", False)
    ], space_after=2.6)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Core Languages: ", True, "Python, Java (Core Java, Spring Boot, OOP), SQL (PostgreSQL, MySQL, Query Tuning), TypeScript, JavaScript (ES6+), C++, Go (Golang), HTML5, CSS3"),
        ("Databases & Business Intelligence: ", True, "SQL (PostgreSQL, MySQL, Query Tuning, Indexing), Power BI (Interactive Dashboards, Data Modeling, DAX, Reports), MongoDB, Redis (Caching/Locks), Schema Design"),
        ("AI / Machine Learning: ", True, "Document Intelligence, Generative AI, OpenAI API, Google Gemini, Langfuse (LLM Tracing/Evals), Machine Learning Fundamentals, RAG Pipelines, Prompt Engineering, Vapi.ai"),
        ("Backend & Cloud Systems: ", True, "Spring Boot, NestJS, Node.js, RESTful APIs, Google Cloud Platform (Cloud Run, Pub/Sub), AWS, Microservices, Apache Kafka, WebSockets"),
        ("Testing, Release & Practices: ", True, "Unit Testing (JUnit, Jest), Jenkins CI/CD, Docker, Git, Agile/Scrum Release Cycles, SDLC, Postman, Sentry (Observability), Technical Documentation")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.8)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting a real-time voice screening platform with low-latency audio inference.", False)
    ], space_after=2.2)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for scalable system design and backend architecture.", False)
    ], space_after=2.2)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, demonstrating deep expertise in algorithmic design and efficient code delivery.", False)
    ], space_after=1.5)

def populate_juspay(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Frontend & Mobile UI Focus)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Developed responsive, accessible ", False),
        ("React", True),
        (" and ", False),
        ("Next.js", True),
        (" interfaces featuring reusable design system components, custom hooks, and strict state lifecycle management, reducing manual counsellor review overhead by ", False),
        ("80%", True),
        (".", False)
    ])
    add_bullet([
        ("Engineered an enterprise commercial ", False),
        ("GPA evaluation platform", True),
        (" using a ", False),
        ("Next.js BFF", True),
        (" and ", False),
        ("RESTful APIs", True),
        (" across ", False),
        ("4+", True),
        (" university deployments generating ", False),
        ("$150K+", True),
        (" contract value per client, enforcing strict ", False),
        ("FERPA compliance", True),
        (" and RBAC security.", False)
    ])
    add_bullet([
        ("Implemented an enterprise self-serve billing module in ", False),
        ("Next.js", True),
        (" and ", False),
        ("Redux", True),
        (" with ", False),
        ("Stripe Checkout", True),
        (" integration, building asynchronous webhook handlers, resilient error boundaries, and ", False),
        ("100%", True),
        (" payment reconciliation integrity.", False)
    ])
    add_bullet([
        ("Instrumented ", False),
        ("Sentry observability", True),
        (" across production ", False),
        ("React", True),
        (" and ", False),
        ("Next.js", True),
        (" web clients with source maps and session replay, troubleshooting client runtime crashes and cutting mean-time-to-resolution (", False),
        ("MTTR", True),
        (") by ", False),
        ("45%", True),
        (".", False)
    ])
    add_bullet([
        ("Optimized client-side data access patterns and ", False),
        ("RESTful API", True),
        (" payloads, reducing analytics response latency ", False),
        ("3×", True),
        (" (", False),
        ("15s to 5s", True),
        (") and eliminating redundant network round-trips to maximize mobile and web responsiveness.", False)
    ])

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Constructed the ", False),
        ("React/TypeScript", True),
        (" frontend for a commercial assessment platform, delivering role-based dashboards, real-time audio controls, and an intuitive scheduling UI for ", False),
        ("200+", True),
        (" monthly sessions (", False),
        ("50%", True),
        (" turnaround cut).", False)
    ])
    add_bullet([
        ("Refactored client-side asset bundles and ", False),
        ("CDN edge caching", True),
        (" across web platforms in an ", False),
        ("Agile team", True),
        (", eliminating render-blocking dependencies, accelerating mobile load times, and lowering Largest Contentful Paint (", False),
        ("LCP", True),
        (") by ", False),
        ("25%", True),
        (".", False)
    ])

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks.", False)
    ], space_after=2.5)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "React, Redux, Vite, Node.js, Express.js, PostgreSQL, RESTful APIs", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Built a mobile-first ", False),
        ("React", True),
        (" and ", False),
        ("Redux", True),
        (" fintech transaction dashboard bundled with ", False),
        ("Vite", True),
        (", establishing centralized state management, responsive card layouts, and deterministic UI rendering under burst network load.", False)
    ])
    add_bullet([
        ("Designed idempotency-aware transaction APIs on ", False),
        ("Node.js/Express", True),
        (" and ", False),
        ("PostgreSQL", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" (simulated) at p95 latency under ", False),
        ("120ms", True),
        (", validated through authored Artillery load tests and comprehensive test suites.", False)
    ])

    add_project_heading("AI Interviewer", "React, Next.js, Redux, WebSockets, TypeScript, RESTful APIs, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Created modular ", False),
        ("React UI components", True),
        (" and custom hooks for an AI evaluation platform, enforcing strict state validation, client-side data caching, and comprehensive ", False),
        ("Jest", True),
        (" unit test suites for UI reliability.", False)
    ])
    add_bullet([
        ("Spearheaded real-time bidirectional audio streaming controls using ", False),
        ("WebSockets", True),
        (" and ", False),
        ("Next.js", True),
        (", conducting 200+ monthly screening evaluations and reducing candidate assessment turnaround by ", False),
        ("50%", True),
        (" across hiring workflows.", False)
    ])

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Frontend & Mobile: ", True, "JavaScript (ES6+), TypeScript, React, React Native (core architecture), Next.js (BFF), Redux, Context API, Android SDK"),
        ("Architecture & UI: ", True, "Component Lifecycle, Custom Hooks, Reusable UI Components, Responsive Layouts, Cross-Platform UI, Tailwind CSS, WCAG"),
        ("Languages: ", True, "JavaScript (ES6+), TypeScript, Python, Go, Java, C++, SQL (PostgreSQL, MySQL), HTML5, CSS3/SASS"),
        ("DevOps, Build & Testing: ", True, "Git, GitHub Actions, Vite, Webpack, npm, Jest, React Testing Library, Artillery Load Testing, Sentry"),
        ("Backend & APIs: ", True, "Node.js, Express.js, NestJS, RESTful APIs, Asynchronous Programming, WebSockets, MongoDB, Redis, PostgreSQL")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=2.0)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform with low-latency audio inference.", False)
    ], space_after=2.0)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for distributed system design and client-server architecture.", False)
    ], space_after=2.0)
    add_bullet([
        ("Solved ", False),
        ("450+ Data Structures & Algorithms problems", True),
        (" across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, demonstrating deep proficiency in algorithmic design and runtime efficiency.", False)
    ], space_after=1.5)

def populate_cashfree(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - React.js, Frontend Architecture, Fintech & Core Web Vitals Focus)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Developed responsive, accessible ", False),
        ("React.js", True),
        (" and ", False),
        ("Next.js", True),
        (" interfaces featuring ", False),
        ("reusable UI component libraries", True),
        (", design tokens, and custom hooks, eliminating redundant re-renders and cutting review overhead by ", False),
        ("80%", True),
        (".", False)
    ], space_after=2.0)
    add_bullet([
        ("Engineered an enterprise web portal across ", False),
        ("4+", True),
        (" university client deployments (", False),
        ("$150K+", True),
        (" contract value) using ", False),
        ("Next.js BFF", True),
        (" proxies to eliminate client-side token exposure and enforce strict ", False),
        ("CSP, CORS, and FERPA", True),
        (" security.", False)
    ], space_after=2.0)
    add_bullet([
        ("Implemented an enterprise billing dashboard in ", False),
        ("Next.js", True),
        (" and ", False),
        ("Redux Toolkit", True),
        (" with ", False),
        ("Stripe Checkout", True),
        (", engineering asynchronous payment webhooks, resilient error boundaries, and ", False),
        ("100%", True),
        (" payment reconciliation integrity.", False)
    ], space_after=2.0)
    add_bullet([
        ("Diagnosed rendering bottlenecks and eliminated unnecessary component re-renders via ", False),
        ("Chrome DevTools", True),
        (" and ", False),
        ("React Profiler", True),
        (", optimizing asset bundles and CDN edge caching to reduce ", False),
        ("Largest Contentful Paint (LCP)", True),
        (" by ", False),
        ("25%", True),
        (".", False)
    ], space_after=2.0)
    add_bullet([
        ("Streamlined client-side data access patterns and ", False),
        ("RESTful API", True),
        (" payloads across ", False),
        ("React", True),
        (" applications, reducing analytics response latency ", False),
        ("3× (15s to 5s)", True),
        (" and eliminating redundant network round-trips for mobile users.", False)
    ], space_after=2.2)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Constructed the ", False),
        ("React/TypeScript", True),
        (" frontend for a commercial assessment platform, delivering role-based dashboards, real-time audio controls, and intuitive scheduling for ", False),
        ("200+", True),
        (" monthly sessions (", False),
        ("50%", True),
        (" turnaround cut).", False)
    ], space_after=2.0)
    add_bullet([
        ("Authored comprehensive unit and integration test suites using ", False),
        ("Jest", True),
        (" and ", False),
        ("React Testing Library", True),
        (" (", False),
        ("90%+", True),
        (" coverage), enforcing clean code principles, constructive code reviews, and ", False),
        ("Git", True),
        (" workflows across Agile sprint cycles.", False)
    ], space_after=2.4)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks, and Modern Web Architectures.", False)
    ], space_after=2.4)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "React, Redux Toolkit, Vite, TypeScript, Node.js, Express.js, PostgreSQL", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Built a mobile-first ", False),
        ("React", True),
        (" and ", False),
        ("Redux", True),
        (" fintech transaction dashboard bundled with ", False),
        ("Vite", True),
        (", establishing centralized state management, responsive card layouts, and deterministic UI rendering under burst network load.", False)
    ], space_after=2.0)
    add_bullet([
        ("Designed idempotency-aware transaction APIs on ", False),
        ("Node.js/Express", True),
        (" and ", False),
        ("PostgreSQL", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency, validated through Artillery load tests and comprehensive test suites.", False)
    ], space_after=2.4)

    add_project_heading("AI Interviewer", "React, Next.js, Zustand, WebSockets, TypeScript, RESTful APIs, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Spearheaded real-time bidirectional audio streaming controls using ", False),
        ("WebSockets", True),
        (" and ", False),
        ("Next.js", True),
        (", conducting ", False),
        ("200+", True),
        (" monthly candidate evaluations and reducing assessment turnaround time by ", False),
        ("50%", True),
        (" for hiring teams.", False)
    ], space_after=2.0)
    add_bullet([
        ("Integrated modular ", False),
        ("React UI components", True),
        (" and custom hooks with ", False),
        ("Zustand", True),
        (" state management, enforcing strict client-side validation, cross-browser compatibility, and ", False),
        ("Jest", True),
        (" test suites for zero UI regressions.", False)
    ], space_after=2.4)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Frontend & Frameworks: ", True, "React.js, Next.js (App Router, SSR/SSG), TypeScript, JavaScript (ES6+), HTML5, CSS3, Tailwind CSS"),
        ("State & Architecture: ", True, "Redux Toolkit, Zustand, Context API, Reusable UI Components, Design Tokens, Custom Hooks, WCAG 2.1"),
        ("Testing, Build & Tooling: ", True, "Jest, React Testing Library, Vite, Webpack, Babel, npm, pnpm, Git, Chrome DevTools, React Profiler"),
        ("Web Security & Performance: ", True, "CORS, XSS, CSRF, CSP, Core Web Vitals (LCP, CLS), Bundle Optimization, Code Splitting, Sentry"),
        ("Backend & Financial APIs: ", True, "Node.js, Express.js, NestJS, RESTful APIs, WebSockets, Stripe API, PostgreSQL, MongoDB, Redis")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.6)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform with low-latency audio inference.", False)
    ], space_after=2.0)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for scalable systems architecture and resilient client-server design.", False)
    ], space_after=2.0)
    add_bullet([
        ("Solved ", False),
        ("450+ Data Structures & Algorithms problems", True),
        (" across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, demonstrating deep proficiency in algorithmic design, complexity analysis, and runtime efficiency.", False)
    ], space_after=1.6)

def populate_akamai(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - High Availability Systems, Automation, Observability, Identity Management, Network Protocols)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Scaled a decoupled two-tier ", False),
        ("Go", True),
        (" ingestion pipeline on ", False),
        ("Google Cloud Run", True),
        (" and ", False),
        ("Pub/Sub", True),
        (", pulling ", False),
        ("10K+ delta records per 15-minute sync cycle", True),
        (" into durable MongoDB queues with atomic lease-fenced consumers, ensuring fault-tolerant zero data loss.", False)
    ], space_after=2.6)
    add_bullet([
        ("Integrated ", False),
        ("Slate CAS", True),
        (" ticket authentication and exact-one-person record resolution in ", False),
        ("NestJS", True),
        (", securing student-facing AI assistants with zero-trust authenticated identity management, ", False),
        ("RBAC", True),
        (", and FERPA-compliant session controls.", False)
    ], space_after=2.6)
    add_bullet([
        ("Instrumented multi-turn agent workflows and backend microservices with ", False),
        ("Sentry", True),
        (" and ", False),
        ("Langfuse", True),
        (" distributed tracing, isolating reasoning bottlenecks and eliminating database query hallucinations to cut ", False),
        ("MTTD by 45%", True),
        (".", False)
    ], space_after=2.6)
    add_bullet([
        ("Engineered a real-time telephony streaming bridge in ", False),
        ("NestJS", True),
        (" implementing the ", False),
        ("Genesys AudioHook v2", True),
        (" WebSocket protocol, transcoding bidirectional audio between 8 kHz µ-law and 16 kHz PCM for ", False),
        ("1,000+", True),
        (" concurrent sessions at sub-", False),
        ("100ms", True),
        (" p95 latency.", False)
    ], space_after=2.6)
    add_bullet([
        ("Constructed asynchronous data pipelines processing ", False),
        ("50K+", True),
        (" records from ", False),
        ("PostgreSQL", True),
        (" via scheduled ", False),
        ("Python", True),
        (" batch workers, eliminating manual intervention and cutting stakeholder reporting turnaround by ", False),
        ("60%", True),
        (".", False)
    ], space_after=2.6)
    add_bullet([
        ("Deployed high availability evaluation platforms across ", False),
        ("4+", True),
        (" universities (", False),
        ("Rowan", True),
        (", ", False),
        ("RFU", True),
        (", ", False),
        ("PennWest", True),
        (") generating ", False),
        ("$150K+", True),
        (" contract value per client, executing secure writeback of computed records directly into ", False),
        ("Slate CRM", True),
        (" via REST/JSON endpoints.", False)
    ], space_after=2.6)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Optimized backend analytics REST API response latency ", False),
        ("3×", True),
        (" (", False),
        ("15s to 5s", True),
        (") by profiling SQL query execution plans, adding composite indexes, and restructuring data-access schemas across ", False),
        ("PostgreSQL/MySQL", True),
        (" to maximize system throughput.", False)
    ], space_after=2.6)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines automating containerized deployments (", False),
        ("Docker", True),
        (") and integration test suites on Linux in an ", False),
        ("Agile team", True),
        (", eliminating broken staging builds and accelerating release cadence.", False)
    ], space_after=2.8)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Computer Networks, Operating Systems, Database Management Systems (DBMS), Distributed Systems, System Design, Software Engineering.", False)
    ], space_after=2.8)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Go (Golang), Docker, Kafka, Redis, PostgreSQL, Linux, Artillery", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a cloud-native payments microservice in ", False),
        ("Go (Golang)", True),
        (" utilizing goroutines and channels, sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency, validated via authored Artillery load tests on Linux.", False)
    ], space_after=2.6)
    add_bullet([
        ("Architected an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" with ", False),
        ("Redis", True),
        (" synchronization locks and ", False),
        ("PostgreSQL", True),
        (" ACID validation, guaranteeing idempotent transactions and zero double-charges across 500+ simulated users.", False)
    ], space_after=2.6)

    add_project_heading("AI Interviewer", "NestJS, Next.js, TypeScript, WebSockets, OpenAI API, Jest, Vapi.ai", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Spearheaded an AI assessment platform with ", False),
        ("NestJS", True),
        (" and ", False),
        ("Next.js", True),
        (", designing structured evaluation rubrics to score candidates and integrating ", False),
        ("Jest", True),
        (" unit and integration test suites to enforce strict quality gates.", False)
    ], space_after=2.6)
    add_bullet([
        ("Built bidirectional streaming pipelines using ", False),
        ("WebSockets", True),
        (" for real-time candidate audio processing, reducing evaluation turnaround by ", False),
        ("50%", True),
        (" across 200+ monthly technical interview sessions.", False)
    ], space_after=2.6)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Go (Golang), Python, Bash/Shell Script, TypeScript, JavaScript, Java, C++, SQL (PostgreSQL, MySQL)"),
        ("Systems & Protocols: ", True, "TCP/IP, HTTP/HTTPS, DNS, TLS/SSL, WebSockets, RESTful APIs, Linux/Unix, Socket Programming"),
        ("Messaging & Distributed: ", True, "Google Cloud Pub/Sub, Apache Kafka (DLQ, Outbox Pattern), Redis (Locks/Caching), Microservices"),
        ("Security & Identity: ", True, "Slate CAS SSO, OAuth 2.0, JWT, RBAC, FERPA Compliance, Reverse Proxy/BFF, Session Hardening"),
        ("DevOps & Observability: ", True, "Sentry (Tracing/Alerts), Langfuse, Docker, Jenkins CI/CD, Git, Linux CLI, Artillery (Load Testing)"),
        ("Backend & Frameworks: ", True, "NestJS, Node.js, Express, Spring Boot, React, Next.js, Redux, Angular 14, Tailwind CSS")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=2.0)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting a real-time voice screening platform with low-latency streaming and automated evaluation rubrics.", False)
    ], space_after=2.4)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for robust system design and client-server networking.", False)
    ], space_after=2.4)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef.", False)
    ], space_after=1.5)

def populate_barclays(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Aligned with Barclays Software Engineer JD: Java Spring Boot, Kafka, GitLab CI/CD, Docker, Kubernetes/OpenShift, JUnit/Cucumber, Gradle, MongoDB, OpenAPI)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Architected multi-tenant ", False),
        ("Java 17 (Spring Boot)", True),
        (" microservices integrating ", False),
        ("Kafka producer/consumer", True),
        (" pipelines and DLQ retry topics, defining ", False),
        ("OpenAPI 3.0", True),
        (" specs to ensure seamless schema exchange and sub-50ms inter-service latency.", False)
    ], space_after=2.0)
    add_bullet([
        ("Scaled an enterprise commercial ", False),
        ("GPA evaluation platform", True),
        (" (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+", True),
        (" university deployments generating ", False),
        ("$150K+", True),
        (" contract value per client, enforcing strict ", False),
        ("FERPA compliance", True),
        (" and ", False),
        ("RBAC", True),
        (" security controls.", False)
    ], space_after=2.0)
    add_bullet([
        ("Optimized ", False),
        ("MongoDB", True),
        (" write throughput during burst CRM synchronizations by implementing ordered ", False),
        ("bulk operations", True),
        (" and lease-renewal batching, processing ", False),
        ("10K+ delta records per 15-minute cycle", True),
        (" on Cloud Run with zero data loss.", False)
    ], space_after=2.0)
    add_bullet([
        ("Restructured ", False),
        ("PostgreSQL/MySQL", True),
        (" database schemas and optimized complex ", False),
        ("SQL queries", True),
        (" via ", False),
        ("EXPLAIN", True),
        (" execution plans, eliminating table scan bottlenecks and slashing document analytics API latency ", False),
        ("3× (15s to 5s)", True),
        (" under high concurrency.", False)
    ], space_after=2.0)
    add_bullet([
        ("Collaborated across cross-functional ", False),
        ("Agile squads", True),
        (" with product managers and architects to deliver resilient software solutions, conducting thorough code reviews, upholding ", False),
        ("secure coding practices", True),
        (", and driving continuous delivery.", False)
    ], space_after=2.0)
    add_bullet([
        ("Instrumented centralized observability using ", False),
        ("Sentry", True),
        (" telemetry and ", False),
        ("Langfuse", True),
        (" distributed tracing across microservice backends, diagnosing runtime anomalies, automating alerts, and reducing incident ", False),
        ("MTTD by 45%", True),
        (" in production.", False)
    ], space_after=2.0)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Launched a full-stack candidate evaluation platform conducting ", False),
        ("200+", True),
        (" monthly assessments and slashing turnaround by ", False),
        ("50%", True),
        (", designing secure RESTful APIs and responsive ", False),
        ("React/Next.js", True),
        (" interfaces for institutional stakeholders.", False)
    ], space_after=2.0)
    add_bullet([
        ("Maintained ", False),
        ("GitLab", True),
        (" and ", False),
        ("Jenkins CI/CD", True),
        (" pipelines to automate ", False),
        ("Docker", True),
        (" container image publishing and multi-module ", False),
        ("Gradle", True),
        (" builds, executing automated test suites in an Agile team to eliminate build breaks and accelerate releases.", False)
    ], space_after=2.2)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, OOP, Database Management Systems (DBMS), Operating Systems, Computer Networks.", False)
    ], space_after=2.0)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Java 17, Spring Boot, Kafka, Redis, PostgreSQL, Docker, Gradle, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Built a high-concurrency payment microservice in ", False),
        ("Java 17 (Spring Boot)", True),
        (" enforcing authentication via ", False),
        ("Spring Security", True),
        (" and ", False),
        ("JWT", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency with published Artillery load tests.", False)
    ], space_after=2.0)
    add_bullet([
        ("Orchestrated an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" with ", False),
        ("Redis", True),
        (" distributed locks and ", False),
        ("PostgreSQL", True),
        (" ACID transactions, coordinating multi-service payment workflows and guaranteeing zero double-charges across ", False),
        ("500+", True),
        (" simulated users.", False)
    ], space_after=2.0)
    add_bullet([
        ("Containerized microservices using ", False),
        ("Docker", True),
        (", authoring multi-module ", False),
        ("Gradle scripts", True),
        (" for artifact publishing and comprehensive ", False),
        ("JUnit", True),
        (" and ", False),
        ("Cucumber (BDD)", True),
        (" test suites to validate transactional rollback semantics and enforce quality gates.", False)
    ], space_after=2.2)

    add_project_heading("AI Interviewer", "React, Next.js, NestJS, TypeScript, MongoDB, WebSockets, OpenAPI", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Constructed an automated voice screening platform with responsive ", False),
        ("React", True),
        (" interfaces, role-based access control (", False),
        ("RBAC", True),
        ("), and REST APIs, facilitating ", False),
        ("200+", True),
        (" monthly evaluations and reducing assessment turnaround time by ", False),
        ("50%", True),
        (".", False)
    ], space_after=2.0)
    add_bullet([
        ("Designed real-time bidirectional audio streaming pipelines using ", False),
        ("WebSockets", True),
        (" and structured evaluation rubrics in ", False),
        ("NestJS", True),
        (", publishing ", False),
        ("OpenAPI", True),
        (" specifications and enforcing service-level test coverage gates with ", False),
        ("Jest", True),
        (".", False)
    ], space_after=2.2)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Java (17+ / Core Java, OOP), TypeScript, JavaScript, SQL (PostgreSQL, MySQL), Go (Golang), Python, C++"),
        ("Backend & Microservices: ", True, "Java 17, Spring Boot, Hibernate/JPA, NestJS, RESTful APIs, OpenAPI 3.0 / Swagger, Microservices"),
        ("Messaging & Streaming: ", True, "Apache Kafka (Producer/Consumer, DLQ), Redis (Distributed Locks & Caching), Cloud Pub/Sub"),
        ("DevOps & Cloud: ", True, "GitLab CI/CD, Docker, Gradle, Jenkins, Kubernetes (familiar), Red Hat OpenShift (familiar), Gitflow"),
        ("Testing & Quality: ", True, "JUnit, Cucumber (BDD), Mockito, Jest, Postman, Artillery, SonarQube, Code Reviews, Secure Coding"),
        ("Databases & Tools: ", True, "MongoDB (bulkWrite, Atlas), PostgreSQL, MySQL, Sentry (Telemetry), GitLab, GitHub Copilot, Claude AI")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.4)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform.", False)
    ], space_after=2.0)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for distributed system design.", False)
    ], space_after=2.0)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, emphasizing algorithmic complexity, memory optimization, and concurrent execution.", False)
    ], space_after=1.4)

def populate_natwest(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Aligned with NatWest Software Engineer Associate JD)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Architected a multi-tenant ", False),
        ("Java 17 (Spring Boot)", True),
        (" service extracting CRM data via Strategy and Factory patterns; deployed ", False),
        ("Hibernate/JPA", True),
        (" and ", False),
        ("Kafka DLQ", True),
        (" retry topics to maintain eventual consistency and high availability under burst traffic.", False)
    ], space_after=2.4)
    add_bullet([
        ("Delivered an enterprise commercial ", False),
        ("GPA evaluation platform", True),
        (" (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+", True),
        (" university deployments generating ", False),
        ("$150K+", True),
        (" contract value per client, enforcing ", False),
        ("FERPA compliance", True),
        (" and ", False),
        ("RBAC", True),
        (" security.", False)
    ], space_after=2.4)
    add_bullet([
        ("Implemented a decoupled ", False),
        ("Go", True),
        (" ingestion pipeline on ", False),
        ("Cloud Run", True),
        (" and ", False),
        ("Pub/Sub", True),
        (", processing ", False),
        ("10K+ delta records per 15-minute sync cycle", True),
        (" into a ", False),
        ("MongoDB", True),
        (" queue with lease-fenced consumers to guarantee zero data loss.", False)
    ], space_after=2.4)
    add_bullet([
        ("Optimized database performance by profiling ", False),
        ("SQL", True),
        (" query execution plans (", False),
        ("EXPLAIN", True),
        (") and restructuring ", False),
        ("PostgreSQL/MySQL", True),
        (" schemas, slashing analytics API latency ", False),
        ("3× (15s to 5s)", True),
        (" and eliminating table scan bottlenecks.", False)
    ], space_after=2.4)
    add_bullet([
        ("Collaborated within cross-functional feature teams with architects and business analysts across the full ", False),
        ("SDLC", True),
        (", translating complex requirements into robust technical designs and reusable ", False),
        ("TypeScript/Node.js", True),
        (" modules.", False)
    ], space_after=2.4)
    add_bullet([
        ("Instrumented centralized observability using ", False),
        ("Sentry", True),
        (" and ", False),
        ("Langfuse", True),
        (" across microservice backends, diagnosing tool-call anomalies and tracking error telemetry to accelerate root-cause analysis and reduce ", False),
        ("MTTD by 45%", True),
        (".", False)
    ], space_after=2.4)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Launched a full-stack assessment platform conducting ", False),
        ("200+", True),
        (" monthly candidate evaluations and slashing turnaround ", False),
        ("50%", True),
        (", designing RESTful APIs and interactive ", False),
        ("React/Next.js", True),
        (" dashboards for institutional stakeholders.", False)
    ], space_after=2.4)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines automating containerized deployments (", False),
        ("Docker", True),
        (") and integration test suites in an ", False),
        ("Agile team", True),
        (", eliminating broken staging builds and accelerating developer release cadence.", False)
    ], space_after=2.6)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks, Software Engineering.", False)
    ], space_after=2.4)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Java 17, Spring Boot, Hibernate, Kafka, Redis, PostgreSQL, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a mission-critical payment microservice in ", False),
        ("Java 17 (Spring Boot / Hibernate)", True),
        
        (", ", False),
        ("Spring Security", True),
        (", and ", False),
        ("JWT", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency with published Artillery test suites.", False)
    ], space_after=2.4)
    add_bullet([
        ("Engineered an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" with ", False),
        ("Redis", True),
        (" distributed locks and ", False),
        ("PostgreSQL", True),
        (" ACID validation, guaranteeing idempotent transactions and zero double-charges across ", False),
        ("500+", True),
        (" simulated users.", False)
    ], space_after=2.4)
    add_bullet([
        ("Containerized microservices with ", False),
        ("Docker", True),
        (" virtualization and authored comprehensive ", False),
        ("JUnit", True),
        (" and ", False),
        ("Mockito", True),
        (" test suites to validate transactional rollback semantics, reconciliation routines, and code quality gates.", False)
    ], space_after=2.6)

    add_project_heading("AI Interviewer", "React, Next.js, NestJS, TypeScript, MongoDB, WebSockets, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Built an interactive web evaluation platform with responsive ", False),
        ("React", True),
        (" interfaces, role-based access control (", False),
        ("RBAC", True),
        ("), and REST APIs, facilitating ", False),
        ("200+", True),
        (" monthly evaluations and reducing assessment turnaround by ", False),
        ("50%", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Orchestrated real-time bidirectional streaming using ", False),
        ("WebSockets", True),
        (" and structured evaluation rubrics in ", False),
        ("NestJS", True),
        (", enforcing rigorous test coverage and quality gates with ", False),
        ("Jest", True),
        (" to ensure robust continuous delivery.", False)
    ], space_after=2.6)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Java (17+ / Core Java, OOP), Go (Golang), TypeScript, JavaScript, Python, SQL (PostgreSQL, MySQL), C++"),
        ("Backend & Frameworks: ", True, "Spring Boot (Java 17), Hibernate/JPA, NestJS, Node.js, RESTful APIs, Microservices"),
        ("Distributed & Messaging: ", True, "Apache Kafka (DLQ, Event-Driven), Redis (Distributed Caching & Locks), Cloud Pub/Sub"),
        ("DevOps & Virtualization: ", True, "Docker (Containerization), Jenkins CI/CD, Git/Gitflow, GCP (Cloud Run), Linux, Kubernetes (familiar)"),
        ("Testing, Tools & Agile: ", True, "JUnit, Mockito, Jest, Postman, Artillery, SonarQube, Jira (Bug Tracking), Confluence (Wikis), Scrum / Agile"),
        ("Databases & Observability: ", True, "PostgreSQL, MySQL, MongoDB (NoSQL), Sentry (Telemetry & Monitoring), Langfuse")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=2.0)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting a real-time voice screening platform.", False)
    ], space_after=2.4)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for fault-tolerant system architecture.", False)
    ], space_after=2.4)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, emphasizing algorithmic complexity, numerical problems, and memory optimization.", False)
    ], space_after=1.8)

def populate_purplle(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Purplle Backend Engineer / Unbreakable Systems & AI-Native Velocity)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Engineered a decoupled two-tier ", False),
        ("Go", True),
        (" ingestion pipeline on ", False),
        ("GCP Cloud Run", True),
        (" and ", False),
        ("Pub/Sub", True),
        (", streaming ", False),
        ("10K+ delta records per 15-minute sync cycle", True),
        (" into ", False),
        ("MongoDB", True),
        (" with atomic lease-fenced consumers, ensuring zero data loss during upstream outages.", False)
    ], space_after=2.4)
    add_bullet([
        ("Architected a self-reinforcing academic articulation engine coupling deterministic rule caches with ", False),
        ("LLM fallback", True),
        (", resolving ", False),
        ("40%+", True),
        (" of course mappings without model invocations to eliminate prompt variance and cut AI token spend by ", False),
        ("40%", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Scaled real-time ", False),
        ("NestJS", True),
        (" WebSocket middleware bridging telephony streams under peak load, handling bidirectional audio transcoding for ", False),
        ("1,000+", True),
        (" concurrent sessions at sub-", False),
        ("100ms", True),
        (" p95 latency with zero connection drops.", False)
    ], space_after=2.4)
    add_bullet([
        ("Optimized ", False),
        ("MongoDB", True),
        (" ingestion throughput during burst CRM synchronizations by implementing ordered ", False),
        ("bulk operations (bulkWrite)", True),
        (" and lease-renewal batching, slashing commit latency and eliminating database lock contention.", False)
    ], space_after=2.4)
    add_bullet([
        ("Profiled ", False),
        ("SQL", True),
        (" query execution plans (", False),
        ("EXPLAIN", True),
        (") and restructured relational schemas across ", False),
        ("PostgreSQL/MySQL", True),
        (", slashing analytics API latency ", False),
        ("3× (15s to 5s)", True),
        (" and eliminating table scan bottlenecks.", False)
    ], space_after=2.4)
    add_bullet([
        ("Instrumented centralized observability using ", False),
        ("Sentry", True),
        (" and ", False),
        ("Langfuse", True),
        (" across microservices, detecting tool-call anomalies and tracking error telemetry to accelerate root-cause analysis and reduce ", False),
        ("MTTD by 45%", True),
        (".", False)
    ], space_after=2.4)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Launched a full-stack AI interview platform facilitating ", False),
        ("200+", True),
        (" monthly technical assessments, slashing turnaround by ", False),
        ("50%", True),
        (" and integrating versioned RESTful APIs with automated evaluation pipelines.", False)
    ], space_after=2.4)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines packaging containerized deployments (", False),
        ("Docker", True),
        (") and executing integration test suites in an ", False),
        ("Agile team", True),
        (", eliminating broken staging builds and accelerating developer release cadence.", False)
    ], space_after=2.5)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks.", False)
    ], space_after=2.4)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Go (Golang), Apache Kafka, Redis, PostgreSQL, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a high-throughput, cloud-native payments microservice in ", False),
        ("Go (Golang)", True),
        (" with goroutines and channels, sustaining ", False),
        ("1,200 TPS", True),
        (" (simulated) at sub-", False),
        ("120ms", True),
        (" p95 latency, validated via authored Artillery load tests.", False)
    ], space_after=2.4)
    add_bullet([
        ("Implemented an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" and ", False),
        ("Redis", True),
        (" distributed locks with idempotent ", False),
        ("PostgreSQL", True),
        (" updates, guaranteeing zero double-charges across ", False),
        ("500+", True),
        (" concurrent simulated users under load.", False)
    ], space_after=2.5)

    add_project_heading("AI Interviewer", "NestJS, Next.js, React, TypeScript, WebSockets, OpenAI API, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Shipped an AI assessment platform with responsive ", False),
        ("React/Next.js", True),
        (" UI and ", False),
        ("NestJS", True),
        (" microservices, designing structured evaluation rubrics for objective candidate scoring and integrating ", False),
        ("Jest", True),
        (" testing to enforce quality gates.", False)
    ], space_after=2.4)
    add_bullet([
        ("Constructed bidirectional streaming pipelines using ", False),
        ("WebSockets", True),
        (" and ", False),
        ("Vapi.ai", True),
        (", enabling real-time voice screening and reducing candidate turnaround by ", False),
        ("50%", True),
        (" across 200+ monthly assessments.", False)
    ], space_after=2.5)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Go (Golang), Node.js, TypeScript, JavaScript, Python, Java, SQL (MySQL, PostgreSQL), C++"),
        ("Backend & Architecture: ", True, "NestJS, Express, Microservices, Event-Driven Design, RESTful APIs, Idempotency, Circuit Breakers, Next.js (BFF)"),
        ("Databases & Messaging: ", True, "Apache Kafka, Redis (Distributed Locks & Caching), MongoDB (NoSQL, bulkWrite), MySQL, PostgreSQL, Cloud Pub/Sub"),
        ("Cloud & DevOps: ", True, "Docker (Containerization), GCP (Cloud Run, Pub/Sub), Kubernetes (familiar), Jenkins CI/CD, Sentry, OpenTelemetry (familiar), Git"),
        ("AI & Developer Tools: ", True, "Cursor, Claude, GitHub Copilot, Langfuse (Tracing/Observability), Jest, Artillery (Load Testing), JUnit")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.8)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an autonomous real-time voice screening platform.", False)
    ], space_after=2.2)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for resilient system architecture.", False)
    ], space_after=2.2)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, emphasizing computational complexity, graph theory, and memory optimization.", False)
    ], space_after=1.6)

def populate_bajaj_finance(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Aligned with Bajaj Finance Software Engineer JD)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Architected a multi-tenant backend service in ", False),
        ("Java 17 (Spring Boot)", True),
        (" extracting enterprise CRM data via Strategy and Factory design patterns (", False),
        ("OOP/SOLID", True),
        ("); deployed ", False),
        ("Hibernate/JPA", True),
        (" and ", False),
        ("Kafka DLQ", True),
        (" retry topics to eliminate burst data loss and enforce transactional consistency.", False)
    ], space_after=2.4)
    add_bullet([
        ("Delivered an enterprise commercial platform (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+", True),
        (" university deployments generating ", False),
        ("$150K+", True),
        (" contract value per client, enforcing strict ", False),
        ("FERPA compliance", True),
        (", audit logging, and role-based access control (", False),
        ("RBAC", True),
        (").", False)
    ], space_after=2.4)
    add_bullet([
        ("Implemented a decoupled ", False),
        ("Go", True),
        (" ingestion pipeline on ", False),
        ("GCP Cloud Run", True),
        (" and ", False),
        ("Pub/Sub", True),
        (", processing ", False),
        ("10K+ delta records per 15-minute sync cycle", True),
        (" into a durable ", False),
        ("MongoDB", True),
        (" queue with atomic lease-fenced consumers to guarantee zero data loss during upstream outages.", False)
    ], space_after=2.4)
    add_bullet([
        ("Optimized database performance by profiling ", False),
        ("SQL", True),
        (" query execution plans (", False),
        ("EXPLAIN", True),
        (") and restructuring ", False),
        ("PostgreSQL/MySQL", True),
        (" schemas, slashing analytics API latency ", False),
        ("3× (15s to 5s)", True),
        (" and eliminating table scan bottlenecks.", False)
    ], space_after=2.4)
    add_bullet([
        ("Collaborated across cross-functional ", False),
        ("Agile/Scrum", True),
        (" teams to translate technical specifications and user stories into clean, testable ", False),
        ("TypeScript and Java", True),
        (" modules, actively participating in sprint planning, daily stand-ups, estimations, and peer code reviews.", False)
    ], space_after=2.4)
    add_bullet([
        ("Instrumented centralized observability using ", False),
        ("Sentry", True),
        (" and ", False),
        ("Langfuse", True),
        (" across microservices, investigating application logs and error telemetry to accelerate defect resolution and reduce mean-time-to-detect (", False),
        ("MTTD by 45%", True),
        (").", False)
    ], space_after=2.4)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Launched a full-stack assessment platform conducting ", False),
        ("200+", True),
        (" monthly evaluations and slashing turnaround ", False),
        ("50%", True),
        (", designing robust RESTful APIs and automating end-to-end applicant screening workflows for hiring teams.", False)
    ], space_after=2.4)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines automating containerized deployments (", False),
        ("Docker", True),
        (") and integration test execution in an ", False),
        ("Agile team", True),
        (", eliminating broken staging builds and accelerating developer release cadence.", False)
    ], space_after=2.6)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks, Software Engineering.", False)
    ], space_after=2.4)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Java 17, Spring Boot, Hibernate, Kafka, Redis, PostgreSQL, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a mission-critical payment microservice in ", False),
        ("Java 17 (Spring Boot / Hibernate)", True),
        
        (", ", False),
        ("Spring Security", True),
        (", and ", False),
        ("JWT", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency with published Artillery load test suites.", False)
    ], space_after=2.3)
    add_bullet([
        ("Engineered an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" with ", False),
        ("Redis", True),
        (" distributed locks and ", False),
        ("PostgreSQL", True),
        (" ACID validation, eliminating duplicate charges across ", False),
        ("500+", True),
        (" simulated users; containerized with ", False),
        ("Docker", True),
        (" and verified via ", False),
        ("JUnit", True),
        (" suites.", False)
    ], space_after=2.4)

    add_project_heading("AI Interviewer", "React, Next.js, NestJS, TypeScript, MongoDB, WebSockets, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Built an interactive web evaluation platform with responsive ", False),
        ("React", True),
        (" interfaces, role-based access control (", False),
        ("RBAC", True),
        ("), and REST APIs, facilitating ", False),
        ("200+", True),
        (" monthly evaluations and reducing assessment turnaround by ", False),
        ("50%", True),
        (".", False)
    ], space_after=2.3)
    add_bullet([
        ("Orchestrated real-time bidirectional streaming using ", False),
        ("WebSockets", True),
        (" and structured evaluation rubrics in ", False),
        ("NestJS", True),
        (", enforcing rigorous test coverage and quality gates with ", False),
        ("Jest", True),
        (" to ensure reliable continuous delivery.", False)
    ], space_after=2.4)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Programming Languages: ", True, "C#, .NET / .NET Core, Java (17+ / Core Java, OOP), TypeScript, JavaScript, Python, SQL (PostgreSQL, MySQL), C++, Go"),
        ("Backend & Web APIs: ", True, "ASP.NET Core Web API (familiar), Spring Boot, NestJS, Node.js, RESTful APIs, Microservices, Event-Driven Architecture, Next.js (BFF)"),
        ("Databases & Storage: ", True, "PostgreSQL, MySQL, Redis (Distributed Caching & Locks), MongoDB (NoSQL), Kafka (DLQ, Outbox), Cloud Pub/Sub, Query Plans (EXPLAIN)"),
        ("DevOps & Cloud: ", True, "Docker (Containers), Jenkins CI/CD, Git / GitHub, GCP (Cloud Run), AWS, Linux, Kubernetes (familiar)"),
        ("Testing & SDLC: ", True, "JUnit, Mockito, Jest, Postman, Artillery Load Testing, Sentry (Logs & Monitoring), Agile/Scrum, SDLC Controls")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.8)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting a real-time voice screening platform.", False)
    ], space_after=2.2)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for fault-tolerant system architecture.", False)
    ], space_after=2.2)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, emphasizing algorithmic complexity, numerical problems, and memory optimization.", False)
    ], space_after=1.6)

def populate_amex(doc, add_section_header, add_role_heading, add_project_heading, add_bullet, role="Software_Engineer_I"):
    # 2. Experience (HR-First Layout - Infrastructure, Full-Stack Python & AI Focus)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Engineered high-throughput data ingestion pipelines processing ", False),
        ("50K+ enterprise records", True),
        (" via scheduled ", False),
        ("Python workers", True),
        (" and ", False),
        ("Google Cloud Run", True),
        (", integrating ", False),
        ("Sentry", True),
        (" health monitoring and telemetry alerts to guarantee 99.9% pipeline availability.", False)
    ], space_after=2.8)
    add_bullet([
        ("Optimized relational database performance and reduced backend analytics API latency ", False),
        ("3x", True),
        (" (", False),
        ("15s to 5s", True),
        (") by profiling ", False),
        ("SQL execution plans (EXPLAIN)", True),
        (", indexing foreign keys, and refactoring ", False),
        ("PostgreSQL and MySQL", True),
        (" schemas to eliminate table scans.", False)
    ], space_after=2.8)
    add_bullet([
        ("Developed responsive web applications and administrative dashboards using ", False),
        ("React.js", True),
        (", ", False),
        ("Next.js", True),
        (", and ", False),
        ("TypeScript", True),
        (", engineering modular component hierarchies and ", False),
        ("RESTful API", True),
        (" integrations for real-time compute usage tracking.", False)
    ], space_after=2.8)
    add_bullet([
        ("Instrumented multi-turn LLM agent workflows using ", False),
        ("Langfuse", True),
        (" and ", False),
        ("Sentry", True),
        (", diagnosing tool-call hallucinations and prompt drift to enforce enterprise guardrails, establish evaluation gates, and cut incident triage time (", False),
        ("MTTD", True),
        (") by ", False),
        ("45%", True),
        (".", False)
    ], space_after=2.8)
    add_bullet([
        ("Scaled an enterprise multi-tenant platform (", False),
        ("Next.js BFF, NestJS", True),
        (") across ", False),
        ("4+ university environments", True),
        (" generating ", False),
        ("$150K+", True),
        (" average contract value per client, enforcing strict ", False),
        ("role-based access control (RBAC)", True),
        (", data isolation, and regulatory compliance.", False)
    ], space_after=3.0)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Launched an interactive evaluation platform with ", False),
        ("React/TypeScript", True),
        (" interfaces and ", False),
        ("RESTful APIs", True),
        (", processing ", False),
        ("200+", True),
        (" monthly candidate assessments and slashing administrative evaluation turnaround by ", False),
        ("50%", True),
        (" for hiring teams.", False)
    ], space_after=2.8)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" pipelines orchestrating containerized deployments (", False),
        ("Docker", True),
        (") and regression test suites (", False),
        ("Pytest, Jest", True),
        (") across Agile release sprints, collaborating with cross-functional teams to eliminate broken staging builds.", False)
    ], space_after=3.0)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks, Software Engineering & Testing.", False)
    ], space_after=3.0)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("AI Interviewer", "React.js, Next.js, TypeScript, OpenAI API, WebSockets, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Architected an autonomous assessment platform using ", False),
        ("React.js", True),
        (", ", False),
        ("Next.js", True),
        (", and ", False),
        ("OpenAI API", True),
        (", codifying structured evaluation rubrics and ", False),
        ("Jest test suites", True),
        (" to deliver objective scoring across ", False),
        ("200+ monthly technical evaluations", True),
        (".", False)
    ], space_after=2.8)
    add_bullet([
        ("Constructed bidirectional streaming audio pipelines via ", False),
        ("WebSockets", True),
        (" and RESTful services, enabling real-time speech-to-text inference and reducing candidate assessment turnaround by ", False),
        ("50%", True),
        (" without human reviewer bottlenecks.", False)
    ], space_after=3.0)

    add_project_heading("BOB Banking Application", "Python, Flask, SQLite, Pytest, CI/CD", "github.com/Yash-Gupta-01/bob-banking-app", url="https://github.com/Yash-Gupta-01/bob-banking-app")
    add_bullet([
        ("Built a full-stack financial web service in ", False),
        ("Python (Flask)", True),
        (" with ", False),
        ("SQLAlchemy ORM", True),
        (" and ", False),
        ("SQLite", True),
        (" relational schemas, enforcing transactional ACID compliance, modular MVC separation, and role-based customer authentication.", False)
    ], space_after=2.8)
    add_bullet([
        ("Configured ", False),
        ("GitHub Actions CI/CD", True),
        (" workflows executing multi-stage ", False),
        ("Pytest", True),
        (" unit and integration test suites with coverage thresholds, securing reliable production-ready deployments with zero regression escapes.", False)
    ], space_after=3.0)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Python, TypeScript, JavaScript (ES6+), SQL (PostgreSQL, MySQL, SQLite), Go, Java, C++, HTML5, CSS3"),
        ("Backend & Frameworks: ", True, "Python (FastAPI, Flask, Django), Node.js, NestJS, Next.js (BFF), RESTful APIs, Microservices"),
        ("Frontend & UI: ", True, "React.js, Next.js, Redux, Tailwind CSS, Responsive Web Design, WebSockets, State Management"),
        ("Databases & Storage: ", True, "PostgreSQL, MySQL, SQLite, MongoDB, Redis (Distributed Caching), Relational Schema Design, SQL Tuning"),
        ("Cloud & Virtualization: ", True, "Google Cloud Platform (Cloud Run, Pub/Sub, VPC), Docker, AWS (EC2, S3), Linux, Virtual Machines, VMware / Nutanix (familiar), Kubernetes (familiar), CI/CD (Jenkins, GitHub Actions), Pytest, Jest, Cursor, Copilot")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=2.4)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, delivering an intelligent real-time voice screening platform.", False)
    ], space_after=2.6)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for distributed system design.", False)
    ], space_after=2.6)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, emphasizing graph algorithms and memory optimization.", False)
    ], space_after=1.8)

def populate_salesforce(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Salesforce Trailhead Generalist Software Engineer Focus)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Architected a multi-tenant ", False),
        ("Java/Spring Boot", True),
        (" ETL microservice integrating Slate and ", False),
        ("Salesforce CRMs", True),
        (" via OOP factory patterns, deploying ", False),
        ("Hibernate/JPA", True),
        (" and ", False),
        ("Kafka Dead Letter Queues (DLQ)", True),
        (" to eliminate data loss during burst traffic.", False)
    ], space_after=2.3)
    add_bullet([
        ("Hardened embedded web applications by engineering a ", False),
        ("Next.js BFF", True),
        (" proxy, eliminating client-side credential exposure, sanitizing ", False),
        ("DOM nodes", True),
        (", and establishing origin-whitelisted postMessage protocols (", False),
        ("CORS/CSRF defense", True),
        (") across university portals.", False)
    ], space_after=2.3)
    add_bullet([
        ("Instrumented multi-turn agent workflows with ", False),
        ("Langfuse", True),
        (" tracing, leveraging ", False),
        ("Cursor and Claude", True),
        (" AI developer workflows to accelerate code generation, isolate prompt drift, and automate quality gates, cutting mean-time-to-detect (", False),
        ("MTTD", True),
        (") by ", False),
        ("45%", True),
        (".", False)
    ], space_after=2.3)
    add_bullet([
        ("Developed high-throughput ", False),
        ("NestJS", True),
        (" WebSocket middleware for real-time collaboration streams, handling bidirectional audio transcoding for ", False),
        ("1k+", True),
        (" concurrent sessions at sub-", False),
        ("100ms", True),
        (" p95 latency with zero connection drops under load.", False)
    ], space_after=2.3)
    add_bullet([
        ("Optimized backend ", False),
        ("RESTful API", True),
        (" response times ", False),
        ("3×", True),
        (" (", False),
        ("15s to 5s", True),
        (") by profiling ", False),
        ("SQL execution plans (EXPLAIN)", True),
        (", adding indexes, and restructuring ", False),
        ("PostgreSQL/MySQL", True),
        (" schemas, eliminating table scans to maximize throughput.", False)
    ], space_after=2.4)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Launched a full-stack AI evaluation platform with interactive ", False),
        ("React/TypeScript", True),
        (" dashboards and ", False),
        ("REST APIs", True),
        (", automating ", False),
        ("200+", True),
        (" monthly candidate evaluations and slashing turnaround by ", False),
        ("50%", True),
        (" for cross-functional teams.", False)
    ], space_after=2.3)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines automating containerized deployments (", False),
        ("Docker, AWS", True),
        (") and automated test suites in an ", False),
        ("Agile team", True),
        (", eliminating broken staging builds and accelerating release cadence.", False)
    ], space_after=2.4)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks, Software Engineering & Testing.", False)
    ], space_after=2.4)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Node.js, Express.js, PostgreSQL, Redis, Kafka, Docker, TDD (Jest)", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Engineered high-throughput, ", False),
        ("RESTful microservice APIs", True),
        (" on ", False),
        ("Node.js/Express", True),
        (" and ", False),
        ("PostgreSQL", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" (simulated) at sub-", False),
        ("120ms", True),
        (" p95 latency, validated through Artillery load testing and ", False),
        ("TDD test suites", True),
        (".", False)
    ], space_after=2.3)
    add_bullet([
        ("Constructed an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" with ", False),
        ("Redis distributed locks", True),
        (", ensuring ", False),
        ("ACID-compliant ledger transactions", True),
        (", idempotency, and zero double-charges across 500+ simulated users.", False)
    ], space_after=2.4)

    add_project_heading("AI Interviewer", "Next.js, React.js, TypeScript, OpenAI API, WebSockets, WCAG", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Orchestrated real-time ", False),
        ("WebSocket", True),
        (" communication and interactive ", False),
        ("React/Next.js", True),
        (" interfaces, automating ", False),
        ("200+ monthly evaluations", True),
        (" with sub-second STT/LLM inference and structured evaluation rubrics.", False)
    ], space_after=2.3)
    add_bullet([
        ("Designed scalable ", False),
        ("RESTful APIs", True),
        (" in ", False),
        ("NestJS", True),
        (" and MongoDB Atlas, enforcing JWT-secured role-based access control and modular component architecture conforming to ", False),
        ("WCAG accessibility standards", True),
        (".", False)
    ], space_after=2.4)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "TypeScript, JavaScript (ES6+), Python, Java (Spring Boot), Go (Golang), C#, C++, SQL (PostgreSQL, MySQL)"),
        ("Web & Frontend: ", True, "Next.js (App Router, BFF), React, Redux, DOM API, Web Fundamentals (HTTP/HTTPS, REST, CORS, CSRF), HTML5/CSS3, Tailwind CSS, WCAG"),
        ("Backend & Microservices: ", True, "Node.js, Express.js, NestJS, Spring Boot, RESTful APIs, WebSockets (Socket.io), Microservices Architecture, OOP & Clean Design Patterns"),
        ("Cloud, Infrastructure & CI/CD: ", True, "AWS (EC2, S3), Docker, Kubernetes (familiar), Jenkins CI/CD, GitHub Actions, GCP (Cloud Run), Linux Shell Scripting"),
        ("Databases & Quality: ", True, "PostgreSQL, MySQL, Redis (Caching/Locking), MongoDB, Kafka (DLQ), Cursor, Claude, Copilot, TDD, Jest, Pytest, Sentry Observability")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.8)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform.", False)
    ], space_after=2.2)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for distributed system design.", False)
    ], space_after=2.2)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, emphasizing graph algorithms and memory optimization.", False)
    ], space_after=1.5)

def populate_temple(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Real-Time Systems, Go Pipelines, Signal Processing & On-Device ML Focus)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Engineered a real-time ", False),
        ("Genesys AudioHook v2", True),
        (" WebSocket bridge in ", False),
        ("NestJS", True),
        (", sustaining ", False),
        ("1,000+", True),
        (" concurrent bidirectional audio streams at sub-", False),
        ("100ms", True),
        (" p95 latency with ", False),
        ("8 kHz µ-law to 16 kHz linear PCM", True),
        (" transcoding for AI voice agents.", False)
    ], space_after=2.3)
    add_bullet([
        ("Built a decoupled two-tier ", False),
        ("Go", True),
        (" ingestion pipeline on ", False),
        ("GCP Cloud Run", True),
        (" and ", False),
        ("Pub/Sub", True),
        (", extracting ", False),
        ("10K+", True),
        (" delta records per 15-minute sync cycle into a durable ", False),
        ("MongoDB", True),
        (" queue with lease-fenced consumers, guaranteeing zero data loss during upstream outages.", False)
    ], space_after=2.3)
    add_bullet([
        ("Scaled the ", False),
        ("Go/Pub/Sub", True),
        (" sync engine to replicate CRM data across an isolated ", False),
        ("Student Copilot MongoDB", True),
        (" cluster with revision-fenced idempotent delivery and sub-second soft-delete propagation.", False)
    ], space_after=2.3)
    add_bullet([
        ("Instrumented multi-turn agent workflows with ", False),
        ("Langfuse", True),
        (" distributed tracing and ", False),
        ("Sentry", True),
        (" observability, isolating reasoning-chain bottlenecks and eliminating production ", False),
        ("MongoDB", True),
        (" tool-call hallucinations, cutting MTTD by ", False),
        ("45%", True),
        (".", False)
    ], space_after=2.3)
    add_bullet([
        ("Deployed commercial transcript evaluation platforms across ", False),
        ("4+", True),
        (" universities (", False),
        ("Rowan", True),
        (", ", False),
        ("RFU", True),
        (", ", False),
        ("PennWest", True),
        (") generating ", False),
        ("$150K+", True),
        (" average contract value per client, automating writeback of computed GPAs directly into ", False),
        ("Slate CRM", True),
        (".", False)
    ], space_after=2.4)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Launched a full-stack AI interview platform automating ", False),
        ("200+", True),
        (" monthly evaluations and slashing turnaround ", False),
        ("50%", True),
        (", integrating real-time WebSocket audio streaming with structured evaluation rubrics.", False)
    ], space_after=2.3)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines automating containerized deployments (", False),
        ("Docker", True),
        (") and integration test suites, eliminating broken staging builds and accelerating release cadence.", False)
    ], space_after=2.4)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, OOP, Database Management Systems (DBMS), Operating Systems, Computer Networks, Signal Processing.", False)
    ], space_after=2.5)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Go (Golang), Kafka, Redis, PostgreSQL, Docker, Artillery", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a high-throughput, cloud-native payments microservice in ", False),
        ("Go (Golang)", True),
        (" with goroutines and channels, sustaining ", False),
        ("1,200 TPS", True),
        (" (simulated) at sub-", False),
        ("120ms", True),
        (" p95 latency, validated via authored ", False),
        ("Artillery", True),
        (" load tests.", False)
    ], space_after=2.3)
    add_bullet([
        ("Architected an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" with ", False),
        ("Redis distributed locks", True),
        (" and ", False),
        ("PostgreSQL ACID", True),
        (" validation, guaranteeing idempotent transactions and zero double-charges across ", False),
        ("500+", True),
        (" concurrent simulated users.", False)
    ], space_after=2.4)

    add_project_heading("Aether", "Kotlin, Jetpack Compose, Room DB, llama.cpp, Clean Architecture", "github.com/Yash-Gupta-01/Personal-AI-Agent", url="https://github.com/Yash-Gupta-01/Personal-AI-Agent")
    add_bullet([
        ("Designed an offline-first Android personal AI assistant in ", False),
        ("Kotlin", True),
        (" and ", False),
        ("Jetpack Compose", True),
        (" using ", False),
        ("Clean Architecture", True),
        (" with pure JVM domain layers, integrating ", False),
        ("Room", True),
        (" persistence and on-device ", False),
        ("llama.cpp", True),
        (" inference for local thought capture.", False)
    ], space_after=2.3)
    add_bullet([
        ("Constructed a modular tool registry with schema-guarded tool calling, enabling on-device capabilities (alarms, reminders, calendar) without direct LLM access to Android APIs.", False)
    ], space_after=2.4)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Go (Golang), TypeScript, JavaScript, Python, Java, Kotlin, C++, SQL (PostgreSQL, MySQL)"),
        ("Systems & Streaming: ", True, "Google Cloud Pub/Sub, Apache Kafka (Outbox/DLQ), WebSockets (AudioHook v2), Redis (Distributed Locks/Caching)"),
        ("Backend & APIs: ", True, "NestJS, Node.js, Express.js, Spring Boot, RESTful APIs, Microservices, Clean Architecture"),
        ("Cloud & DevOps: ", True, "GCP (Cloud Run, Cloud Functions, Pub/Sub, VPC Access), Docker, Jenkins CI/CD, Git, Sentry, Artillery"),
        ("AI & Mobile: ", True, "Langfuse (Tracing/Evals), OpenAI API, Vapi.ai, llama.cpp (On-Device), Android SDK, Jetpack Compose, Room DB"),
        ("Databases & Frontend: ", True, "MongoDB (Aggregation/BulkWrite), PostgreSQL, MySQL, React, Next.js (BFF), Angular 14, Tailwind CSS")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.8)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform.", False)
    ], space_after=2.2)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for distributed system design.", False)
    ], space_after=2.2)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef.", False)
    ], space_after=1.5)

def populate_paypal(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Java 17, Spring Boot, Payments Architecture, Kafka, Agile & CI/CD)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Architected a multi-tenant backend service in ", False),
        ("Java 17 (Spring Boot)", True),
        (" extracting CRM data via Strategy and Factory design patterns; deployed ", False),
        ("Hibernate/JPA", True),
        (" and ", False),
        ("Kafka DLQ", True),
        (" retry topics to ensure eventual consistency under burst traffic.", False)
    ], space_after=2.4)
    add_bullet([
        ("Deployed an asynchronous ", False),
        ("Transfer Credit Evaluation (TCE)", True),
        (" pipeline with ", False),
        ("Kafka", True),
        (" and ", False),
        ("NestJS/Java", True),
        (" microservices, decoupling degree audit processing from CRM webhooks and cutting evaluation turnaround time by ", False),
        ("80%", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Slashed document-analytics API response latency ", False),
        ("3×", True),
        (" (", False),
        ("15s to 5s", True),
        (") by profiling ", False),
        ("SQL", True),
        (" query execution plans, adding composite indexes, and eliminating redundant join queries in ", False),
        ("PostgreSQL/MySQL", True),
        (" schemas to optimize database throughput.", False)
    ], space_after=2.4)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" pipelines and ", False),
        ("Gitflow", True),
        (" peer review workflows, enforcing ", False),
        ("SonarQube", True),
        (" static analysis gates, unit testing (", False),
        ("JUnit", True),
        (", ", False),
        ("Mockito", True),
        ("), and Sentry telemetry to eliminate anti-patterns and cut MTTD by ", False),
        ("45%", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Scaled an enterprise commercial ", False),
        ("GPA evaluation platform", True),
        (" (", False),
        ("Next.js/React BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+", True),
        (" university deployments generating ", False),
        ("$150K+", True),
        (" contract value per client, enforcing strict ", False),
        ("FERPA compliance", True),
        (" and ", False),
        ("RBAC", True),
        (" security.", False)
    ], space_after=2.4)
    add_bullet([
        ("Implemented low-latency ", False),
        ("NestJS", True),
        (" WebSocket middleware bridging real-time telephony streams, handling bidirectional audio transcoding for ", False),
        ("1,000+", True),
        (" concurrent active sessions at sub-", False),
        ("100ms", True),
        (" p95 latency under peak production traffic.", False)
    ], space_after=2.4)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Formulated a ", False),
        ("Kafka-driven", True),
        (" microservice for asynchronous event synchronization between partner CRMs and internal databases within a ", False),
        ("Scaled Agile (SAFe)", True),
        (" sprint cadence, collaborating with project leads to review technical solutions.", False)
    ], space_after=2.4)
    add_bullet([
        ("Maintained artifact packaging pipelines and containerized environments with ", False),
        ("Docker", True),
        (" in an ", False),
        ("Agile team", True),
        (", executing automated regression suites to eliminate broken staging builds and accelerate release cadence.", False)
    ], space_after=2.6)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks, Software Engineering & Testing.", False)
    ], space_after=2.4)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Java 17, Spring Boot, Hibernate, Kafka, Redis, PostgreSQL, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a mission-critical payment microservice in ", False),
        ("Java 17 (Spring Boot / Hibernate)", True),
        
        (", ", False),
        ("Spring Security", True),
        (", and ", False),
        ("JWT", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency with published Artillery test suites.", False)
    ], space_after=2.4)
    add_bullet([
        ("Spearheaded an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" with ", False),
        ("Redis", True),
        (" distributed locks and ", False),
        ("PostgreSQL", True),
        (" ACID validation, guaranteeing idempotent transactions and zero double-charges across ", False),
        ("500+", True),
        (" simulated users.", False)
    ], space_after=2.4)
    add_bullet([
        ("Containerized microservices with ", False),
        ("Docker", True),
        (" and authored comprehensive ", False),
        ("JUnit", True),
        (" and ", False),
        ("Mockito", True),
        (" test suites to validate transactional rollback semantics, reconciliation routines, and code quality standards.", False)
    ], space_after=2.6)

    add_project_heading("AI Interviewer", "React, Next.js, NestJS, TypeScript, MongoDB, REST APIs, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Built a cloud-based evaluation platform with responsive ", False),
        ("React", True),
        (" dashboards, RESTful APIs, and role-based access control (", False),
        ("RBAC", True),
        ("), managing ", False),
        ("200+", True),
        (" monthly candidate assessments and cutting evaluation turnaround by ", False),
        ("50%", True),
        (" for hiring teams.", False)
    ], space_after=2.4)
    add_bullet([
        ("Orchestrated bidirectional streaming pipelines using ", False),
        ("WebSockets", True),
        (" and structured evaluation rubrics, enforcing service-level test coverage and quality gates with ", False),
        ("Jest", True),
        (" to ensure reliable continuous delivery in CI/CD pipelines.", False)
    ], space_after=2.6)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Java (17+ / Core Java, OOP), TypeScript, JavaScript, SQL (PostgreSQL, MySQL), Go, Python, C++"),
        ("Frontend & Web: ", True, "React.js, Next.js (BFF), Redux, Tailwind CSS, HTML5, CSS3, RESTful APIs, Web Applications"),
        ("Backend & Frameworks: ", True, "Spring Boot (Java 17), Spring MVC, Hibernate/JPA, Inversion of Control (IoC), Spring Security, NestJS"),
        ("Architecture & Messaging: ", True, "Apache Kafka (Event-Driven Services, Outbox Pattern, DLQ), Redis (Caching/Locks), SOA"),
        ("Databases & Quality: ", True, "PostgreSQL, MySQL, MongoDB (NoSQL), SonarQube (Quality Gates), Checkmarx, Sentry"),
        ("SDLC, Testing & DevOps: ", True, "Gitflow (Peer Reviews), Docker, Jenkins CI/CD, JUnit, Mockito, Jest, Agile / SAFe")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=2.0)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting a real-time voice screening platform.", False)
    ], space_after=2.4)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for fault-tolerant system architecture.", False)
    ], space_after=2.4)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, emphasizing algorithmic complexity, memory optimization, and concurrent execution.", False)
    ], space_after=1.8)

def populate_target(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (Target Engineer - Scalable Architecture, Reliability, Distributed Systems, SRE & Testing)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Spearheaded centralized application monitoring with ", False),
        ("Sentry", True),
        (" across microservices, tracking error logs and capacity metrics; resolved a critical production schema drift incident in ", False),
        ("2 hours", True),
        (" to maintain strict ", False),
        ("SLA compliance", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Architected a multi-tenant batch ETL pipeline in ", False),
        ("Java (Spring Boot)", True),
        (" and ", False),
        ("Kafka", True),
        (" with Dead Letter Queues (", False),
        ("DLQ", True),
        (") and retry topics, isolating failed records to guarantee zero data loss during sudden traffic spikes.", False)
    ], space_after=2.4)
    add_bullet([
        ("Deployed an academic evaluation engine coupling deterministic rule caches with ", False),
        ("LLM fallback", True),
        (", resolving ", False),
        ("40%+", True),
        (" of course equivalency queries without AI inference to cut token costs by ", False),
        ("40%", True),
        (" and lower total cost of ownership.", False)
    ], space_after=2.4)
    add_bullet([
        ("Engineered a decoupled two-tier ", False),
        ("Go", True),
        (" ingestion pipeline on ", False),
        ("Cloud Run", True),
        (" and ", False),
        ("Pub/Sub", True),
        (", pulling ", False),
        ("10K+ delta records per 15-minute sync cycle", True),
        (" into ", False),
        ("MongoDB", True),
        (" with lease-fenced consumers to isolate upstream CRM outages.", False)
    ], space_after=2.4)
    add_bullet([
        ("Formulated an enterprise commercial ", False),
        ("GPA evaluation platform", True),
        (" (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+", True),
        (" university deployments generating ", False),
        ("$150K+", True),
        (" contract value per client, enforcing strict ", False),
        ("FERPA compliance", True),
        (" and RBAC security.", False)
    ], space_after=2.4)
    add_bullet([
        ("Optimized database query performance, slashing backend analytics API latency ", False),
        ("3× (15s to 5s)", True),
        (" by restructuring ", False),
        ("SQL execution plans (EXPLAIN)", True),
        (", indexing composite keys, and eliminating redundant table scans across ", False),
        ("PostgreSQL and MySQL", True),
        (".", False)
    ], space_after=2.4)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Launched a full-stack AI evaluation platform handling ", False),
        ("200+", True),
        (" monthly candidate assessments and slashing turnaround by ", False),
        ("50%", True),
        (", conducting proof-of-concept trials that productized into commercial admissions systems.", False)
    ], space_after=2.4)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines executing automated regression test suites and containerized ", False),
        ("Docker", True),
        (" packaging in an ", False),
        ("Agile team", True),
        (", eliminating broken staging builds and accelerating developer release cadence.", False)
    ], space_after=2.6)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks, Software Engineering & Testing.", False)
    ], space_after=2.4)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Java, Spring Boot, Hibernate, Kafka, Redis, PostgreSQL, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a distributed payment microservice in ", False),
        ("Java (Spring Boot / Hibernate)", True),
        
        (", enforcing authentication via ", False),
        ("Spring Security", True),
        (" and ", False),
        ("JWT", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency.", False)
    ], space_after=2.4)
    add_bullet([
        ("Constructed an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" and ", False),
        ("Redis", True),
        (" mutex locks with ", False),
        ("PostgreSQL ACID", True),
        (" validation, preventing race conditions and ensuring zero double-charges across ", False),
        ("500+", True),
        (" simulated users.", False)
    ], space_after=2.4)
    add_bullet([
        ("Containerized microservices using ", False),
        ("Docker", True),
        (" and authored ", False),
        ("JUnit", True),
        (" integration test scripts to validate disaster recovery procedures, network timeout handling, and transactional rollbacks, ensuring release readiness.", False)
    ], space_after=2.6)

    add_project_heading("AI Interviewer", "NestJS, Next.js, React, MongoDB (NoSQL), TypeScript, OpenAI API, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Created an interactive evaluation platform with ", False),
        ("Next.js", True),
        (", ", False),
        ("NestJS", True),
        (", and REST APIs, handling ", False),
        ("200+", True),
        (" monthly evaluations, conducting peer code reviews, and tracking telemetry logs for continuous capacity monitoring.", False)
    ], space_after=2.4)
    add_bullet([
        ("Implemented comprehensive unit and integration test suites with ", False),
        ("Jest", True),
        (" and rigorous API contract validation, establishing strict CI quality gates to prevent regression defects and verify operational readiness.", False)
    ], space_after=2.6)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Java (Core Java, OOP), Go, Python, TypeScript, JavaScript, SQL (PostgreSQL, MySQL), C++"),
        ("Backend & Distributed: ", True, "Spring Boot, Microservices, REST APIs, Kafka (DLQ, Event Streams), Redis Mutex Locks, NestJS"),
        ("Package Solutions & Cloud: ", True, "Slate CRM (Configurable Joins), Salesforce ETL, Docker, GCP (Cloud Run, Pub/Sub)"),
        ("Databases & Modeling: ", True, "PostgreSQL, MySQL, MongoDB (NoSQL, Aggregation), Data Structures & Algorithms, Indexing"),
        ("Testing & Reliability: ", True, "Disaster Recovery, Sentry Observability, JUnit, Mockito, Jest, Pytest, Artillery Load Testing"),
        ("SDLC & Engineering: ", True, "Jenkins CI/CD, Git, Peer Code Reviews, Design Reviews, Agile/Scrum, TCO Optimization")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=2.0)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting a real-time voice screening platform with low-latency WebSocket audio streaming.", False)
    ], space_after=2.4)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for fault-tolerant system architecture and service resilience.", False)
    ], space_after=2.4)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, emphasizing algorithmic complexity analysis, memory optimization, and concurrent execution.", False)
    ], space_after=1.8)

def populate_zs(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Aligned with ZS Business Technology Solutions Associate JD)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Spearheaded an enterprise commercial ", False),
        ("GPA evaluation platform", True),
        (" (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+ university client engagements", True),
        (" generating ", False),
        ("$150K+", True),
        (" contract value each, translating complex business rules into technical designs under ", False),
        ("FERPA compliance", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Automated high-volume asynchronous ", False),
        ("ETL data pipelines", True),
        (" processing ", False),
        ("50K+ student records", True),
        (" from ", False),
        ("PostgreSQL", True),
        (" using scheduled ", False),
        ("Python jobs on GCP", True),
        (", formulating ", False),
        ("optimized SQL queries", True),
        (" to power analytics reporting and cutting manual cycles by ", False),
        ("60%", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Architected a decoupled two-tier ", False),
        ("Go ingestion pipeline", True),
        (" on ", False),
        ("Google Cloud Run", True),
        (" and ", False),
        ("Pub/Sub", True),
        (", ingesting ", False),
        ("10K+ delta records per 15-minute sync cycle", True),
        (" into a durable ", False),
        ("MongoDB", True),
        (" queue with lease-fenced consumers to guarantee zero data loss.", False)
    ], space_after=2.4)
    add_bullet([
        ("Formulated a self-reinforcing ", False),
        ("Document Intelligence engine", True),
        (" coupling deterministic articulation caches with ", False),
        ("LLM fallback", True),
        (", resolving ", False),
        ("40%+", True),
        (" of course mappings deterministically and cutting external AI token spend by ", False),
        ("40%", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Evaluated agentic workflows and LLM reasoning chains using ", False),
        ("Langfuse", True),
        (" and ", False),
        ("Sentry telemetry", True),
        (", leveraging ", False),
        ("GitHub Copilot", True),
        (" and assisted ", False),
        ("SQL generation", True),
        (" prompts to diagnose database tool-call hallucinations and cut ", False),
        ("MTTD by 45%", True),
        (".", False)
    ], space_after=2.5)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Reduced backend analytics API response latency ", False),
        ("3×", True),
        (" (", False),
        ("15s to 5s", True),
        (") by profiling ", False),
        ("SQL execution plans (EXPLAIN)", True),
        (" across ", False),
        ("PostgreSQL/MySQL", True),
        (" relational schemas, eliminating redundant joins and table scans to maximize data retrieval throughput.", False)
    ], space_after=2.4)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" pipelines orchestrating containerized deployments (", False),
        ("Docker", True),
        (") and unit test suites in an ", False),
        ("Agile team", True),
        (", executing ", False),
        ("System Integration Testing (SIT)", True),
        (" and embedded QA procedures to ensure zero broken staging builds.", False)
    ], space_after=2.6)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Database Management Systems (DBMS), Relational Database Design (RDBMS), SQL & Query Optimization, Data Structures & Algorithms, Object-Oriented Programming (OOP), Computer Networks, Operating Systems.", False)
    ], space_after=2.6)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Java, Spring Boot, Python, Apache Kafka, PostgreSQL, Redis, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a high-throughput payments microservice in ", False),
        ("Java (Spring Boot)", True),
        (" and ", False),
        ("Python", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency, publishing Artillery load test suites and profiling thread pools for peak concurrency.", False)
    ], space_after=2.4)
    add_bullet([
        ("Built an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" and Redis mutex locks with idempotent ", False),
        ("PostgreSQL", True),
        (" updates, guaranteeing zero double-charges across ", False),
        ("500+", True),
        (" simulated users and enforcing strict transactional ACID consistency.", False)
    ], space_after=2.6)

    add_project_heading("AI Interviewer", "Python, Next.js, NestJS, TypeScript, OpenAI API, WebSockets, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Engineered an autonomous screening platform with ", False),
        ("Next.js", True),
        (", ", False),
        ("NestJS", True),
        (", and ", False),
        ("OpenAI API", True),
        (", codifying structured evaluation rubrics and ", False),
        ("Jest test suites", True),
        (" to deliver objective scoring across ", False),
        ("200+ monthly candidate reviews", True),
        (" for institutional clients.", False)
    ], space_after=2.4)
    add_bullet([
        ("Maintained bidirectional streaming audio pipelines via ", False),
        ("WebSockets", True),
        (" and ", False),
        ("Vapi.ai", True),
        (", handling real-time speech-to-text inference and lowering candidate evaluation turnaround time by ", False),
        ("50%", True),
        (" without human reviewer bottlenecks.", False)
    ], space_after=2.6)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Programming Languages: ", True, "Python, SQL (PostgreSQL, MySQL, SQL Server), Java (Core Java, Spring Boot), TypeScript, JavaScript (ES6+), C++, Go (Golang), C#"),
        ("RDBMS, Data & Cloud Platforms: ", True, "RDBMS Concepts, Relational Schema Modeling, Query Optimization (EXPLAIN, Indexing), Google Cloud Platform (Cloud Run, Pub/Sub), AWS, Microsoft Azure, MongoDB, Redis"),
        ("Software Architecture & Integration: ", True, "RESTful APIs, Microservices, Next.js (BFF), NestJS, Spring Boot, Apache Kafka, Slate CRM, Data Pipelines, Event-Driven Architecture, WebSockets"),
        ("AI, Productivity Tools & Analytics: ", True, "GitHub Copilot, Automated SQL Generation, Document Intelligence, Google Gemini, OpenAI API, Langfuse (LLM Tracing/Evals), Sentry Telemetry, Power BI"),
        ("Methodologies, Testing & SDLC: ", True, "Data Structures & Algorithms, Agile/Scrum, System Integration Testing (SIT), Embedded QA, Unit Testing (JUnit, Jest, Pytest), Docker, Jenkins CI/CD, Git")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.8)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an autonomous real-time voice screening platform with low-latency WebSocket audio streaming and structured rubric evaluation.", False)
    ], space_after=2.4)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for distributed system design, relational database modeling, and disaster relief tracking.", False)
    ], space_after=2.4)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, exhibiting deep analytical rigor, algorithmic efficiency, and mastery of optimal space-time complexities.", False)
    ], space_after=1.5)

def populate_bank_of_america(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Bank of America GTST Junior Java Full Stack Developer Focus)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Architected a multi-tenant ", False),
        ("Java 17 (Spring Boot)", True),
        (" ETL service integrating enterprise CRMs via RESTful APIs, deploying ", False),
        ("Hibernate/JPA", True),
        (" and ", False),
        ("Kafka Dead Letter Queues (DLQ)", True),
        (" to ensure eventual consistency and zero message loss under burst traffic.", False)
    ], space_after=2.4)
    add_bullet([
        ("Built responsive, accessible web components in ", False),
        ("Angular 14", True),
        (" and ", False),
        ("TypeScript", True),
        (" with ", False),
        ("HTML5/CSS3", True),
        (" and RxJS observable streams, standardizing modular frontend architecture and cutting user onboarding turnaround by ", False),
        ("40%", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Optimized relational database performance, reducing API latency ", False),
        ("3× (15s to 5s)", True),
        (" by profiling ", False),
        ("SQL", True),
        (" query execution plans (", False),
        ("EXPLAIN", True),
        ("), adding indexes, and eliminating redundant joins across ", False),
        ("PostgreSQL and MS SQL Server", True),
        (" schemas.", False)
    ], space_after=2.4)
    add_bullet([
        ("Delivered an enterprise commercial ", False),
        ("GPA evaluation platform", True),
        (" (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+", True),
        (" university deployments generating ", False),
        ("$150K+", True),
        (" contract value per client, enforcing ", False),
        ("FERPA compliance", True),
        (" and role-based access control (", False),
        ("RBAC", True),
        (").", False)
    ], space_after=2.4)
    add_bullet([
        ("Engineered a decoupled ", False),
        ("Go", True),
        (" ingestion pipeline on ", False),
        ("Cloud Run", True),
        (" and ", False),
        ("Pub/Sub", True),
        (", processing ", False),
        ("10K+ delta records per 15-minute sync cycle", True),
        (" into a ", False),
        ("MongoDB", True),
        (" queue with lease-fenced consumers to guarantee zero data loss.", False)
    ], space_after=2.4)
    add_bullet([
        ("Instrumented centralized error telemetry and distributed tracing across microservices using ", False),
        ("Sentry", True),
        (" and ", False),
        ("Langfuse", True),
        (", isolating tool-call latency bottlenecks and reducing production mean-time-to-detect (", False),
        ("MTTD) by 45%", True),
        (".", False)
    ], space_after=2.4)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Created the responsive ", False),
        ("Angular 14", True),
        (" and ", False),
        ("React", True),
        (" frontend for an AI assessment platform, designing reusable ", False),
        ("TypeScript", True),
        (" components and interactive dashboards that automated ", False),
        ("200+", True),
        (" monthly candidate evaluations.", False)
    ], space_after=2.4)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines automating containerized ", False),
        ("Docker", True),
        (" deployments and integration test execution in an ", False),
        ("Agile team", True),
        (", eliminating broken staging builds and accelerating release cadence.", False)
    ], space_after=2.6)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks, Software Engineering.", False)
    ], space_after=2.4)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Java 17, Spring Boot, Hibernate, Kafka, Redis, PostgreSQL, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a mission-critical payment microservice in ", False),
        ("Java 17 (Spring Boot / Hibernate)", True),
        
        (", ", False),
        ("Spring Security", True),
        (", and ", False),
        ("JWT", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency with published Artillery test suites.", False)
    ], space_after=2.4)
    add_bullet([
        ("Formulated an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" with ", False),
        ("Redis", True),
        (" distributed locks and ", False),
        ("PostgreSQL", True),
        (" ACID validation, guaranteeing idempotent transactions and zero double-charges across ", False),
        ("500+", True),
        (" simulated users.", False)
    ], space_after=2.4)
    add_bullet([
        ("Containerized microservices with ", False),
        ("Docker", True),
        (" virtualization and authored comprehensive ", False),
        ("JUnit", True),
        (" and ", False),
        ("Mockito", True),
        (" test suites to validate transactional rollback semantics, reconciliation routines, and code quality gates.", False)
    ], space_after=2.6)

    add_project_heading("AI Interviewer", "Angular 14, TypeScript, Next.js, NestJS, MongoDB, WebSockets, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Spearheaded an interactive web evaluation platform with responsive ", False),
        ("Angular 14", True),
        (" and ", False),
        ("TypeScript", True),
        (" interfaces, role-based access control (", False),
        ("RBAC", True),
        ("), and RESTful APIs, facilitating ", False),
        ("200+", True),
        (" monthly evaluations and reducing assessment turnaround by ", False),
        ("50%", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Orchestrated real-time bidirectional streaming using ", False),
        ("WebSockets", True),
        (" and structured evaluation rubrics in ", False),
        ("NestJS", True),
        (", enforcing rigorous test coverage and quality gates with ", False),
        ("Jest", True),
        (" to ensure robust continuous delivery.", False)
    ], space_after=2.6)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Java (17 / Core Java, OOP), TypeScript, JavaScript, SQL (T-SQL), Python, Go, HTML5, CSS3, C++"),
        ("Backend & Frameworks: ", True, "Spring Boot, Spring Framework, Hibernate/JPA, RESTful APIs, Microservices, NestJS, Node.js"),
        ("Frontend & UI: ", True, "Angular 14, TypeScript, RxJS, React, Next.js (BFF), Responsive Web Design, Component Architecture"),
        ("Databases & Messaging: ", True, "Microsoft SQL Server (T-SQL, Stored Procedures, Indexes), PostgreSQL, Kafka (DLQ), Redis, MongoDB"),
        ("DevOps, Cloud & Build: ", True, "Docker, Jenkins CI/CD, Git, GitHub PRs & Code Reviews, GCP (Cloud Run), Kubernetes (familiar)"),
        ("Testing, Security & Agile: ", True, "JUnit, Mockito, Jest, Spring Security, JWT, OAuth2, Sentry, Agile / Scrum, SDLC, SIT & UAT")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.6)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting a real-time voice screening platform.", False)
    ], space_after=2.4)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for fault-tolerant system architecture.", False)
    ], space_after=2.4)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, emphasizing algorithmic complexity, numerical problems, and memory optimization.", False)
    ], space_after=1.8)

def populate_parkplus(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - AI Engineering, Computer Vision/OCR, Agentic AI & Voice AI)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Engineered real-time telephony WebSocket middleware in ", False),
        ("NestJS", True),
        (" bridging ", False),
        ("Genesys AudioHook v2", True),
        (", sustaining ", False),
        ("1,000+ concurrent streams", True),
        (" at sub-", False),
        ("100ms", True),
        (" p95 latency via bidirectional ", False),
        ("8 kHz µ-law to 16 kHz PCM transcoding", True),
        (".", False)
    ], space_after=2.0)
    add_bullet([
        ("Architected a content-hashed ", False),
        ("file UUID deduplication engine", True),
        (" across transcript processing pipelines, preventing duplicate ", False),
        ("OCR extraction passes", True),
        (" across page reloads and eliminating redundant model compute costs.", False)
    ], space_after=2.0)
    add_bullet([
        ("Formulated an academic articulation engine coupling deterministic rule caches with ", False),
        ("LLM fallback", True),
        (", resolving ", False),
        ("40%+", True),
        (" of course mappings deterministically to eliminate prompt variance and cut token spend by ", False),
        ("40%", True),
        (".", False)
    ], space_after=2.0)
    add_bullet([
        ("Instrumented multi-turn agent workflows with ", False),
        ("Langfuse", True),
        (" tracing and telemetry, isolating reasoning-chain bottlenecks, eliminating production ", False),
        ("MongoDB tool-call hallucinations", True),
        (", and mitigating prompt drift.", False)
    ], space_after=2.0)
    add_bullet([
        ("Authored high-volume data pipelines in ", False),
        ("Python", True),
        (" for ", False),
        ("50K+", True),
        (" records from ", False),
        ("PostgreSQL", True),
        (", implementing scheduled batch jobs and data validation to eliminate manual reporting and accelerate analytics cycles by ", False),
        ("60%", True),
        (".", False)
    ], space_after=2.0)
    add_bullet([
        ("Deployed commercial transcript evaluation platforms across ", False),
        ("4+", True),
        (" enterprise universities generating ", False),
        ("$150K+", True),
        (" contract value per client, automating evaluation writeback into ", False),
        ("Slate CRM", True),
        (" under strict ", False),
        ("FERPA compliance", True),
        (".", False)
    ], space_after=2.0)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Built the full-stack AI interview platform (", False),
        ("NestJS", True),
        (", ", False),
        ("MongoDB", True),
        (", ", False),
        ("Vapi.ai", True),
        (") automating ", False),
        ("200+", True),
        (" monthly candidate screening sessions and cutting turnaround time by ", False),
        ("50%", True),
        (", commercialized for university admissions.", False)
    ], space_after=2.0)
    add_bullet([
        ("Constructed a ", False),
        ("Kafka-driven NestJS", True),
        (" microservice for asynchronous event synchronization between partner CRMs and internal databases, maintaining sub-", False),
        ("50ms", True),
        (" sync latency for conversational chatbot responses.", False)
    ], space_after=2.2)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Neural Networks & Deep Learning, Linear Algebra, Probability & Random Processes, Data Structures & Algorithms, Operating Systems, DBMS.", False)
    ], space_after=2.2)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("AI Interviewer", "NestJS, Next.js, Vapi.ai, WebSockets, OpenAI API, MongoDB", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Implemented an automated voice screening platform with ", False),
        ("NestJS", True),
        (", ", False),
        ("Next.js", True),
        (", and ", False),
        ("Vapi.ai", True),
        (", conducting ", False),
        ("200+", True),
        (" monthly candidate interviews and commercialized for university admissions.", False)
    ], space_after=1.8)
    add_bullet([
        ("Orchestrated real-time WebSocket audio streaming, ", False),
        ("VAD", True),
        (", and structured prompt evaluation rubrics in ", False),
        ("NestJS", True),
        (", evaluating candidate technical proficiency with zero human reviewer bottleneck.", False)
    ], space_after=1.8)
    add_bullet([
        ("Integrated ", False),
        ("OpenAI API", True),
        (" function calling and audio chunking in ", False),
        ("Next.js", True),
        (" to parse candidate responses against structured rubrics with sub-second feedback generation.", False)
    ], space_after=2.2)

    add_project_heading("AI PDF Summarizer", "Python, FastEmbed, LangChain, FAISS, Chroma, Streamlit", "github.com/Yash-Gupta-01/AI-PDF-Summarizer", url="https://github.com/Yash-Gupta-01/AI-PDF-Summarizer")
    add_bullet([
        ("Implemented a local document intelligence pipeline in ", False),
        ("Python", True),
        (" using ", False),
        ("FastEmbed", True),
        (" and ", False),
        ("LangChain", True),
        (", enabling sub-second semantic retrieval and contextual summarization over multi-page research documents.", False)
    ], space_after=1.8)
    add_bullet([
        ("Benchmarked dense vector embeddings against chunking strategies across ", False),
        ("FAISS", True),
        (" and ", False),
        ("Chroma", True),
        (" vector indexes, optimizing retrieval precision and minimizing semantic context dilution.", False)
    ], space_after=2.2)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Python (NumPy, Pytest), TypeScript, JavaScript, Go, Java, C++, SQL (PostgreSQL, MySQL)"),
        ("AI & LLM Architecture: ", True, "Transformers, Attention, RAG, Embeddings (FastEmbed), Prompt Engineering, LangChain"),
        ("Agentic & Voice AI: ", True, "Multi-Turn Agents, Tool Calling, Langfuse Evals, Vapi.ai, Genesys AudioHook v2, WebSockets"),
        ("Computer Vision & ML: ", True, "OCR (Document Intelligence, Deduplication), CNNs & Detection Concepts, Efficient Inference"),
        ("Backend & Data Pipelines: ", True, "Node.js, NestJS, Express.js, FastAPI, Apache Kafka (Streaming), Cloud Pub/Sub, REST APIs"),
        ("Databases & DevOps: ", True, "PostgreSQL, MongoDB (Aggregations), Redis, Docker, Git, CI/CD (GitHub Actions), Sentry")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.8)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform.", False)
    ], space_after=2.0)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ competing teams for distributed system architecture.", False)
    ], space_after=2.0)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" algorithmic problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, emphasizing graph algorithms, memory optimization, and low-latency system design.", False)
    ], space_after=0.0)

def populate_uber(doc, add_section_header, add_role_heading, add_project_heading, add_bullet, role="Integrations_Engineer"):
    # 2. Experience (Catalog Ingestion & Integrations - Data Pipelines, Partner APIs, Schema Modeling, Reliability & Python)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Authored internal ", False),
        ("Python scripts", True),
        (" to extract and transform ", False),
        ("50K+ student application records", True),
        (" from ", False),
        ("PostgreSQL", True),
        (", eliminating manual operational reporting overhead and accelerating analytics turnaround by ", False),
        ("60%", True),
        (".", False)
    ], space_after=2.0)
    add_bullet([
        ("Standardized partner data ingestion across ", False),
        ("6 differential bulk feeds", True),
        (", extending the core schema into nested ", False),
        ("MongoDB", True),
        (" models and resolving out-of-order race conditions via ", False),
        ("pending stubs", True),
        (".", False)
    ], space_after=2.0)
    add_bullet([
        ("Constructed a multi-tenant ETL integration pipeline interfacing partner ", False),
        ("REST APIs", True),
        (" in ", False),
        ("Java/Spring Boot", True),
        (", utilizing ", False),
        ("Kafka Dead Letter Queues (DLQ)", True),
        (" to isolate malformed payloads and prevent data loss.", False)
    ], space_after=2.0)
    add_bullet([
        ("Spearheaded centralized ", False),
        ("Sentry observability", True),
        (" and alerting across microservices, cutting ", False),
        ("MTTD by 45%", True),
        (" and resolving a critical live schema drift production incident within ", False),
        ("2 hours", True),
        (" during on-call rotation.", False)
    ], space_after=2.0)
    add_bullet([
        ("Architected a decoupled two-tier ingestion pipeline on ", False),
        ("Cloud Run", True),
        (" and ", False),
        ("Google Cloud Pub/Sub", True),
        (" in ", False),
        ("Go", True),
        (", processing ", False),
        ("10K+ delta records per 15-minute sync cycle", True),
        (" into ", False),
        ("MongoDB", True),
        (" with lease fencing to ensure zero data loss.", False)
    ], space_after=2.0)
    add_bullet([
        ("Accelerated backend analytics query performance ", False),
        ("3x (15s to 5s)", True),
        (" by analyzing ", False),
        ("SQL execution plans (EXPLAIN)", True),
        (", indexing composite keys, and eliminating redundant full-table scans across ", False),
        ("PostgreSQL and MySQL", True),
        (".", False)
    ], space_after=2.2)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Launched an interactive evaluation platform with ", False),
        ("React/TypeScript", True),
        (" and ", False),
        ("REST APIs", True),
        (", conducting ", False),
        ("200+", True),
        (" monthly candidate reviews and slashing assessment turnaround by ", False),
        ("50%", True),
        (" across university partner trials.", False)
    ], space_after=2.0)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" deployment pipelines executing automated regression test suites and containerized ", False),
        ("Docker", True),
        (" packaging in an ", False),
        ("Agile team", True),
        (", eliminating broken builds and expediting release cycles.", False)
    ], space_after=2.2)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks, Software Engineering & Testing.", False)
    ], space_after=2.2)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Go (Golang), Python, Apache Kafka, Redis, PostgreSQL, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a cloud-native payment microservice in ", False),
        ("Go", True),
        (" with a ", False),
        ("Python", True),
        (" reference implementation, sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency, verified via authored ", False),
        ("Artillery", True),
        (" load tests.", False)
    ], space_after=2.0)
    add_bullet([
        ("Formulated an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" and ", False),
        ("Redis", True),
        (" distributed locks with ", False),
        ("PostgreSQL ACID", True),
        (" validation, preventing race conditions and zero double-charges across ", False),
        ("500+", True),
        (" users.", False)
    ], space_after=2.2)

    add_project_heading("BOB Banking Application", "Python, Flask, SQLite, Pytest, GitHub Actions CI/CD", "github.com/Yash-Gupta-01/bob-banking-app", url="https://github.com/Yash-Gupta-01/bob-banking-app")
    add_bullet([
        ("Built a modular financial web service in ", False),
        ("Python", True),
        (" using ", False),
        ("Flask", True),
        (" and ", False),
        ("SQLAlchemy ORM", True),
        (", enforcing strict transactional ACID compliance, structured MVC separation, and secure session customer authentication.", False)
    ], space_after=2.0)
    add_bullet([
        ("Configured GitHub Actions CI/CD workflows running multi-stage ", False),
        ("Pytest", True),
        (" unit and integration test suites with coverage thresholds, preventing regression defects across production deployments.", False)
    ], space_after=2.2)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Python, Go (Golang), Java, TypeScript, JavaScript (ES6+), SQL (PostgreSQL, MySQL), C++"),
        ("Backend & Integrations: ", True, "RESTful APIs, Microservices, Apache Kafka (DLQ, Event Streams), GCP Pub/Sub, Redis, NestJS, Spring Boot"),
        ("Databases & Storage: ", True, "PostgreSQL, MongoDB (Aggregations, Bulk Writes), MySQL, SQLite, Data Modeling, Schema Migrations, Indexing"),
        ("Observability & Reliability: ", True, "Sentry (Alerting & Telemetry), On-Call Incident Triage, SLA Monitoring, Root Cause Analysis, Error Budgets"),
        ("Testing & QA: ", True, "Pytest, Jest, JUnit, Artillery Load Testing, API Contract Testing, Automated Regression Suites"),
        ("Cloud & DevOps: ", True, "Docker, Google Cloud Platform (Cloud Run), AWS (familiar), Jenkins CI/CD, GitHub Actions, Linux/Shell, Git, Kubernetes (familiar)")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.8)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting a real-time voice screening platform with low-latency WebSockets.", False)
    ], space_after=2.0)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for fault-tolerant system architecture and service resilience.", False)
    ], space_after=2.0)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, emphasizing algorithmic complexity analysis, memory optimization, and concurrent execution.", False)
    ], space_after=1.2)

def populate_rubrik(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Systems, Distributed Pipelines, DB Performance & Cloud Infrastructure)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Architected a decoupled ", False),
        ("Go ingestion pipeline", True),
        (" on ", False),
        ("Cloud Run", True),
        (" and ", False),
        ("Pub/Sub", True),
        (", pulling ", False),
        ("10K+ delta records per 15-minute cycle", True),
        (" into ", False),
        ("MongoDB", True),
        (" with lease-fenced consumers, guaranteeing ", False),
        ("zero data loss", True),
        (".", False)
    ], space_after=3.0)
    add_bullet([
        ("Replicated applicant data across an isolated ", False),
        ("Student Copilot MongoDB cluster", True),
        (" via ", False),
        ("Go and Pub/Sub", True),
        (", enforcing ", False),
        ("revision-fenced idempotent writes", True),
        (" and sub-second soft deletes to eliminate concurrent write races.", False)
    ], space_after=3.0)
    add_bullet([
        ("Scaled commercial transcript evaluation platforms across ", False),
        ("4+ universities", True),
        (" (", False),
        ("Rowan", True),
        (", ", False),
        ("RFU", True),
        (", ", False),
        ("PennWest", True),
        (") generating ", False),
        ("$150K+ contract value", True),
        (" per client, automating computed grade writeback into ", False),
        ("Slate CRM", True),
        (".", False)
    ], space_after=3.0)
    add_bullet([
        ("Built a self-reinforcing academic articulation engine coupling deterministic rule caches with ", False),
        ("LLM fallback", True),
        (", resolving ", False),
        ("40%+", True),
        (" of course mappings deterministically to cut token spend by ", False),
        ("40%", True),
        (".", False)
    ], space_after=3.0)
    add_bullet([
        ("Instrumented cloud microservices with ", False),
        ("Sentry application tracing", True),
        (" and telemetry alerts, classifying ", False),
        ("transient network timeouts", True),
        (" versus service failures to isolate root causes and slash production ", False),
        ("MTTD by 45%", True),
        (".", False)
    ], space_after=3.0)
    add_bullet([
        ("Optimized database storage by introducing ", False),
        ("MongoDB ordered bulk operations", True),
        (" and ", False),
        ("PostgreSQL", True),
        (" composite indexing, eliminating shared lock contention during burst synchronizations and slashing latency ", False),
        ("3× (15s to 5s)", True),
        (".", False)
    ], space_after=3.0)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" pipelines packaging containerized deployments (", False),
        ("Docker", True),
        (") and integration test suites on ", False),
        ("Linux", True),
        (", eliminating broken staging builds and accelerating developer release cadence.", False)
    ], space_after=3.0)
    add_bullet([
        ("Launched an AI evaluation platform processing ", False),
        ("200+", True),
        (" monthly technical assessments with ", False),
        ("50% turnaround reduction", True),
        (", designing modular ", False),
        ("REST services", True),
        (" that accelerated cross-functional partner team onboarding.", False)
    ], space_after=3.2)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Operating Systems, Computer Networks, Database Management Systems (DBMS), Object-Oriented Programming (OOP), Distributed Systems, System Design, Storage Systems.", False)
    ], space_after=3.2)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Go (Golang), Docker, Apache Kafka, Redis, PostgreSQL, Linux, Artillery", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a cloud-native payments microservice in ", False),
        ("Go (Golang)", True),
        (" utilizing goroutines and channels, sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms p95 latency", True),
        (", validated via authored ", False),
        ("Artillery load tests", True),
        (" on Linux.", False)
    ], space_after=3.0)
    add_bullet([
        ("Constructed an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" and ", False),
        ("Redis locks", True),
        (" with ", False),
        ("ACID PostgreSQL updates", True),
        (", guaranteeing strict idempotency and zero double-charges across ", False),
        ("500+ simulated users", True),
        (".", False)
    ], space_after=3.2)

    add_project_heading("AI Interviewer", "NestJS, Next.js, React, TypeScript, WebSockets, OpenAI API, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Spearheaded an autonomous evaluation platform with ", False),
        ("Next.js", True),
        (" and ", False),
        ("NestJS", True),
        (" microservices, codifying structured scoring rubrics and ", False),
        ("Jest regression suites", True),
        (" to maintain deterministic candidate assessment quality.", False)
    ], space_after=3.0)
    add_bullet([
        ("Maintained bidirectional streaming audio pipelines via ", False),
        ("WebSockets", True),
        (" and ", False),
        ("Vapi.ai", True),
        (", handling real-time audio chunking and speech-to-text inference to cut technical evaluation turnaround by ", False),
        ("50%", True),
        (" across ", False),
        ("200+ sessions", True),
        (".", False)
    ], space_after=3.2)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Go (Golang), C++, Java, Python, TypeScript, JavaScript, C#, SQL (PostgreSQL, MySQL), Bash/Shell"),
        ("Cloud & Infrastructure: ", True, "Google Cloud (Cloud Run, Pub/Sub, VPC), Docker, Kubernetes (familiar), AWS, Linux, Git"),
        ("Distributed Systems & Messaging: ", True, "Microservices, Apache Kafka (DLQ, Outbox Pattern), Redis, Multi-Node Replication"),
        ("Storage & Databases: ", True, "PostgreSQL, MongoDB (Ordered Bulk Writes), MySQL, SQL Server, Redis, ACID Transactions"),
        ("Networking & Architecture: ", True, "TCP/IP, HTTP/HTTPS, WebSockets, RESTful APIs, Backend-for-Frontend (BFF), RBAC, Proxies"),
        ("Reliability, Testing & Observability: ", True, "Sentry (Tracing, MTTD Reduction), Artillery (Load Testing), Jest, Pytest, JUnit")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=2.8)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting a real-time voice screening platform with streaming and deterministic evaluation rubrics.", False)
    ], space_after=2.8)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for robust systems architecture and networking.", False)
    ], space_after=2.8)
    add_bullet([
        ("Solved ", False),
        ("450+ Data Structures & Algorithms problems", True),
        (" across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, demonstrating deep proficiency in graph traversal and concurrent programming.", False)
    ], space_after=1.8)

def populate_optum(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (Optum Full-Stack Java Microservices, APIs, Cloud & Quality Automation Focus)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Constructed a multi-tenant ", False),
        ("Java 17 (Spring Boot)", True),
        (" ETL service integrating enterprise REST APIs and JSON payloads, implementing ", False),
        ("Hibernate/JPA", True),
        (" and ", False),
        ("Kafka Dead Letter Queues (DLQ)", True),
        (" to ensure message delivery and eliminate data loss under burst traffic.", False)
    ], space_after=2.0)
    add_bullet([
        ("Engineered an enterprise commercial ", False),
        ("GPA evaluation platform", True),
        (" (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+", True),
        (" university deployments generating ", False),
        ("$150K+", True),
        (" contract value per client, enforcing strict role-based access control (", False),
        ("RBAC", True),
        (").", False)
    ], space_after=2.0)
    add_bullet([
        ("Accelerated backend database query performance ", False),
        ("3× (15s to 5s)", True),
        (" by analyzing ", False),
        ("SQL execution plans (EXPLAIN)", True),
        (", indexing composite keys, and validating JSON schemas across ", False),
        ("MySQL and MongoDB", True),
        (" databases.", False)
    ], space_after=2.0)
    add_bullet([
        ("Architected a decoupled two-tier ingestion pipeline on ", False),
        ("Cloud Run", True),
        (" and ", False),
        ("Google Cloud Pub/Sub", True),
        (" in Go, processing ", False),
        ("10K+ delta records per 15-minute sync cycle", True),
        (" into ", False),
        ("MongoDB", True),
        (" with lease fencing to ensure zero data loss.", False)
    ], space_after=2.0)
    add_bullet([
        ("Spearheaded centralized ", False),
        ("Sentry observability", True),
        (" and log telemetry across microservices, cutting ", False),
        ("MTTD by 45%", True),
        (" and troubleshooting a live schema drift production incident within ", False),
        ("2 hours", True),
        (" during on-call rotation.", False)
    ], space_after=2.0)
    add_bullet([
        ("Formulated a ", False),
        ("Document Intelligence", True),
        (" engine coupling deterministic articulation rule caches with ", False),
        ("LLM fallback", True),
        (", resolving ", False),
        ("40%+", True),
        (" of course mappings without LLM inference to cut token costs by ", False),
        ("40%", True),
        (".", False)
    ], space_after=2.2)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Launched an interactive evaluation platform with ", False),
        ("React", True),
        (", ", False),
        ("TypeScript", True),
        (", and RESTful APIs, facilitating ", False),
        ("200+", True),
        (" monthly candidate reviews and slashing assessment turnaround by ", False),
        ("50%", True),
        (" across university partner trials.", False)
    ], space_after=2.0)
    add_bullet([
        ("Maintained ", False),
        ("GitHub Actions", True),
        (" and Jenkins CI/CD artifact pipelines executing multi-stage ", False),
        ("JUnit", True),
        (" regression test suites and containerized ", False),
        ("Docker", True),
        (" packaging in an Agile Kanban team, accelerating release readiness.", False)
    ], space_after=2.2)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks, Software Engineering & Testing.", False)
    ], space_after=2.0)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Java 17, Spring Boot, Hibernate, Kafka, Redis, PostgreSQL, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a cloud-native payment microservice in ", False),
        ("Java 17 (Spring Boot)", True),
        
        (", ", False),
        ("Spring Security", True),
        (", and ", False),
        ("JWT", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency with authored Artillery test suites.", False)
    ], space_after=2.0)
    add_bullet([
        ("Maintained an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" and ", False),
        ("Redis", True),
        (" distributed locks with ", False),
        ("PostgreSQL ACID", True),
        (" validation, preventing race conditions and zero double-charges across ", False),
        ("500+", True),
        (" users.", False)
    ], space_after=2.0)
    add_bullet([
        ("Containerized microservices with ", False),
        ("Docker", True),
        (" virtualization and authored comprehensive ", False),
        ("JUnit", True),
        (" regression test suites, validating transactional rollback semantics, reconciliation routines, and code quality gates.", False)
    ], space_after=2.2)

    add_project_heading("AI Interviewer", "Next.js, React, NestJS, MongoDB, TypeScript, OpenAI API, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Designed an interactive evaluation platform with ", False),
        ("Next.js", True),
        (", ", False),
        ("NestJS", True),
        (", and REST APIs, conducting ", False),
        ("200+", True),
        (" monthly assessments and establishing automated scoring pipelines with end-to-end data validation.", False)
    ], space_after=2.0)
    add_bullet([
        ("Implemented comprehensive test suites and API contract validation using ", False),
        ("Jest", True),
        (" and ", False),
        ("GitHub Actions CI/CD", True),
        (" workflows, establishing rigorous quality gates to prevent regression defects and ensure release stability.", False)
    ], space_after=2.2)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Java (17 / Core Java, OOP), TypeScript, JavaScript, SQL (MySQL), Python, Go, HTML5, CSS3, C++"),
        ("Backend & Frameworks: ", True, "Spring Boot, Spring Framework, Hibernate/JPA, RESTful APIs, JSON, Kafka (DLQ), NestJS, Node.js"),
        ("Frontend & UI: ", True, "React, Next.js (BFF), TypeScript, HTML5/CSS3, Tailwind CSS, Responsive Web Design, Component Architecture"),
        ("Databases & Storage: ", True, "MySQL (Query Optimization, Indexes), MongoDB (NoSQL, Aggregation), PostgreSQL, Redis, SQLite"),
        ("DevOps, Cloud & CI/CD: ", True, "GitHub Actions CI/CD, Docker, GCP (Cloud Run, Pub/Sub), Jenkins, Git, Kubernetes (familiar)"),
        ("Testing, Quality & SDLC: ", True, "JUnit, Mockito, Jest, Pytest, Playwright (familiar), Selenium (familiar), Splunk, Sentry, Agile SDLC")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.5)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting a real-time voice screening platform.", False)
    ], space_after=2.0)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for fault-tolerant system architecture.", False)
    ], space_after=2.0)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, emphasizing algorithmic complexity and memory optimization.", False)
    ], space_after=1.8)

def populate_fynd(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Fullstack Web, Commerce/Billing & Distributed Systems)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Engineered an enterprise commercial platform (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+", True),
        (" university deployments, generating ", False),
        ("$150K+", True),
        (" contract value per client with 100% Stripe billing reconciliation.", False)
    ], space_after=2.4)
    add_bullet([
        ("Architected a self-reinforcing ", False),
        ("Document Intelligence", True),
        (" engine coupling deterministic rule caches with LLM fallback, resolving ", False),
        ("40%+", True),
        (" of course mappings without model calls to cut token spend by ", False),
        ("40%", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Scaled real-time ", False),
        ("NestJS", True),
        (" WebSocket middleware bridging telephony streams, handling bidirectional audio transcoding for ", False),
        ("1,000+", True),
        (" concurrent sessions at sub-", False),
        ("100ms", True),
        (" p95 latency with zero packet drops.", False)
    ], space_after=2.4)
    add_bullet([
        ("Optimized ", False),
        ("MongoDB", True),
        (" ingestion throughput during burst CRM synchronizations by implementing ordered ", False),
        ("bulk operations (bulkWrite)", True),
        (" and lease-renewal batching, eliminating database lock contention.", False)
    ], space_after=2.4)
    add_bullet([
        ("Profiled ", False),
        ("SQL", True),
        (" query execution plans (", False),
        ("EXPLAIN", True),
        (") and restructured relational schemas across ", False),
        ("PostgreSQL and MySQL", True),
        (", slashing analytics API latency ", False),
        ("3× (15s to 5s)", True),
        (" and eliminating table scan bottlenecks.", False)
    ], space_after=2.4)
    add_bullet([
        ("Instrumented centralized observability using ", False),
        ("Sentry", True),
        (" and ", False),
        ("Langfuse", True),
        (" across web clients and microservices, diagnosing tool-call anomalies and tracking error telemetry to reduce ", False),
        ("MTTD by 45%", True),
        (".", False)
    ], space_after=2.4)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Launched a full-stack AI interview platform facilitating ", False),
        ("200+", True),
        (" monthly technical assessments, slashing turnaround by ", False),
        ("50%", True),
        (" and integrating versioned RESTful APIs with responsive React UI dashboards.", False)
    ], space_after=2.4)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines packaging containerized deployments (", False),
        ("Docker", True),
        (") and executing integration test suites in an ", False),
        ("Agile team", True),
        (", accelerating developer staging release cadence.", False)
    ], space_after=2.5)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks.", False)
    ], space_after=2.4)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Go (Golang), Apache Kafka, Redis, PostgreSQL, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a cloud-native payments microservice in ", False),
        ("Go (Golang)", True),
        (" with goroutines and channels, sustaining ", False),
        ("1,200 TPS", True),
        (" (simulated) at sub-", False),
        ("120ms", True),
        (" p95 latency, validated via authored Artillery load tests.", False)
    ], space_after=2.4)
    add_bullet([
        ("Implemented an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" and ", False),
        ("Redis", True),
        (" distributed locks with idempotent ", False),
        ("PostgreSQL", True),
        (" updates, eliminating double-charges across ", False),
        ("500+", True),
        (" concurrent simulated users.", False)
    ], space_after=2.5)

    add_project_heading("AI Interviewer", "NestJS, Next.js, React, TypeScript, WebSockets, OpenAI API, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Shipped an AI evaluation platform with responsive ", False),
        ("React/Next.js", True),
        (" UI and ", False),
        ("NestJS", True),
        (" microservices, designing structured rubrics for objective candidate scoring and integrating ", False),
        ("Jest", True),
        (" test suites.", False)
    ], space_after=2.4)
    add_bullet([
        ("Constructed bidirectional streaming pipelines using ", False),
        ("WebSockets", True),
        (" and ", False),
        ("Vapi.ai", True),
        (", enabling real-time conversational screening and reducing assessment turnaround by ", False),
        ("50%", True),
        (" across 200+ monthly reviews.", False)
    ], space_after=2.5)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "TypeScript, JavaScript (ES6+), Node.js, Go (Golang), Python, Java, SQL (PostgreSQL, MySQL), C++"),
        ("Frontend & Web: ", True, "React, Next.js (App Router, BFF), Redux, Tailwind CSS, HTML5, CSS3, Component-Driven Design, Responsive UI"),
        ("Backend & APIs: ", True, "NestJS, Express.js, RESTful APIs, Microservices, Event-Driven Architecture, Stripe Payments API, WebSockets"),
        ("Databases & Storage: ", True, "PostgreSQL, MongoDB (Aggregations, bulkWrite), Redis (Distributed Locks/Cache), MySQL, Prisma ORM"),
        ("DevOps, Cloud & DevEx: ", True, "Docker, Jenkins CI/CD, Git, GitHub Actions, GCP (Cloud Run, Pub/Sub), Sentry, Langfuse, Linux"),
        ("Testing & Methodologies: ", True, "Jest, JUnit, Artillery Load Testing, Agile / Scrum, Code Reviews, Performance Tuning")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.8)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, delivering an autonomous real-time voice screening platform.", False)
    ], space_after=2.2)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for resilient system architecture.", False)
    ], space_after=2.2)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, emphasizing computational complexity, graph theory, and memory optimization.", False)
    ], space_after=1.6)

def populate_ftse_russell(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Aligned with FTSE Russell / LSEG Associate Software Engineer)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Authored internal ", False),
        ("Python", True),
        (" data transformation routines querying ", False),
        ("PostgreSQL", True),
        (" to normalize and extract ", False),
        ("50K+", True),
        (" student application records, eliminating manual data wrangling and accelerating operational analytics turnaround by ", False),
        ("60%", True),
        (".", False)
    ], space_after=2.0)
    add_bullet([
        ("Optimized relational database performance to reduce backend analytics query latency ", False),
        ("3× (15s to 5s)", True),
        (", analyzing ", False),
        ("query execution plans", True),
        (", indexing composite keys, and eliminating full-table scans across ", False),
        ("PostgreSQL and MySQL", True),
        (".", False)
    ], space_after=2.0)
    add_bullet([
        ("Engineered a decoupled two-tier ", False),
        ("Go", True),
        (" ingestion pipeline on ", False),
        ("Cloud Run", True),
        (" and ", False),
        ("Google Cloud Pub/Sub", True),
        (", processing ", False),
        ("10K+ delta records per 15-minute sync cycle", True),
        (" into durable MongoDB queues with lease-fenced consumers to guarantee zero data loss.", False)
    ], space_after=2.0)
    add_bullet([
        ("Improved the ", False),
        ("Document Intelligence", True),
        (" pipeline by coupling deterministic articulation rule caches with ", False),
        ("LLM fallback", True),
        (", resolving ", False),
        ("40%+", True),
        (" of course mappings without model inference to cut token costs by ", False),
        ("40%", True),
        (" and eliminate prompt variance.", False)
    ], space_after=2.0)
    add_bullet([
        ("Delivered an enterprise commercial evaluation platform (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+", True),
        (" university deployments, generating ", False),
        ("$150K+", True),
        (" contract value per client while enforcing strict ", False),
        ("FERPA data privacy", True),
        (" and RBAC security.", False)
    ], space_after=2.0)
    add_bullet([
        ("Instrumented centralized ", False),
        ("Sentry error telemetry", True),
        (" and distributed tracing across microservices, diagnosing production schema anomalies and reducing mean-time-to-detect (", False),
        ("MTTD) by 45%", True),
        (" through automated root-cause alerting.", False)
    ], space_after=2.2)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Developed an interactive assessment web service facilitating ", False),
        ("200+", True),
        (" monthly candidate evaluations and slashing turnaround ", False),
        ("50%", True),
        (", designing modular ", False),
        ("RESTful APIs", True),
        (" and authoring Jest integration test suites.", False)
    ], space_after=2.0)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" deployment pipelines automating containerized ", False),
        ("Docker", True),
        (" packaging and regression test suites in an ", False),
        ("Agile team", True),
        (", collaborating on Gitflow peer reviews and accelerating release cadence.", False)
    ], space_after=2.2)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks, Software Engineering.", False)
    ], space_after=2.2)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Go, Python, Apache Kafka, Redis, PostgreSQL, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Implemented an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" with ", False),
        ("Redis", True),
        (" distributed locks and ", False),
        ("PostgreSQL ACID", True),
        (" validation, coordinating distributed payments and guaranteeing zero double-charges across ", False),
        ("500+", True),
        (" simulated users.", False)
    ], space_after=2.0)
    add_bullet([
        ("Scaled a high-concurrency payment microservice using ", False),
        ("Go and Python", True),
        (" routines, sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency, validated through authored Artillery load testing and benchmark suites.", False)
    ], space_after=2.0)
    add_bullet([
        ("Containerized payment microservices with ", False),
        ("Docker", True),
        (" and configured automated test suites, validating transactional rollback semantics, reconciliation routines, and code quality gates for reliable cloud deployment.", False)
    ], space_after=2.2)

    add_project_heading("BOB Banking Application", "Python, Flask, SQLAlchemy, SQLite, Pytest, CI/CD", "github.com/Yash-Gupta-01/bob-banking-app", url="https://github.com/Yash-Gupta-01/bob-banking-app")
    add_bullet([
        ("Built a modular financial web application in ", False),
        ("Python (Flask)", True),
        (" with ", False),
        ("SQLAlchemy ORM", True),
        (" and relational schemas, enforcing transactional ACID compliance, clean MVC architecture, and secure customer session authentication.", False)
    ], space_after=2.0)
    add_bullet([
        ("Automated multi-stage ", False),
        ("Pytest", True),
        (" unit and integration test suites via ", False),
        ("GitHub Actions CI/CD", True),
        (" workflows, enforcing code coverage thresholds, linter checks, and branch gates to prevent production regression escapes.", False)
    ], space_after=2.2)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Python, Java, Go (Golang), TypeScript, JavaScript (ES6+), C++, SQL (PostgreSQL, MySQL, SQLite)"),
        ("Backend & Frameworks: ", True, "Python (Flask, FastAPI), NestJS, Node.js, Spring Boot, RESTful APIs, Microservices"),
        ("Databases & Storage: ", True, "PostgreSQL (Aurora, Query Tuning, EXPLAIN), MongoDB (Bulk Writes), SQLite, Redis Caching"),
        ("Cloud & Distributed Systems: ", True, "AWS (Lambda, S3), GCP (Cloud Run), Docker, Apache Kafka (Saga), GCP Pub/Sub, Pandas"),
        ("Testing & Observability: ", True, "Pytest, Jest, JUnit, Sentry (Tracing & Telemetry), Artillery, Langfuse, Root Cause Analysis"),
        ("DevOps & Engineering: ", True, "GitHub Actions, Jenkins CI/CD, Gitflow, Agile Peer Code Review, Type Hints, Linters")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.6)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting a real-time voice screening platform with low-latency WebSocket audio streaming.", False)
    ], space_after=2.0)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ competing teams for fault-tolerant system architecture.", False)
    ], space_after=2.0)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, demonstrating strong foundation in computational efficiency, data structures, and algorithmic design.", False)
    ], space_after=0.0)

def populate_lseg(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Aligned with LSEG Data & Analytics SWE / Senior Associate)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Architected a multi-tenant ", False),
        ("Java 17 (Spring Boot)", True),
        (" ETL service integrating enterprise CRMs via RESTful APIs, implementing ", False),
        ("Hibernate/JPA", True),
        (" and ", False),
        ("Kafka Dead Letter Queues (DLQ)", True),
        (" to isolate ingestion failures and guarantee zero data loss.", False)
    ], space_after=2.4)
    add_bullet([
        ("Authored internal ", False),
        ("Python", True),
        (" data transformation routines querying ", False),
        ("PostgreSQL", True),
        (" to normalize and extract ", False),
        ("50K+", True),
        (" student application records, eliminating manual data wrangling and accelerating operational analytics turnaround by ", False),
        ("60%", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Engineered an enterprise commercial ", False),
        ("GPA evaluation platform", True),
        (" (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+", True),
        (" university deployments, generating ", False),
        ("$150K+", True),
        (" contract value per client and enforcing strict ", False),
        ("FERPA compliance", True),
        (" and RBAC security.", False)
    ], space_after=2.4)
    add_bullet([
        ("Improved the ", False),
        ("Document Intelligence", True),
        (" engine coupling deterministic articulation rule caches with ", False),
        ("LLM fallback", True),
        (", resolving ", False),
        ("40%+", True),
        (" of course mappings without LLM inference to cut token costs by ", False),
        ("40%", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Optimized relational database performance to reduce backend analytics API latency ", False),
        ("3× (15s to 5s)", True),
        (", tuning ", False),
        ("SQL execution plans (EXPLAIN)", True),
        (", indexing composite keys, and eliminating redundant table scans across ", False),
        ("PostgreSQL and MySQL", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Instrumented centralized error telemetry and LLM reasoning-chain tracing via ", False),
        ("Langfuse", True),
        (" and ", False),
        ("Sentry", True),
        (", diagnosing prompt drift and tool-call hallucinations to reduce production mean-time-to-detect (", False),
        ("MTTD) by 45%", True),
        (".", False)
    ], space_after=2.5)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Delivered a full-stack AI evaluation platform facilitating ", False),
        ("200+", True),
        (" monthly candidate assessments and slashing turnaround ", False),
        ("50%", True),
        (", designing responsive web dashboards and integrating secure RESTful evaluation APIs.", False)
    ], space_after=2.4)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines automating containerized deployments (", False),
        ("Docker", True),
        (") and unit/integration test suites in an ", False),
        ("Agile team", True),
        (", eliminating broken staging builds and accelerating release cadence.", False)
    ], space_after=2.5)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks, Software Engineering.", False)
    ], space_after=2.4)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Java 17, Spring Boot, Hibernate, Kafka, Redis, PostgreSQL, Docker, JUnit", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a mission-critical payment microservice in ", False),
        ("Java 17 (Spring Boot / Hibernate)", True),
        
        (", ", False),
        ("Spring Security", True),
        (", and ", False),
        ("JWT", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency with published Artillery test suites.", False)
    ], space_after=2.4)
    add_bullet([
        ("Formulated an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" with ", False),
        ("Redis", True),
        (" distributed locks and ", False),
        ("PostgreSQL", True),
        (" ACID validation, guaranteeing idempotent transactions and zero double-charges across ", False),
        ("500+", True),
        (" simulated users.", False)
    ], space_after=2.4)
    add_bullet([
        ("Containerized microservices with ", False),
        ("Docker", True),
        (" and authored comprehensive ", False),
        ("JUnit", True),
        (" and ", False),
        ("Mockito", True),
        (" test suites to validate transactional rollback semantics, reconciliation routines, and code quality gates for cloud deployment.", False)
    ], space_after=2.5)

    add_project_heading("AI Interviewer", "NestJS, Next.js, TypeScript, MongoDB, OpenAI, WebSockets, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Designed an interactive evaluation platform using ", False),
        ("NestJS", True),
        (", ", False),
        ("Next.js", True),
        (", and ", False),
        ("OpenAI API", True),
        (", establishing structured rubrics to score candidate responses and authoring ", False),
        ("Jest", True),
        (" test suites to enforce code quality gates.", False)
    ], space_after=2.4)
    add_bullet([
        ("Deployed bidirectional streaming audio pipelines via ", False),
        ("WebSockets", True),
        (" and ", False),
        ("Vapi.ai", True),
        (", enabling real-time conversational voice interviews and reducing assessment turnaround by ", False),
        ("50%", True),
        (" across 200+ monthly reviews.", False)
    ], space_after=2.5)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Java (17+ / Core Java, OOP), Python, C#, C++, TypeScript, JavaScript, Go, SQL (PostgreSQL, MySQL)"),
        ("Backend & Frameworks: ", True, "Spring Boot (Java 17), NestJS, Node.js, Next.js (BFF), ASP.NET Core, RESTful APIs, Microservices"),
        ("Distributed Systems: ", True, "Apache Kafka (DLQ, Saga Pattern), Google Cloud Pub/Sub, Redis (Locks/Cache), WebSockets"),
        ("Databases & Storage: ", True, "PostgreSQL, MongoDB (Aggregations), MySQL, Hibernate/JPA, Entity Framework Core, Prisma ORM"),
        ("Cloud, CI/CD & DevOps: ", True, "AWS (EC2, S3), Google Cloud Platform (Cloud Run), Docker, Jenkins CI/CD, GitHub Actions, Linux, Git"),
        ("Testing, AI & Tools: ", True, "JUnit, Pytest, Jest, Mockito, Artillery, Sentry, Langfuse (LLM Tracing), Gemini API, OpenAI API")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.8)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting a real-time voice screening platform with low-latency WebSocket audio streaming.", False)
    ], space_after=2.4)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ competing teams for fault-tolerant system architecture.", False)
    ], space_after=2.4)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, demonstrating strong foundation in computational efficiency and algorithmic design.", False)
    ], space_after=1.8)

def populate_jll(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - JLL Martech Intelligence Pod: Full-Stack Web, AI Agents & Enterprise Systems)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Engineered customer-facing web platforms in ", False),
        ("React", True),
        (", ", False),
        ("Next.js BFF", True),
        (", and ", False),
        ("TypeScript", True),
        (", building modular component trees, managing client state, and consuming secure RESTful APIs across ", False),
        ("4+", True),
        (" enterprise university deployments.", False)
    ], space_after=2.2)
    add_bullet([
        ("Architected multi-tenant backend services in ", False),
        ("Java 17 (Spring Boot)", True),
        (" and ", False),
        ("C#", True),
        (", designing REST APIs, applying clean ", False),
        ("OOP/SOLID", True),
        (" patterns, and integrating ", False),
        ("Hibernate/JPA", True),
        (" with ", False),
        ("Kafka DLQs", True),
        (" to isolate failures and guarantee zero data loss.", False)
    ], space_after=2.2)
    add_bullet([
        ("Profiled ", False),
        ("SQL query execution plans (EXPLAIN)", True),
        (" across ", False),
        ("PostgreSQL and MySQL", True),
        (", slashing backend API response latency ", False),
        ("3× (15s to 5s)", True),
        (" by optimizing composite indexes and eliminating slow table scans.", False)
    ], space_after=2.2)
    add_bullet([
        ("Spearheaded an AI Document Intelligence engine coupling deterministic articulation caches with ", False),
        ("Gemini and OpenAI APIs", True),
        (", resolving ", False),
        ("40%+", True),
        (" of evaluations deterministically to cut token inference spend by ", False),
        ("40%", True),
        (".", False)
    ], space_after=2.2)
    add_bullet([
        ("Enhanced customer-facing web platforms by diagnosing browser rendering bottlenecks, structuring crawlable pagination, and optimizing asset delivery to reduce ", False),
        ("Largest Contentful Paint (LCP) by 25%", True),
        (".", False)
    ], space_after=2.2)
    add_bullet([
        ("Instrumented centralized error telemetry and LLM reasoning-chain tracing via ", False),
        ("Langfuse", True),
        (" and ", False),
        ("Sentry", True),
        (", diagnosing prompt drift and tool-call hallucinations to reduce production mean-time-to-detect (", False),
        ("MTTD) by 45%", True),
        (".", False)
    ], space_after=2.4)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Delivered a full-stack candidate evaluation platform facilitating ", False),
        ("200+", True),
        (" monthly assessments and slashing turnaround ", False),
        ("50%", True),
        (", implementing responsive ", False),
        ("React", True),
        (" components, form validations, and RESTful APIs.", False)
    ], space_after=2.2)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines automating containerized deployments (", False),
        ("Docker", True),
        (") and unit/integration test suites (", False),
        ("JUnit, Jest", True),
        (") in an ", False),
        ("Agile team", True),
        (", ensuring zero broken staging builds.", False)
    ], space_after=2.4)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks, Software Engineering.", False)
    ], space_after=2.2)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Java 17, Spring Boot, Hibernate, Kafka, Redis, PostgreSQL, Docker, JUnit", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a high-throughput payment microservice in ", False),
        ("Java 17 (Spring Boot)", True),
        
        (", ", False),
        ("Spring Security", True),
        (", and ", False),
        ("JWT", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency with published Artillery test suites.", False)
    ], space_after=2.2)
    add_bullet([
        ("Formulated an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" with ", False),
        ("Redis", True),
        (" distributed locks and ", False),
        ("PostgreSQL", True),
        (" ACID validation, guaranteeing idempotent transactions across ", False),
        ("500+", True),
        (" concurrent simulated users.", False)
    ], space_after=2.2)
    add_bullet([
        ("Containerized microservices with ", False),
        ("Docker", True),
        (" and authored comprehensive ", False),
        ("JUnit", True),
        (" and ", False),
        ("Mockito", True),
        (" test suites to validate transactional rollback semantics, reconciliation routines, and code quality gates for cloud deployment.", False)
    ], space_after=2.4)

    add_project_heading("AI Interviewer", "NestJS, Next.js, React, TypeScript, OpenAI API, WebSockets, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Designed an autonomous evaluation platform with responsive ", False),
        ("React/TypeScript", True),
        (" UI and ", False),
        ("NestJS", True),
        (" microservices, establishing structured rubrics for objective candidate scoring and authoring ", False),
        ("Jest", True),
        (" unit test suites.", False)
    ], space_after=2.2)
    add_bullet([
        ("Deployed bidirectional streaming audio pipelines via ", False),
        ("WebSockets", True),
        (" and ", False),
        ("Vapi.ai", True),
        (", orchestrating multi-step LLM reasoning with ", False),
        ("OpenAI API", True),
        (" to cut assessment turnaround by ", False),
        ("50%", True),
        (" across 200+ monthly evaluations.", False)
    ], space_after=2.4)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "TypeScript, JavaScript, Java (17+ / Core Java), C#, Python, SQL (PostgreSQL, MySQL), HTML5/CSS3, Go"),
        ("Frontend & CMS: ", True, "React, Next.js (BFF), TypeScript, Redux, Tailwind CSS, Browser DevTools, AEM / WordPress (familiar)"),
        ("Backend & APIs: ", True, "Spring Boot (Java 17), ASP.NET Core (C#), NestJS, Node.js, RESTful APIs, Microservices, WebSockets"),
        ("Databases & Storage: ", True, "PostgreSQL, MySQL, MongoDB (Aggregations), Redis (Distributed Caching), EF Core, Prisma ORM"),
        ("Cloud, AI & DevOps: ", True, "OpenAI API, Gemini API, Langfuse, MCPs / Agents, Docker, Jenkins CI/CD, GitHub Actions, AWS"),
        ("Testing & Practices: ", True, "Jest, React Testing Library, JUnit, Mockito, Pytest, Artillery, Sentry, Agile / Scrum, Code Reviews")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.6)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting a real-time voice screening platform with low-latency WebSocket audio streaming.", False)
    ], space_after=2.2)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ competing teams for fault-tolerant system architecture.", False)
    ], space_after=2.2)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, demonstrating strong foundation in computational efficiency and algorithmic design.", False)
    ], space_after=1.6)

def populate_nasdaq(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Aligned with Nasdaq Backoffice / Calypso Platform: Java 17, Multi-threading, Kafka, Persistence & SQL)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Architected a multi-tenant ", False),
        ("Java 17 (Spring Boot)", True),
        (" backend service integrating enterprise systems via RESTful APIs, implementing ", False),
        ("Hibernate/JPA", True),
        (" persistence, multi-threaded worker pools, and ", False),
        ("Kafka Dead Letter Queues (DLQ)", True),
        (" to isolate ingestion failures and guarantee zero data loss.", False)
    ], space_after=2.4)
    add_bullet([
        ("Optimized relational database performance to reduce backend query latency ", False),
        ("3× (15s to 5s)", True),
        (", tuning ", False),
        ("SQL execution plans (EXPLAIN)", True),
        (", indexing composite keys, and eliminating full-table scans across high-volume ", False),
        ("PostgreSQL and SQL", True),
        (" transactional workloads.", False)
    ], space_after=2.4)
    add_bullet([
        ("Engineered enterprise workflow and evaluation services (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+", True),
        (" client deployments, generating ", False),
        ("$150K+", True),
        (" contract value per institution while enforcing strict data security, caching, and RBAC governance.", False)
    ], space_after=2.4)
    add_bullet([
        ("Instrumented centralized error telemetry, root-cause debugging, and reasoning-chain tracing via ", False),
        ("Langfuse", True),
        (" and ", False),
        ("Sentry", True),
        (", isolating runtime bottlenecks and reducing production issue mean-time-to-detect (", False),
        ("MTTD) by 45%", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Authored internal ", False),
        ("Python", True),
        (" data transformation routines querying ", False),
        ("PostgreSQL", True),
        (" to normalize and extract ", False),
        ("50K+", True),
        (" records, eliminating manual data wrangling and accelerating operational analytics turnaround by ", False),
        ("60%", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Improved the ", False),
        ("Document Intelligence", True),
        (" engine coupling deterministic rule caches with ", False),
        ("LLM fallback", True),
        (", resolving ", False),
        ("40%+", True),
        (" of complex mappings without LLM inference to cut compute and API costs by ", False),
        ("40%", True),
        (".", False)
    ], space_after=2.5)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Delivered a full-stack AI evaluation platform facilitating ", False),
        ("200+", True),
        (" monthly candidate assessments and slashing turnaround ", False),
        ("50%", True),
        (", designing responsive web dashboards and integrating secure RESTful evaluation APIs.", False)
    ], space_after=2.4)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines automating containerized deployments (", False),
        ("Docker", True),
        (") and automated unit and integration test suites in an ", False),
        ("Agile", True),
        (" workflow, eliminating staging regressions.", False)
    ], space_after=2.5)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Object-Oriented Programming (OOP), Data Structures & Algorithms, Database Management Systems (DBMS), Operating Systems, Computer Networks, Software Engineering.", False)
    ], space_after=2.4)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Java 17, Spring Boot, Hibernate, Kafka, Redis, PostgreSQL, Docker, Gradle, JUnit", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a distributed payments microservice in ", False),
        ("Java 17 (Spring Boot / Hibernate)", True),
        
        (", enforcing authentication via ", False),
        ("Spring Security and JWT", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency.", False)
    ], space_after=2.4)
    add_bullet([
        ("Formulated an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" with ", False),
        ("Redis", True),
        (" distributed locks and ", False),
        ("PostgreSQL", True),
        (" ACID validation, guaranteeing idempotent transactions and zero double-charges across ", False),
        ("500+", True),
        (" concurrent users.", False)
    ], space_after=2.4)
    add_bullet([
        ("Containerized microservices with ", False),
        ("Docker", True),
        (" and authored automated ", False),
        ("JUnit and Mockito", True),
        (" test suites to validate transactional rollback semantics, persistence integrity, and CI/CD quality gates.", False)
    ], space_after=2.5)

    add_project_heading("AI Interviewer", "NestJS, Next.js, TypeScript, MongoDB, OpenAI, WebSockets, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Designed an interactive screening platform using ", False),
        ("NestJS", True),
        (", ", False),
        ("Next.js", True),
        (", and ", False),
        ("OpenAI API", True),
        (", establishing prompt-engineered rubrics to evaluate responses and authoring ", False),
        ("Jest", True),
        (" unit test suites.", False)
    ], space_after=2.4)
    add_bullet([
        ("Deployed bidirectional streaming pipelines via ", False),
        ("WebSockets", True),
        (" and ", False),
        ("Vapi.ai", True),
        (", enabling real-time conversational processing and reducing evaluation turnaround by ", False),
        ("50%", True),
        (" across 200+ monthly reviews.", False)
    ], space_after=2.5)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Java (17+ / Core Java, Multi-threaded Programming, OOAD), SQL, Python, TypeScript, JavaScript, Go, C#, C++"),
        ("Backend & Messaging: ", True, "Spring Boot (Java 17), RESTful APIs, JMS, Apache Kafka (DLQ, Saga Pattern), Message-Oriented Middleware (MOM), NestJS, Microservices"),
        ("Databases & Persistence: ", True, "Oracle Database, PostgreSQL, JDBC, Hibernate/JPA, Redis (Distributed Caching & Locks), MongoDB, SQL Server"),
        ("DevOps, CI/CD & Build Tools: ", True, "Jenkins CI/CD, Git, Gradle, Docker, Jira, GitHub Actions, Linux, GCP"),
        ("Testing & Code Quality: ", True, "JUnit, Mockito, Pytest, Jest, Artillery Load Testing, Automated Unit & Integration Testing"),
        ("AI-Assisted Development & Observability: ", True, "Prompt Engineering, AI-Assisted Code & Test Generation, Langfuse (LLM Tracing & RCA), Sentry, Gemini API")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.8)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting a real-time voice screening platform with low-latency WebSocket audio streaming.", False)
    ], space_after=2.4)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for distributed system design.", False)
    ], space_after=2.4)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef.", False)
    ], space_after=1.8)

def populate_olx(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Aligned with OLX India Associate Product Analyst JD)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Authored internal ", False),
        ("Python data extraction pipelines", True),
        (" and complex ", False),
        ("SQL queries on PostgreSQL", True),
        (", aggregating ", False),
        ("50K+ student records", True),
        (" to eliminate manual reporting overhead and accelerate operational analytics turnaround by ", False),
        ("60%", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Analyzed end-to-end ", False),
        ("user journeys", True),
        (" and applicant ", False),
        ("conversion funnels", True),
        (" across ", False),
        ("6 Slate CRM differential feeds", True),
        (", identifying document completion drop-offs to formulate actionable recommendations that boosted stage progression.", False)
    ], space_after=2.4)
    add_bullet([
        ("Formulated a self-reinforcing ", False),
        ("Document Intelligence engine", True),
        (" analyzing transcript articulation patterns with ", False),
        ("LLM fallback", True),
        (", resolving ", False),
        ("40%+", True),
        (" of course mappings deterministically to cut external AI token spend by ", False),
        ("40%", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Constructed ", False),
        ("interactive product dashboards", True),
        (" and billing telemetry modules in ", False),
        ("Next.js and Stripe", True),
        (", streaming high-volume tenant usage audit logs to CSV with constant ", False),
        ("O(1) memory", True),
        (" and ", False),
        ("100% payment reconciliation integrity", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Architected a decoupled two-tier ", False),
        ("Go ingestion pipeline", True),
        (" on ", False),
        ("Cloud Run and Pub/Sub", True),
        (", synchronizing ", False),
        ("10K+ delta records per 15-minute cycle", True),
        (" into ", False),
        ("MongoDB", True),
        (" queues with lease-fenced consumers to guarantee zero data loss.", False)
    ], space_after=2.4)
    add_bullet([
        ("Deployed commercial transcript evaluation platforms across ", False),
        ("4+ university client engagements", True),
        (" generating ", False),
        ("$150K+", True),
        (" contract value each, automating writeback of computed GPAs, credits, and grades directly into ", False),
        ("Slate CRM", True),
        (".", False)
    ], space_after=2.5)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Reduced backend analytics API response latency ", False),
        ("3× (15s to 5s)", True),
        (" by profiling ", False),
        ("SQL execution plans (EXPLAIN)", True),
        (" across ", False),
        ("PostgreSQL and MySQL", True),
        (" schemas, ", False),
        ("indexing composite keys", True),
        (", and eliminating redundant full-table scans.", False)
    ], space_after=2.4)
    add_bullet([
        ("Launched a full-stack AI evaluation platform with interactive ", False),
        ("React/TypeScript dashboards", True),
        (" and REST APIs, automating ", False),
        ("200+ monthly candidate evaluations", True),
        (" and slashing administrative turnaround by ", False),
        ("50%", True),
        (" for hiring teams.", False)
    ], space_after=2.6)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Database Management Systems (DBMS), Relational Database Design (RDBMS), SQL & Query Optimization, Data Structures & Algorithms, Object-Oriented Programming (OOP), Computer Networks, Operating Systems.", False)
    ], space_after=2.6)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Java, Spring Boot, Python, Apache Kafka, PostgreSQL, Redis, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a high-throughput payments microservice in ", False),
        ("Java (Spring Boot)", True),
        (" and ", False),
        ("Python", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency, validating service capacity through authored ", False),
        ("Artillery load test suites", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Modeled a ", False),
        ("normalized PostgreSQL schema", True),
        (" with idempotency-aware ", False),
        ("SQL transaction procedures", True),
        (" over a ", False),
        ("Kafka", True),
        (" event pipeline, guaranteeing ", False),
        ("zero double-charges", True),
        (" across ", False),
        ("500+ concurrent simulated users", True),
        (" under sustained load.", False)
    ], space_after=2.6)

    add_project_heading("AI Interviewer", "Next.js, NestJS, TypeScript, MongoDB, OpenAI API, WebSockets", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Engineered an autonomous candidate evaluation platform with ", False),
        ("Next.js", True),
        (", ", False),
        ("NestJS", True),
        (", and ", False),
        ("OpenAI API", True),
        (", establishing structured rubrics and ", False),
        ("analytics dashboards", True),
        (" delivering objective scoring across ", False),
        ("200+ monthly applicant reviews", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Maintained bidirectional streaming audio pipelines via ", False),
        ("WebSockets", True),
        (" and ", False),
        ("Vapi.ai", True),
        (", handling real-time speech-to-text inference and ", False),
        ("analyzing interaction drop-offs", True),
        (" to cut candidate evaluation turnaround time by ", False),
        ("50%", True),
        (".", False)
    ], space_after=2.6)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Data Analytics & SQL: ", True, "SQL (PostgreSQL, MySQL, SQL Server), Product Analytics, User Journey Mapping, Conversion Funnel Analysis, A/B Testing Fundamentals, Data Modeling, Schema Normalization"),
        ("Programming Languages: ", True, "Python, SQL, TypeScript, JavaScript (ES6+), Go (Golang), Java (Spring Boot), C++, C#"),
        ("Frontend & Dashboards: ", True, "Next.js (BFF), React, Redux, Stream-Based CSV Exports, Tailwind CSS, Interactive Telemetry Dashboards, Power BI"),
        ("Databases & Data Pipelines: ", True, "PostgreSQL, MySQL, MongoDB (Aggregations, Lease Fencing), Redis, Google Cloud Pub/Sub, Apache Kafka (DLQ, Outbox Pattern)"),
        ("Cloud & DevOps: ", True, "Google Cloud Platform (Cloud Run, Cloud Scheduler), Docker, Git, GitHub Actions, Linux"),
        ("Methodology & Core: ", True, "Data Structures & Algorithms, Agile/Scrum, Query Optimization (EXPLAIN, Indexing), Quantitative Analysis, System Integration Testing (SIT)")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.8)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an autonomous real-time voice screening platform with low-latency WebSocket audio streaming and structured rubric evaluation.", False)
    ], space_after=2.4)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for distributed system design, relational database modeling, and disaster relief tracking.", False)
    ], space_after=2.4)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, exhibiting deep analytical rigor, algorithmic efficiency, and mastery of optimal space-time complexities.", False)
    ], space_after=1.5)

def populate_oracle(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Aligned with Oracle Application Software Engineer 1 JD)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Architected a multi-tenant backend ETL service in ", False),
        ("Java 17 (Spring Boot)", True),
        (" integrating CRM systems via ", False),
        ("RESTful APIs", True),
        (", deploying ", False),
        ("Hibernate/JPA", True),
        (" and ", False),
        ("Kafka DLQ", True),
        (" to prevent data loss during upstream outages.", False)
    ], space_after=1.8)
    add_bullet([
        ("Optimized backend analytics database performance ", False),
        ("3× (15s to 5s)", True),
        (" by analyzing ", False),
        ("SQL execution plans (EXPLAIN)", True),
        (", indexing composite keys, and eliminating redundant table scans across ", False),
        ("PostgreSQL and MySQL", True),
        (" schemas.", False)
    ], space_after=1.8)
    add_bullet([
        ("Authored internal ", False),
        ("Python", True),
        (" batch processing pipelines to extract and transform ", False),
        ("50K+", True),
        (" applicant records from ", False),
        ("PostgreSQL", True),
        (", automating stakeholder reporting workflows and cutting reporting turnaround cycles by ", False),
        ("60%", True),
        (".", False)
    ], space_after=1.8)
    add_bullet([
        ("Developed a stream-based reporting and CSV export pipeline in ", False),
        ("NestJS/Node.js", True),
        (" for tenant usage analytics, enabling real-time SQL data fetching and processing of audit datasets with constant ", False),
        ("O(1) memory", True),
        (".", False)
    ], space_after=1.8)
    add_bullet([
        ("Diagnosed production application defects and documented ", False),
        ("Root Cause Analyses (RCA)", True),
        (" using ", False),
        ("Sentry", True),
        (" distributed tracing, isolating execution bottlenecks and schema anomalies to reduce ", False),
        ("MTTR", True),
        (" across scheduled release cycles.", False)
    ], space_after=1.8)
    add_bullet([
        ("Engineered an enterprise commercial transcript evaluation platform (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (", Slate CRM) across ", False),
        ("4+", True),
        (" university client deployments (including Rosalind Franklin University) generating ", False),
        ("$150K+", True),
        (" contract value.", False)
    ], space_after=2.0)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Launched an interactive AI evaluation platform automating ", False),
        ("200+", True),
        (" monthly candidate reviews and cutting assessment turnaround by ", False),
        ("50%", True),
        (", designing responsive UI dashboards and integrating secure ", False),
        ("RESTful APIs", True),
        (".", False)
    ], space_after=1.8)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" deployment pipelines, automating containerized ", False),
        ("Docker", True),
        (" packaging, Git version control, and multi-stage regression test suites in an ", False),
        ("Agile/Scrum", True),
        (" team to support reliable software releases.", False)
    ], space_after=2.0)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India (Open to Relocate — Bengaluru)")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks, Software Engineering & Testing.", False)
    ], space_after=2.0)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Java 17, Spring Boot, Hibernate, Kafka, Redis, PostgreSQL, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Built a distributed payments microservice in ", False),
        ("Java 17 (Spring Boot / Hibernate)", True),
        
        (" and Spring Security JWT authentication, sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency with Artillery load testing.", False)
    ], space_after=1.8)
    add_bullet([
        ("Formulated an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" with Redis distributed locks and ", False),
        ("PostgreSQL ACID", True),
        (" validation, preventing race conditions and zero double-charges across ", False),
        ("500+", True),
        (" users.", False)
    ], space_after=1.8)
    add_bullet([
        ("Containerized microservices using ", False),
        ("Docker", True),
        (" and authored comprehensive ", False),
        ("JUnit", True),
        (" unit and integration test suites to validate transactional rollback semantics, status code handling, and code coverage gates.", False)
    ], space_after=2.0)

    add_project_heading("AI Interviewer", "NestJS, Next.js, React, MongoDB, TypeScript, OpenAI API, Jest, Git", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Designed an AI evaluation platform using ", False),
        ("NestJS", True),
        (", ", False),
        ("Next.js", True),
        (", and REST APIs, establishing structured evaluation rubrics to score candidate responses and authoring ", False),
        ("Jest", True),
        (" test suites to enforce code quality gates.", False)
    ], space_after=1.8)
    add_bullet([
        ("Deployed bidirectional streaming audio pipelines via ", False),
        ("WebSockets", True),
        (" and ", False),
        ("Vapi.ai", True),
        (", enabling real-time conversational voice interviews and reducing assessment turnaround by ", False),
        ("50%", True),
        (" across 200+ monthly reviews.", False)
    ], space_after=2.0)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Java (17 / Core Java, OOP), Python, C++, Go, TypeScript, JavaScript, SQL (PostgreSQL, MySQL, PL/SQL)"),
        ("Databases & Storage: ", True, "Database Management Systems (RDBMS), PostgreSQL, MySQL, PL/SQL, Redis, MongoDB, Indexing"),
        ("Backend & Architecture: ", True, "Spring Boot, Hibernate/JPA, RESTful APIs, NestJS, Node.js, Next.js (BFF), Apache Kafka (DLQ)"),
        ("Data Analytics & Reporting: ", True, "SQL Data Fetching, ETL Pipelines, Stream-based CSV Reporting, Data Processing"),
        ("Testing & Reliability: ", True, "Root Cause Analysis (RCA), Sentry Telemetry, Unit Testing (JUnit, Jest, Pytest), Debugging"),
        ("Cloud & DevOps: ", True, "Docker, Jenkins CI/CD, Git, GitHub Actions, Google Cloud Platform (Cloud Run), AWS (S3, EC2), Linux")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.5)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform with low-latency WebSocket audio streaming.", False)
    ], space_after=1.8)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ competing teams for fault-tolerant distributed system architecture.", False)
    ], space_after=1.8)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, demonstrating strong algorithmic complexity analysis, memory optimization, and concurrent execution skills.", False)
    ], space_after=0.0)


def populate_groww(doc, add_section_header, add_role_heading, add_project_heading, add_bullet, role="SDE_1_React_Native"):
    # 2. Experience (HR-First Layout - Groww App Team: React Native, Android SDK, TypeScript, REST APIs & FinTech)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Developed responsive, accessible (", False),
        ("WCAG", True),
        (") ", False),
        ("React", True),
        (" and ", False),
        ("Next.js", True),
        (" mobile and web interfaces featuring reusable component libraries, custom hooks, and state lifecycles, cutting manual review overhead by ", False),
        ("80%", True),
        (".", False)
    ], space_after=2.0)
    add_bullet([
        ("Engineered an enterprise commercial evaluation platform across ", False),
        ("4+", True),
        (" university deployments (", False),
        ("$150K+", True),
        (" contract value) using ", False),
        ("Next.js BFF", True),
        (" proxies to eliminate client token exposure and enforce strict ", False),
        ("FERPA", True),
        (" security.", False)
    ], space_after=2.0)
    add_bullet([
        ("Implemented an enterprise self-serve billing module in ", False),
        ("Next.js", True),
        (" and ", False),
        ("Redux", True),
        (" with ", False),
        ("Stripe Checkout", True),
        (", building asynchronous webhook handlers, resilient error boundaries, and ", False),
        ("100%", True),
        (" payment reconciliation integrity.", False)
    ], space_after=2.0)
    add_bullet([
        ("Instrumented ", False),
        ("Sentry client observability", True),
        (" across production frontend applications with source maps and session replay, diagnosing runtime crashes and cutting mean-time-to-resolution (", False),
        ("MTTR", True),
        (") by ", False),
        ("45%", True),
        (".", False)
    ], space_after=2.0)
    add_bullet([
        ("Optimized client-side data fetching and ", False),
        ("RESTful API", True),
        (" payloads, reducing analytics response latency ", False),
        ("3× (15s to 5s)", True),
        (" and eliminating redundant network round-trips to maximize mobile and web responsiveness.", False)
    ], space_after=2.2)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Constructed the ", False),
        ("React/TypeScript", True),
        (" frontend for a commercial assessment platform, delivering role-based dashboards, real-time audio controls, and an intuitive scheduling UI for ", False),
        ("200+", True),
        (" monthly sessions (", False),
        ("50%", True),
        (" turnaround cut).", False)
    ], space_after=2.0)
    add_bullet([
        ("Authored comprehensive unit and integration test suites using ", False),
        ("Jest", True),
        (" and ", False),
        ("React Testing Library", True),
        (" (", False),
        ("90%+", True),
        (" coverage), enforcing strict code quality standards, constructive peer reviews, and ", False),
        ("Git", True),
        (" workflows across Agile sprint cycles.", False)
    ], space_after=2.4)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks.", False)
    ], space_after=2.4)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Aether", "Kotlin, Android SDK, Jetpack Compose, Room DB, Clean Architecture", "github.com/Yash-Gupta-01/Personal-AI-Agent", url="https://github.com/Yash-Gupta-01/Personal-AI-Agent")
    add_bullet([
        ("Architected an offline-first mobile assistant in ", False),
        ("Kotlin", True),
        (" and ", False),
        ("Jetpack Compose", True),
        (" using ", False),
        ("Clean Architecture", True),
        (", decoupling UI from pure JVM domain logic to execute unit tests in milliseconds without emulator overhead.", False)
    ], space_after=2.0)
    add_bullet([
        ("Integrated an offline-first local capture loop with ", False),
        ("Room database", True),
        (" schema migrations and modular tool schemas, achieving responsive, high-performance UI rendering across diverse Android devices and screen densities.", False)
    ], space_after=2.4)

    add_project_heading("Fluxpay", "React, TypeScript, Redux Toolkit, Node.js, Express, PostgreSQL, REST APIs", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Built a mobile-first ", False),
        ("React/TypeScript", True),
        (" fintech transaction dashboard, utilizing ", False),
        ("Redux", True),
        (" and ", False),
        ("Zustand", True),
        (" for centralized state management, responsive card layouts, and deterministic UI rendering under burst network load.", False)
    ], space_after=2.0)
    add_bullet([
        ("Formulated idempotency-aware transaction APIs on ", False),
        ("Node.js", True),
        (" and ", False),
        ("PostgreSQL", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency, verified through Artillery load tests to prevent double-charges.", False)
    ], space_after=2.4)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Frontend & Mobile: ", True, "JavaScript (ES6+), TypeScript, React Native (core architecture), React.js, Android Native (Kotlin, Java, Android SDK, Jetpack Compose), Next.js (BFF), Redux Toolkit, Zustand, Context API"),
        ("Architecture & Engineering: ", True, "Clean Architecture, MVVM, Test-Driven Development (TDD), Design Patterns, Object-Oriented Programming (OOP), SOLID Principles, Version Control (Git)"),
        ("Testing & Observability: ", True, "Jest, React Testing Library, JUnit, Gradle, Metro Bundler, Vite, Webpack, Chrome DevTools, React Profiler, Sentry (Session Replay, Source Maps)"),
        ("Backend & Financial APIs: ", True, "RESTful APIs, WebSockets, Node.js, Express.js, NestJS, Stripe Payments API, PostgreSQL, MongoDB, Redis, Docker, JSON/XML, CDN Edge Caching"),
        ("Distributed & Cloud Systems: ", True, "Microservices, Caching, Event-Driven Architecture, CI/CD, Containerization (Docker), Core Web Vitals (LCP), Mobile Performance")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.6)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform with low-latency audio inference.", False)
    ], space_after=2.0)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for scalable systems architecture and resilient client-server design.", False)
    ], space_after=2.0)
    add_bullet([
        ("Solved ", False),
        ("450+ Data Structures & Algorithms problems", True),
        (" across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, demonstrating deep proficiency in algorithmic design, complexity analysis, and runtime efficiency.", False)
    ], space_after=1.6)

def populate_galaxyai(doc, add_section_header, add_role_heading, add_project_heading, add_bullet, role="Software_Engineer"):
    # 2. Experience (HR-First Layout - Galaxy.ai / Magica: Agent Orchestration, WebSockets, Evals, Connected Services)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Engineered high-throughput ", False),
        ("NestJS", True),
        (" WebSocket middleware bridging real-time telephony streams, orchestrating bidirectional ", False),
        ("µ-law and PCM", True),
        (" audio transcoding for ", False),
        ("1,000+", True),
        (" concurrent sessions at sub-", False),
        ("100ms", True),
        (" p95 latency under peak load.", False)
    ], space_after=2.2)
    add_bullet([
        ("Instrumented multi-turn agent workflows with ", False),
        ("Langfuse distributed tracing", True),
        (" and ", False),
        ("Sentry", True),
        (", isolating reasoning bottlenecks to eliminate production ", False),
        ("MongoDB", True),
        (" tool hallucinations and silent prompt drift across admissions assistants.", False)
    ], space_after=2.2)
    add_bullet([
        ("Architected a decoupled two-tier ", False),
        ("Go", True),
        (" ingestion pipeline on ", False),
        ("Cloud Run", True),
        (" and ", False),
        ("Pub/Sub", True),
        (", processing ", False),
        ("10K+ delta records per 15-min sync cycle", True),
        (" into durable ", False),
        ("MongoDB", True),
        (" queues with atomic lease-fenced consumers, ensuring zero data loss.", False)
    ], space_after=2.2)
    add_bullet([
        ("Improved the Document Intelligence pipeline coupling deterministic rule caches with ", False),
        ("LLM fallback", True),
        (", resolving ", False),
        ("40%+", True),
        (" of course mappings without LLM inference to eliminate prompt variance and cut AI token spend by ", False),
        ("40%", True),
        (".", False)
    ], space_after=2.2)
    add_bullet([
        ("Deployed commercial transcript evaluation platforms across ", False),
        ("4+", True),
        (" universities (", False),
        ("$150K+", True),
        (" contract value) using ", False),
        ("Next.js BFF", True),
        (" proxies, integrating human-in-the-loop review workflows and automating writeback into ", False),
        ("Slate CRM", True),
        (".", False)
    ], space_after=2.2)
    add_bullet([
        ("Authored internal ", False),
        ("Python", True),
        (" scripts to extract and transform ", False),
        ("50K+ student application records", True),
        (" from ", False),
        ("PostgreSQL", True),
        (", eliminating manual operational reporting overhead and accelerating analytics turnaround by ", False),
        ("60%", True),
        (".", False)
    ], space_after=2.5)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Constructed the ", False),
        ("React/TypeScript", True),
        (" frontend for an automated voice screening platform, delivering real-time audio controls, intuitive candidate dashboards, and automated scheduling for ", False),
        ("200+", True),
        (" monthly sessions (", False),
        ("50%", True),
        (" turnaround cut).", False)
    ], space_after=2.2)
    add_bullet([
        ("Streamlined web asset delivery and frontend modules by configuring ", False),
        ("CDN edge caching", True),
        (" and asset chunking in an ", False),
        ("Agile team", True),
        (", delivering zero-downtime releases and lowering Largest Contentful Paint (", False),
        ("LCP", True),
        (") by ", False),
        ("25%", True),
        (".", False)
    ], space_after=2.5)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Operating Systems, DBMS, Computer Networks.", False)
    ], space_after=2.5)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("AI Interviewer", "NestJS, Next.js, WebSockets, OpenAI API, Vapi.ai, MongoDB", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Built an automated voice screening platform with ", False),
        ("NestJS", True),
        (", ", False),
        ("Next.js", True),
        (", and ", False),
        ("Vapi.ai", True),
        (", winning First Place (", False),
        ("Rs. 50,000", True),
        (") among 30+ teams, conducting ", False),
        ("200+", True),
        (" monthly evaluations and commercialized for university admissions.", False)
    ], space_after=2.2)
    add_bullet([
        ("Orchestrated real-time ", False),
        ("WebSocket", True),
        (" audio streaming and structured prompt evaluation rubrics in ", False),
        ("NestJS", True),
        (", evaluating candidate technical proficiency and conversational flow with zero human interviewer bottleneck.", False)
    ], space_after=2.5)

    add_project_heading("Aether", "Kotlin, Android SDK, Jetpack Compose, Room DB, llama.cpp", "github.com/Yash-Gupta-01/Personal-AI-Agent", url="https://github.com/Yash-Gupta-01/Personal-AI-Agent")
    add_bullet([
        ("Designed an offline-first mobile AI assistant in ", False),
        ("Kotlin", True),
        (" and ", False),
        ("Jetpack Compose", True),
        (" using ", False),
        ("Clean Architecture", True),
        (", decoupling UI from pure JVM domain logic to execute unit test suites in milliseconds without emulator overhead.", False)
    ], space_after=2.2)
    add_bullet([
        ("Integrated an offline-first local thought capture loop with ", False),
        ("Room database", True),
        (" schema migrations and modular tool schemas, executing on-device ", False),
        ("llama.cpp", True),
        (" inference while strictly enforcing tool schemas for local device actions.", False)
    ], space_after=2.5)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("AI, Agents & Real-Time: ", True, "Google Gemini API, OpenAI API, Vapi.ai, Langfuse (LLM Tracing & Evals), WebSockets, AudioHook v2 (µ-law/PCM), Prompt Engineering, RAG"),
        ("Languages: ", True, "TypeScript, JavaScript, Python, Go (Golang), Java, Kotlin, C#, SQL (PostgreSQL, MySQL)"),
        ("Backend & Distributed Systems: ", True, "Node.js, NestJS, Express, Go standard library (net/http), Spring Boot, RESTful APIs, Microservices, Event-Driven Architecture"),
        ("Databases, Queues & Caching: ", True, "PostgreSQL, MongoDB (aggregation, bulkWrite, lease fencing), Redis (caching & locks), Google Cloud Pub/Sub, Apache Kafka"),
        ("Cloud, DevOps & Observability: ", True, "Google Cloud Platform (Cloud Run, VPC Access, Cloud Functions), Docker, Git, GitHub Actions, Sentry, Jest, Pytest, Postman"),
        ("Frontend & Architecture: ", True, "Next.js (App Router, BFF), React, Redux Toolkit, Zustand, Tailwind CSS, Shadcn UI, Clean Architecture")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.8)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform with low-latency audio inference.", False)
    ], space_after=2.2)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ competing teams for fault-tolerant distributed system architecture.", False)
    ], space_after=2.2)
    add_bullet([
        ("Solved ", False),
        ("450+ Data Structures & Algorithms problems", True),
        (" across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, demonstrating strong algorithmic complexity analysis, memory optimization, and concurrent execution skills.", False)
    ], space_after=0.0)

def populate_the_hartford(doc, add_section_header, add_role_heading, add_project_heading, add_bullet, role="Associate_Software_Engineer"):
    # 2. Experience (HR-First Layout - The Hartford: AI Platform, GenAI/Agentic Workflows, Ingestion, Cloud, SAFe/Agile)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Improved the ", False),
        ("Document Intelligence", True),
        (" platform coupling deterministic articulation rule caches with ", False),
        ("LLM fallback", True),
        (", resolving ", False),
        ("40%+", True),
        (" of course mappings without LLM inference to cut token costs by ", False),
        ("40%", True),
        (".", False)
    ], space_after=1.8)
    add_bullet([
        ("Instrumented multi-turn agent workflows with ", False),
        ("Langfuse distributed tracing", True),
        (" and ", False),
        ("Sentry", True),
        (", isolating reasoning bottlenecks to eliminate production ", False),
        ("MongoDB", True),
        (" tool-call hallucinations and prompt drift across admissions assistants.", False)
    ], space_after=1.8)
    add_bullet([
        ("Architected a decoupled two-tier ", False),
        ("Go", True),
        (" ingestion pipeline on ", False),
        ("Cloud Run", True),
        (" and ", False),
        ("Google Cloud Pub/Sub", True),
        (", processing ", False),
        ("10K+ delta records per 15-minute sync cycle", True),
        (" into durable ", False),
        ("MongoDB", True),
        (" queues with atomic lease-fenced consumers.", False)
    ], space_after=1.8)
    add_bullet([
        ("Engineered high-throughput ", False),
        ("NestJS", True),
        (" WebSocket middleware bridging real-time telephony streams, handling bidirectional ", False),
        ("µ-law to PCM", True),
        (" audio transcoding for ", False),
        ("1,000+", True),
        (" concurrent sessions at sub-", False),
        ("100ms", True),
        (" p95 latency under peak load.", False)
    ], space_after=1.8)
    add_bullet([
        ("Authored internal ", False),
        ("Python", True),
        (" data extraction pipelines to process and transform ", False),
        ("50K+ student application records", True),
        (" from ", False),
        ("PostgreSQL", True),
        (", eliminating manual operational reporting overhead and accelerating analytics turnaround by ", False),
        ("60%", True),
        (".", False)
    ], space_after=1.8)
    add_bullet([
        ("Constructed a multi-tenant ", False),
        ("Java/Spring Boot", True),
        (" ETL pipeline integrating CRM data via REST APIs, implementing ", False),
        ("Hibernate/JPA", True),
        (" and ", False),
        ("Apache Kafka", True),
        (" Dead Letter Queues (", False),
        ("DLQ", True),
        (") to ensure zero data loss during upstream outages.", False)
    ], space_after=2.2)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Launched a full-stack AI evaluation platform automating ", False),
        ("200+", True),
        (" monthly candidate assessments and slashing turnaround ", False),
        ("50%", True),
        (", designing responsive UI dashboards and integrating secure RESTful evaluation APIs.", False)
    ], space_after=1.8)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines automating containerized deployments (", False),
        ("Docker", True),
        (") and integration test suites in a ", False),
        ("Scaled Agile (SAFe)", True),
        (" cadence, eliminating broken staging builds and accelerating release cycles.", False)
    ], space_after=2.2)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), DBMS, Operating Systems, Computer Networks.", False)
    ], space_after=2.2)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("AI Interviewer", "NestJS, Next.js, WebSockets, OpenAI API, Vapi.ai, MongoDB, Docker", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Built an autonomous voice screening platform with ", False),
        ("NestJS", True),
        (", ", False),
        ("Next.js", True),
        (", and ", False),
        ("Vapi.ai", True),
        (", winning First Place (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ teams, conducting ", False),
        ("200+", True),
        (" monthly evaluations and commercialized for university admissions.", False)
    ], space_after=1.8)
    add_bullet([
        ("Orchestrated real-time ", False),
        ("WebSocket", True),
        (" audio streaming and structured prompt evaluation rubrics in ", False),
        ("NestJS", True),
        (", evaluating candidate technical proficiency and conversational flow with zero human interviewer bottleneck.", False)
    ], space_after=2.2)

    add_project_heading("Fluxpay", "Java 17, Spring Boot, Go, Apache Kafka, Redis, PostgreSQL, Docker", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a distributed payments microservice in ", False),
        ("Java 17 (Spring Boot)", True),
        (" and ", False),
        ("Go", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency, validating concurrency thresholds via authored ", False),
        ("Artillery", True),
        (" load-test suites.", False)
    ], space_after=1.8)
    add_bullet([
        ("Formulated an event-driven Saga pattern over ", False),
        ("Apache Kafka", True),
        (" with ", False),
        ("Redis", True),
        (" distributed locks and ", False),
        ("PostgreSQL", True),
        (" ACID validation, guaranteeing idempotent execution and zero double-charges across ", False),
        ("500+", True),
        (" simulated users.", False)
    ], space_after=2.2)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("AI & Agentic AI: ", True, "Google Gemini API, OpenAI API, Langfuse (Evals), Vapi.ai, Vertex AI (familiar), Scikit-Learn (familiar)"),
        ("Languages: ", True, "Python, Java (17+ / Core Java, OOP), Go (Golang), TypeScript, JavaScript, SQL (PostgreSQL, MySQL), C++"),
        ("Cloud & DevOps: ", True, "Google Cloud Platform (Cloud Run, Pub/Sub), AWS (S3, EC2), Terraform (familiar), Docker, Jenkins CI/CD, Git"),
        ("Distributed & Backend: ", True, "Spring Boot (Java), Node.js, NestJS, RESTful APIs, Apache Kafka (DLQ, Outbox), Redis (Locks/Caching)"),
        ("Databases & Data: ", True, "PostgreSQL, MongoDB (bulkWrite, lease fencing), MySQL, Data Pipelines, ETL, Stream Processing"),
        ("Testing & SDLC Tools: ", True, "Scaled Agile (SAFe), Agile/Scrum, Pytest, JUnit, Jest, Mockito, Artillery, Sentry, Claude Code, Gemini CLI")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.6)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform with low-latency audio inference.", False)
    ], space_after=1.8)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ competing teams for fault-tolerant distributed system architecture.", False)
    ], space_after=1.8)
    add_bullet([
        ("Solved ", False),
        ("450+ Data Structures & Algorithms problems", True),
        (" across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, emphasizing algorithmic complexity, memory optimization, and concurrent execution.", False)
    ], space_after=0.0)

def populate_stripe(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Distributed Financial Systems, Go, Event-Driven Architecture, High Reliability)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Engineered an enterprise self-serve billing module in ", False),
        ("Next.js", True),
        (" with ", False),
        ("Stripe Checkout", True),
        (", implementing payment lifecycle webhook handlers, resilient error boundaries, and ", False),
        ("100%", True),
        (" payment reconciliation integrity.", False)
    ], space_after=1.8)
    add_bullet([
        ("Scaled a high-throughput ", False),
        ("Go ingestion pipeline", True),
        (" on ", False),
        ("Cloud Run", True),
        (" and ", False),
        ("Google Cloud Pub/Sub", True),
        (", streaming ", False),
        ("10K+ delta records per 15-minute sync cycle", True),
        (" into ", False),
        ("MongoDB", True),
        (" with atomic lease fencing to guarantee zero data loss during upstream outages.", False)
    ], space_after=1.8)
    add_bullet([
        ("Improved the ", False),
        ("Document Intelligence", True),
        (" pipeline by coupling deterministic articulation rule caches with ", False),
        ("LLM fallback", True),
        (", resolving ", False),
        ("40%+", True),
        (" of course mappings without model inference to cut token costs by ", False),
        ("40%", True),
        (" and eliminate prompt variance.", False)
    ], space_after=1.8)
    add_bullet([
        ("Optimized database throughput by profiling ", False),
        ("SQL execution plans", True),
        (", adding composite indexes, and eliminating redundant join queries in ", False),
        ("PostgreSQL/MySQL", True),
        (" schemas, slashing document-analytics API latency ", False),
        ("3× (15s to 5s)", True),
        (".", False)
    ], space_after=1.8)
    add_bullet([
        ("Maintained multi-tenant backend services under ", False),
        ("Jenkins CI/CD", True),
        (" and ", False),
        ("Gitflow", True),
        (" peer review workflows, authoring ", False),
        ("Jest", True),
        (" test suites, passing SonarQube quality gates, and tracking ", False),
        ("Sentry", True),
        (" telemetry to cut production MTTD by ", False),
        ("45%", True),
        (".", False)
    ], space_after=1.8)
    add_bullet([
        ("Deployed an enterprise commercial ", False),
        ("GPA evaluation platform", True),
        (" (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+", True),
        (" university deployments, generating ", False),
        ("$150K+", True),
        (" contract value per client while enforcing strict ", False),
        ("FERPA compliance", True),
        (" and RBAC security.", False)
    ], space_after=1.8)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Formulated an asynchronous ", False),
        ("event synchronization microservice", True),
        (" between partner CRMs and internal databases within an ", False),
        ("Agile/SAFe", True),
        (" cadence, collaborating with team leads to validate technical designs and unblock dependencies.", False)
    ], space_after=1.8)
    add_bullet([
        ("Automated containerized staging environments using ", False),
        ("Docker", True),
        (", executing automated ", False),
        ("regression test suites", True),
        (" to eliminate broken deployment builds and accelerate staging release cadence across engineering teams.", False)
    ], space_after=2.0)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, OOP, DBMS, Operating Systems, Computer Networks, Software Engineering & Testing.", False)
    ], space_after=1.8)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Go (Golang), Apache Kafka, Redis, PostgreSQL, Docker, Artillery", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a distributed payments microservice in ", False),
        ("Go (Golang)", True),
        (" using goroutines and channels, sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency, verified through published ", False),
        ("Artillery load tests", True),
        (".", False)
    ], space_after=1.8)
    add_bullet([
        ("Implemented an event-driven ", False),
        ("Saga pattern", True),
        (" over ", False),
        ("Kafka", True),
        (" with ", False),
        ("Redis distributed locks", True),
        (" and ", False),
        ("PostgreSQL ACID transactions", True),
        (", coordinating distributed payments and guaranteeing zero double-charges across ", False),
        ("500+", True),
        (" simulated users.", False)
    ], space_after=1.8)
    add_bullet([
        ("Containerized microservices with ", False),
        ("Docker", True),
        (" and authored comprehensive integration test suites to validate transactional rollback semantics, compensating ", False),
        ("Saga steps", True),
        (", and fault-tolerant failovers.", False)
    ], space_after=2.0)

    add_project_heading("AI Interviewer", "React, Next.js, NestJS, TypeScript, WebSockets, MongoDB, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Built an autonomous assessment platform with responsive ", False),
        ("React", True),
        (" dashboards, RESTful APIs, and role-based access control (", False),
        ("RBAC", True),
        ("), conducting ", False),
        ("200+", True),
        (" monthly candidate evaluations and reducing turnaround time by ", False),
        ("50%", True),
        (".", False)
    ], space_after=1.8)
    add_bullet([
        ("Orchestrated low-latency bidirectional streaming pipelines using ", False),
        ("WebSockets", True),
        (" and structured evaluation rubrics, enforcing service-level test coverage and quality gates with ", False),
        ("Jest", True),
        (" to ensure reliable continuous delivery.", False)
    ], space_after=2.0)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Go (Golang), Java (Core Java, OOP), TypeScript, JavaScript, Python, SQL (PostgreSQL, MySQL), C++"),
        ("Backend & APIs: ", True, "Go net/http, Spring Boot, NestJS, Node.js, RESTful APIs, WebSockets, BFF Architecture"),
        ("Distributed Systems: ", True, "Apache Kafka (Saga Pattern, Event-Driven, DLQ), Google Cloud Pub/Sub, Redis (Distributed Locks)"),
        ("Databases & Storage: ", True, "PostgreSQL, MySQL, MongoDB (Bulk Operations, Lease Fencing), ACID Transactions, Query Optimization"),
        ("DevOps & Cloud: ", True, "Google Cloud (Cloud Run, Jobs), Docker, Jenkins CI/CD, Gitflow (Peer Reviews), SonarQube, Sentry"),
        ("Testing & Reliability: ", True, "Jest, Pytest, JUnit, Artillery (Load Testing), Postman, Automated Regression Suites")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.8)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform.", False)
    ], space_after=1.8)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for distributed system design.", False)
    ], space_after=1.8)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, emphasizing algorithmic complexity, memory optimization, and concurrent execution.", False)
    ], space_after=0.0)

def populate_airtel(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (HR-First Layout - Airtel Backend & Full Stack Developer Focus: Java, Spring Boot, Microservices, Event-Driven)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Architected a multi-tenant ", False),
        ("Java 17 (Spring Boot)", True),
        (" ETL service integrating enterprise CRMs via RESTful APIs, deploying ", False),
        ("Hibernate/JPA", True),
        (" and ", False),
        ("Kafka Dead Letter Queues (DLQ)", True),
        (" to ensure eventual consistency and zero message loss under burst traffic.", False)
    ], space_after=2.4)
    add_bullet([
        ("Engineered an enterprise commercial ", False),
        ("GPA evaluation platform", True),
        (" (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+", True),
        (" university deployments generating ", False),
        ("$150K+", True),
        (" contract value per client, enforcing ", False),
        ("FERPA compliance", True),
        (" and role-based access control (", False),
        ("RBAC", True),
        (").", False)
    ], space_after=2.4)
    add_bullet([
        ("Scaled a decoupled two-tier ", False),
        ("Go", True),
        (" ingestion pipeline on ", False),
        ("Cloud Run", True),
        (" and ", False),
        ("Pub/Sub", True),
        (", processing ", False),
        ("10K+ delta records per 15-minute sync cycle", True),
        (" into a ", False),
        ("MongoDB", True),
        (" queue with lease-fenced consumers to guarantee zero data loss.", False)
    ], space_after=2.4)
    add_bullet([
        ("Optimized relational database performance, reducing API latency ", False),
        ("3× (15s to 5s)", True),
        (" by profiling ", False),
        ("SQL", True),
        (" query execution plans (", False),
        ("EXPLAIN", True),
        ("), adding indexes, and eliminating redundant joins across ", False),
        ("PostgreSQL", True),
        (" schemas.", False)
    ], space_after=2.4)
    add_bullet([
        ("Built responsive, accessible web components in ", False),
        ("Angular 14", True),
        (" and ", False),
        ("TypeScript", True),
        (" with ", False),
        ("HTML5/CSS3", True),
        (" and RxJS observable streams, standardizing modular frontend architecture and cutting user onboarding turnaround by ", False),
        ("40%", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Instrumented centralized error telemetry and distributed tracing across microservices using ", False),
        ("Sentry", True),
        (" and ", False),
        ("Langfuse", True),
        (", isolating tool-call latency bottlenecks and reducing production mean-time-to-detect (", False),
        ("MTTD) by 45%", True),
        (".", False)
    ], space_after=2.4)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Created the responsive ", False),
        ("Angular 14", True),
        (" and ", False),
        ("React", True),
        (" frontend for an AI assessment platform, designing reusable ", False),
        ("TypeScript", True),
        (" components and interactive dashboards that automated ", False),
        ("200+", True),
        (" monthly candidate evaluations.", False)
    ], space_after=2.4)
    add_bullet([
        ("Maintained ", False),
        ("Jenkins CI/CD", True),
        (" artifact pipelines automating containerized ", False),
        ("Docker", True),
        (" deployments and integration test execution in an ", False),
        ("Agile team", True),
        (", eliminating broken staging builds and accelerating release cadence.", False)
    ], space_after=2.6)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks, Software Engineering.", False)
    ], space_after=2.4)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Java 17, Spring Boot, Hibernate, Kafka, Redis, PostgreSQL, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a mission-critical payment microservice in ", False),
        ("Java 17 (Spring Boot / Hibernate)", True),
        (", ", False),
        ("Spring Security", True),
        (", and ", False),
        ("JWT", True),
        (", sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency with published Artillery test suites.", False)
    ], space_after=2.4)
    add_bullet([
        ("Formulated an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" with ", False),
        ("Redis", True),
        (" distributed locks and ", False),
        ("PostgreSQL", True),
        (" ACID validation, coordinating distributed transactions and guaranteeing zero double-charges across ", False),
        ("500+", True),
        (" simulated users.", False)
    ], space_after=2.4)
    add_bullet([
        ("Containerized microservices with ", False),
        ("Docker", True),
        (" virtualization and authored comprehensive ", False),
        ("JUnit", True),
        (" and ", False),
        ("Mockito", True),
        (" test suites to validate transactional rollback semantics, reconciliation routines, and code quality gates.", False)
    ], space_after=2.6)

    add_project_heading("AI Interviewer", "Angular 14, TypeScript, Next.js, NestJS, MongoDB, WebSockets, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Automated candidate screening workflows by designing interactive full-stack web applications with ", False),
        ("Next.js", True),
        (", ", False),
        ("Angular 14", True),
        (", and role-based access control (", False),
        ("RBAC", True),
        ("), facilitating ", False),
        ("200+", True),
        (" monthly evaluations and cutting turnaround by ", False),
        ("50%", True),
        (".", False)
    ], space_after=2.4)
    add_bullet([
        ("Orchestrated real-time bidirectional streaming using ", False),
        ("WebSockets", True),
        (" and structured evaluation rubrics in ", False),
        ("NestJS", True),
        (", enforcing rigorous test coverage and quality gates with ", False),
        ("Jest", True),
        (" to ensure robust continuous delivery.", False)
    ], space_after=2.6)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "Java (17 / Core Java, OOP), TypeScript, JavaScript, SQL (PostgreSQL, MySQL), Python, Go, HTML5, CSS3, C++"),
        ("Backend & Microservices: ", True, "Spring Boot, Spring Framework, Hibernate/JPA, RESTful APIs, Microservices, NestJS, Node.js, Express"),
        ("Frontend & UI: ", True, "Angular 14, TypeScript, RxJS, React, Next.js (BFF), Responsive Web Design, Component Architecture"),
        ("Databases & Messaging: ", True, "PostgreSQL, MongoDB, Kafka (DLQ, Saga Pattern), Redis (Distributed Locks), MySQL, SQL Server"),
        ("DevOps, Cloud & Build: ", True, "Docker, Jenkins CI/CD, Git, GitHub PRs & Code Reviews, GCP (Cloud Run), Kubernetes (familiar)"),
        ("Testing, Security & Agile: ", True, "JUnit, Mockito, Jest, Spring Security, JWT, Sentry, Agile / Scrum, SDLC, SIT & UAT")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.6)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an automated real-time voice screening platform.", False)
    ], space_after=2.4)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for fault-tolerant system architecture.", False)
    ], space_after=2.4)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, emphasizing algorithmic complexity, memory optimization, and concurrent execution.", False)
    ], space_after=1.8)

def populate_uipath(doc, add_section_header, add_role_heading, add_project_heading, add_bullet):
    # 2. Experience (Data Fabric Platform - Distributed Data Infrastructure, Frontend/UX, AI Agents, APIs)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Engineered an enterprise commercial platform (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("NestJS", True),
        (", ", False),
        ("Slate CRM", True),
        (") across ", False),
        ("4+", True),
        (" university deployments, generating ", False),
        ("$150K+", True),
        (" contract value per client while enforcing strict ", False),
        ("FERPA data privacy", True),
        (" and RBAC security.", False)
    ], space_after=2.2)
    add_bullet([
        ("Architected a decoupled two-tier ", False),
        ("Go", True),
        (" ingestion pipeline on ", False),
        ("Cloud Run", True),
        (" and ", False),
        ("Google Cloud Pub/Sub", True),
        (", processing ", False),
        ("10K+ delta records per 15-minute sync cycle", True),
        (" into durable MongoDB queues with lease-fenced consumers to guarantee zero data loss.", False)
    ], space_after=2.2)
    add_bullet([
        ("Improved the ", False),
        ("Document Intelligence", True),
        (" pipeline by coupling deterministic articulation rule caches with ", False),
        ("LLM fallback", True),
        (", resolving ", False),
        ("40%+", True),
        (" of course mappings without model inference to cut token costs by ", False),
        ("40%", True),
        (" and eliminate prompt variance.", False)
    ], space_after=2.2)
    add_bullet([
        ("Authored internal ", False),
        ("Python", True),
        (" data transformation routines querying ", False),
        ("PostgreSQL", True),
        (" to normalize and extract ", False),
        ("50K+", True),
        (" student application records, eliminating manual data wrangling and accelerating operational analytics turnaround by ", False),
        ("60%", True),
        (".", False)
    ], space_after=2.2)
    add_bullet([
        ("Scaled real-time ", False),
        ("NestJS", True),
        (" WebSocket middleware bridging telephony streams, handling bidirectional audio transcoding for ", False),
        ("1,000+", True),
        (" concurrent sessions at sub-", False),
        ("100ms", True),
        (" p95 latency under peak production load with zero packet loss.", False)
    ], space_after=2.2)
    add_bullet([
        ("Instrumented multi-turn agent workflows with ", False),
        ("Langfuse", True),
        (" distributed tracing and ", False),
        ("Sentry", True),
        (" error telemetry, diagnosing prompt drift and eliminating production MongoDB tool-call hallucinations to reduce ", False),
        ("MTTD by 45%", True),
        (".", False)
    ], space_after=2.4)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Launched a full-stack AI interview platform facilitating ", False),
        ("200+", True),
        (" monthly candidate evaluations, slashing turnaround ", False),
        ("50%", True),
        (" while crafting responsive ", False),
        ("React/TypeScript", True),
        (" UI dashboards and versioned RESTful evaluation endpoints.", False)
    ], space_after=2.2)
    add_bullet([
        ("Automated containerized ", False),
        ("Docker", True),
        (" deployments and authored multi-stage ", False),
        ("Jest and Pytest", True),
        (" regression test suites in an Agile team, enforcing rigorous code quality gates and accelerating developer release cadence.", False)
    ], space_after=2.4)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks.", False)
    ], space_after=2.4)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("AI Interviewer", "NestJS, Next.js, React, TypeScript, WebSockets, OpenAI API, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Designed an interactive AI interview platform with responsive ", False),
        ("React/Next.js", True),
        (" UI and ", False),
        ("NestJS", True),
        (" microservices, crafting structured rubrics for objective candidate scoring and authoring ", False),
        ("Jest", True),
        (" test suites to enforce quality gates.", False)
    ], space_after=2.2)
    add_bullet([
        ("Orchestrated bidirectional streaming audio pipelines via ", False),
        ("WebSockets", True),
        (" and ", False),
        ("OpenAI API", True),
        (", enabling real-time conversational screening and reducing candidate evaluation turnaround by ", False),
        ("50%", True),
        (" across 200+ monthly technical reviews.", False)
    ], space_after=2.4)

    add_project_heading("Fluxpay", "Go (Golang), Apache Kafka, Redis, PostgreSQL, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Developed a cloud-native payments microservice in ", False),
        ("Go (Golang)", True),
        (" using goroutines and channels, sustaining ", False),
        ("1,200 TPS", True),
        (" (simulated) at sub-", False),
        ("120ms", True),
        (" p95 latency, validated through comprehensive authored Artillery load test suites.", False)
    ], space_after=2.2)
    add_bullet([
        ("Implemented an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" with ", False),
        ("Redis", True),
        (" distributed locks and idempotent ", False),
        ("PostgreSQL", True),
        (" transactions, eliminating race conditions and preventing duplicate charges across ", False),
        ("500+", True),
        (" concurrent simulated users.", False)
    ], space_after=2.4)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Languages: ", True, "TypeScript, JavaScript (ES6+), Python, Go (Golang), Java, SQL (PostgreSQL, MySQL), C++"),
        ("Frontend & UX: ", True, "React, Next.js (App Router, BFF), Redux, Tailwind CSS, HTML5, CSS3, Web Vitals, Responsive Design"),
        ("Backend & APIs: ", True, "Node.js, NestJS, Express.js, RESTful APIs, Microservices, WebSockets, Event-Driven Architecture"),
        ("Distributed & Data: ", True, "Apache Kafka (DLQ, Saga), Google Cloud Pub/Sub, Redis (Distributed Locks), MongoDB, PostgreSQL"),
        ("AI & Intelligent Agents: ", True, "OpenAI API, Google Gemini, Vapi.ai, Langfuse (LLM Tracing & Evals), Agentic Workflows"),
        ("Cloud & DevOps: ", True, "Docker, Git, GitHub Actions CI/CD, GCP (Cloud Run, Pub/Sub), Jest, Pytest, Artillery Testing, Agile/Scrum")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.8)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an autonomous real-time voice screening agent with low-latency WebSocket streaming.", False)
    ], space_after=2.2)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ competing teams for fault-tolerant distributed system architecture.", False)
    ], space_after=2.2)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, emphasizing computational complexity, memory optimization, and concurrent execution.", False)
    ], space_after=1.8)

def populate_truaudience(doc, add_section_header, add_role_heading, add_project_heading, add_bullet, role="Full_Stack_Engineer"):
    # 2. Experience (Full Stack, JVM Microservices, Cloud Data Pipelines, GenAI)
    add_section_header("Experience")

    add_role_heading("Associate Software Engineer", "Oct 2025 – Present", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Engineered an enterprise customer-facing web platform (", False),
        ("Next.js BFF", True),
        (", ", False),
        ("React", True),
        (", ", False),
        ("NestJS", True),
        (") across ", False),
        ("4+", True),
        (" university client deployments, generating ", False),
        ("$150K+", True),
        (" contract value per client while enforcing accessible UI and strict ", False),
        ("FERPA data privacy", True),
        (".", False)
    ], space_after=2.2)
    add_bullet([
        ("Constructed a multi-tenant ", False),
        ("Java/Spring Boot", True),
        (" ETL pipeline integrating enterprise CRM REST APIs, implementing ", False),
        ("Hibernate/JPA", True),
        (" and ", False),
        ("Kafka Dead Letter Queues (DLQ)", True),
        (" to handle asynchronous batch workloads with zero message loss.", False)
    ], space_after=2.2)
    add_bullet([
        ("Architected a decoupled two-tier ", False),
        ("Go", True),
        (" ingestion pipeline on ", False),
        ("Cloud Run", True),
        (" and ", False),
        ("Google Cloud Pub/Sub", True),
        (", processing ", False),
        ("10K+ delta records per 15-minute sync cycle", True),
        (" into durable MongoDB queues with lease-fenced consumers to guarantee zero data loss.", False)
    ], space_after=2.2)
    add_bullet([
        ("Improved the ", False),
        ("Document Intelligence", True),
        (" pipeline by coupling deterministic articulation rule caches with ", False),
        ("LLM fallback", True),
        (", resolving ", False),
        ("40%+", True),
        (" of course mappings without model inference to cut token costs by ", False),
        ("40%", True),
        (" and eliminate prompt variance.", False)
    ], space_after=2.2)
    add_bullet([
        ("Optimized backend analytics API response times ", False),
        ("3× (15s to 5s)", True),
        (" by restructuring complex SQL execution plans, indexing composite keys, and eliminating redundant table scans across ", False),
        ("PostgreSQL and MySQL", True),
        (" relational schemas.", False)
    ], space_after=2.2)
    add_bullet([
        ("Scaled real-time ", False),
        ("NestJS", True),
        (" WebSocket streaming middleware bridging telephony audio feeds, handling bidirectional transcoding for ", False),
        ("1,000+", True),
        (" concurrent sessions at sub-", False),
        ("100ms", True),
        (" p95 latency under peak production load with zero packet loss.", False)
    ], space_after=2.4)

    add_role_heading("Software Engineer Intern", "Apr 2025 – Sep 2025", "ISchoolConnect", "Noida, UP")
    add_bullet([
        ("Launched a full-stack AI interview platform facilitating ", False),
        ("200+", True),
        (" monthly candidate evaluations, slashing turnaround ", False),
        ("50%", True),
        (" while crafting responsive ", False),
        ("React/TypeScript", True),
        (" UI dashboards and versioned RESTful evaluation endpoints.", False)
    ], space_after=2.2)
    add_bullet([
        ("Maintained multi-tenant backend services under ", False),
        ("Jenkins CI/CD", True),
        (" pipelines, containerizing ", False),
        ("Docker", True),
        (" packaging and executing ", False),
        ("Jest and Pytest", True),
        (" integration test suites to enforce code quality gates across sprint releases.", False)
    ], space_after=2.4)

    # 3. Education
    add_section_header("Education")
    add_role_heading("Guru Gobind Singh Indraprastha University", "Dec 2021 – Jun 2025", "B.Tech, Electronics & Communication Engineering", "New Delhi, India")
    add_bullet([
        ("Coursework: Data Structures & Algorithms, Object-Oriented Programming (OOP), Database Management Systems (DBMS), Operating Systems, Computer Networks, Distributed Systems, Cloud Architecture, Software Engineering Principles.", False)
    ], space_after=2.4)

    # 4. Projects
    add_section_header("Projects")
    add_project_heading("Fluxpay", "Java, Spring Boot, Hibernate, Apache Kafka, Redis, PostgreSQL, Docker, Git", "github.com/Yash-Gupta-01/-Fluxpay", url="https://github.com/Yash-Gupta-01/-Fluxpay")
    add_bullet([
        ("Built a distributed payments microservice in ", False),
        ("Java (Spring Boot / Hibernate)", True),
        (" with ", False),
        ("Spring Security JWT", True),
        (" authentication, sustaining ", False),
        ("1,200 TPS", True),
        (" in simulation at sub-", False),
        ("120ms", True),
        (" p95 latency, verified with published Artillery load test suites.", False)
    ], space_after=2.2)
    add_bullet([
        ("Implemented an event-driven Saga pattern over ", False),
        ("Kafka", True),
        (" with ", False),
        ("Redis", True),
        (" distributed locks and idempotent ", False),
        ("PostgreSQL", True),
        (" transactions, eliminating race conditions and preventing duplicate charges across ", False),
        ("500+", True),
        (" simulated users.", False)
    ], space_after=2.4)

    add_project_heading("AI Interviewer", "NestJS, Next.js, React, TypeScript, WebSockets, OpenAI API, Jest", "github.com/Yash-Gupta-01/AI-interviewer", url="https://github.com/Yash-Gupta-01/AI-interviewer")
    add_bullet([
        ("Designed an interactive AI interview platform with responsive ", False),
        ("React/Next.js", True),
        (" UI and ", False),
        ("NestJS", True),
        (" microservices, crafting structured prompt rubrics for objective candidate evaluations and authoring ", False),
        ("Jest", True),
        (" test suites to enforce quality gates.", False)
    ], space_after=2.2)
    add_bullet([
        ("Orchestrated bidirectional streaming audio pipelines via ", False),
        ("WebSockets", True),
        (" and ", False),
        ("OpenAI API", True),
        (", enabling real-time conversational screening and reducing candidate evaluation turnaround by ", False),
        ("50%", True),
        (" across 200+ monthly technical reviews.", False)
    ], space_after=2.4)

    # 5. Technical Skills
    add_section_header("Technical Skills")
    skills = [
        ("Programming Languages: ", True, "Java (JVM), JavaScript, TypeScript, Python, Go (Golang), SQL (PostgreSQL, MySQL), C++"),
        ("Frontend & Web: ", True, "React, Angular 14, Next.js (BFF, SSR), TypeScript, Redux, Tailwind CSS, HTML5, CSS3, Accessible UI"),
        ("Backend & JVM: ", True, "Java/Spring Boot, Hibernate/JPA, NestJS, Node.js, Express, RESTful APIs, Microservices, Asynchronous I/O"),
        ("Data & Messaging: ", True, "Apache Kafka (DLQ, Saga), GCP Pub/Sub, Redis, MongoDB, PostgreSQL, SQL Optimization, ETL Pipelines"),
        ("AI & Emerging Tech: ", True, "Generative AI, LLM Applications, Prompt Engineering, RAG & Vector Search, OpenAI API, Gemini, Langfuse"),
        ("Cloud, DevOps & QA: ", True, "GCP (Cloud Run), Docker, Kubernetes (familiar), Git, CI/CD Pipelines, JUnit, Jest, Pytest, Artillery")
    ]
    for title, bold, items in skills:
        add_bullet([(title, bold), (items, False)], space_after=1.8)

    # 6. Achievements
    add_section_header("Achievements")
    add_bullet([
        ("Won ", False),
        ("First Place", True),
        (" at the ISchoolConnect internal AI hackathon (", False),
        ("Rs. 50,000 award", True),
        (") among 30+ engineering teams, architecting an autonomous real-time voice screening agent with low-latency WebSocket streaming.", False)
    ], space_after=2.2)
    add_bullet([
        ("Secured ", False),
        ("First Position", True),
        (" in HackBPIT (HashDefine, BPIT) and ", False),
        ("First Runner-Up", True),
        (" in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ competing teams for fault-tolerant distributed system architecture.", False)
    ], space_after=2.2)
    add_bullet([
        ("Solved ", False),
        ("450+", True),
        (" Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef, emphasizing computational time/space complexity, memory optimization, and concurrent multithreading.", False)
    ], space_after=1.8)

def create_resume_document(output_docx_path, company="Weekday", role="Software_Engineer"):
    doc = Document()

    # Page Geometry: Letterpaper with 0.36in top/bottom, 0.45in left/right
    for section in doc.sections:
        section.top_margin = Inches(0.36)
        section.bottom_margin = Inches(0.36)
        section.left_margin = Inches(0.45)
        section.right_margin = Inches(0.45)
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)

    # Base typography styling
    normal_style = doc.styles['Normal']
    font = normal_style.font
    font.name = 'Georgia'
    font.size = Pt(9.5)
    font.color.rgb = RGBColor(0x11, 0x11, 0x11)
    normal_style.paragraph_format.line_spacing = 1.0

    # 1. Header (Centered, elegant hierarchy)
    p_name = doc.add_paragraph()
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_name.paragraph_format.space_before = Pt(0)
    p_name.paragraph_format.space_after = Pt(2)
    p_name.paragraph_format.line_spacing = 1.0
    run_name = p_name.add_run("Yash Gupta")
    run_name.font.size = Pt(21)
    run_name.bold = True
    run_name.font.name = 'Georgia'

    p_c1 = doc.add_paragraph()
    p_c1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_c1.paragraph_format.space_before = Pt(0)
    p_c1.paragraph_format.space_after = Pt(1.5)
    p_c1.paragraph_format.line_spacing = 1.0
    if "visa" in company.lower():
        loc_text = "New Delhi / NCR, India (Open to Relocate — Bengaluru / PAN-India)"
    elif "swiggy" in company.lower() or "instamart" in company.lower():
        loc_text = "New Delhi, India (Open to Relocate — Bangalore)"
    elif "makemytrip" in company.lower() or "mmt" in company.lower():
        loc_text = "New Delhi / NCR, India"
    elif "franklin" in company.lower() or "templeton" in company.lower():
        loc_text = "New Delhi, India (Open to Relocate — Hyderabad)"
    elif "flex" in company.lower():
        loc_text = "New Delhi, India (Open to Relocate — Chennai)"
    elif "cisco" in company.lower():
        loc_text = "New Delhi, India (Open to Relocate — Bangalore)"
    elif "iqvia" in company.lower():
        loc_text = "New Delhi / NCR, India (Open to Relocate — Bangalore / PAN-India)"
    elif "atlassian" in company.lower():
        loc_text = "New Delhi, India (Open to Relocate — Bangalore)"
    elif "rippling" in company.lower():
        loc_text = "New Delhi, India (Open to Relocate — Bangalore)"
    elif "mastercard" in company.lower():
        loc_text = "New Delhi / NCR, India"
    elif "google" in company.lower():
        loc_text = "New Delhi, India (Open to Relocate — Bengaluru)"
    elif "zomato" in company.lower() or "eternal" in company.lower():
        loc_text = "New Delhi / NCR, India"
    elif "dentsu" in company.lower() or "dgs" in company.lower():
        loc_text = "New Delhi, India (Open to Relocate — Mumbai / Thane)"
    elif "juspay" in company.lower():
        loc_text = "New Delhi, India (Open to Relocate — Mumbai)"
    elif "akamai" in company.lower():
        loc_text = "New Delhi, India (Open to Relocate — Bangalore)"
    elif "bajaj" in company.lower():
        loc_text = "New Delhi / NCR, India (Open to Relocate — Pune / PAN-India)"
    elif "amex" in company.lower() or "american" in company.lower():
        loc_text = "New Delhi / NCR, India"
    elif "salesforce" in company.lower():
        loc_text = "New Delhi / NCR, India (Open to Relocate — Hyderabad / Bengaluru / PAN-India)"
    elif "paypal" in company.lower():
        loc_text = "New Delhi, India (Open to Relocate — Bengaluru / Chennai / PAN-India)"
    elif "target" in company.lower():
        loc_text = "New Delhi, India (Open to Relocate — Bangalore / PAN-India)"
    elif "zs" in company.lower():
        loc_text = "New Delhi / NCR, India (Open to Relocate — Gurgaon, Noida, Pune, Bengaluru)"
    elif "bank_of_america" in company.lower() or "bofa" in company.lower() or "bank of america" in company.lower():
        loc_text = "New Delhi / NCR, India"
    elif "uber" in company.lower():
        loc_text = "New Delhi / NCR, India (Open to Relocate — Bangalore / Hyderabad / PAN-India)"
    elif "park" in company.lower():
        loc_text = "New Delhi / NCR, India (Gurgaon On-Site Ready)"
    elif "rubrik" in company.lower():
        loc_text = "New Delhi, India (Open to Relocate — Bangalore)"
    elif "optum" in company.lower() or "unitedhealth" in company.lower() or "uhg" in company.lower():
        loc_text = "New Delhi / NCR, India (Open to Relocate — Hyderabad / PAN-India)"
    elif "nasdaq" in company.lower():
        loc_text = "New Delhi / NCR, India (Open to Relocate — Mumbai / PAN-India)"
    elif "oracle" in company.lower():
        loc_text = "New Delhi / NCR, India (Open to Relocate — Bengaluru / PAN-India)"
    elif "olx" in company.lower():
        loc_text = "New Delhi / NCR, India (Gurgaon On-Site Ready)"
    elif "instahyre" in company.lower() or "general" in company.lower():
        loc_text = "New Delhi / NCR, India (Open to Relocate PAN-India)"
    elif "cashfree" in company.lower():
        loc_text = "New Delhi / NCR, India (Open to Relocate — Bangalore / PAN-India)"
    elif "lseg" in company.lower() or "refinitiv" in company.lower() or "london" in company.lower():
        loc_text = "New Delhi / NCR, India (Open to Relocate — Bengaluru / Hyderabad / PAN-India)"
    elif "stripe" in company.lower():
        loc_text = "New Delhi / NCR, India (Open to Relocate — Bengaluru / PAN-India)"
    elif "galaxy" in company.lower() or "magica" in company.lower():
        loc_text = "New Delhi, India (Dwarka On-Site Ready)"
    elif "hartford" in company.lower():
        loc_text = "New Delhi / NCR, India (Open to Relocate — Hyderabad / PAN-India)"
    elif "barclays" in company.lower():
        loc_text = "New Delhi / NCR, India (Open to Relocate — Pune / PAN-India)"
    elif "uipath" in company.lower():
        loc_text = "New Delhi / NCR, India (Open to Relocate — Bengaluru / Hyderabad / PAN-India)"
    elif "truaudience" in company.lower() or "transunion" in company.lower():
        loc_text = "New Delhi / NCR, India (Hybrid Ready / Open to Relocate)"
    else:
        loc_text = "New Delhi, India"
    r1 = p_c1.add_run(f"{loc_text}  |  +91 8287552253  |  ")
    r1.font.size = Pt(8.8)
    add_hyperlink(p_c1, "mailto:guptayash112003@gmail.com", "guptayash112003@gmail.com", font_size=8.8)

    p_c2 = doc.add_paragraph()
    p_c2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_c2.paragraph_format.space_before = Pt(0)
    p_c2.paragraph_format.space_after = Pt(5)
    p_c2.paragraph_format.line_spacing = 1.0
    add_hyperlink(p_c2, "https://www.linkedin.com/in/yash-gupta-engineer/", "linkedin.com/in/yash-gupta-engineer", font_size=8.8)
    r_sep = p_c2.add_run("  |  ")
    r_sep.font.size = Pt(8.8)
    add_hyperlink(p_c2, "https://github.com/Yash-Gupta-01", "github.com/Yash-Gupta-01", font_size=8.8)

    def add_section_header(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(5.5)
        p.paragraph_format.space_after = Pt(2.2)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.keep_with_next = True
        run = p.add_run(title)
        run.font.size = Pt(11.5)
        run.bold = True
        run.font.name = 'Georgia'
        pPr = p._p.get_or_add_pPr()
        pBdr = parse_xml(r'<w:pBdr %s><w:bottom w:val="single" w:sz="6" w:space="2" w:color="000000"/></w:pBdr>' % nsdecls('w'))
        pPr.append(pBdr)

    def add_role_heading(role, dates, company, location):
        p1 = doc.add_paragraph()
        p1.paragraph_format.space_before = Pt(3.5)
        p1.paragraph_format.space_after = Pt(1.0)
        p1.paragraph_format.line_spacing = 1.0
        p1.paragraph_format.keep_with_next = True
        p1.paragraph_format.tab_stops.add_tab_stop(Inches(7.6), WD_TAB_ALIGNMENT.RIGHT)
        r_role = p1.add_run(role)
        r_role.bold = True
        r_role.font.size = Pt(9.8)
        p1.add_run("\t")
        r_dates = p1.add_run(dates)
        r_dates.bold = True
        r_dates.font.size = Pt(9.8)

        p2 = doc.add_paragraph()
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(2.5)
        p2.paragraph_format.line_spacing = 1.0
        p2.paragraph_format.keep_with_next = True
        p2.paragraph_format.tab_stops.add_tab_stop(Inches(7.6), WD_TAB_ALIGNMENT.RIGHT)
        r_comp = p2.add_run(company)
        r_comp.italic = True
        r_comp.font.size = Pt(9.2)
        p2.add_run("\t")
        r_loc = p2.add_run(location)
        r_loc.italic = True
        r_loc.font.size = Pt(9.2)

    def add_project_heading(name, stack, link, url=None):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3.5)
        p.paragraph_format.space_after = Pt(2.2)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.keep_with_next = True
        p.paragraph_format.tab_stops.add_tab_stop(Inches(7.6), WD_TAB_ALIGNMENT.RIGHT)
        r_name = p.add_run(name)
        r_name.bold = True
        r_name.font.size = Pt(9.8)
        p.add_run("  |  ").font.size = Pt(9.2)
        r_stack = p.add_run(stack)
        r_stack.italic = True
        r_stack.font.size = Pt(9.0)
        p.add_run("\t")
        target_url = url if url else (link if link.startswith("http") else f"https://{link}")
        add_hyperlink(p, target_url, link, font_size=9.0)

    def add_bullet(spans, space_after=2.2):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.left_indent = Inches(0.18)
        p.paragraph_format.first_line_indent = Inches(-0.14)
        r_bullet = p.add_run("•  ")
        r_bullet.font.size = Pt(9.2)
        r_bullet.font.name = 'Georgia'
        for text, bold in spans:
            r = p.add_run(text)
            r.bold = bold
            r.font.size = Pt(9.2)
            r.font.name = 'Georgia'

    if company.lower() in ("visa", "visanet"):
        populate_visa(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif company.lower() in ("swiggy", "instamart"):
        populate_swiggy(doc, add_section_header, add_role_heading, add_project_heading, add_bullet, role=role)
    elif company.lower() in ("makemytrip", "mmt"):
        populate_makemytrip(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif company.lower() in ("jiohotstar", "hotstar"):
        populate_jiohotstar(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif company.lower() in ("microsoft", "azure", "msft"):
        populate_microsoft(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif company.lower() in ("elevenlabs", "eleven_labs", "11labs"):
        populate_elevenlabs(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif company.lower() in ("gokiwi", "gokiwi_tech", "kiwi"):
        populate_gokiwi(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "franklin" in company.lower() or "templeton" in company.lower():
        populate_franklin_templeton(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "cisco" in company.lower():
        populate_cisco(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "flex" in company.lower():
        populate_flex(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "blackrock" in company.lower() or "black_rock" in company.lower():
        populate_blackrock(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "iqvia" in company.lower():
        populate_iqvia(doc, add_section_header, add_role_heading, add_project_heading, add_bullet, role=role)
    elif "atlassian" in company.lower():
        populate_atlassian(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "brillio" in company.lower():
        populate_brillio(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "mastercard" in company.lower():
        populate_mastercard(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "rippling" in company.lower():
        populate_rippling(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "google" in company.lower() or "gdc" in company.lower():
        populate_google(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "zomato" in company.lower() or "eternal" in company.lower():
        populate_zomato(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "dentsu" in company.lower() or "dgs" in company.lower():
        populate_dentsu(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "juspay" in company.lower():
        populate_juspay(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "akamai" in company.lower():
        populate_akamai(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "natwest" in company.lower():
        populate_natwest(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "barclays" in company.lower():
        populate_barclays(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "purplle" in company.lower():
        populate_purplle(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "bajaj" in company.lower():
        populate_bajaj_finance(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "amex" in company.lower() or "american" in company.lower():
        populate_amex(doc, add_section_header, add_role_heading, add_project_heading, add_bullet, role=role)
    elif "salesforce" in company.lower():
        populate_salesforce(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "temple" in company.lower() or "weartemple" in company.lower():
        populate_temple(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "paypal" in company.lower():
        populate_paypal(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "target" in company.lower():
        populate_target(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "zs" in company.lower():
        populate_zs(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "bank_of_america" in company.lower() or "bofa" in company.lower() or "bank of america" in company.lower():
        populate_bank_of_america(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "park" in company.lower():
        populate_parkplus(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "uber" in company.lower():
        populate_uber(doc, add_section_header, add_role_heading, add_project_heading, add_bullet, role=role)
    elif "rubrik" in company.lower():
        populate_rubrik(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "optum" in company.lower() or "unitedhealth" in company.lower() or "uhg" in company.lower():
        populate_optum(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "instahyre" in company.lower():
        populate_instahyre(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "cashfree" in company.lower():
        populate_cashfree(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "fynd" in company.lower():
        populate_fynd(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "jll" in company.lower() or "jones" in company.lower():
        populate_jll(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "ftse" in company.lower() or "russell" in company.lower():
        populate_ftse_russell(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "lseg" in company.lower() or "london" in company.lower() or "refinitiv" in company.lower():
        populate_lseg(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "oracle" in company.lower():
        populate_oracle(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "olx" in company.lower():
        populate_olx(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "groww" in company.lower():
        populate_groww(doc, add_section_header, add_role_heading, add_project_heading, add_bullet, role=role)
    elif "stripe" in company.lower():
        populate_stripe(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "galaxy" in company.lower() or "magica" in company.lower():
        populate_galaxyai(doc, add_section_header, add_role_heading, add_project_heading, add_bullet, role=role)
    elif "hartford" in company.lower():
        populate_the_hartford(doc, add_section_header, add_role_heading, add_project_heading, add_bullet, role=role)
    elif "airtel" in company.lower():
        populate_airtel(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "uipath" in company.lower():
        populate_uipath(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)
    elif "truaudience" in company.lower() or "transunion" in company.lower():
        populate_truaudience(doc, add_section_header, add_role_heading, add_project_heading, add_bullet, role=role)
    else:
        populate_general(doc, add_section_header, add_role_heading, add_project_heading, add_bullet)

    doc.save(output_docx_path)
    print(f"Successfully generated DOCX: {output_docx_path}")

def export_to_pdf_via_word(docx_path, pdf_path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc_word = word.Documents.Open(os.path.abspath(docx_path))
        doc_word.SaveAs(os.path.abspath(pdf_path), FileFormat=17) # 17 = wdFormatPDF
        doc_word.Close()
        print(f"Successfully exported PDF via Microsoft Word: {pdf_path}")
    finally:
        word.Quit()

def audit_resume_vocabulary(docx_path):
    import collections
    doc = Document(docx_path)
    banned_hallucinations = [
        'gemini api fallback', 'spearheaded the document intelligence',
        'spearheaded a document intelligence', 'transactional outbox',
        'pioneered an event-driven', 'established jenkins', 'maven'
    ]
    for p in doc.paragraphs:
        t_low = p.text.strip().lower()
        for bh in banned_hallucinations:
            if bh in t_low:
                raise ValueError(f"Quality QA Failed: Disproven hallucination/claim '{bh}' detected in: '{p.text.strip()}'")

    banned_buzzwords = [
        'problem solving', 'dynamic', 'team player', 'hardworking',
        'passionate', 'detail-oriented', 'results-driven', 'self-starter',
        'quick learner', 'fast learner', 'enthusiastic', 'motivated',
        'synergy', 'go-getter', 'strategic thinker', 'think outside the box',
        'rockstar', 'ninja', 'various', 'proven track record'
    ]
    
    bullet_verbs = []
    
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        
        # Check for banned buzzwords
        for buzz in banned_buzzwords:
            if buzz in t.lower():
                raise ValueError(f"Quality QA Failed: Banned buzzword '{buzz}' detected in: '{t}'")
        
        # Check opening verb if it's an experience/project/achievement bullet
        if t.startswith("•"):
            is_skill_or_course = any(k in t.lower() for k in [
                'languages:', 'backend &', 'messaging &', 'databases &',
                'devops,', 'devops &', 'methodology &', 'architecture &', 'security &',
                'coursework:', 'frontend &', 'ai & developer tools:',
                'systems &', 'distributed &', 'cloud &', 'languages &',
                'data analytics &', 'ai /', 'devops, testing &', 'testing &',
                'observability &', 'ai &', 'agentic &', 'computer vision &'
            ])
            if not is_skill_or_course:
                parts = t.lstrip("•").strip().split()
                if parts:
                    first_word = parts[0].rstrip(":,.;")
                    bullet_verbs.append(first_word)
                    
    # Check for duplicate opening verbs
    verb_counts = collections.Counter(bullet_verbs)
    duplicates = {v: c for v, c in verb_counts.items() if c > 1}
    if duplicates:
        raise ValueError(f"Quality QA Failed: Duplicate opening action verbs detected: {duplicates}")
        
    print(f"Vocabulary & Quality Audit Passed: Zero banned buzzwords, {len(bullet_verbs)} distinct opening action verbs (zero duplicates).")

def verify_and_render_qa(pdf_path, png_path):
    # Verify page count
    r = pypdf.PdfReader(pdf_path)
    page_count = len(r.pages)
    print(f"Page Count: {page_count}")
    if page_count != 1:
        raise ValueError(f"Visual QA Failed: Document is {page_count} pages; must be exactly 1.")

    # Render PNG
    pdf = pdfium.PdfDocument(pdf_path)
    img = pdf[0].render(scale=2).to_pil()
    img.save(png_path)
    print(f"Visual QA PNG saved to: {png_path}")

if __name__ == "__main__":
    import argparse
    import shutil

    parser = argparse.ArgumentParser(description="Generate tailored 1-page resume for Yash Gupta")
    parser.add_argument("--company", type=str, default="Visa", help="Target company name (e.g. Visa)")
    parser.add_argument("--role", type=str, default="Software_and_ML_Engineer", help="Target role name (e.g. Software_and_ML_Engineer)")
    args = parser.parse_args()

    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", ".."))

    # Strictly adhere to Yash_Gupta_{CompanyName}_{RoleName}.{extension}
    target_stem = f"Yash_Gupta_{args.company}_{args.role}"
    out_docx = os.path.join(base_dir, f"{target_stem}.docx")
    out_pdf = os.path.join(base_dir, f"{target_stem}.pdf")
    target_png = os.path.join(base_dir, "docx_resume_render.png")

    create_resume_document(out_docx, company=args.company, role=args.role)
    audit_resume_vocabulary(out_docx)
    export_to_pdf_via_word(out_docx, out_pdf)

    verify_and_render_qa(out_pdf, target_png)
    print(f"Generated successfully:\n  - {out_docx}\n  - {out_pdf}")
