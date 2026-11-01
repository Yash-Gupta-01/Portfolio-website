---
name: resume-tailoring
description: Tailor and build an executive-grade 1-page résumé (strictly DOCX and PDF, no LaTeX) that maximizes both ATS score and human-recruiter attention for a specific job description. Reuses battle-tested bullet points from past resumes (bullet-vault.md) to ensure proven phrasing, verified metrics, and flawless formatting.
---

# Résumé Tailoring (1-Page Executive Format — DOCX & PDF Only)

Build and tailor a high-precision, executive-grade résumé that accomplishes two goals simultaneously:
1. **Passes automated ATS checks** (100% extractable text, high keyword match, standardized headers).
2. **Captivates human recruiters and hiring managers** in a ~10-second scan (HR-first section layout, bolded metrics, zero widow lines, balanced full-page vertical fill).

> **MANDATE: FORMAT & TOOLCHAIN**
> - **Outputs**: Strictly generate **`.docx`** (Word OpenXML) and **`.pdf`** (via native Word export).
> - **Exact File Count & Naming Contract**: Strictly produce **exactly two files** per generation: `Yash_Gupta_{CompanyName}_{RoleName}.docx` and `Yash_Gupta_{CompanyName}_{RoleName}.pdf`. NEVER generate alias copies, duplicate mirrors (e.g. `_Resume`), intermediate files, or fallback duplicates.
> - **NO LaTeX**: Do NOT generate or compile `.tex` files moving forward.
> - **Bullet Reuse Protocol**: Always check the **Bullet Vault** (`resources/bullet-vault.md`) from past resumes first. If an existing battle-tested bullet provides the required fit and value for the target JD, **REUSE IT DIRECTLY** rather than writing new phrasing from scratch.

---

## 1. Ground Truth & Honesty Constraints

Always consult the candidate's canonical experience bank and bullet vault:
👉 [resources/experience-bank.md](./resources/experience-bank.md)
👉 [resources/bullet-vault.md](./resources/bullet-vault.md)

### Non-Negotiable Honesty Rules:
- **Document Intelligence & AI Fallback Mandate (Strict Truthfulness & LLD Scope)**:
  - In production at ISchoolConnect, the Document Intelligence pipeline used **deterministic articulation rule caches with LLM fallback** (resolving 40%+ of course mappings without model inference to cut token costs by 40%).
  - **STRICTLY BAN "Gemini API" in Document Intelligence**: Production Document Intelligence NEVER used Gemini API. Gemini API was used solely in a personal side project (*AI Architect Assistant*). Conflating Gemini API into production Document Intelligence is a hallucination. Always write **`LLM fallback`** or **`model inference`**.
  - **STRICTLY BAN "Spearheaded" for Document Intelligence**: Yash improved the deterministic caching and fallback subsystem; claiming he "spearheaded" the platform falsely implies he designed the complete LLD of a massive enterprise system. Always write **`Improved`** (*"Improved the Document Intelligence pipeline..."*).
- **Fluxpay Architecture Mandate (Strict Saga Pattern Rule)**:
  - Fluxpay implemented the **Saga pattern** (coordinating distributed transactions across microservices with compensating steps over Kafka, Redis distributed locks, and PostgreSQL ACID transactions).
  - **STRICTLY BAN "transactional outbox pattern"**: NEVER claim transactional outbox pattern for Fluxpay. Always state **`Saga pattern`** and **`compensating Saga steps`**.
  - **STRICTLY BAN "Pioneered"**: Never write *"Pioneered an event-driven..."*. Always use **`Implemented`** or **`Developed`**.
- **Build Tool & Dependency Purge (Strict Zero-Maven Rule)**:
  - Yash has **NEVER used Maven**.
  - **STRICTLY BAN "Maven"**: NEVER mention "Maven" anywhere on the resume (not in Fluxpay, not in internships, not in skills, not in Java tooling).
- **CI/CD & DevOps Ownership Mandate (Strict Team Collaboration Rule)**:
  - Yash did **NOT** establish the company's Jenkins CI/CD pipelines from scratch.
  - **STRICTLY BAN "Established Jenkins CI/CD"**: Yash worked within and maintained backend services under existing team Jenkins CI/CD and Gitflow peer review workflows. Always write **`Maintained`** or **`Collaborated`** (*"Maintained multi-tenant backend services under Jenkins CI/CD and Gitflow peer review workflows..."*).
- **Inflated "Power Verbs" Ban**:
  - Strictly ban `Spearheaded`, `Pioneered`, and `Established` for early-career ownership claims. Use grounded, truthful engineering verbs: *Improved, Engineered, Implemented, Developed, Scaled, Deployed, Maintained, Automated, Built, Orchestrated*.
- **Title**: Strictly **Associate Software Engineer** — never Senior, Lead, or inflated.
- **Employer**: **ISchoolConnect** (sister brand EDMO displays as ISchoolConnect).
- **Commercial GPA Platform**: Full-stack **Next.js BFF**, **NestJS**, **Slate CRM**, deployed across **4+** enterprise universities and generating **$150K+** in commercial contract value per client. **NEVER mention Kafka in the GPA calculation.**
- **Kafka Usage**: Kafka is documented strictly for the **Transfer Credit Evaluation (TCE)** pipeline, the **Java/Spring Boot ETL DLQ**, and the **Fluxpay** microservice.
- **Kubernetes**: Listed as `Kubernetes (familiar)` only, without supporting project bullets.
- **Certifications**: Zero cloud certifications held (never claim AWS/GCP certified).
- **Never fabricate**: If the JD requires a technology the candidate lacks, openly flag the gap rather than inventing it.
- **GPA Omission Mandate**: NEVER include candidate's GPA (7.9/10.0) anywhere on the résumé. Display location (`New Delhi, India`) in the right-aligned column of Education instead.
- **Education Institution Mandate**: Strictly display candidate's university as **Guru Gobind Singh Indraprastha University** (degree-granting state university for global enterprise ATS recognition) — never use Bhagwan Parshuram Institute of Technology.
- **Verified Project GitHub URLs Mandate**: Always use direct repository URLs and embed active clickable hyperlinks in Word and PDF:
  - **Fluxpay**: `https://github.com/Yash-Gupta-01/-Fluxpay` (displayed as `github.com/Yash-Gupta-01/-Fluxpay`)
  - **AI Interviewer**: `https://github.com/Yash-Gupta-01/AI-interviewer` (displayed as `github.com/Yash-Gupta-01/AI-interviewer`)
  - **AI Architect Assistant**: `https://github.com/Yash-Gupta-01/AI-Architect-Assistant` (displayed as `github.com/Yash-Gupta-01/AI-Architect-Assistant`)
  - **AI PDF Summarizer**: `https://github.com/Yash-Gupta-01/AI-PDF-Summarizer` (displayed as `github.com/Yash-Gupta-01/AI-PDF-Summarizer`)
  - **Aether (Personal AI Agent / Second Brain)**: `https://github.com/Yash-Gupta-01/Personal-AI-Agent` (displayed as `github.com/Yash-Gupta-01/Personal-AI-Agent`)
  - **Request Management System (RMS)**: `https://github.com/Yash-Gupta-01/Request-Management-System2` (displayed as `github.com/Yash-Gupta-01/Request-Management-System2`)
  - **BOB Banking Application**: `https://github.com/Yash-Gupta-01/bob-banking-app` (displayed as `github.com/Yash-Gupta-01/bob-banking-app`)
  - **Realtime Chat & Discussion Platform**: `https://github.com/Yash-Gupta-01/discussion-app` (displayed as `github.com/Yash-Gupta-01/discussion-app`)
  - **HackBPIT Disaster Relief Platform**: `https://github.com/Yash-Gupta-01/bpit_hackathon` (displayed as `github.com/Yash-Gupta-01/bpit_hackathon`)
  - NEVER use the generic profile root `github.com/Yash-Gupta-01` as a project repository link.
- **Jenkins & LCP Separation Mandate (Strict Non-Conflation Rule)**: NEVER conflate Jenkins CI/CD pipelines with Largest Contentful Paint (LCP) or Core Web Vitals. Jenkins is strictly for automated test execution, containerized Docker packaging, and staging release cadence. LCP optimization is strictly frontend web performance (CDN caching, eliminating render-blocking scripts, asset minification). Conflating them into "maintaining Jenkins pipelines cutting LCP" is an absurd technical error that destroys engineering credibility in interviews.
- **Ask Before Writing (Domain Inquiries Authorized)**: If the target JD demands experience in a specific technology, system, or domain where knowledge is not currently present in `experience-bank.md` or `bullet-vault.md`, the agent is **explicitly authorized and encouraged to ask the candidate clarifying questions first** before generating bullets. Confirm whether they have real, defensible hands-on experience, a side project, or working familiarity to draw from, or if it is an authentic gap. Never assume or guess.
- **Action Verb & Vocabulary Diversity Mandate (Strict Zero-Repetition Rule)**:
  - **No Duplicate Opening Verbs**: Never reuse the same opening action verb across bullets anywhere on the resume (e.g. NEVER use "Engineered" multiple times; strictly diversify across *Architected, Deployed, Built, Scaled, Implemented, Spearheaded, Optimized, Established, Containerized, Maintained*).
  - **No Repeated Technical Descriptors**: Never repeat generic technical descriptors or buzzwords across bullets (e.g. strictly at most 1 occurrence of "automated/automating", "distributed", "streamlined", "high-throughput").
  - **No Intra-Bullet Redundancy**: Never repeat the same root word within the same bullet (e.g. never write "automating deployments and automated test suites").
  - **Mandatory Deduplication Pass**: After selecting bullets from the Bullet Vault, always perform a mandatory vocabulary deduplication pass to ensure no single verb or modifier repeats across the document.
- **Strict Cliché & Vague Buzzword Ban**:
  - **Never use vague subjective buzzwords**: Strictly ban terms flagged by recruiters and ATS checkers as empty fillers: *"problem solving"*, *"dynamic"*, *"team player"*, *"fast learner" / "quick learner"*, *"passionate"*, *"detail-oriented"*, *"results-driven"*, *"strategic thinker"*, *"go-getter"*, *"synergy"*.
  - **Replace with Evidence & Hard Engineering**: Replace subjective claims with specific technical implementations, architectural decisions, and verified metrics (e.g., replace *"concurrent problem solving"* with *"memory optimization and concurrent execution"*; replace *"dynamic rubric scoring"* with *"structured evaluation rubrics"*).

---

## 2. Layout & Visual Standards

### Section Hierarchy (HR-First Layout):
Recruiters and non-technical HR screeners review experience first. Never put Technical Skills on top.
1. **Header** (Name, Location, Phone, Email, LinkedIn, GitHub — with active hyperlinks)
2. **Experience (TOP)** (Current Associate Software Engineer role, followed by Intern sub-role)
3. **Education** (Degree, College, Location, Coursework — NO GPA)
4. **Projects** (High-signal builds: Fluxpay, AI Interviewer — with verified direct repo hyperlinks)
5. **Technical Skills (BOTTOM)** (Categorized index: Languages, Distributed Systems, Databases, Testing, Cloud, Fundamentals, Frontend)
6. **Achievements** (Hackathon wins, CP ratings)

### Typography & Spacing:
- **Typography**: **Georgia** (native Word OpenXML, 9.5pt body, 21pt title, 11.5pt section headings).
- **Margins**: 0.45in top/bottom, 0.5in left/right (Letterpaper).
- **Page Fill**: Content must fill the entire page vertically down to ~0.45in from the bottom edge. Never leave a 1-to-2 inch empty void at the bottom.
- **Line-Fitting Law**: Every wrapping bullet must occupy **85–95% of its second line**. Never allow a trailing "widow stub" of 1–3 words.

### Mandatory Achievements Structure (Full Vertical Page Fill Mandate):
- **ALWAYS split Achievements into strictly 3 separate, distinct bullets** (NEVER combine the hackathons into a single bullet):
  1. **ISchoolConnect Internal AI Hackathon**: Won First Place (Rs. 50,000 award) among 30+ engineering teams, architecting an automated real-time voice screening platform.
  2. **Inter-College Hackathons**: Secured First Position in HackBPIT (HashDefine, BPIT) and First Runner-Up in CodeGeeks Hackathon (DTU, 2023), recognized among 100+ teams for distributed system design.
  3. **Competitive Programming / DSA**: Solved 450+ Data Structures & Algorithms problems across LeetCode, GeeksforGeeks, Codeforces, and CodeChef.
- Breaking Achievements into 3 bullets is essential to achieve a rich, full-page vertical fill without leaving empty void at the bottom, while fitting cleanly on strictly 1 page.

---

## 3. Workflow & Past Resumes Indexing

### Step 1: Index Past Resumes (Whenever new resumes are added)
Whenever you drop resumes into the `past_resumes/` folder (or root directory):
```powershell
python .agents/skills/resume-tailoring/scripts/index_past_resumes.py
```
This automatically extracts, deduplicates, and categorizes all bullet variations into [resources/bullet-vault.md](./resources/bullet-vault.md).

### Step 2: Tailor for a Target JD (Clarify Gaps & Reuse First)
1. Ingest the target Job Description (JD) and identify its core thesis, required technologies, and experience band.
2. **Clarify Missing Domain Knowledge**: If the JD emphasizes a technology, tool, or domain (e.g. GraphQL, specific AWS services, gRPC, Celery) where experience is not recorded in `experience-bank.md`, **ask the candidate directly** before generating bullets. Confirm if they have real, defensible experience with it or if it is an authentic gap.
3. **Query the Bullet Vault**: Search [resources/bullet-vault.md](./resources/bullet-vault.md) to select battle-tested S-Tier and A-Tier bullets matching the JD keywords.
4. **Draft / Integrate**: Reuse proven phrasing from the vault directly. Only craft new phrasing if an authentic technical requirement or verified new detail has no existing bullet.
5. **Deduplication & Vocabulary Variety Pass**: Run a mandatory global scan across all selected bullets to eliminate repetitive opening action verbs (e.g. NEVER repeat "Engineered") and repetitive modifiers (e.g. "automated", "distributed"). Diversify synonyms to keep executive impact high.

### Step 3: Build & Export (DOCX & PDF)
Run the generator script with target company and role:
```powershell
python .agents/skills/resume-tailoring/scripts/build_resume.py --company <CompanyName> --role <RoleName>
```
This script:
1. Generates `Yash_Gupta_<CompanyName>_<RoleName>_Resume.docx` (strictly adhering to `Yash_Gupta_{CompanyName}_{RoleName}.{extension}`) with Georgia typography, OpenXML bottom borders, and right-aligned tab stops (7.5in).
2. **Automated Vocabulary & Cliché Audit**: Programmatically verifies zero banned buzzwords (*problem solving, dynamic, team player*, etc.) and asserts 100% unique opening action verbs across bullets (failing the build if any duplicate verb or cliché is detected).
3. Automates Microsoft Word via COM to export `Yash_Gupta_<CompanyName>_<RoleName>_Resume.pdf`.
4. Renders a 300-DPI visual preview image (`docx_resume_render.png`) using `pypdfium2`.
5. Asserts that page count is **strictly 1**.

### Step 4: Deliver
Deliver both the `.docx` and `.pdf` files directly with clickable links.
