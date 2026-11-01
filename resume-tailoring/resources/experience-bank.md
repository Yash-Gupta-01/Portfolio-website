# Candidate Master Experience Bank & Inventory: Yash Gupta

This file is the single canonical source of truth for Yash Gupta's career history, verified engineering metrics, architectural designs, and technical stack. Every detail here is backed by production deployments, verified Jira issue logs, and benchmarked repositories.

> **Non-Negotiable Honesty Constraints for Agents & Subagents:**
> - **Current Title**: Associate Software Engineer only — never Senior, Lead, or inflated.
> - **Employer**: ISchoolConnect (EDMO is a sister brand; both roles display as ISchoolConnect).
> - **CRM**: Slate CRM (Salesforce is only an ETL data source, not the primary CRM).
> - **Commercial GPA Platform**: Full-stack **Next.js BFF**, **NestJS**, **Slate CRM**, deployed across **4+** enterprise universities and generating **$150K+** average commercial contract value per client ($450K+ aggregate across Rowan, RFU, and PennWest; credit/package-based usage model). **NEVER mention Kafka in the GPA calculation.**
> - **Kafka Usage**: Kafka is documented strictly for the **Transfer Credit Evaluation (TCE)** pipeline, the **Java/Spring Boot ETL DLQ**, and the **Fluxpay** microservice.
> - **Go (Golang) Production Role**: Production Go code powers the high-throughput **Slate delta ingestion connector** on Google Cloud Run, pulling **10K+ delta records per 15-minute sync cycle** into MongoDB queues and decoupling processing via Pub/Sub and lease-fenced consumers.
> - **Document Intelligence (DI)**: Consolidated processing engine unifying GPA Calculator and Transfer Credit Evaluation (TCE); uses content-derived SHA/UUID hashing for single-pass OCR deduplication and a self-reinforcing deterministic articulation cache resolving **40%+** of course mappings without LLM calls.
> - **Angular**: Modern Angular 14 (core/CLI 14.1.0) — never legacy AngularJS v1.x.
> - **Kubernetes**: Listed as `Kubernetes (familiar)` only, without supporting project bullets.
> - **Cloud Certifications**: NONE held — never claim GCP/AWS certifications.
> - **Agent Frameworks**: Experience is with Google Gemini, OpenAI, Vapi.ai, and Langfuse (observability/tracing). Do not claim LangGraph/CrewAI unless explicitly instructed.
> - **Product Distinction**: **GPA evaluation** and **TCE (Transfer Credit Evaluation)** are distinct products at ISchoolConnect — never conflate them.
> - **Never invent or extrapolate** numbers or technologies. Frame personal projects strictly as personal/side projects.
> - **GPA Omission Mandate**: NEVER include candidate's college GPA (7.9/10.0) on the résumé. Display location (`New Delhi, India`) in the right-aligned column instead.
> - **Education Institution Mandate**: Strictly display candidate's university as **Guru Gobind Singh Indraprastha University** (degree-granting state university for global enterprise ATS recognition) — never use Bhagwan Parshuram Institute of Technology.
> - **Exact Project GitHub URLs Mandate**: Always use direct repository URLs with active hyperlinks:
>   - **Fluxpay**: `https://github.com/Yash-Gupta-01/-Fluxpay` (displayed as `github.com/Yash-Gupta-01/-Fluxpay`)
>   - **AI Interviewer**: `https://github.com/Yash-Gupta-01/AI-interviewer` (displayed as `github.com/Yash-Gupta-01/AI-interviewer`)
>   - **AI Architect Assistant**: `https://github.com/Yash-Gupta-01/AI-Architect-Assistant` (displayed as `github.com/Yash-Gupta-01/AI-Architect-Assistant`)
>   - **AI PDF Summarizer**: `https://github.com/Yash-Gupta-01/AI-PDF-Summarizer` (displayed as `github.com/Yash-Gupta-01/AI-PDF-Summarizer`)
>   - **Aether (Personal AI Agent / Second Brain)**: `https://github.com/Yash-Gupta-01/Personal-AI-Agent` (displayed as `github.com/Yash-Gupta-01/Personal-AI-Agent`)
>   - **Request Management System (RMS)**: `https://github.com/Yash-Gupta-01/Request-Management-System2` (displayed as `github.com/Yash-Gupta-01/Request-Management-System2`)
>   - **BOB Banking Application**: `https://github.com/Yash-Gupta-01/bob-banking-app` (displayed as `github.com/Yash-Gupta-01/bob-banking-app`)
>   - **Realtime Chat Platform**: `https://github.com/Yash-Gupta-01/discussion-app` (displayed as `github.com/Yash-Gupta-01/discussion-app`)
>   - **HackBPIT Disaster Relief Platform**: `https://github.com/Yash-Gupta-01/bpit_hackathon` (displayed as `github.com/Yash-Gupta-01/bpit_hackathon`)
>   - **NEVER** use the generic profile root `github.com/Yash-Gupta-01` as the project repository link.
>   - **Notice Period & Employment Status Mandate (Strict Rule)**: Yash is currently employed full-time as Associate Software Engineer at ISchoolConnect with an official notice period of **60 days**. NEVER write, state, or imply "available to start immediately", "immediate joiner", or zero-notice availability in any email, LinkedIn message, outreach note, resume, or recruiter communication. If asked about availability, state an official **60-day notice period** (or "serving a standard 60-day notice period").

---

## 1. Candidate Contact & Identity

- **Full Name**: Yash Gupta
- **Current Status**: Full-time Associate Software Engineer at ISchoolConnect (currently employed)
- **Notice Period**: Strictly **60 days** (official contractual notice period; NEVER claim or write "available to start immediately")
- **Target Role Level**: Early-career SWE / SDE-1 / Software Engineer / Backend Engineer (0–2 years experience band)
- **Current Location**: New Delhi / NCR, India (NCR-first / Gurugram GCC preference, open PAN-India)
- **Email**: guptayash112003@gmail.com
- **Phone**: +91 8287552253
- **LinkedIn**: [linkedin.com/in/yash-gupta-engineer/](https://www.linkedin.com/in/yash-gupta-engineer/)
- **GitHub**: [github.com/Yash-Gupta-01](https://github.com/Yash-Gupta-01)
- **File Name Convention**: Strictly `Yash_Gupta_{CompanyName}_{RoleName}.{extension}` (e.g. `Yash_Gupta_Visa_Software_and_ML_Engineer_Resume.docx` and `Yash_Gupta_Visa_Software_and_ML_Engineer_Resume.pdf`).

---

## 2. Technical Stack & Defensibility Matrix

### A. Primary / Interview-Defensible (Production & Deep Mastery)
*Technologies Yash has built production/complex systems with and can defend thoroughly in technical interviews:*
- **Languages**: TypeScript, JavaScript, Python, Go (Golang), Java (Spring Boot), C#, C++, SQL, Kotlin.
- **Backend Frameworks**: Node.js, NestJS, Express, Spring Boot (Java), Go standard library (`net/http`), ASP.NET Core 8.0, Flask (Python), Django (Python).
- **Databases & Caching**: PostgreSQL, MongoDB (NoSQL - complex indexing, bulk writes, aggregation, transactions), MySQL, SQL Server, SQLite, Redis (distributed locking, caching), Prisma ORM, Entity Framework Core.
- **Messaging & Streaming**: Google Cloud Pub/Sub, Apache Kafka (consumer/producer, Saga pattern, DLQ), WebSockets (real-time audio & telemetry, Socket.io).
- **Cloud & Infrastructure (GCP & AWS)**: Google Cloud Platform (Cloud Run, Cloud Run Jobs, Cloud Functions, Cloud Pub/Sub, Cloud Scheduler, Serverless VPC Access, Static NAT Egress IPs, MongoDB Atlas Network Peering/Allowlisting), Docker, AWS (S3, EC2), Jenkins CI/CD, Git, GitHub Actions.
- **Enterprise Integrations & CRM**: Slate CRM (Configurable Joins, Query Web Services, Source Formats XML/JSON, Slate Reader Custom Tabs, Communication Logs, CAS SSO ticket validation), Salesforce (ETL data source), Genesys Cloud (AudioHook v2 WebSocket protocol), Stripe Payments API.
- **Security & Architecture**: Backend-for-Frontend (BFF in Next.js), Clean Architecture (Android / JVM), Origin-Constrained iframe postMessage communication, DOM PII sanitization, Bounded transactional lease fencing, Idempotent revision-fenced writes, FERPA compliance.
- **Document Intelligence & AI**: Google Gemini API, OpenAI API, Vapi.ai, Langfuse (tracing/observability/evals), Document Intelligence OCR engine, content-derived file UUID hashing, compounding deterministic articulation rule caching, RAG, on-device inference (`llama.cpp`).
- **Frontend & Mobile**: React, Next.js (App Router, API Routes, SSR), Angular 14, Redux, Zustand, Tailwind CSS, Shadcn UI, Android SDK (Kotlin, Jetpack Compose, Room Database).
- **Testing & Observability**: Sentry (distributed tracing & alerting), Artillery (load testing), Pytest, Jest, JUnit, Postman.

### B. Working Familiarity (Secondary / Explicitly Tagged)
*Technologies touched in minor components or study — tag as `(familiar)`:*
- `Kubernetes (familiar)` (No project bullets; knowledge of pods, deployments, services primitives only).
- `Celery (familiar)`
- `FastAPI (familiar)`
- `Banner ERP (familiar)` (Integration boundary scoped for financial aid/matriculation).

### C. Explicit Gaps / Do NOT Claim
- Cloud certifications (AWS/GCP certified).
- Heavy Kafka administration/infra internals (experience is as application engineer: producer, consumer, DLQ, Saga pattern).
- LangGraph / CrewAI (uses Langfuse for tracing/evals and direct LLM API/orchestration).
- Legacy AngularJS v1.x (only modern Angular 14).

---

## 3. Professional Experience

### ISchoolConnect — Noida, India
**Timeline:**
- **Associate Software Engineer**: Oct 2025 – Present (Full-time, converted from internship)
- **Software Engineer Intern**: Apr 2025 – Sep 2025 *(First project was migrating a legacy NestJS Students API document downloading service; assigned from backlog ticket HLTB-213)*

**Core Thesis / Scope**:
Full-stack and backend systems engineering across higher-education enterprise AI platforms, mission-critical Slate CRM integrations, high-throughput Go data connectors, real-time telephony voice AI (Genesys AudioHook v2), Document Intelligence OCR/articulation engines, and multi-tenant ETL pipelines serving universities including Rowan University, Rosalind Franklin University (RFU), and Pennsylvania Western University (PennWest).

---

### Production Engineering Domains & Battle-Tested Achievements

#### Domain 1: Slate CRM Go Connector, Durable Queues & Distributed Ingestion
*High-throughput, distributed data pipelines written in Go and deployed on Google Cloud Run to ingest Slate CRM data.*

- **Two-Tier Decoupled Producer/Consumer Architecture (Go / Cloud Run / Pub/Sub / MongoDB)**:
  - *Context*: Replaced fragile, record-by-record ingestion with a resilient two-tier distributed pipeline in Go.
  - *Producer Scale & Cadence*: A scheduled **Cloud Run Job** triggers every **15 minutes** via Cloud Scheduler, extracting **10,000+ applicant and person delta records** across Slate Configurable Joins endpoints. It buffers raw deltas into a durable MongoDB queue collection with time-window export recovery and publishes payload-free **Google Cloud Pub/Sub** wake-up messages. Extraction and ingestion reliably finish within the 15-minute window.
  - *Consumer*: An authenticated **Cloud Run** consumer claims queued batches using atomic **MongoDB lease fencing** (`leaseExpiresAt`, lease renewal, bounded transactional batch claims), enriches student dossiers, and performs idempotent upserts into `applicantsData`.
  - *Reliability Invariants*: Added release-versus-failure error classification (distinguishing transient 401/403/429/5xx network timeouts from permanent schema rejections) to prevent poison-pill drops, ensuring zero data loss across consumer restarts or Slate rate-limit spikes.
  - *Draft Bullet*: Engineered a decoupled two-tier **Go** ingestion pipeline on **Cloud Run** and **Pub/Sub**, ingesting **10K+ delta records per 15-minute sync cycle** into a durable **MongoDB** queue with lease-fenced consumers, ensuring zero data loss during upstream outages.
  - *HAMS*: H: Go, Cloud Run, Pub/Sub, MongoDB lease fencing, Slate CRM; A: engineered decoupled ingestion pipeline; M: 10K+ delta records / 15-min cycle, zero data loss; S: high-throughput distributed pipeline.

- **Six Slate Configurable Joins Differential Bulk Exports**:
  - *Context*: Slate CRM provides Configurable Joins web services. Yash configured and exposed 6 dedicated differential JSON queries:
    1. `Persons`: Demographics, identity, Slate GUID, names, external reference IDs.
    2. `Devices`: Email addresses, phone numbers, mobile identifiers.
    3. `Applications`: Degree program, entry term, lifecycle status, stage.
    4. `Checklist`: Document submission requirements, statuses, fulfillment dates.
    5. `Activities`: Student interaction events, campus visits, status modifications.
    6. `Messages`: Communication log, sent/received emails, SMS history.
  - *Transformation & Out-of-Order Race Resolution*: Reworked the connector from N+1 record-by-record HTTP calls into bulk group-by-person processing. Modeled a nested person-rooted document structure in MongoDB (`persons` / `applications` / `interactions`).
  - *Race Recovery*: Handled out-of-order child records (e.g. checklist or device updates arriving before the person record) via pending-person stubs and subsequent reconciliation.
  - *Draft Bullet*: Architected bulk synchronization for Slate CRM across **6 Configurable Joins** differential feeds in **Go**, aggregating persons, applications, and communications into person-rooted **MongoDB** documents with pending-stub race resolution.
  - *HAMS*: H: Go, Slate Configurable Joins, MongoDB, data modeling; A: architected bulk synchronization; M: 6 differential feeds, person-rooted aggregation; S: multi-entity CRM synchronization.

- **MongoDB Throughput Optimization & Ordered Bulk Operations**:
  - *Context*: Ingestion bottlenecked under high write contention when processing thousands of applicant updates.
  - *Implementation*: Replaced single-document writes with ordered **MongoDB bulk operations** (`bulkWrite`), introduced batch-level mapping counters to reduce shared lock contention, and instrumented command/connection-pool timings with **Sentry** alerts.
  - *Networking & Infrastructure*: Deployed via automated CI/CD to **Google Cloud Run**, configuring Serverless VPC Access with static NAT egress IPs for secure allowlisting on MongoDB Atlas.
  - *Draft Bullet*: Optimized **MongoDB** write throughput during burst CRM synchronizations by implementing ordered **bulk operations** and lease-renewal batching, cutting write latency and eliminating database lock contention.
  - *HAMS*: H: MongoDB bulkWrite, Go, Cloud Run, VPC Access; A: optimized write throughput; M: ordered bulk operations, eliminated lock contention; S: database performance tuning.

- **Internal Python Data Tooling & PostgreSQL Analytics Extraction**:
  - *Context*: Operations and admissions counseling teams required clean, aggregated student application dossiers from the operational PostgreSQL database without manual spreadsheet wrangling or ad-hoc SQL exports.
  - *Implementation*: Authored internal Python scripts and batch data utilities to query PostgreSQL, normalize applicant records, and transform datasets for downstream reporting and analytics.
  - *Scale & Metric*: Processed **50K+ student application records**, eliminating manual operational data wrangling and accelerating internal reporting turnaround by **60%**.
  - *Draft Bullet*: Authored internal **Python scripts** to extract and transform **50K+ student application records** from **PostgreSQL**, eliminating manual operational reporting overhead and accelerating analytics turnaround by **60%**.
  - *HAMS*: H: Python, PostgreSQL, data extraction/transformation, internal reporting; A: authored internal Python scripts; M: 50K+ student records, 60% turnaround acceleration; S: internal data engineering & operational reporting.
  - *Defensibility & Boundary Invariant*: This was internal data extraction tooling and batch reporting, NOT a customer-facing or partner-facing core ingestion platform. Never rebrand this as external partner ingestion.

---

#### Domain 2: Student Copilot & Multi-Cluster Pipeline Architecture
*Extending enterprise CRM synchronization to real-time student-facing AI copilot clusters.*

- **Multi-Cluster Student Copilot Pipeline (Go / MongoDB / Pub/Sub)**:
  - *Architecture*: Evolved the operational Advisor Copilot Slate pipeline to feed a dedicated, isolated Student Copilot MongoDB cluster via feature-flagged delivery.
  - *Revision-Fenced Delivery*: Implemented pending source revision tracking, projecting exactly one document per application, and performing revision-fenced idempotent writes and soft deletes to keep student-facing portals synchronized without overwriting concurrent user updates.
  - *Admissions Identity & Integrity*: Resolved edge-case anomalies: eliminated ambiguous duplicate application reference IDs (ECI-2234), filtered out past-term records from current portal views, and decoupled intermediate payment statuses from final admissions decisions in LLM conversational responses (ECI-2233).
  - *Draft Bullet*: Scaled the **Go/Pub/Sub** sync engine to replicate Slate applicant data across an isolated **Student Copilot MongoDB cluster**, enforcing revision-fenced idempotent writes and sub-second soft-delete propagation.
  - *HAMS*: H: Go, Pub/Sub, MongoDB multi-cluster, idempotency; A: scaled sync engine; M: multi-cluster replication, revision-fenced delivery; S: distributed data replication.

- **Slate CAS (Central Authentication Service) & Student Identity Verification**:
  - *Implementation*: Implemented identity-gated access validating Slate CAS single sign-on tickets, ensuring strict exact-one-person record resolution before exposing authenticated student profiles or copilot conversational tools.
  - *Draft Bullet*: Integrated **Slate CAS** ticket validation and exact-one-person record matching in **NestJS**, securing student-facing AI assistants with zero-trust authenticated identity resolution.
  - *HAMS*: H: Slate CAS, SSO, NestJS, identity verification; A: integrated ticket validation; M: exact-one-person matching, zero-trust security; S: enterprise identity gating.

---

#### Domain 3: Embedded Slate Security, Next.js BFF & FERPA Privacy Compliance
*Enterprise security, cross-origin iframe hardening, and PII protection for Slate Reader integrations.*

- **Next.js Server-Side API Proxies (BFF Architecture)**:
  - *Problem*: Early prototype embedded widgets exposed internal backend endpoints and API credentials in the client browser bundle.
  - *Solution*: Re-architected the embedded GPA Calculator frontend in **Next.js** using a Backend-for-Frontend (**BFF**) pattern. Shifted all downstream API calls, authentication tokens, and service credentials into server-side Next.js route handlers shielded by route middleware.
  - *Draft Bullet*: Hardened embedded **Slate CRM** web applications by engineering a **Next.js BFF** proxy, eliminating client-side API credential exposure and enforcing server-side session authentication across university portals.
  - *HAMS*: H: Next.js BFF, API proxies, Slate CRM, middleware; A: hardened embedded web apps; M: eliminated credential exposure, server-side session control; S: frontend security architecture.

- **Origin-Constrained iframe postMessage & DOM PII Sanitization**:
  - *Cross-Origin Hardening*: Secured Slate Reader iframe embedding by restricting `postMessage` senders and receivers to explicitly allowlisted origins (strictly banning wildcard `*`), coupled with strict parent-origin validation.
  - *PII & Privacy Protection*:
    - Stripped student-identifying query parameters from iframe source URLs.
    - Completely removed client-side `localStorage` and insecure cookies holding applicant records.
    - Sanitized the rendered DOM to prevent student PII leakage across Reader tab navigation.
    - Resolved process collisions where distinct same-named transcripts across different applicants previously collided.
  - *Draft Bullet*: Enforced strict **FERPA compliance** across embedded Slate iframes by replacing wildcard messaging with origin-whitelisted **postMessage** protocols and purging student PII from DOM and client storage.
  - *HAMS*: H: FERPA, iframe postMessage, Next.js, origin validation, PII scrubbing; A: enforced compliance; M: origin-whitelisted protocol, zero PII leakage; S: browser application security.

---

#### Domain 4: Document Intelligence (DI), OCR Deduplication & Deterministic TCE Articulation
*Unifying document OCR and making academic course-to-course articulation repeatable, cost-effective, and compounding.*

- **Unified Document Intelligence (DI) Consolidation**:
  - *Architecture*: Merged the standalone GPA Calculator backend into **Document Intelligence (DI)**, creating a single shared service for both GPA calculation and Transfer Credit Evaluation (TCE).
  - *Lifecycle*: Consolidated session verification, organization configuration, process records, calculation routes, and document webhooks under the DI API, retiring fragmented legacy backends.
  - *Draft Bullet*: Consolidated fragmented GPA and credit evaluation services into a unified **Document Intelligence (DI)** platform, streamlining document webhooks and calculation pipelines into a single high-throughput backend.
  - *HAMS*: H: Document Intelligence, NestJS, microservice consolidation, webhooks; A: consolidated fragmented services; M: single shared platform, retired legacy backends; S: service consolidation.

- **Deterministic Content-Derived File UUIDs (OCR Deduplication Engine)**:
  - *Problem*: Re-uploading or reloading transcripts triggered expensive redundant OCR and LLM extraction passes.
  - *Solution*: Implemented deterministic content-derived UUID generation (content hashing) on uploaded files. Identical transcripts are instantly recognized across frontend and backend flows, reconciling uploads after page reloads and guaranteeing that OCR and extraction execute exactly once per distinct file.
  - *Draft Bullet*: Engineered a content-hashed **file UUID deduplication engine** across transcript processing pipelines, preventing duplicate OCR extraction runs across page reloads and eliminating redundant compute costs.
  - *HAMS*: H: Content hashing, UUID generation, OCR pipeline, document processing; A: engineered deduplication engine; M: 100% duplicate OCR prevention, zero redundant passes; S: compute optimization.

- **Self-Reinforcing Deterministic Articulation Caching (Pre-LLM Course Lookup)**:
  - *Problem*: Relying solely on LLMs for course-to-course transfer equivalency mapping produced non-deterministic results and racked up expensive API token costs for known courses.
  - *Solution*: Re-architected articulation mapping to execute a two-stage evaluation: first checking confirmed articulation rules, followed by an immediate database lookup of previously verified historical mappings (matching source school, target school, program, and course code).
  - *Compounding Architecture*: The system functions as a self-reinforcing flywheel: every human counselor confirmation permanently adds to the cache. Currently, **~40% of all incoming course equivalency queries are resolved deterministically** without calling the LLM, cutting AI token spend by **40%+** and guaranteeing 100% repeatable transfer decisions for previously processed courses.
  - *Draft Bullet*: Architected a self-reinforcing academic articulation engine coupling deterministic rule caches with **LLM fallback**, resolving **40%+** of course mappings deterministically to eliminate prompt variance and cut AI token spend by **40%**.
  - *HAMS*: H: Deterministic rule caching, LLM orchestration, course articulation, MongoDB/PostgreSQL; A: architected articulation engine; M: 40%+ deterministic resolution, 40% token cost reduction; S: cost-effective self-reinforcing AI design.

- **Enterprise University Deployments & Slate Writeback**:
  - *Deployments*: Successfully deployed across **Rowan University** (Rowan Undergrad & Rowan Grad production), **Rosalind Franklin University (RFU)**, and **PennWest University** (California, Clarion, Edinboro campuses).
  - *Commercial Value*: Generating **$150K+ average commercial contract value** per institution under a credit-metered package pricing model ($450K+ aggregate).
  - *Slate Writeback*: Engineered automated writeback of evaluated GPA, quality points, credits, and course-level grades directly into RFU and Rowan Slate instances via XML/JSON endpoints.
  - *Edge Cases*: Added full support for Sophia Learning non-traditional transcripts, percentage-to-4.0 scale conversions, and HIL (Human-in-the-Loop) faculty review workflows.
  - *Draft Bullet*: Deployed commercial transcript evaluation platforms across **4+** universities (**Rowan**, **RFU**, **PennWest**) generating **$150K+** average contract value per client, automating writeback of computed GPAs and credits directly into **Slate CRM**.
  - *HAMS*: H: Slate CRM writeback, Next.js, NestJS, XML/JSON integration; A: deployed evaluation platforms; M: 4+ universities, $150K+ avg contract value; S: enterprise SaaS deployment.

---

#### Domain 5: Real-Time Telephony, Voice AI & Genesys Cloud AudioHook
*Low-latency WebSocket streaming, audio codec transcoding, and voice agent orchestration under real-world peak load.*

- **Genesys Cloud AudioHook v2 WebSocket Bridge**:
  - *Protocol Engineering*: Designed and built low-latency WebSocket middleware implementing the **Genesys AudioHook v2** specification. Handled connection handshakes, heartbeat ping/pong keepalives, message sequencing, and codec negotiation.
  - *Real-World Peak Load*: Architected to handle real-world peak call center loads, sustaining **1,000+ concurrent audio sessions** with sub-**100ms** p95 latency.
  - *Audio Transcoding*: Orchestrated bidirectional audio streaming, converting between telephony **8 kHz µ-law (G.711u)** and **16 kHz linear PCM** required by AI voice agent engines (**Vapi.ai** / OpenAI Realtime).
  - *Session Isolation & Escalation*: Enforced strict call-session isolation, preventing audio packet cross-talk across concurrent calls, and implemented context-rich call summaries for smooth escalation to human contact center agents.
  - *Draft Bullet*: Engineered a real-time **Genesys AudioHook v2** WebSocket bridge in **NestJS**, sustaining **1,000+** concurrent telephony streams at sub-**100ms** p95 latency under real-world peak load via bidirectional **8 kHz µ-law to 16 kHz PCM** transcoding.
  - *HAMS*: H: Genesys AudioHook v2, WebSockets, NestJS, audio transcoding (µ-law/PCM); A: engineered real-time audio bridge; M: 1,000+ concurrent sessions under peak load, sub-100ms p95; S: real-time telephony engineering.

- **Langfuse LLM Observability & Root-Cause Debugging**:
  - *Implementation*: Integrated **Langfuse** distributed tracing across agentic workflows. Traced multi-turn tool-call executions and reasoning chains to identify and eliminate MongoDB query hallucinations and prompt drift in admissions chatbots.
  - *Draft Bullet*: Instrumented multi-turn agent workflows with **Langfuse** distributed tracing, isolating reasoning-chain bottlenecks and eliminating production **MongoDB** tool-call hallucinations.
  - *HAMS*: H: Langfuse, LLM observability, MongoDB, prompt engineering; A: instrumented agent workflows; M: isolated bottlenecks, eliminated hallucinations; S: AI observability & reliability.

---

#### Domain 6: Enterprise Billing, Usage Streaming & Email Document Extractor
*Self-serve Stripe checkout, stream-based reporting, and automated attachment processing.*

- **Stripe Checkout & Credit Balance System (Next.js / Stripe)**:
  - *Billing Platform*: Developed an enterprise self-serve billing module featuring credit balance summaries, dynamic package tier purchases, Stripe Checkout session redirects, and payment lifecycle webhook handlers with 100% payment reconciliation integrity.
  - *Resilience*: Implemented robust React error boundaries and fallback states to ensure graceful degradation during external payment gateway or network downtime.
  - *Draft Bullet*: Implemented an enterprise billing engine in **Next.js** with **Stripe Checkout** integration, supporting credit package purchases, resilient fallback error boundaries, and **100%** payment reconciliation integrity.
  - *HAMS*: H: Next.js, Stripe API, webhooks, error boundaries; A: implemented billing engine; M: 100% payment reconciliation integrity; S: fintech billing integration.

- **Stream-Based CSV Export for High-Volume Usage Logs**:
  - *Data Streaming*: Built a stream-based CSV export backend and admin frontend for massive usage logs, streaming chunks directly from the database to HTTP responses to prevent server memory bloat on large date-range queries.
  - *Draft Bullet*: Developed a stream-based **CSV export pipeline** in **Node.js/NestJS** for tenant usage analytics, enabling real-time streaming of large audit datasets with constant O(1) server memory consumption.
  - *HAMS*: H: Node.js, NestJS, stream-based CSV export, memory optimization; A: developed streaming export pipeline; M: constant O(1) memory, zero heap overflow; S: backend data streaming.

- **Email Document Extractor Reliability & Better-Match Algorithm**:
  - *Attachment Processing*: Maintained and stabilized the Email Document Extractor, an automated service polling institutional mailboxes to extract transcript attachments.
  - *Student Matching*: Built an interactive "Better Match" candidate resolution UI and scoring table to reconcile ambiguous transcript attachments to student profiles, with state persistence across browser tabs and reloads.
  - *Draft Bullet*: Stabilized an automated **Email Document Extractor** processing institutional mailbox attachments, implementing resilient state persistence and student candidate-matching algorithms.
  - *HAMS*: H: Email ingestion, document parsing, React, Next.js, state persistence; A: stabilized document extractor; M: resilient state persistence, resolved ambiguous attachments; S: document workflow automation.

---

#### Domain 7: Slate Native Messaging & Communication Logs
*Direct CRM communication pipelines without third-party email/SMS provider lock-in.*

- **Slate Native Email & SMS Import Pipelines**:
  - *Native Delivery*: Configured PennWest JSON import Source Formats and Communication Log mappings directly within Slate, eliminating dependencies on standalone SMS/email SaaS providers.
  - *Identifier Mapping*: Keyed all outbound communication logs directly to the student's unique Slate Person GUID. Supported Rowan advisor personalized email payloads and managed SMS/WhatsApp opt-out status flags.
  - *Draft Bullet*: Engineered native **Slate CRM** communication pipelines using custom JSON Source Formats and Slate GUID mappings, synchronizing outbound emails, SMS, and WhatsApp dispatch directly into institutional logs.
  - *HAMS*: H: Slate CRM, JSON Source Formats, communication logs, SMS/WhatsApp; A: engineered native communication pipelines; M: Slate GUID mapping, eliminated third-party provider lock-in; S: CRM-native messaging.

---

#### Domain 8: Agentic Interview Platform, University Admissions & Technical Web Engineering
*Agentic interview screening platform commercialized for university admissions, and technical SEO remediation.*

- **Agentic Interview Screening Platform & Commercial University Admissions (Next.js / Shadcn / Redux / MongoDB Atlas)**:
  - *Platform Evolution*: Built an end-to-end autonomous agentic interview screening platform with Next.js, Redux, Tailwind CSS, Shadcn UI, and MongoDB Atlas. Equipped with an interactive conversational agent, JWT authentication, and role-based access control, the system was designed to autonomously conduct multi-turn technical and behavioral interviews.
  - *Product Commercialization*: Repurposed and productized as a commercial admissions interview platform actively pitched and deployed for universities (such as **Morgan State University**), automating high-volume student screening.
  - *Draft Bullet*: Built an agentic interview screening platform in **Next.js** and **MongoDB Atlas** commercialized for university admissions (deployed for clients including **Morgan State**), automating multi-turn student applicant evaluations.
  - *HAMS*: H: Next.js, Redux, MongoDB Atlas, Shadcn, agentic AI; A: built interview platform; M: commercialized for universities (Morgan State), automated multi-turn evaluations; S: agentic AI product engineering.

- **Technical SEO, Crawlability & Web Performance**:
  - *Remediation*: Conducted extensive SEO optimization across iSchoolConnect/iSchoolPrep web platforms: eliminated invalid internal `nofollow` attributes, replaced soft-404 responses with true HTTP 404 status codes, structured crawlable paginated content, injected Strapi FAQ JSON-LD schemas, and reduced LCP by **25%** via asset optimizations.
  - *Draft Bullet*: Enhanced technical web architecture across corporate domains, structuring crawlable pagination, injecting **JSON-LD** schema markup, and optimizing asset delivery to reduce **LCP by 25%**.
  - *HAMS*: H: Technical SEO, JSON-LD, crawlability, Core Web Vitals, LCP; A: enhanced web architecture; M: 25% LCP reduction, eliminated soft-404s; S: frontend performance & SEO.

---

## 4. Notable Projects & Open Source (Verified GitHub Inventory)

### 1. AI Interviewer (Hackathon 1st Place Winner & University Admissions Commercial Product)
- **Nature**: Internal Hackathon Winner (1st Prize, ₹50,000 award) / Commercialized Admissions Product
- **Repository**: [github.com/Yash-Gupta-01/AI-interviewer](https://github.com/Yash-Gupta-01/AI-interviewer)
- **Stack**: NestJS, Next.js, MongoDB, Vapi.ai, WebSockets, OpenAI API, Tailwind CSS
- **Live / Scale**: Automated **200+** monthly candidate screening interviews; reduced assessment turnaround time by **50%**; commercialized for university admissions interviews (pitched/sold to institutions like **Morgan State University**).
- **Business Angle**: High-throughput screening automation delivering immediate objective evaluations to hiring managers and university admissions committees.
- **Technical Architecture Angle**: Low-latency voice interaction loop using WebSockets and real-time audio chunking, orchestrating STT, dynamic prompting, and structured evaluation rubrics.
- **Key Bullets**:
  - *S-Tier*: Won First Place (Rs. 50,000) among 30+ teams by engineering an automated voice screening platform with NestJS, Next.js, and Vapi.ai, conducting 200+ monthly interviews and commercialized for university admissions.
  - *Architecture*: Orchestrated real-time WebSocket audio streaming and dynamic prompt evaluation rubrics in NestJS, evaluating candidate technical proficiency with zero human interviewer bottleneck.

### 2. Fluxpay (Distributed Payments Microservice)
- **Nature**: Flagship Distributed Systems / Learning-by-Building Project
- **Repository**: [github.com/Yash-Gupta-01/-Fluxpay](https://github.com/Yash-Gupta-01/-Fluxpay)
- **Stack Variants**: Primary implementation in **Go (Golang)**; complementary reference implementations in **Java/Spring Boot**, **Node/Express**, and **Python/Flask**.
- **Performance & Scale**: Benchmarked at **1,200 TPS** with p95 response times under **120ms** via authored Artillery load tests.
- **Architecture Highlights**: Idempotent payment processing, distributed locking with **Redis**, Saga pattern pattern with **Kafka**, and ACID-compliant ledger updates in **PostgreSQL**.
- **Key Bullets**:
  - *Go Variant*: Developed a high-throughput, cloud-native payments microservice in Go (Golang) with goroutines and channels, sustaining 1,200 TPS (simulated) at sub-120ms p95 latency, validated via authored Artillery load tests.
  - *Java Variant*: Developed a distributed payment microservice in Java (Spring Boot / Hibernate) enforcing authentication via Spring Security and JWT, sustaining 1,200 TPS in simulation at sub-120ms p95 latency.
  - *Distributed Architecture*: Architected an event-driven Saga pattern pattern over Kafka and Redis distributed locks with PostgreSQL ACID validation, guaranteeing idempotent transactions and zero double-charges across 500+ simulated users.

### 3. Aether: Offline-First Android "Second Brain" AI Assistant (Personal AI Agent)
- **Nature**: Offline-First Edge AI & Mobile Architecture
- **Repository**: [github.com/Yash-Gupta-01/Personal-AI-Agent](https://github.com/Yash-Gupta-01/Personal-AI-Agent)
- **Stack**: Kotlin, Android SDK (Java 21), Jetpack Compose, Room Database (exported schema JSON migrations), Clean Architecture (`:app`, `:domain`, `:data`, `:llm`, `:tools`), on-device inference (`llama.cpp`), Speech-to-Text (`SpeechToText`).
- **Core Engineering Highlights**:
  - Implemented strict **Clean Architecture** where the `:domain` layer is pure JVM with zero Android dependencies (`import android.*` rejected at compile time via `ArchitectureTest`), allowing domain unit tests to run in milliseconds without an emulator or device.
  - Built an offline-first **capture loop**: `Voice/Text -> Transcript -> Memory (UNCLASSIFIED) -> Room DB -> Dashboard` with provenance recording (`MANUAL_VOICE` vs typed text for downstream reasoning trust scores).
  - Architected a modular tool registry where on-device capabilities (alarms, reminders, calendar) declare strict schemas without LLM direct access to Android APIs.
- **Key Bullet**:
  - *Architecture & Mobile*: Architected **Aether**, an offline-first Android personal AI assistant in **Kotlin** and **Jetpack Compose** using **Clean Architecture** and pure JVM domain layers, integrating **Room** persistence and on-device **llama.cpp** inference for local thought capture.

### 4. Request Management System (RMS)
- **Nature**: Full-Stack Enterprise Internal Tools & Workflow State Machine
- **Repository**: [github.com/Yash-Gupta-01/Request-Management-System2](https://github.com/Yash-Gupta-01/Request-Management-System2)
- **Stack**: **ASP.NET Core 8.0**, **C#**, **Entity Framework Core**, **SQL Server**, **Tailwind CSS**.
- **Core Engineering Highlights**:
  - Built a multi-role enterprise request tracking portal with authenticated guest and administrator interfaces in **ASP.NET Core 8.0 MVC / Web APIs**.
  - Engineered a state-machine lifecycle tracking model (`RequestDetails` and `RequestStatusHistory`) recording complete chronological audit trails across statuses (*Yet to Start*, *In Progress*, *About to Finish*, *Finished*).
  - Utilized Entity Framework Core with Code-First Migrations, LINQ queries, and relational schema constraints for high-integrity request management.
- **Key Bullet**:
  - *C# / .NET Enterprise*: Developed an enterprise Request Management System in **ASP.NET Core 8.0** and **C#** with **Entity Framework Core** and **SQL Server**, engineering role-based access control and state-machine audit trails across request lifecycles.

### 5. BOB Banking Web Application
- **Nature**: Full-Stack Financial Web Architecture & Automated CI/CD
- **Repository**: [github.com/Yash-Gupta-01/bob-banking-app](https://github.com/Yash-Gupta-01/bob-banking-app)
- **Stack**: **Python**, **Flask**, **SQLite / SQLAlchemy**, **Bootstrap 5**, **Pytest**, **GitHub Actions CI/CD**.
- **Core Engineering Highlights**:
  - Implemented a 3-tier financial web application with secure customer authentication, session management, and ACID-compliant account balance updates (deposits and withdrawals).
  - Designed clean modular separation across `app.py`, `auth.py`, `models.py`, and `routes.py` with database seed utilities.
  - Configured automated **GitHub Actions CI/CD** (`banking-app-ci.yml`) executing multi-stage **Pytest** unit and integration suites with code coverage validation.
- **Key Bullet**:
  - *Python / CI-CD*: Engineered a full-stack banking web service in **Python/Flask** and **SQLAlchemy**, enforcing transaction ACID validation and configuring automated **GitHub Actions CI/CD** pipelines with **Pytest** test suites.

### 6. AI Architect Assistant
- **Nature**: Personal Full-Stack Side Project
- **Repository**: [github.com/Yash-Gupta-01/AI-Architect-Assistant](https://github.com/Yash-Gupta-01/AI-Architect-Assistant)
- **Stack**: Next.js (App Router + API Routes), TypeScript, Google Gemini API, Google Nano Banana image diffusion model, Tailwind CSS.
- **Functionality**: Guided architectural requirement-gathering conversational assistant generating interactive structural recommendations and 2D floor-plan visualizations.
- **Key Bullet**: Built an AI architectural assistant in **Next.js** using **Google Gemini** and image diffusion models, orchestrating multi-turn requirement extraction to generate interactive floor-plan recommendations.

### 7. AI PDF Summarizer
- **Nature**: Personal AI / Document Intelligence Side Project
- **Repository**: [github.com/Yash-Gupta-01/AI-PDF-Summarizer](https://github.com/Yash-Gupta-01/AI-PDF-Summarizer)
- **Stack**: Python, FastEmbed, LangChain, Streamlit, Vector Database.
- **Functionality**: Intelligent PDF summarization and semantic Q&A extraction pipeline utilizing lightweight embeddings and semantic vector search.
- **Key Bullet**: Implemented a local document intelligence pipeline in **Python** using **FastEmbed** and **LangChain**, enabling sub-second semantic retrieval and contextual summarization over multi-page research documents.

### 8. Real-Time Chat & Discussion Platform
- **Nature**: Full-Stack Real-Time Communication Platform
- **Repository**: [github.com/Yash-Gupta-01/discussion-app](https://github.com/Yash-Gupta-01/discussion-app)
- **Stack**: **React**, **Node.js**, **Express**, **Socket.io**, **MongoDB**, **Zustand**, **Tailwind CSS**, **DaisyUI**, **Cloudinary**.
- **Core Engineering Highlights**:
  - Engineered bidirectional real-time messaging using **Socket.io** WebSockets with broadcast presence tracking (online/offline user statuses).
  - Implemented JWT authentication and authorization with secure cookie handling.
  - Managed lightweight client state using **Zustand** and integrated Cloudinary image upload pipelines.
- **Key Bullet**: Built a full-stack real-time communication platform using the **MERN** stack and **Socket.io**, implementing WebSocket event streaming for instant messaging, presence tracking, and **Zustand** state management.

### 9. Disaster Relief & Volunteer Coordination Platform (HackBPIT 1st Place Winner)
- **Nature**: Hackathon Winning Platform (1st Place, HackBPIT)
- **Repository**: [github.com/Yash-Gupta-01/bpit_hackathon](https://github.com/Yash-Gupta-01/bpit_hackathon)
- **Stack**: **Python**, **Django**, **SQLite**, **Bootstrap**, **HTML5/CSS3**.
- **Functionality**: Emergency response coordination platform enabling district-based volunteer matching and real-time relief tracking during disaster recovery operations.
- **Key Bullet**: Won First Place at HackBPIT by developing a disaster relief coordination platform in **Python/Django**, matching volunteers and resources across geographic districts with relational database indexing.

---

## 5. Education, Hackathons & Honors

### Education
- **Degree**: Bachelor of Technology (B.Tech) in Electronics & Communication Engineering
- **Institution**: Guru Gobind Singh Indraprastha University (Mandate: strictly display as "Guru Gobind Singh Indraprastha University" on resumes and job applications; BPIT is the affiliated campus)
- **Graduation**: Dec 2021 – Jun 2025
- **Location**: New Delhi, India
- **GPA Policy**: **NEVER INCLUDE GPA ON RÉSUMÉ** (Explicit candidate instruction: omit 7.9 GPA; display location `New Delhi, India` in right-aligned column instead).
- **Secondary Education**: Class XII: **88.8%** | Class X: **84.6%**
- **Core Technical Coursework**: Data Structures & Algorithms, Object-Oriented Programming, Database Management Systems, Computer Networks, Operating Systems.

### Hackathons, Competitions & Problem Solving
- **1st Place Winner**: ISchoolConnect Internal AI Hackathon (won **₹50,000** prize for AI Interviewer platform among 30+ teams).
- **1st Place Winner**: HackBPIT Hackathon (HashDefine, BPIT) — Disaster relief platform in Python/Django ([github.com/Yash-Gupta-01/bpit_hackathon](https://github.com/Yash-Gupta-01/bpit_hackathon)).
- **Runner-up**: CodeGeeks Hackathon (Delhi Technological University - DTU 2023).
- **Competitive Programming**: **450+** Data Structures & Algorithms problems solved across LeetCode, GeeksforGeeks, Codeforces, and CodeChef.

> **Mandatory Formatting Rule**: When generating resumes, Achievements MUST ALWAYS be split into strictly **3 separate bullets** (Bullet 1: ISchoolConnect AI Hackathon, Bullet 2: HackBPIT & DTU CodeGeeks, Bullet 3: 450+ DSA problems). NEVER collapse or merge them into 1 or 2 bullets. This ensures full-page vertical fill down to the bottom margin.

---

## 6. Strategic Tailoring Rules by Target Role

When tailoring the 1-page executive résumé for specific job descriptions:

1. **Backend / Distributed Systems (e.g., Go/Kafka/PBCs/Fintech - Razorpay, Pine Labs, Zeta, Cisco, Atlassian)**:
   - **Lead with**: Go (Golang), Cloud Run, Google Cloud Pub/Sub, MongoDB lease fencing, Kafka, PostgreSQL, Redis, Java/Spring Boot.
   - **Highlight**:
     - Slate Go Connector (10K+ delta records / 15-min sync cycle, decoupled Cloud Run producer/consumer, ordered bulk MongoDB writes, Pub/Sub wake-ups, lease fencing).
     - Fluxpay microservice (Go/goroutines, 1,200 TPS, sub-120ms p95, Saga pattern pattern).
     - Student Copilot multi-cluster replication pipeline with revision-fenced idempotent delivery.
   - **De-emphasize**: Frontend UI and technical SEO.

2. **Enterprise GCC Roles (e.g., American Express, Mastercard, Optum, Expedia, Barclays, Franklin Templeton)**:
   - **Lead with**: Java/Spring Boot, Node.js/NestJS, MongoDB bulk operations, Kafka DLQ, Sentry tracing, SQL optimization (3× latency reduction), and enterprise CRM integration.
   - **Highlight**:
     - Enterprise stability, lease-fenced distributed consumers, and zero data loss retry policies.
     - Multi-tenant data pipelines, 100% payment reconciliation in Stripe, and FERPA-compliant BFF proxies.
     - Fluxpay Java/Spring Boot implementation with Spring Security and JWT.

3. **C# / .NET Enterprise Roles (e.g., Microsoft, Finastra, EY, Deloitte, Invesco)**:
   - **Lead with**: C#, ASP.NET Core 8.0, Entity Framework Core, SQL Server, RESTful APIs, OOP / SOLID principles.
   - **Highlight**:
     - Request Management System (ASP.NET Core 8.0, C#, EF Core code-first migrations, SQL Server, state machine lifecycle audit tracking).
     - Full-stack integration with Tailwind CSS and clean MVC architecture.

4. **Python Backend & Microservices Roles (e.g., Zeta, Swiggy, Zomato, CRED, FinTech startups)**:
   - **Lead with**: Python, Flask, Django, PostgreSQL, SQLite, Pytest, Docker, GitHub Actions CI/CD.
   - **Highlight**:
     - BOB Banking App (Python, Flask, SQLAlchemy, automated Pytest suites, GitHub Actions CI/CD).
     - Disaster Relief Volunteer Platform (Python/Django, relational indexing).
     - Document Intelligence & AI pipelines (Python, FastEmbed, LangChain).

5. **Full-Stack / Web Engineering (e.g., Google Web App Engineer, Next.js/React roles, Rippling, Swiggy)**:
   - **Lead with**: Next.js (BFF, SSR, API routes), React, Angular 14, TypeScript, Tailwind CSS, Shadcn UI, Redux.
   - **Highlight**:
     - Embedded Slate Reader security architecture (server-side Next.js proxies, origin-constrained `postMessage`, DOM PII sanitization).
     - Self-serve Stripe billing dashboard, stream-based CSV exports with O(1) memory, and AI Interviewer / Admissions screening platform.
     - Realtime Chat App (Socket.io WebSockets, Zustand state management).
     - Core Web Vitals optimization (-25% LCP).

6. **GenAI / LLM, Voice & Mobile Edge AI Roles (e.g., AI Labs, Conversational AI, Voice Agent roles)**:
   - **Lead with**: Real-time telephony voice AI (Genesys AudioHook v2, WebSockets, 8 kHz µ-law to 16 kHz PCM transcoding, sub-100ms p95 under peak load), Langfuse tracing, Document Intelligence OCR deduplication, on-device AI (Kotlin, `llama.cpp`), and RAG.
   - **Highlight**:
     - Self-reinforcing deterministic course articulation caching (40%+ resolution rate compounding over time, cutting token costs by 40%+).
     - AI Interviewer voice screening platform (₹50k hackathon win, 200+ monthly interviews, commercialized for university admissions like Morgan State).
     - Aether (Offline-first Android second-brain assistant in Kotlin with Clean Architecture and on-device `llama.cpp` inference).
     - Elimination of MongoDB tool-call hallucinations and prompt drift via Langfuse tracing.
