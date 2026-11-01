---
name: email-outreach
description: Draft cold and warm job-search emails for Yash Gupta's switch from ISchoolConnect to a product-based company or MNC (NCR-first, 10+ LPA) — recruiter cold emails, hiring-manager outreach, referral-routed application emails, and follow-ups. Use whenever Yash wants to email someone about a role — "write a cold email to this recruiter," "email the hiring manager for X," "draft a referral email for the Google req," "follow up on my application," or when he pastes a job post / recruiter contact and asks what to send. Trigger even without the word "skill." Prefer this over generic email drafting because it enforces his real profile, his honesty constraints, and a channel strategy tuned to his targets.
---

# Email Outreach (recruiter / hiring-manager / referral emails)

A cold job email has one job: in the time it takes to read the first two lines, make a busy recruiter or engineer decide you're worth a reply. It earns that with a sharp subject, a one-line tldr, a short paragraph of *specific shipped work*, and one clear ask. Length is the enemy; specificity is the lever.

## Who this is for (do not invent a different persona)

Yash is an **early-career SWE (~1 year)**, **Associate Software Engineer at ISchoolConnect** (Noida), **B.Tech ECE, BPIT, 2025**. He's making a **lateral move to a PBC/MNC**, NCR-first, **10+ LPA**. His narrative is **production backend + real-time/agentic AI**: NestJS, Node.js, Kafka, Redis, WebSocket, payments, voice-AI orchestration. He is **not** a student or new-grad intern-seeker, and **not** a senior engineer.

## The anti-pattern (read this before writing anything)

There is a well-known cold email — the "Soham Parekh" email — that is structurally excellent and substantively a cautionary tale. It opens with a crisp tldr, names impressive companies, and claims sweeping full-stack range across early teams. It became infamous because the breadth was inflated and the candidate was later exposed for overstating and overcommitting. **Borrow its skeleton; never its substance.**

Concretely, that means:
- **No invented tenure.** Yash has ~1 year of professional experience. Never write "5 years" or imply more.
- **No fabricated breadth.** Don't list technologies he hasn't used to pad range. **gRPC is a genuine gap. Kubernetes is skills-list only (no deep project). FastAPI/Celery are "familiar," not primary.** If an email needs one of these to fit, flag the gap — don't fake it.
- **No company-name inflation.** Employer is **ISchoolConnect** (never EDMO). CRM is **Slate** (not Salesforce). Angular is **14** (not AngularJS). Personal projects (**Fluxpay, AI Architect Assistant**) are labeled personal — never passed off as employment.
- **Notice Period (Strict Mandate).** Yash is currently working full-time with a contractual notice period of **60 days**. NEVER write or imply "available to start immediately", "immediate joiner", or 0-day start in any email or outreach. If availability is mentioned, strictly state: "serving a standard 60-day notice period" (or "60 days notice").
Always consult the canonical records: [experience-bank.md](../resume-tailoring/resources/experience-bank.md) and [bullet-vault.md](../resume-tailoring/resources/bullet-vault.md).

A specific, modest, *true* email beats an impressive, inflated one — because everything in it has to survive an interview.

## Anatomy of the email

**1. Subject line — a hook, not a label.**
Make it scannable and specific. Good patterns:
- `Backend SWE (Node.js/NestJS/Kafka) — interested in [role] at [company]`
- `ISchoolConnect engineer → [company] [team]: quick intro`
Avoid `Job application` or `Resume attached`. Lead with the searchable stack keyword + the target.

**2. tldr line (optional but strong for cold sends).**
One sentence at the very top so a skimmer gets the point without reading the body: *"tldr — early-career backend engineer (Node/NestJS/Kafka, real-time + payments) looking to move into [company]'s [team]; résumé attached."*

**3. Opening — who he is, in one line.** No throat-clearing. "I'm Yash, a software engineer at ISchoolConnect" — then straight to relevance.

**4. Proof paragraph — 2-3 sentences of specific, shipped work.** This is the whole email. Pick the 1-2 strongest proof points that match the role's thesis (see channel strategy). Use real metrics. Examples to draw from:
- NestJS middleware bridging Genesys telephony with Vapi.ai over WebSocket — real-time voice AI, 1k+ concurrent sessions.
- Event-driven GPA / Transfer-Credit evaluation pipelines in Slate CRM with Kafka.
- 3× analytics API speedup (15s → 5s) by restructuring data-access patterns.
- Next.js + Stripe subscription billing with 100% payment integrity.
- Fluxpay (personal) — payments microservice, 1,200 TPS simulated, p95 <120ms.

**5. The ask — one, concrete, easy.** "Would you be open to a quick screen this week?" or "Is this role still open, and would my background be a fit to put forward?" Make the next step a single yes/no.

**6. Sign-off with links.** Name, LinkedIn (`linkedin.com/in/yash-gupta-engineer`), GitHub (`github.com/Yash-Gupta-01`). Attach the **role-tailored résumé variant** (built via the resume-tailoring skill), not a generic one.

## Channel strategy (which proof points lead)

Match the proof paragraph to where the email is going:

- **Cold GCC / product-company email** (Gurugram GCCs — AmEx, Optum, Mastercard, S&P Global, Fiserv, Pine Labs; product cos — Razorpay, Zeta, Groww, Cred, Postman, Sprinklr): lead with **Node.js / NestJS / Kafka / payments / real-time systems**. This is his sharpest, most differentiated narrative.
- **Google referral / big-tech DSA-heavy req:** foreground **C++ / Python / distributed-systems** signal and DSA depth (450+ problems), matching their stated minimum qualifications — but only what's true.
- **AI/GenAI-forward role:** lead with the **voice-AI orchestration, agentic GPA/TCE pipelines, Langfuse debugging** work.

## Referral-routed email (different from a cold ask)

When someone has agreed to refer him, the email to *them* is a thank-you-plus-ammunition: a 3-line summary they can paste, the tailored résumé, and the exact role/Job ID. Make their job trivial. When the referral email goes *to the company* (some portals ask for a note), it's warmer and can name the referrer.

## Length and tone

- **Under ~150 words for a cold send.** If it's longer, cut.
- Confident peer, not supplicant and not faux-senior. He has real, shippable work — the tone should match an engineer who knows it.
- No desperation ("I'll take anything"), no entitlement ("I expect a response"). "Thanks either way" is a strong, low-pressure close.
- Indian recruiting context: it's fine to be direct about NCR location preference and that he's actively interviewing.

## Output format

Produce the **single best email** for the situation — subject line + body + sign-off — ready to send, with the résumé-attach point noted. Only give labeled variants if the user explicitly wants to A/B two angles (e.g., results-led vs. curiosity-led subject). Always flag: (a) any place he must drop in a real, specific hook, and (b) any JD requirement he doesn't meet, so he can decide how to address it honestly rather than letting the email imply a skill he lacks.
