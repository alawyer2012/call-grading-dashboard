# Call Grading Dashboard — Agent Context

## What This Is

A self-contained HTML dashboard comparing AI model call-grading scores against human grader scores for Leasing Center calls. Supports **two lead types** (Cold Leads and Warm Leads) with independent run tracking for each. The goal is to tune the AI model until it aligns with human graders (target: 90%+ agreement, <3% avg score delta).

## Dual Lead Type Architecture (added June 8, 2026)

The dashboard uses a **top-level lead-type selector** (Cold Leads | Warm Leads) above the run selector. Each lead type has its own:
- Data: `coldRuns`/`warmRuns`, `coldAnswerData`/`warmAnswerData`, `coldRunAnswerData`/`warmRunAnswerData`
- Config: `leadConfig.cold`/`leadConfig.warm` (scoring formula, question count, labels)
- Matrix questions: `coldMatrixQuestions`/`warmMatrixQuestions`

Active references (`runs`, `answerData`, `runAnswerData`, `matrixQuestions`) are swapped by `switchLeadType()`.

### Cold Leads
- **21 scored questions, 80 weighted points**, 3 DQs (FHA, Secure info, No contact)
- Score = (earned / 80) × 100%, with 20% reduction per DQ
- **Rebuild script:** `rebuild_from_spreadsheet.py`
- **Source data:** `~/Downloads/20 Call Cold Lead Comparison (N).xlsx` (current: `(9).xlsx`)

### Cold Lead Run 8.0 (July 20, 2026) — Fatal objection protocol
- **Tab:** `AI 8` → Run 8.0 (full 80-pt card); also Run 8.0-sim (same answers, Run 7 exclusions)
- **Change:** Fatal objection — agents not penalized when they had no legitimate chance to gather an answer
- **Results:** Run 8.0 = **84.8%** (+1.0 vs Run 6); Run 8.0-sim = **88.1%** (flat vs Run 7)
- **Note:** Spreadsheet also has an `AI 7` tab not yet imported (dashboard Run 7 remains the exclusion sim on AI 6)
- Lenient errors improved (24→20); strict stuck at 44 on full card. Conversational still 50%.

### Warm Leads
- **14 scored questions, 52 weighted points** (from AI Warm Lead Fundamentals tab)
- Score = unweighted Yes count / 13 (text_email_perm excluded as conditional)
- DQ penalties NOT applied (SecureInfo column has inverted semantics — "Yes" = compliant)
- **Rebuild script:** `rebuild_warm_from_spreadsheet.py`
- **Source data:** `~/Downloads/20 Call Warm Lead Comparison.xlsx`

### How to Add Data for a New Lead Type Run
1. Download fresh `.xlsx` from Google Sheet
2. Add new tab entry to `AI_TABS` in the appropriate rebuild script
3. Run the rebuild script (cold or warm)
4. Run the other rebuild script too (they don't interfere with each other's data)
5. Manually add recommendations and rootCause for the new run
6. Deploy: `git add -A && git commit -m "update" && git push`

**File:** `/Users/alawyer/Entrata PM/Dashboard/call-grading/index.html`  
**Live URL:** https://alawyer2012.github.io/call-grading-dashboard/  
**GitHub Repo:** https://github.com/alawyer2012/call-grading-dashboard  
**Serve locally:** `cd "/Users/alawyer/Entrata PM/Dashboard/call-grading" && python3 -m http.server 8890`  
**Deploy:** `cd "/Users/alawyer/Entrata PM/Dashboard/call-grading" && git add -A && git commit -m "update" && git push`  
(GitHub Pages auto-deploys on push to master. No tokens, credits, or CLI auth needed.)

---

## The Scoring Model — Cold Leads

**Total: 80 weighted points** across 21 regular questions (Yes/No). Score = earned_points / 80 × 100%.

### Question Weights (from AI QA 2026 spreadsheet):

| Weight | Question |
|--------|----------|
| 9 | Feature/amenity benefit selling (HIGHEST — 11.25% of total) |
| 6 | Open-ended questions (asked 2+ distinct) |
| 6 | Acknowledged caller's question / willingness to help |
| 5 | Urgency language (if pricing discussed) |
| 5 | Required disclaimers stated |
| 5 | Tour scheduled OR clear next step |
| 4 | Greeting (property name + intro) |
| 4 | Rapport building / sentence variety |
| 4 | Offered a tour |
| 3 | Name usage (asked for + used) |
| 3 | Conversational info gathering |
| 3 | Phone number gathered |
| 3 | Occupants confirmed |
| 3 | Pets confirmed |
| 3 | How heard about community |
| 3 | Inclusive language (we/us) |
| 3 | Pricing value (if applicable) |
| 2 | Text/email permission |
| 2 | Email address gathered |
| 2 | Pet breed (if applicable) |
| 2 | Closing / confirm next steps |

### Disqualifiers (20% reduction each if triggered):
- FHA violation (if YES → score × 0.80)
- Secure info disclosed (if YES → score × 0.80)
- No contact attempt despite caller interest (if YES → score × 0.80)

Multiple disqualifiers stack: 2 triggered = score × 0.60.

---

## The Scoring Model — Warm Leads

**14 scored questions, unweighted** (each question = 1 point out of 13). `text_email_perm` is excluded (conditional — "If this was never offered, always answer YES"). Score = Yes count / 13 × 100%.

### Question Weights (from AI Warm Lead Fundamentals tab):

| Weight | Question |
|--------|----------|
| 6 | Open-ended questions (caller's stage in leasing process) |
| 6 | Acknowledged caller / took ownership |
| 5 | Pricing disclaimer / legal language |
| 5 | Required disclaimers stated |
| 4 | Greeting (property name + intro) |
| 4 | Rapport / sentence variety |
| 4 | Tour offer |
| 3 | Name usage (asked for + used) |
| 3 | Conversational info gathering |
| 3 | Phone number gathered |
| 3 | Inclusive language (we/us) |
| 2 | Text/email permission (conditional — excluded from scoring) |
| 2 | Email address gathered |
| 2 | Closing / confirm next steps |

### Disqualifiers:
- FHA violation (DQ) — never triggered in benchmark set (all No)
- Secure info (DQ) — inverted semantics: "Yes" = agent was compliant (all Yes in benchmark). NOT applied to scoring.

### Key differences from Cold Leads:
- No: Occupants, Pets, Pet breed, How heard, Pricing value, Feature/amenity, Tour/next step, No-contact DQ
- Scoring: Unweighted (count-based) vs weighted points
- Total: 13 applicable questions vs 21

### Warm Lead Benchmark (Run 1 — baseline, June 8, 2026):
- **76.2% overall question-level agreement** (198/260 answers match, 13 scored questions × 20 calls)
- **62 total disagreements** across all calls
- **11.5% average absolute score delta**
- **42 strict errors** (AI=No, Human=Yes — AI under-crediting)
- **20 lenient errors** (AI=Yes, Human=No — AI over-crediting)
- 2:1 strict/lenient ratio — AI is predominantly too strict on warm leads
- **Data correction (June 8 PM):** secure_info DQ column corrected from "Yes" to "No" on all 20 calls (was a data entry error). Source file updated to `(1).xlsx`.
- **Note:** Earlier CONTEXT.md versions listed 70.7% / 280 comparisons — that included `text_email_perm` (excluded) and pre-correction DQ data.

### Warm Lead Call IDs in the benchmark set:
264935375, 264953560, 264953958, 264954253, 264994068, 264999130, 265063727, 265114364, 265115677, 265323578, 265335415, 265371349, 265406913, 265486931, 265502366, 265522418, 265537337, 265546780, 265567044, 265570194

### Warm Lead Transcript Files (downloaded June 8, 2026):
`/Users/alawyer/Downloads/warm-lead-transcripts/` — 20 JSON files (Deepgram format, 2-channel audio). Channel 0 = caller, Channel 1 = agent. Files include word-level timestamps, confidence scores, and paragraph-level interleaved dialogue.

---

## Warm Lead Transcript Analysis (June 8, 2026)

### Call Pattern Analysis
Warm leads are fundamentally **service calls, not sales calls**:
- **55% mention callbacks** (11/20) — callers are existing prospects or applicants
- **55% mention applications** (11/20) — callers are already in the leasing pipeline
- **35% mention tours** (7/20) — only a third involve tour scheduling
- Callers often already know specific people at the property (Tia, Ashley, Andrew, Jacob, Tara)
- Call durations: 81s–545s, with many under 2 minutes
- Agent talk share is consistently higher than caller (agents drive the interaction)

### Disagreement Patterns — Transcript Evidence

**1. CONVERSATIONAL (12 total: 9 lenient, 3 strict) — HIGHEST IMPACT**
- **Lenient (AI=Yes, Human=No):** AI credits rapid-fire Q&A as "conversational." In Call 264935375, agent asks 14 questions with 11 consecutive (name? application? phone? email? working with Tia? merging from where?). Pure information gathering, zero personalization. Human says No — correctly.
- In Call 264994068, agent asks 18 questions in an 8-exchange block. Same pattern.
- **Strict (AI=No, Human=Yes):** In Call 265502366, agent has brief but natural exchange ("I'm happy to cancel that for you. Would you like to reschedule?"). Agent adapts to caller's situation. Human credits conversational; AI doesn't.
- **Key insight:** For warm leads, many calls are essentially callback-arrangement flows. The agent collects name → phone → email → reason → arranges callback. This is *administrative*, not conversational. The AI can't distinguish the two.

**2. NAME USAGE (9 total: 7 strict, 2 lenient)**
- **Strict:** AI misses names used once or in data-gathering context:
  - Call 264935375: "And have you completed an application, **Sabrina**?" — clear name usage, AI missed it.
  - Call 265115677: "Thank you, **Levon**. And your phone number, please?" — clear name usage.
  - Call 265406913: "And how do you spell your last name, **Marcia**?" — asking about name IS using it.
  - Call 265323578: Agent tries to pull up under caller's name but never uses it conversationally.
- **Root cause:** Warm lead agents use names once (confirmatory) rather than repeatedly (sales-style). AI requires repeated use.

**3. PHONE NUMBER (6 strict)**
- Call 264935375: Agent asks "What's your phone number and email?" → Caller gives "305 2995514." → AI says No.
- Call 265371349: Agent asks "May I have, please, your phone number?" → Caller starts giving it ("301...") → AI says No.
- **Root cause:** Phone numbers are often already on file for warm leads. Agent may confirm existing number or collect it as part of a multi-info request ("name, email, and phone?"). AI requires isolated, explicit phone-only collection.

**4. PRICING DISCLAIMER (5 strict)**
- Call 265114364: Agent says "contact us today to receive up to 1 month free" — promotional pricing, human credits it.
- Call 264954253: Agent discusses bedroom sizes and occupancy limits — implicit pricing context.
- **Root cause:** For warm leads, pricing context appears in existing-relationship language ("your special," "the rate"). AI requires formal disclaimer wording.

**5. INCLUSIVE LANGUAGE (5 strict, 0 lenient)**
- Call 265063727: Agent says "**we do** help with the leasing process," "**our** property," "**our** end" — 3 clear instances. AI missed all.
- Call 265570194: Agent says "**our** team" — 1 instance. AI missed it.
- 3 of 5 calls (264935375, 264953958, 265115677) genuinely have NO we/our in agent speech. Human may be applying a more relaxed standard for warm leads.
- **Root cause:** Mixed. AI misses legitimate uses in some calls, but human over-credits in others.

**6. RAPPORT (7 total: 5 strict, 2 lenient)**
- **Strict:** AI misses empathy-driven rapport:
  - Call 265115677: "I am sorry. Unfortunately, I do not have access..." + "I am **definitely** taking your message and arranging a callback" — empathy + commitment.
  - Call 265570194: "I'm **so sorry**" + "**No worries**" + "I **understand** that" — genuine empathy.
- **Lenient:** AI over-credits generic pleasantries:
  - Call 265546780: "no worries" and "thank you so much" — friendly but not rapport.
- **Key insight:** For warm leads, rapport means acknowledging the caller's existing relationship and urgency, not just being friendly.

**7. CLOSING (3 strict, 0 lenient)**
- Call 264953560: Ends with "You have a good day. Thank you for calling. Bye." — warm but not formal.
- Call 265570194: Agent says "Is there anything else I can help you with?" then confirms callback — IS a next step.
- **Root cause:** Warm lead closings are service-oriented ("I'll arrange a callback," "I'll send this over") rather than sales-oriented ("Let me schedule that tour for you").

### Meta-Insights

1. **Warm leads are service calls.** The rubric is designed for sales (cold leads). Service calls have different conversational patterns, rapport markers, and closing structures.
2. **Conversational is the single biggest problem** (12 disagreements, 45% of all lenient errors). The AI credits administrative Q&A as conversational. This is the highest-ROI fix.
3. **The strict/lenient split is almost 50/50** (42/40). Very different from cold leads (65% strict). Warm lead fixes need to be balanced.
4. **Name usage differs from cold leads.** Warm leads have names on file. One-time confirmatory use ("Thank you, Levon") should count.
5. **Phone collection differs.** "Is this a good callback number?" or multi-info requests ("name, email, and phone?") are phone collection for warm leads.

---

## The Data Set — Cold Leads

**20 cold-lead calls** graded by both human and AI on the same rubric.

### Current Performance (Corrected — against 'New Manual' tab, June 10, 2026):
- **81.9% overall question-level agreement** (344/420 answers match) — Run 1 baseline
- **76 total disagreements** across all calls
- **13.2% average absolute score delta** (weighted)
- **55 strict errors** (AI=No, Human=Yes — AI under-crediting agents, 72%)
- **21 lenient errors** (AI=Yes, Human=No — AI over-crediting agents, 28%)
- Runs 1-4 all cluster at 81.9-82.1% (stable model, ~75 disagreements)

**⚠️ Previous reports showed 67.1% agreement — this was due to 100 incorrect answers in the old 'Manual' tab. Corrected June 10, 2026.**

**Methodology:** Both human and AI scores are independently calculated from raw Yes/No answers using the same formula: (earned points / 80) × 100%, with 20% reductions per disqualifier. Human answers come from the Manual tab, AI answers from the AI tab of the `20 Call Cold Lead Comparison` Google Sheet.

**Data correction (June 5 PM):** Previous sessions extracted data from `Untitled spreadsheet.xlsx` with incorrect column mappings, producing 100 wrong human answer values out of 480. All data was re-extracted from the authoritative source (`20 Call Cold Lead Comparison (1).xlsx`, Manual + AI tabs). Score validation: 19/20 computed human scores exactly match the spreadsheet's Overall Score column; one 5-point discrepancy on Call 265311653 (computed 90%, sheet says 85%).

**Note:** Call 263494762 has incomplete AI data (only 4 of 21 scored questions answered by AI). Its AI score of 28.75% is unreliable and skews the average delta.

### Call IDs in the benchmark set:
263494762, 264909961, 265017569, 265039609, 265085947, 265103498, 265106571, 265121798, 265198083, 265212524, 265255755, 265292177, 265300729, 265309940, 265311653, 265329801, 265336988, 265340388, 265390645, 265426567

### Transcript files (extracted from Google Drive):
`/Users/alawyer/Downloads/call-transcripts/` — 20 text files, one per call (extracted from Deepgram JSON).

### Source spreadsheet (authoritative):
- `/Users/alawyer/Downloads/20 Call Cold Lead Comparison (1).xlsx` — Downloaded from Google Sheet ID `1SeNuw9lI43bibffZGmpjHlWeRtYsepbswXwWPCgRzEc`
  - **Manual tab**: Human grades. Col 1=Overall Score, Col 2=Call Id, Cols 3–26=Yes/No answers. 20 data rows.
  - **AI tab**: AI Run 1 grades. Col 1=call_id, Cols 2–25=Yes/No answers, Col 26=error. 20 data rows.
  - **AI 2 tab**: Placeholder for Run 2 (currently empty — waiting for Monica/Myles changes).
  - **AI 3, AI 4 tabs**: Placeholders for future runs.
- `/Users/alawyer/Downloads/AI_ QA 2026.xlsx` — Tab "AI Lead Fundamentals" has all questions, prompts, protocols, definitions (Column 7 = Protocols)

**⚠️ Do NOT use** `/Users/alawyer/Downloads/Untitled spreadsheet.xlsx` — that file has a different layout (human/AI on same sheet at different row offsets) and produced incorrect data in earlier sessions.

---

## Dashboard Architecture (multi-run, updated June 5, 2026)

### Two-tier navigation
- **Tier 1 (Run Selector):** `Overview` | `Run 1` | `Run 2` | ... — prominent bar at top
  - When only 1 run exists, lands directly on Run 1 (no Overview)
  - When 2+ runs exist, Overview is default landing
  - Latest run gets a green "Latest" badge, first run gets gray "Baseline" badge
- **Tier 2 (Content Tabs):** scoped to selected run → `Question Alignment` | `Per-Call Scores` | `Matrix` | `Root Causes` | `Recommendations`

### Overview tab (visible with 2+ runs)
- 4 hero metric cards: Agreement %, Avg Delta, Total Disagreements, Error Bias — each with delta arrows vs prior run
- Trend line charts: Agreement % over runs (with 90% target line), Disagreement count over runs (total + strict + lenient)
- Run History table: clickable rows → jump to any run

### Per-run content (5 tabs each, identical structure)

#### 1. Question Alignment
- 4 stat tiles: overall agreement, total disagreements, avg weighted delta, target
- Expanded key findings callout
- Expandable table: each question row shows weight, agreement rate, score impact, direction
- Click any row → accordion drops down with transcript evidence per call

#### 2. Per-Call Scores
- Bar chart showing score delta per call
- Expandable table sorted by largest delta: Human Score, AI Score (exact, calculated from weighted scorecard), Delta, Direction
- Click any call → accordion shows every question disagreed on with weight, AI answer, human answer, explanation
- "HOW TO IMPROVE THIS SCORE" recommendation at bottom of each expanded call

#### 3. Matrix
- Full grid: Call IDs (rows) × All 24 questions (columns)
- Cell colors: **Red bold** = AI too strict, **Amber bold** = AI too lenient, Plain = agreed
- **Column header colors** (based on 8+ total disagreements for that question):
  - **Red header** = majority of disagreements are AI too strict
  - **Amber header** = majority of disagreements are AI too lenient
  - **Plain header** = fewer than 8 disagreements or no dominant direction

#### 4. Root Causes
- Bias tiles (strict / lenient counts)
- Two cards: "AI too lenient" and "AI too strict" with detailed explanations
- Transcription evidence section with specific examples
- Callout on transcription mitigations

#### 5. Recommendations
- Recommendations with **current protocol** and **recommended protocol** with changes in red bold
- Protocols sourced from `AI_ QA 2026.xlsx` → "AI Lead Fundamentals" → Column 7
- Green-bordered boxes labeled "Recommended Protocol — Copy This" for copy-paste into model
- **Owner pills on each rec:**
  - **Monica** (purple pill) — protocol changes
  - **Myles** (green pill) — prompt/AI changes

### Current Recommendations:
1. **Expand recognition for occupants, pets, how-heard** — Monica — est. ~20 fixes
2. **Redefine conversational and rapport with negative examples** — Monica — est. ~15 fixes
3. **Reclassify open-ended question detection** — Monica — est. ~8 fixes
4. **Lower the bar for feature/amenity and inclusive language** — Monica — est. ~12 fixes
5. **Strengthen transcription handling with question-specific guidance** — Myles — est. ~10 fixes (appends to existing TRANSCRIPT NOISE section in prompt)
6. **Add protocol-prioritization instruction to STEP 3** — Myles — may fix multiple errors
7. **Investigate: favorable-default may be causing leniency** — Myles — investigation item

---

## The LLM Scoring Prompt (provided by Myles, June 5)

The model uses a system prompt with these key sections:
- **TRANSCRIPT NOISE** — already tells the model to infer intent from context, handle phonetic misspellings. Rec 5 appends question-specific guidance here.
- **STEP 1: READ THE ENTIRE TRANSCRIPT** — standard
- **STEP 2: IDENTIFY THE AGENT** — phonetic name matching for agent identification (already robust)
- **STEP 3: ANSWER QUESTIONS** — tells model to analyze transcript and answer. Rec 6 adds protocol-prioritization instruction here.
- **CONDITIONAL / NOT-APPLICABLE QUESTIONS** — auto-passes questions when the condition doesn't apply ("select the most favorable/positive answer"). Rec 7 investigates whether this causes false leniency.
- **STEP 4: FORMAT YOUR RESPONSE** — JSON output format

Key insight: The question protocols (from the spreadsheet) are passed as input alongside the questions, NOT embedded in this system prompt. So Monica's protocol changes (Recs 1–4) go into the question definitions, while Myles's changes (Recs 5–7) go into the system prompt.

---

## Key Findings Summary

### Where AI is too STRICT (90 errors):
1. **Open-ended questions (6 pts, 9 strict)** — AI grammar-parses too literally; "What size?" gets rejected.
2. **Pricing value (3 pts, 9 strict)** — AI misses pricing value language in various phrasings.
3. **Occupants (3 pts, 8 strict)** — AI requires a direct question; misses when caller volunteers info.
4. **Feature/amenity (9 pts, 8 strict)** — AI requires explicit benefit language; humans credit clear mentions.
5. **Pets (3 pts, 7 strict)** — Same pattern as occupants — AI misses volunteered info.
6. **How heard (3 pts, 7 strict)** — AI misses various phrasings of the question.
7. **Name usage (3 pts, 6 strict)** — Exact string matching fails on transcription garbles.
8. **Tour offer (4 pts, 6 strict)** — AI requires exact "tour" wording; misses "visit"/"showing."
9. **Inclusive language (3 pts, 6 strict)** — AI misses "We do have..." and "Our community..." patterns.

### Where AI is too LENIENT (48 errors):
1. **Conversational gathering (3 pts, 9 lenient)** — AI credits scripted checklist-style exchanges.
2. **Rapport building (4 pts, 8 lenient)** — AI counts filler enthusiasm as rapport.
3. **Tour/next step (5 pts, 4 lenient)** — AI credits vague "I'll send you info" as a next step.
4. **Occupants (3 pts, 4 lenient)** — AI over-credits mentions without proper confirmation.
5. **Pets (3 pts, 4 lenient)** — AI over-credits mentions without proper confirmation.
6. **How heard (3 pts, 4 lenient)** — AI over-credits without proper confirmation.

---

## Previous Conversations

- [Call grading comparison](d5b4170a-43be-48d2-a7ec-08faa9a829ee) — First analysis with equal weighting
- [Deep dive + dashboard build](fac8e933-e238-4527-8b37-e1395905a4ff) — Downloaded transcripts, built dashboard, started Netlify deploy
- [Dashboard overhaul + prompt review](fac8e933-e238-4527-8b37-e1395905a4ff) — Matrix header coloring, recommendations rewrite with real protocols, Myles prompt analysis, Netlify deploy pipeline, multi-run architecture scaffolded
- [Score correction + Run 2 prep](earlier June 5 session) — Added `answerData` constant, rewrote matrix builder. BUT: extracted data from the WRONG file (`Untitled spreadsheet.xlsx`), producing 100 incorrect human answers. This was caught and fixed in the next session.
- [Data integrity fix](June 5 PM session) — Full re-extraction from authoritative source. Metrics changed from 83.8% → 67.1% agreement. CONTEXT.md created with lessons learned.
- [Run 2 extraction + deep analysis](June 8 AM session) — Downloaded `(2).xlsx` with AI 2 tab. Rebuilt script for multi-run support. Run 2 results: 69.0% (+1.9%). Deep per-question diff showed only 26/480 answers changed. Diagnosed: model ignores additive protocol text, only responds to threshold simplification. Reviewed actual prompt from Myles — identified positioning issues. Updated Run 2 recommendations with prompt-aware analysis. Restored Run 1 recommendations. Updated CONTEXT.md.

---

## What's Next

1. **Waiting on Myles** to apply the two positioning changes to the prompt (move VOLUNTEERED INFO + NEGATIVE EXAMPLE OVERRIDE above the numbered steps, add forward reference in step 1)
2. **Waiting on Monica** to simplify how-heard protocol (remove "Evaluate the entire transcript before marking NO")
3. **Re-run model** on the same 20 calls → results go in the `AI 3` tab of the Google Sheet
4. **Austin downloads** the updated sheet to `~/Downloads/`
5. **Process Run 3:** Run `rebuild_from_spreadsheet.py` (add `AI 3` to the `AI_TABS` list in the script)
6. **Redeploy** — `cd "/Users/alawyer/Entrata PM/Dashboard/call-grading" && npx netlify-cli deploy --dir=. --prod --site b7f024c5-ca46-427e-89e6-801ba0ada830`
7. **Expected outcome:** 75-80% agreement (currently 69%). Main gains from volunteered-info (strict) and negative-example override (lenient).
8. Iterate — Run 4+ will need question-specific transcription guidance for remaining strict errors (pricing value, phone, tour offer)

### Progression strategy (crawl → walk → run):
- **Run 2 (done):** Protocol threshold simplification → proved that lowering YES bars works (+1.9%)
- **Run 3 (next):** Structural prompt additions → proving model can follow processing rules (+6-11% expected)
- **Run 4+:** Apply both levers to all remaining questions → converge on 90%

### How to Add a New Run

The `rebuild_from_spreadsheet.py` script handles everything. To add Run 3:

1. Add a new entry to the `AI_TABS` list at the top of the script:
   ```python
   {"tab": "AI 3", "id": 3, "label": "Run 3", "date": "...",
    "description": "Structural prompt additions — processing order, volunteered info rule, negative example override.",
    "changes": "Added protocol processing ORDER, VOLUNTEERED INFORMATION RULE, NEGATIVE EXAMPLE OVERRIDE to system prompt"}
   ```
2. Update `XLSX_PATH` if the filename changes (currently `(2).xlsx`)
3. Run the script: `python3 rebuild_from_spreadsheet.py`
4. The script reads all AI tabs, computes metrics, and patches the HTML automatically
5. Manually add recommendations and rootCause for the new run (the script outputs empty arrays for these)
6. Deploy: `git add -A && git commit -m "update" && git push`

### Future enhancements (discussed but not built):
- What-If Simulator (toggle fixes on/off, see projected impact)
- Export fix briefs per recommendation as one-pagers for Jira
- Call-type classification (removed from recs for now — needs more design work)
- Per-question trend sparklines in Overview (agreement rate for each question across runs)

---

## Run 2 Results & Analysis (June 8, 2026)

### Performance: 67.1% → 69.0% (+1.9%)

| Metric | Run 1 | Run 2 | Delta |
|--------|-------|-------|-------|
| Agreement | 67.1% | 69.0% | +1.9% |
| Disagreements | 138 | 130 | -8 |
| Strict errors | 90 | 83 | -7 |
| Lenient errors | 48 | 47 | -1 |
| Avg score delta | 23.7% | 22.4% | -1.3 pts |

### What worked (threshold simplification):
- **Inclusive language (6→1, -5):** Protocol + prompt guidance (double fix)
- **Feature/amenity (8→5, -3):** Protocol + prompt guidance (double fix)
- **Open-ended (9→7, -2):** Protocol broadened

### What didn't work (additive protocol text):
- **Occupants (8→8):** "Also mark YES if volunteered" — model ignored
- **Conversational (9→9 lenient):** Negative examples added — model ignored
- **Rapport (8→8 lenient):** Negative examples added — model ignored
- **How heard (7→9, +2 worse):** Protocol expansion confused model

### Root cause diagnosis:
The model only changed 26 of 480 answers (5.4%). Of those, 18 helped, 8 hurt (net +10).
The model responds to threshold simplification but ignores additive protocol text.
This is a prompt architecture issue — the model needs structural processing instructions.

### Key insight: Protocol changes only work when paired with matching prompt guidance.
Feature/amenity and inclusive language both improved because they got BOTH a simpler protocol
AND matching entries in QUESTION SPECIFIC TRANSCRIPTION GUIDANCE. Occupants/pets/how-heard
got protocol expansions but no prompt support — zero effect.

---

## The LLM Scoring Prompt (Run 2 version — from Myles, June 8)

Key sections and what they do:

- **TRANSCRIPT NOISE** — Generic transcription error handling (phonetic misspellings, context inference)
- **QUESTION SPECIFIC TRANSCRIPTION GUIDANCE** — Feature/amenity pattern matching + inclusive language scanning. Only these two questions have guidance; name usage, phone, pricing value do NOT.
- **STEP 1: READ THE ENTIRE TRANSCRIPT** — Standard
- **STEP 2: ANSWER QUESTIONS** — Protocol-following instructions. Tells model to follow protocols exactly, not substitute judgment. "If the protocol says 'Mark NO if...' mark NO."
  - 5-step numbered procedure (read protocol, identify sections, consider explicit/implicit, cross-reference, choose answer)
- **STEP 3: FORMAT YOUR RESPONSE** — JSON output format

**What was added for Run 2 (vs Run 1):**
- QUESTION SPECIFIC TRANSCRIPTION GUIDANCE section (feature/amenity + inclusive language)
- Expanded step 1 to say "Read and fully understand the FULL protocol text — not just the first sentence"
- "If the protocol says 'Mark NO if...' mark NO even if behavior seems reasonable" (protocol-prioritization)

**What was removed for Run 2:**
- CONDITIONAL / NOT-APPLICABLE section ("select most favorable answer") — gone entirely
- STEP 2: IDENTIFY THE AGENT — removed (old prompt had 4 steps, new has 3)

**What's being added for Run 3 (pending Myles):**
- VOLUNTEERED INFORMATION RULE (after step 1 expansion)
- NEGATIVE EXAMPLE OVERRIDE (after volunteered info)
- Positioning: both should move ABOVE the numbered steps, not after them
- Forward reference in step 1: "mark YES — unless a negative example also matches"
- (Suggested) NAME USAGE entry in QUESTION SPECIFIC TRANSCRIPTION GUIDANCE

---

## Data Pipeline — How to Process a New Run

### What Austin provides
A fresh download of the Google Sheet (`1SeNuw9lI43bibffZGmpjHlWeRtYsepbswXwWPCgRzEc`) as `.xlsx`, dropped in `~/Downloads/`. The sheet has separate tabs:
- **Manual tab**: Human grades (ground truth). Col 1=Overall Score, Col 2=Call Id, Cols 3–26=Yes/No answers. 20+ data rows (50 as of June 8, but only 20 match the benchmark set).
- **AI tab**: Run 1 AI grades. Col 1=call_id, Cols 2–25=Yes/No answers, Col 26=error. 20 data rows.
- **AI 2 tab**: Run 2 AI grades. Same column structure as AI tab. 20 data rows.
- **AI 3 tab**: Placeholder for Run 3 (currently empty).
- **AI 4 tab**: Placeholder for Run 4.

The Manual (human) answers stay the same across runs (they're ground truth). Each AI run gets its own tab.

### Source spreadsheet (authoritative):
- **Current file:** `/Users/alawyer/Downloads/20 Call Cold Lead Comparison (5).xlsx`
- **Human scores tab:** `New Manual` (NOT `Manual` — the original Manual tab had 100 incorrect answers)
- Previous files: `(1).xlsx` was Run 1 only. `(2).xlsx` added AI 2 tab. `(3).xlsx` same as (2). `(4).xlsx` added AI 3 tab. `(5).xlsx` added AI 4 tab + corrected `New Manual` tab.
- **⚠️ Do NOT use the `Manual` tab** — it has 100 incorrect answers across 19 calls. Always use `New Manual`.
- **⚠️ Do NOT use** `/Users/alawyer/Downloads/Untitled spreadsheet.xlsx`

### Column-to-question mapping (validated — 19/20 human scores match spreadsheet)

**Manual tab (cols 3–26, Col 1=Overall Score, Col 2=Call Id):**

| Col | Question | Weight |
|-----|----------|--------|
| 3 | Open-ended questions | 6 |
| 4 | Pricing disclaimer / urgency | 5 |
| 5 | Text/email permission | 2 |
| 6 | Required disclaimers | 5 |
| 7 | Feature/amenity benefit | 9 |
| 8 | Tour offer | 4 |
| 9 | Email gathered | 2 |
| 10 | Pricing value language | 3 |
| 11 | Greeting | 4 |
| 12 | Tour/next step | 5 |
| 13 | Acknowledged caller | 6 |
| 14 | Inclusive language | 3 |
| 15 | Name usage | 3 |
| 16 | Conversational gathering | 3 |
| 17 | Rapport / variety | 4 |
| 18 | Phone number | 3 |
| 19 | Occupants | 3 |
| 20 | Pets | 3 |
| 21 | How heard | 3 |
| 22 | Pet breed | 2 |
| 23 | FHA (DQ) | — |
| 24 | Secure info (DQ) | — |
| 25 | No contact (DQ) | — |
| 26 | Closing | 2 |

**AI tabs (all use the same layout — Col 1=call_id, cols 2–25 answers, col 26=error):**

| Col | Question | Weight |
|-----|----------|--------|
| 2 | FHA (DQ) | — |
| 3 | Secure info (DQ) | — |
| 4 | No contact (DQ) | — |
| 5 | Greeting | 4 |
| 6 | Name usage | 3 |
| 7 | Conversational gathering | 3 |
| 8 | Rapport / variety | 4 |
| 9 | Phone number | 3 |
| 10 | Occupants | 3 |
| 11 | Pets | 3 |
| 12 | Pet breed | 2 |
| 13 | How heard | 3 |
| 14 | Open-ended questions | 6 |
| 15 | Pricing disclaimer / urgency | 5 |
| 16 | Text/email permission | 2 |
| 17 | Required disclaimers | 5 |
| 18 | Feature/amenity benefit | 9 |
| 19 | Tour offer | 4 |
| 20 | Email gathered | 2 |
| 21 | Pricing value language | 3 |
| 22 | Tour/next step | 5 |
| 23 | Acknowledged caller | 6 |
| 24 | Inclusive language | 3 |
| 25 | Closing | 2 |
| 26 | Error (ignore) | — |

**Important:** The Manual and AI tabs have different column orders. The extraction script (`rebuild_from_spreadsheet.py`) maps each column to the correct question key independently for each tab.

### Score calculation formula
```
earned = sum of weights for all questions where answer = "Yes"
base_pct = (earned / 80) × 100
dq_count = number of disqualifiers where answer = "Yes"
final_score = base_pct × max(0, 1 - 0.20 × dq_count)
```

### Processing steps for a new run
1. Download fresh `.xlsx` from Google Sheet
2. Add new tab entry to `AI_TABS` list in `rebuild_from_spreadsheet.py`
3. Update `XLSX_PATH` if filename changed
4. Run: `cd "/Users/alawyer/Entrata PM/Dashboard/call-grading" && python3 rebuild_from_spreadsheet.py`
5. Script automatically: reads all tabs, validates human scores, computes per-run metrics, generates `runAnswerData` + `runs` array, patches HTML
6. Manually add recommendations and rootCause for the new run
7. Redeploy: `cd "/Users/alawyer/Entrata PM/Dashboard/call-grading" && git add -A && git commit -m "update" && git push`

**Extraction script:** `rebuild_from_spreadsheet.py` handles multi-run extraction. It outputs empty `recommendations: []` and `rootCause: {}` for each run — these must be filled manually with the analysis.

---

## Lessons Learned (read this first if you're a new agent)

### 1. Always extract from the Google Sheet tabs, never from "Untitled spreadsheet"
The workspace may have multiple `.xlsx` files in `~/Downloads/`. **Only use `20 Call Cold Lead Comparison (N).xlsx`** (latest version number) which has separate `Manual` and `AI` tabs. The file `Untitled spreadsheet.xlsx` has a different layout and produced 100 wrong answer values when used.

### 2. The Manual and AI tabs have different column orders
The Manual tab puts questions in one order (open-ended first, DQs near the end). The AI tab puts DQs first, then questions in a different sequence. You CANNOT assume the same column index maps to the same question across tabs. The extraction script has the correct mapping — use it.

### 3. Never hand-write data arrays that can be computed from source data
If data can be computed from `answerData`, compute it — don't hand-write it. The extraction script does this correctly.

### 4. Validate by clicking through the dashboard, not just checking totals
Always spot-check individual calls against the Matrix tab and the raw `answerData`.

### 5. Score validation is necessary but not sufficient
Matching computed scores to the spreadsheet's "Overall Score" column validates the column-to-weight mapping but NOT that the correct answers were extracted from the correct file.

### 6. Use `rebuild_from_spreadsheet.py` for any data changes
Don't manually edit `answerData`, `calls`, or `questions` — run the script. To add a new run, add the tab to `AI_TABS` and re-run.

### 7. Recommendations and rootCause are NOT generated by the script
The script outputs empty arrays for these. They must be manually written after analyzing the run's results. This is intentional — the analysis requires comparing runs, understanding what changed in protocols/prompt, and diagnosing why.

### 8. Protocol changes only work when paired with prompt guidance
Learned from Run 2: expanding a protocol ("Also mark YES if...") produces zero effect unless the system prompt also tells the model HOW to process that expansion. Feature/amenity + inclusive language improved because they got both simpler protocols AND QUESTION SPECIFIC TRANSCRIPTION GUIDANCE entries. Occupants/pets/how-heard got protocol expansions but no prompt support — zero effect.

### 9. The model responds to threshold simplification, not nuance addition
Simplifying the YES criteria (lower bar, "single instance sufficient", "clear mention sufficient") works. Adding negative examples, additional conditions, or phrasing alternatives does not — the model processes the first paragraph of a protocol and stops.

### 10. Always verify the human scoring data is correct
The original 'Manual' tab had 100 incorrect answers across 19 of 20 calls (some calls had up to 10 wrong answers). This made it appear the model was at 67% when it was actually at 82%. Three full runs of prompt engineering were directed at a phantom problem. **Always use the 'New Manual' tab.** If human scoring seems unreliable, validate against the source-of-truth before blaming the model.

### 11. Deploy via GitHub Pages, not Netlify
Netlify free-tier credits run out. The dashboard is now hosted on GitHub Pages (auto-deploys on push to master). Deploy command: `git add -A && git commit -m "update" && git push`. No CLI auth, no tokens, no credit limits.

---

## Technical Notes

- Single self-contained HTML file (~170KB with 2 runs)
- Uses Chart.js 4.4.7 from CDN (only external dependency)
- Light mode theme
- All data is embedded in JavaScript (no API calls)
- `runs` array holds all run data — each run is a self-contained object with meta, questions, calls, recommendations, rootCause, keyFindings
- `runAnswerData` object holds per-run raw answers: `runAnswerData[runId][callId][questionKey] = [human, ai]`
- `answerData` global still exists for backward compat (points to Run 1 data)
- `buildMatrixData(run)` uses `runAnswerData[run.id]` to render correct per-run matrix
- Two-tier navigation: run selector (tier 1) → content tabs (tier 2)
- Overview tab renders dynamically when 2+ runs exist (trend charts, delta cards, run history table)
- Chart instances tracked in `chartInstances` object and destroyed before re-render (prevents canvas reuse errors)
- Accordion/expand pattern used on Question Alignment and Per-Call Scores tabs — unique IDs use `r{runId}-q{i}` / `r{runId}-c{i}` to avoid collisions across runs
- Per-call score field is `ai` (not `aiEst`) — true calculated score from raw data
- Human scores validated against Manual tab "Overall Score" column: 19/20 exact match (Call 265311653 has a 5-point discrepancy: computed 90%, sheet says 85%)
- CSS classes: `.run-selector`, `.run-tab`, `.run-container` (tier 1), `.tabs`, `.tab`, `.tab-content` (tier 2), `.protocol-current` (gray box), `.protocol-new` (green-bordered), `.change` (red bold text), `.pill-owner-monica` (purple), `.pill-owner-myles` (green), `.header-strict` (red column header), `.header-lenient` (amber column header), `.overview-hero`, `.overview-card`, `.overview-metric`, `.overview-delta` (overview)
