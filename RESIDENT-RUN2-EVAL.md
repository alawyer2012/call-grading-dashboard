# Resident Run 2.0 — Evaluation

> Dashboard: https://alawyer2012.github.io/call-grading-dashboard/ (Residents → Overview / Run 2.0)
> Source grades: `~/Downloads/20 Call Resident Comparison (1).xlsx` (New Manual + AI + AI 2)
> Definitions: `AI_ QA 2026 (6).xlsx` → AI Resident Fundamentals (blue = Run 2.0 updates)
> Transcripts (original 12): `~/Downloads/{call_id}.json`

**Official set:** same 12 IDs as Run 1.0  
**Snapshot:** **89.2%** scored agreement (107/120) · 13 disagreements (6 strict / 7 lenient) · 10.7% avg score delta  
**Vs Run 1.0:** +10.9pp agreement · disagreements 26→13 · strict 20→6

---

## Did the Run 1.0 blue-cell updates work?

| Rec | Question | R1 disagrees | R2 disagrees | Verdict |
|-----|----------|--------------|--------------|---------|
| 1 | Hold permission | 7 strict | **0** | Cleared |
| 2 | Ownership / acknowledged | 4 strict | **0** | Cleared on this set |
| 3 | Closing | 5 strict | **1** | Almost — leftover is 269768640 |
| 4 | Secure-info DQ | 3 lenient | **0** | Cleared |

We missed 90% by **1 percentage point**. Bias flipped from 77% strict to slightly lenient.

Perfect agreement: 269245474, 271873957, 272352194, 272543647 (was only 272352194).

---

## Remaining disagreements (original 12)

| Call | Direction | Questions | Notes |
|------|-----------|-----------|-------|
| **269768640** | 4 strict | contact, **unit (regression)**, open-ended, closing | Worst leftover. Hold+ownership now fixed. Crosstalk; unit was asked; callback refused → contact N/A. |
| 269788776 | 1 lenient | open-ended | New. Greeting + “in what regards” + identity counted as 2 probes. Human No. |
| 272063323 | 1 lenient | open-ended | New. Same pattern on a fully-explained inspection callback. |
| 270245082 | 2 lenient | name, neutral | Unchanged. Wendy/Keith third party; unprofessional agent rant. |
| 270543247 | 1 strict | name | Unchanged. Company name / mangled “Prasvi.” |
| 270842045 | 1S + 1L | greeting, validate | Unchanged. Crosstalk greeting; operational “Understood.” |
| 271912426 | 1 lenient | validate | Unchanged. “I’ll get this checked” ≠ empathy. |
| 272053386 | 1 lenient | neutral | Unchanged. Short calm call; possible grader noise. |

---

## Shadow 20-call set (not official yet)

AI 2 graded 8 extra IDs that now have human scores:  
265908165, 266239119, 269829113, 270583790, 273100258, 273290320, 273900663, 274349154

| Slice | Agreement | Disagreements | Notes |
|-------|-----------|---------------|-------|
| Original 12 | 89.2% | 13 | Official Run 2.0 |
| New 8 | 85.0% | 12 | 270583790, 273100258, 274349154 = perfect |
| All 20 | 87.5% | 25 | Open-ended 7, validate 5 |

Watch: **269829113** (H 81 / AI 45) — ownership missed on unseen data. Do not lock 20 until Run 3.0 clears 90% on the original 12.

Transcript intake (new 8): 269829113 ✅ · 265908165 ✅ · 273290320 ✅ · 266239119 ✅ · 273900663 ✅ · 270583790 ✅ · 273100258 ✅ · 274349154 ✅ · **all 8 in**

---

## New-ID call reviews

### 269829113 — ✅ transcript in
- **Scores:** Human 81.0% · AI 45.2% · Δ −35.7 (largest miss on the new 8)
- **Agent:** Manisha · **Caller:** Amira Molina, apt 423 · callback 503-583-3209
- **Type:** ~123s — caller wanted to **leave an office voicemail** about lease / Homebody / signing; spouse Ezra Molina also on the lease
- **File:** `~/Downloads/269829113.json`
- **Note:** Recording **ends while the caller is still talking** (“I’m just wondering”). Last agent line is “Is it about the renewal of the lease?” — possible truncated capture.

| Question | Dir | Notes | Implication for Run 3.0 |
|----------|-----|-------|-------------------------|
| Ownership | Strict | Agent: **“I can just forward you a voice mail as well.”** Clear first-person action. Not in the Run 2.0 callback/note example list. | Add voicemail-forward language to ownership positives: “I’ll forward a voicemail / place a voicemail for the office.” |
| Open-ended | Strict | Real probe: **“What is the voice mail you want me to place?”** Greeting “How can I help you?” + closed “Is it about the renewal?” Identity questions don’t count. Human Yes with arguably 1 real diagnostic + greeting. | Don’t tighten open-ended so hard that a genuine “what message should I leave?” probe still fails. Greeting still shouldn’t count alone. |
| Closing | Strict | No wrap-up in the file (no “I’ll send this / anything else”). Human still Yes — likely inferred from the voicemail-forward offer. | If transcript is truncated, closing Yes is generous. If complete, human credited implied next step. Don’t over-require a formal close when the agent already committed to placing the voicemail. |
| Name usage | Agree No | Asked for name; never used “Amira.” Both correct. | — |
| Validate | Agree No | Technical difficulties mentioned; neither credited empathy. Fine. | — |

### 265908165 — ✅ transcript in
- **Scores:** Human 57.1% · AI 80.0% · Δ +22.9 (AI higher; both applied secure-info DQ)
- **Agent:** Veronica · **Caller:** Michael Pronto, apt 210 · 813-312-3357
- **Type:** ~229s — wants on-site **courtesy officer / CSO number** for ongoing upstairs noise; office had already closed
- **File:** `~/Downloads/265908165.json`

| Question | Dir | Notes | Implication for Run 3.0 |
|----------|-----|-------|-------------------------|
| Open-ended | **Lenient** | Caller stated the ask immediately (“number to the community service officer”). Agent: greeting + unit/phone + closed “Do you have the number?” No second diagnostic probe. Human No is right. | Confirms Rec 1. Don’t auto-Yes “already explained.” Greeting + identity ≠ 2 probes. |
| Validate concern | **Lenient** | Real concern (“constant noise… never being fixed”). Agent was operational, then at close: “I’m so sorry for the inconvenience you called.” Human still No. | Generic closing apology ≠ validating the stated issue. Empathy must address the **concern** (noise / not getting help), not “sorry you had to call.” |
| Secure-info DQ | Agree **Yes** | Agent gave CSO mobile **813-992-9423** + name John Fulton. Both flagged DQ. | True positive — not a disagreement. Keep: on-site staff personal/direct numbers can be secure. |

### 273290320 — ✅ transcript in
- **Scores:** Human 71.4% · AI 90.5% · Δ +19.0 (AI higher)
- **Agent:** Vaishnavi · **Caller:** Jessica Monty, apt 1524 A
- **Type:** ~123s after-hours **lockout** — property does not handle lockouts; locksmith $75 + key $50 at caller’s expense
- **File:** `~/Downloads/273290320.json`

| Question | Dir | Notes | Implication for Run 3.0 |
|----------|-----|-------|-------------------------|
| Contact info | **Strict** | No callback being arranged — agent only explained lockout protocol. Never asked for a phone. Human Yes = N/A. AI No. Same pattern as 269768640. | If the agent is only explaining policy (not taking a message / arranging a callback), contact-info is N/A → Yes. Don’t require a phone on a lockout-protocol call. |
| Open-ended | **Lenient** | Caller stated it immediately (“lost my key… can’t get in”). Agent: greeting + name + “let me check the protocol.” Human No. | Same as 265908165. Rec 1 holds. |
| Validate concern | **Lenient** | Agent: **“I’m so sorry that you have been locked up from the apartment.”** That *is* explicit empathy on the actual issue. Human still No. | Do **not** train the model to withhold this. Human No fights the written protocol (concern + sorry → Yes, *or* no emotional concern → N/A Yes). Split validate-lenient cases: operational “I’ll check” (271912426) vs real apology (this call). |

### 266239119 — ✅ transcript in
- **Scores:** Human 61.0% · AI 51.4% · Δ −9.5
- **Agent:** Aki · **Caller:** unnamed third party (uses “sir”) asking about apt **0703**
- **Type:** ~185s — **third-party occupancy / wellness-check fishing.** Friend/husband not responding for a month after an argument; wants names on the lease (Karim Charania / Simran Bahmani). Agent confirms the unit “belongs to Karim…” then “these 2 persons.”
- **File:** `~/Downloads/266239119.json`
- **Both flagged secure-info DQ Yes** — disclosing leaseholder names to a third party. Correct. Not a disagreement.

| Question | Dir | Notes | Implication for Run 3.0 |
|----------|-----|-------|-------------------------|
| Name usage | **Strict** | Agent asked for the *friend’s* name, never the caller’s. Used “sir,” never a personal name. Caller never gave their own name. Human Yes is generous. AI No is closer to the protocol. | Rec 3 holds: caller’s personal name only. Asking who lives in 0703 ≠ name usage. Don’t loosen this to match the human Yes. |
| Closing | **Lenient** | Agent: “any other things or any message… pass on to my colleague?” then truncated “Is there any.” Human No. | Human likely withheld because the call is inappropriate, not because there was no close. AI Yes is protocol-correct. Don’t punish a real “anything else / take a message” close on ugly calls. |
| Validate concern | **Strict** | Agent: “Oh, I’m so sorry, for the shovel” (garbled — likely “sorry for the trouble”) + “we are not regarding this wellness check channel.” Human Yes, AI No. | Transcription miss. Treat phonetic “sorry for the shovel/trouble” as empathy. Opposite of 273290320 (clear sorry that human refused). |

### 273900663 — ✅ transcript in
- **Scores:** Human 83.3% · AI 100.0% · Δ +16.7 (only miss: open-ended)
- **Agent:** Addie · **Caller:** Jimmy Zambrano, apt 9111 · 704-806-8354
- **Type:** ~82s — wants Jacqueline; “problem with my tube,” doesn’t want to explain to the call center. Callback arranged.
- **File:** `~/Downloads/273900663.json`

| Question | Dir | Notes | Implication for Run 3.0 |
|----------|-----|-------|-------------------------|
| Open-ended | **Lenient** | Real diagnostic: **“May I know what is this regards to, Jimmy?”** Plus greeting “How may I assist you?” Identity (unit/phone) and “anything else?” don’t count. That’s **one** probe. Human No. | Textbook Rec 1. Same as 269788776 / 272063323 / 265908165. Greeting + one “what is this regarding?” ≠ two probes. Don’t auto-Yes because he then explained a tube problem. |
| Ownership / closing / validate | Agree Yes | “I’ll be requesting the callback”; “I’m so sorry about that delay, Jimmy”; “Is there anything I can help you with?” | Model already gets the Run 2.0 wins right on this call. |

### 270583790 — ✅ transcript in (perfect agreement)
- **Scores:** Human 76.2% · AI 76.2% · Δ 0
- **Agent:** Vaishnavi · **Caller:** Elizabeth Hernandez, apt 508 · 432-260-6070
- **Type:** ~74s — pool listed as open but dirty. Centralized team takes a message / callback.
- **File:** `~/Downloads/270583790.json`

Positive control — do **not** regress these:

| Question | Both | Why it matters for Run 3.0 |
|----------|------|----------------------------|
| Open-ended | **No** | Caller stated the issue in sentence one. Agent only collected identity. This is the *correct* No — unlike 265908165 / 273900663 where AI over-credited Yes. Rec 1 must not break this. |
| Name usage | **No** | Asked for name, never used “Elizabeth.” Rec 3 ask+use both required. Keep. |
| Validate | **Yes** | “Oh, I am so sorry to hear that.” Empathy on the actual issue (dirty pool). Keep crediting this (contrast: human refused a similar sorry on 273290320). |
| Ownership | **Yes** | “I can pass along this message… arrange a callback.” Run 2.0 language working. |
| Hold | **Yes** | “Let me just get the check” = lookup, not a hold. Run 2.0 hold fix working. |
| Closing | **Yes** | “I have added on the notes and arranged a callback.” Next step stated; file may truncate before “anything else.” |

### 273100258 — ✅ transcript in (perfect agreement, 100%)
- **Scores:** Human 100% · AI 100% · Δ 0
- **Agent:** Manisha (LC) + Jason (onsite maintenance joined) · **Caller:** David Haughey, apt 2035 · 317-363-3898
- **Type:** ~365s — maintenance in the unit with no notice (Ring camera); no new work order. Agent conferences maintenance. June toilet/floor WO.
- **File:** `~/Downloads/273100258.json`

Positive control — the *other* shape of a resident call (not short callback routing):

| Question | Both | Why it matters for Run 3.0 |
|----------|------|----------------------------|
| Open-ended | **Yes** | Caller explained up front, **and** agent asked real follow-ups: “no WO but they were in your apartment without notification?” + “have you talked to maintenance / any active work order?” Rec 1 must still Yes when there are two diagnostic probes. Don’t let “already explained → No” over-fire. |
| Hold | **Yes** | Explicit: “Can I place your call on hold?” Caller: “Sure.” Then a real hold to conference Jason. Contrast with lookup language. |
| Name | **Yes** | Asked and used “David.” |
| Ownership | **Yes** | “I’ll definitely help you with that” + actually got maintenance on the line. |
| Neutral | **Yes** | Jason’s “big guy / army / nervous” banter did **not** ding the LC agent. Don’t score onsite third parties as the agent. |

### 274349154 — ✅ transcript in (perfect agreement, 100%)
- **Scores:** Human 100% · AI 100% · Δ 0
- **Agent:** Candice · **Caller:** Isabel Nichols, bldg 7101 · 720-892-8178 · boyfriend Robert Hill (Ford F-150, gray, arriving 1pm)
- **Type:** ~119s — guest/vehicle registration stuck pending office review. Central team takes a message.
- **File:** `~/Downloads/274349154.json`

Positive control — short callback that still earns open-ended Yes because the agent actually probed:

| Question | Both | Why it matters for Run 3.0 |
|----------|------|----------------------------|
| Open-ended | **Yes** | Ask stated up front, then agent asked real follow-ups: guests vs guest parking? boyfriend’s name? make/model/color? already there or a date? Rec 1 Yes when there are 2+ diagnostic probes after the greeting. Pair with 273100258. |
| Name | **Yes** | “Who am I speaking with?” + used “Isabelle.” |
| Ownership | **Yes** | “I can definitely get a message over to our leasing professionals.” |
| Validate | **Yes** | “I do apologize for the inconvenience or the delay” **up front**, tied to the registration delay. Contrast 265908165’s generic closing “sorry you called.” Specific + early counts; generic close does not. |
| Closing | **Yes** | “Is there anything else I can help you with today?” + “I’ll get that message right on over.” |

---

## Run 3.0 recs — refined after all 20 transcripts

Official run is still the original 12. These 8 only change *how* we write the protocols.

1. **Open-ended (P0)** — Don’t count greeting or identity. Don’t auto-Yes “already explained.” Need **two diagnostic probes after the greeting**. Keep Yes when those probes exist (273100258, 274349154). Keep No on greeting + identity only (270583790 — already correct; 265908165 / 273900663 / 269788776 — AI over-Yes). Don’t lose 269829113’s real “what voicemail should I place?”
2. **Ownership** — Add **voicemail-forward** (“I can forward you a voicemail”) from 269829113. Callback/note language from Run 2.0 still works (270583790, 273900663, 274349154).
3. **Validate** — Split, don’t blunt:
   - Operational “I’ll check / understood” → No (271912426, 270842045)
   - Generic closing “sorry you had to call” → No (265908165)
   - Empathy on the issue (“sorry to hear that” / “sorry for the delay” up front / “sorry you’re locked out”) → Yes (270583790, 274349154, 273290320 — don’t match the human No on the lockout sorry)
   - Garbled “sorry for the shovel/trouble” → Yes (266239119)
4. **Contact N/A** — Policy-only / no callback (273290320 lockout, 269768640 refused callback) → Yes. Don’t require a phone.
5. **Name** — Caller personal name, ask + use. Ask-without-use = No (270583790). Friend/leaseholder names don’t count (266239119). Don’t loosen to match that human Yes.
6. **Hold** — Already good. Lookup ≠ hold (270583790). Real “can I put you on hold?” = Yes (273100258).
7. **Don’t grade third parties** as the LC agent (Jason on 273100258).
8. **Still: same 12 first**, then lock 20. Shadow 20 is 87.5% until Rec 1 lands.

Then re-score the **same 12**, then promote 20.
