# Resident Run 1.0 — AI vs Human Disconnect Log

> Working log for call-by-call transcript review. Use this to drive protocol / model fixes.
> Dashboard: https://alawyer2012.github.io/call-grading-dashboard/ (Residents → Run 1.0)
> Source grades: `~/Downloads/20 Call Resident Comparison.xlsx`
> Definitions: `AI_ QA 2026 (5).xlsx` → AI Resident Fundamentals
> Transcripts: `~/Downloads/<call_id>.json`

**Run snapshot:** 12 matching calls · 78.3% scored agreement · 26 disagreements (20 strict / 6 lenient) · 15.1% avg score delta

**Review status:** All 12 transcripts reviewed (2026-08-11).

---

## Issue themes (ranked by frequency)

| ID | Theme | Direction | Calls | Priority | Likely fix |
|----|-------|-----------|-------|----------|------------|
| **T2** | Hold false positive — “Let me check / Hang on / bear with me / Let me see” treated as hold w/o permission even when no hold occurs (or permission *was* asked) | Strict (AI under-credits) | **7** | **P0** | Only score hold when an actual hold happens; “check/look up” ≠ hold; if agent asks “can I put you on hold?” → Yes |
| **T5** | Closing — human credits next-step explanation / natural end; AI wants formal “anything else?” (or misses it amid crosstalk) | Strict | **5** | **P0** | Credit confirmed next steps + natural close; don’t require perfect “anything else” wording when caller already winds down |
| **T1** | Ownership / acknowledgment — misses real ownership language (`Let me…`, `I can arrange callback`, `I'll send this over`, `I'll get this checked`) | Strict | **4** | **P0** | Expand positive examples beyond canned “I can definitely help / personally follow up” |
| **T6** | Secure-info DQ false positive — AI flags Yes when human says No (often after phone gather/readback) | Lenient (AI over-flags DQ) | **3** | **P0** | Clarifying: confirming caller’s *own* phone/unit back is NOT secure-info disclosure |
| **T7** | Validate-concern over-credit — AI=Yes when human=No (no real concern, or weak empathy) | Lenient | **2** | **P1** | Require explicit concern + explicit validation phrase; else Yes only when no concern (don’t invent credit) |
| **T8** | Name usage mismatch — ask/use rules ambiguous on non-resident callers, company names, wrong-name use | Mixed | **2** | **P1** | Clarify: must use *caller’s* name; company name ≠ personal name; echo alone may not count |
| **T4** | Open-ended — under-credits real probes OR over-credits weak probes | Mixed | **2** | **P1** | Count “How may I assist / what is this regarding?”; apply “already explained” exception; don’t credit closed probes as open |
| **T3** | Contact-info N/A — callback offered/refused or identity gather confused with phone confirm | Strict | **1** | **P1** | If no update/callback being sent → Yes; gathering name/unit ≠ contact confirm |
| **T9** | Greeting miss on crosstalk / truncated open | Strict | **1** | **P2** | Don’t fail greeting when property+name present but overlapped |
| **T10** | Neutral-language over-credit on quirky/unprofessional tone | Lenient | **2** | **P2** | Elevated/odd agent tone without de-escalation need → not automatic Yes; or only auto-Yes when caller was calm |

---

## Proposed fix backlog for Run 2.0 (draft)

### Must-fix before next resident run
1. **Hold protocol rewrite (T2)** — Explicit: lookup language is not a hold. Hold = agent places caller on hold / asks to hold / extended mute. If no hold event → Yes. If “Can I put you on a brief hold?” and caller agrees → Yes.
2. **Ownership expansion (T1)** — Credit first-person action commitments: arrange callback, send to front office/onsite, put notes, get specialist to call, check status.
3. **Secure-info negative examples (T6)** — Reading back caller-provided phone/unit/name is NOT a DQ. DQ = disclosing *other* residents’ data, credentials, account numbers, or confidential property info.
4. **Closing clarification (T5)** — Credit when next steps are stated and call ends cooperatively; “anything else?” is sufficient but not the only path; caller “okay/thanks” after next steps can satisfy.

### Should-fix
5. **Validate-concern guardrails (T7)** — Don’t credit empathy phrases when no concern was expressed *and* human expects N/A handling differently — align N/A→Yes vs requiring a phrase.
6. **Name-usage edge cases (T8)** — Non-resident callers, business names, wrong-name use.
7. **Open-ended tuning (T4)** — Both under- and over-credit showing up; tighten definition with resident examples.

### Open decisions for Austin
- [ ] Is confirming the caller’s own phone number ever a secure-info DQ? (**Recommend: never**)
- [ ] Does “Hang on / give me a second” without an actual hold require permission? (**Recommend: no — not a hold**)
- [ ] For closing: keep merged question, or split next-steps vs final Q again (45-pt card)?
- [ ] Name usage when caller is a third party (e.g. Wendy calling about Keith’s package)?

---

## Theme × call matrix

| Call | T1 Own | T2 Hold | T5 Close | T6 Secure | T7 Validate | T8 Name | T4 Open | Other |
|------|--------|---------|----------|-----------|-------------|---------|---------|-------|
| 269768640 | S | S | S | | | | S | T3 contact S |
| 270543247 | S | S | | L | | S | L | |
| 269788776 | S | S | S | | | | | |
| 270245082 | | S | | | | L | | T10 neutral L |
| 271873957 | S | S | | L | | | | |
| 270842045 | | | | | L | | | T9 greeting S |
| 272053386 | | | S | | | | | T10 neutral L |
| 271912426 | | | | L | L | | | |
| 272063323 | | S | S | | | | | |
| 269245474 | | | S | | | | | |
| 272543647 | | S | | | | | | |
| 272352194 | — perfect agreement — | | | | | | | |

S = strict (AI=No/H=Yes) · L = lenient (AI=Yes/H=No)

---

## Call reviews

### 269768640 — ✅
- **Scores:** Human 100% · AI 50% · Δ −50
- **Agent:** Ankit · **Caller:** Yolanda Webb (wants Jennifer)
- **Type:** ~90s transfer/callback request
- **File:** `~/Downloads/269768640.json`
- **Disagreements (all strict):**

| Question | Pts | Human rationale | AI failure | Theme |
|----------|-----|-----------------|------------|-------|
| Open-ended | 7 | “How may I assist?” + “what was this call in regards?”; need already stated | Missed probes / already-explained exception | T4 |
| Ownership | 5 | “Let me check on her”; “I can arrange… callback” | Canned-phrase matching | T1 |
| Contact info | 4 | Callback refused → N/A→Yes | Required phone because callback discussed | T3 |
| Closing | 3 | Explained can’t transfer / Jennifer unavailable | Wanted formal “anything else?” | T5 |
| Hold | 2 | No hold; “Let me check” = lookup | Treated check as hold | T2 |

---

### 270543247 — ✅
- **Scores:** Human 82.5% · AI 61.0% · Δ −22
- **Agent:** Ankit · **Caller:** John / Pressley Ridge (looking for Barry/Tammy)
- **Type:** ~124s third-party / receptionist lookup → callback arranged
- **File:** `~/Downloads/270543247.json`
- **Disagreements:**

| Question | Dir | Notes | Theme |
|----------|-----|-------|-------|
| Name usage | Strict | Agent asks name; gets company “Pressley Ridge”; later says “Prasvi” (mangled). Human Yes is soft — AI No may be fairer. Clarify third-party/company names. | T8 |
| Open-ended | **Lenient** | AI Yes / Human No — AI over-credited probes on a simple “is Barry available?” call | T4 |
| Ownership | Strict | Clear: “best thing I can do, I can arrange callback”; “I’ll send this over to the front office”; “ask them to call you back” | T1 |
| Hold | Strict | “Just bear with me” / “Let me see” — no real hold | T2 |
| Secure info DQ | **Lenient** | AI Yes / Human No — agent only read back caller’s phone `443-617-3360`. False DQ. | T6 |

---

### 269788776 — ✅
- **Scores:** Human 82.5% · AI 59.5% · Δ −24
- **Agent:** Rohit · **Caller:** David Lindsey unit 201
- **Type:** ~94s “reach the office” → callback for lease-update
- **File:** `~/Downloads/269788776.json`
- **Disagreements (all strict):**

| Question | Notes | Theme |
|----------|-------|-------|
| Ownership | “What I can do is… let the property know”; “I will just arrange a callback”; “put down the notes” | T1 |
| Closing | States next step (“someone from the property will give you a call”) + caller “Okay. Thank you.” — no explicit “anything else?” | T5 |
| Hold | No hold language beyond normal pacing | T2 |

---

### 270245082 — ✅
- **Scores:** Human 85% · AI 95.2% · Δ +14 (AI *higher*)
- **Agent:** Matthew · **Caller:** Wendy (third party re: Keith Sargent / package, apt 119)
- **Type:** ~395s long, echoey/crosstalk-heavy transcript; quirky agent banter
- **File:** `~/Downloads/270245082.json`
- **Disagreements:**

| Question | Dir | Notes | Theme |
|----------|-----|-------|-------|
| Name usage | Lenient | AI Yes / Human No — agent uses “Wendy” but she’s not the resident; calling about Keith. Human may require resident-name handling. | T8 |
| Hold | Strict | Multiple “Hang on / give me a second” — human still Yes (no formal hold / permission implied?). AI No. Reinforces hold over-detection. | T2 |
| Neutral language | Lenient | AI Yes / Human No — agent rants about software (“anger swirling”, “I beat Silicon”). Human correctly withheld credit for professional neutral tone. | T10 |

**Note:** Transcript quality is poor (heavy echo). Model may be confused by duplicated channels.

---

### 271873957 — ✅
- **Scores:** Human 100% · AI 66.7% · Δ −33
- **Agent:** Shauna · **Caller:** Britney Baker unit 232H
- **Type:** ~92s move-out walkthrough + carpet cleaning referral
- **File:** `~/Downloads/271873957.json`
- **Disagreements:**

| Question | Dir | Notes | Theme |
|----------|-----|-------|-------|
| Ownership | Strict | “Let me ask one of the lesion specialists to give you a call and get you all set up” — clear ownership | T1 |
| Hold | Strict | No hold language at all | T2 |
| Secure info DQ | Lenient | AI Yes / Human No — gathered phone `757617617` + unit. False DQ on routine identity collect. | T6 |

---

### 270842045 — ✅
- **Scores:** Human 87.5% · AI 90.5% · Δ +2.4
- **Agent:** (Hanover at Edgewood) · **Caller:** Kaya Sullivan unit 623
- **Type:** ~154s elevator failure / urgent loading dock — **explicit hold with permission**
- **File:** `~/Downloads/270842045.json`
- **Disagreements:**

| Question | Dir | Notes | Theme |
|----------|-----|-------|-------|
| Greeting | Strict | Property + “How can I help you?” present but overlapped with caller “hold one second.” AI missed greeting. | T9 |
| Validate concern | Lenient | AI Yes / Human No — agent was operational (“Understood”, got maintenance en route) but human didn’t credit empathy validation. Possible human strictness on phrase list. | T7 |

**Positive signal:** Agent said “Can I put you on a brief hold?” → caller “Please. Yes.” — both sides agreed hold=Yes here (not in disagreement list). Good contrast vs T2 false positives elsewhere.

---

### 272053386 — ✅
- **Scores:** Human 82.2% · AI 85.7% · Δ +4.8
- **Agent:** Matthew · **Caller:** Julie Richardson apt 910
- **Type:** ~63s lease-renewal callback request (crosstalk/echo)
- **File:** `~/Downloads/272053386.json`
- **Disagreements:**

| Question | Dir | Notes | Theme |
|----------|-----|-------|-------|
| Closing | Strict | Agent sets callback (“I’ll let them know… call you as soon as they can. Is that cool?”) then “Take care” — human Yes; AI wanted stronger close | T5 |
| Neutral language | Lenient | AI Yes / Human No — short calm call; human No is surprising. May be grader noise or reaction to “hectic at work” phrasing. | T10 |

---

### 271912426 — ✅
- **Scores:** Human 77.5% · AI 72.4% · Δ −6
- **Agent:** Sally/Hailey · **Caller:** Jennifer (parcel locker issue)
- **Type:** ~137s package/parcel locker not opening — urgent same-day
- **File:** `~/Downloads/271912426.json`
- **Disagreements (both lenient):**

| Question | Dir | Notes | Theme |
|----------|-----|-------|-------|
| Validate concern | Lenient | Caller frustrated (“happened multiple times”, “need it today”). Agent operational (“I’ll get this checked”) — AI credited validation; human did not (wanted explicit empathy phrase). | T7 |
| Secure info DQ | Lenient | AI Yes / Human No — phone readback `309-287-2769`. Same false DQ pattern. | T6 |

---

### 272063323 — ✅
- **Scores:** Human 82.5% · AI 71.4% · Δ −12
- **Agent:** Rohit · **Caller:** Mia Grotto Duran unit 1137
- **Type:** ~90s fire-extinguisher inspection while out of town → property callback
- **File:** `~/Downloads/272063323.json`
- **Disagreements (both strict):**

| Question | Notes | Theme |
|----------|-------|-------|
| Closing | Next steps explained (“someone from the property will get back”); caller “Sounds good. Thank you.” — no “anything else?” | T5 |
| Hold | “Let me check what can be done” — lookup, not hold | T2 |

---

### 269245474 — ✅
- **Scores:** Human 100% · AI 92.9% · Δ −7
- **Agent:** Ivy · **Caller:** Sathya Sai Manush Malapudi unit 10202
- **Type:** ~92s mailbox key ready? → callback
- **File:** `~/Downloads/269245474.json`
- **Disagreements:**

| Question | Notes | Theme |
|----------|-------|-------|
| Closing | Strict | Next steps stated (“check when keys ready, get back to you”); call trails off mid-sentence on transcript — AI No, human Yes | T5 |

---

### 272543647 — ✅
- **Scores:** Human 100% · AI 95.2% · Δ −5
- **Agent:** Ankit · **Caller:** Robel/Rodel (move-in status / money order)
- **Type:** ~156s callback request with clear “anything else?” close
- **File:** `~/Downloads/272543647.json`
- **Disagreements:**

| Question | Notes | Theme |
|----------|-------|-------|
| Hold | Strict | “Don’t worry. Let me see.” / status check language — no hold. Closing actually good (“Is there anything else…?”). | T2 |

---

### 272352194 — ✅ (perfect agreement)
- **Scores:** Human 84.4% · AI 84.4% · Δ 0
- **Agent:** Avery · **Caller:** Ivan Fuentes (move-in tomorrow)
- **Type:** ~74s callback for move-in next steps
- **File:** `~/Downloads/272352194.json`
- **Notes:** Clean ownership (“I can definitely get someone to reach out”), phone gather, explicit “anything else?” — useful **positive control** for what the model already gets right.

---

## Synthesis for fix discussion

**If we only fixed T2 + T1 + T6 + T5**, we’d address the bulk of Run 1.0’s 26 disagreements:
- Hold alone appears in **7/11** disagreement calls
- Closing in **5**
- Ownership in **4**
- Secure-info false DQ in **3**

**Resident call pattern:** Most benchmark calls are short centralized-team **callback/routing** flows (reach office, package, move-out, inspection) — not sales. Lead-call phrase libraries transfer poorly; resident protocols need callback-native examples.

**Next step when ready:** Walk the open decisions above, then draft Run 2.0 protocol diffs for AI Engineering.
