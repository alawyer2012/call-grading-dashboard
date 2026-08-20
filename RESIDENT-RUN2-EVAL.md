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

No transcripts yet for the 8 new IDs.

---

## Run 3.0 must-fix (copy-paste is on the dashboard Recs tab)

1. **Open-ended** — don’t count greeting or identity questions; don’t auto-Yes “already explained” on simple routing.
2. **Validate concern** — operational “Understood / I’ll get this checked” is not empathy. (No blue cell in Run 2, so it didn’t move.)
3. **Name usage** — caller’s personal name only; company / third-party resident names don’t count.
4. **Neutral language** — unprofessional/quirky tone is not an auto-Yes just because the caller wasn’t escalated.
5. **269768640 + 270842045** — transcript-noise: asked-for unit counts; refused callback → contact Yes; overlapped greeting still counts.

Then re-score the **same 12**, then promote 20.
