# Resident Run 3.0 — Evaluation

> Dashboard: https://alawyer2012.github.io/call-grading-dashboard/ (Residents → Overview / Run 3.0)
> Source grades: `~/Downloads/20 Call Resident Comparison (3).xlsx` (New Manual + AI 3)
> Same 20 call IDs as Run 2.0

**Official set:** all 20 matching AI 3 + human IDs
**Snapshot:** **87.0%** scored agreement (174/200) · 26 disagreements (18 strict / 8 lenient) · 9.9% avg score delta
**Original 12:** **85.8%** (103/120) · 17 disagreements (12S / 5L) · 10.7% avg delta · −5.0pp vs Run 2.0’s 90.8%
**New 8:** **88.8%** (71/80) · 9 disagreements (6S / 3L) · 8.8% avg delta · +3.8pp vs Run 2.0’s 85.0%

Missed the 90% / ≤20-disagreement target. Strict errors 12 → 18. Perfect agreement 9 → 4. Mean AI score 86.7% → 79.9%; mean human stayed 86.3%.

---

## Did the Run 2.0 recs land?

| Rec | Intent | Verdict |
|-----|--------|---------|
| 1 Open-ended | Fix 5 leftovers. Keep ops Yes and 2+ probe Yes. | Partial / over-fired. 5 leftovers cleared; 10 new strict Nos. |
| 2 Ownership | Credit voicemail-forward (269829113). | That call fixed. New miss on 273290320 lockout protocol. |
| 3 Validate | Split operational vs real apology. Keep lockout sorry. | Did not move. 5 → 6 disagrees. Lockout sorry still AI Yes. |
| 4 Contact N/A | Policy-only / refused callback → Yes. | Landed. Both 269768640 and 273290320 contact stricts cleared. |
| 5 Name | Don’t loosen 266239119. | Held there. New strict on 271912426. 16/20. |
| 6 Hold | Don’t regress lookup ≠ hold. | Still 20/20. |

---

## Question movement vs Run 2.0

| Question | R2 | R3 | Δ | Notes |
|----------|----|----|---|-------|
| Open-ended | 15/20 · 2S/3L | 10/20 · 10S | −25pp | The miss. Cleared 5 leftovers, dinged 10 prior Yeses. |
| Contact info | 18/20 · 2S | **20/20** | +10pp | N/A on 269768640 and 273290320. |
| Unit number | 19/20 · 1S | **20/20** | +5pp | 269768640 street-address miss cleared. |
| Closing | 17/20 · 2S/1L | 18/20 · 2S | +5pp | 269768640 cleared. New miss on 273100258. |
| Name | 17/20 | 16/20 · 3S/1L | −5pp | New strict on 271912426. |
| Validate | 15/20 · 1S/4L | 14/20 · 1S/5L | −5pp | None of the four lenient leftovers moved. |
| Secure-info DQ | 20/20 | 19/20 · 1S | −5pp | Dropped 265908165 CSO mobile (true positive in R2). |
| Ownership | 19/20 | 19/20 | 0 | Voicemail-forward fixed; lockout miss is new. |
| Greeting / Neutral / Hold / FHA | unchanged | unchanged | 0 | Hold and FHA still perfect. |

---

## The 10 new open-ended stricts (H Yes / AI No)

All of these agreed Yes in Run 2.0.

| Call | Why it matters |
|------|----------------|
| **269788776** | Ops-corrected Yes. Greeting + “what is this regarding?” Rec 1 keep. |
| **272063323** | Same ops case. Rec 1 keep. |
| **273100258** | Positive control. Caller explained, then two diagnostic probes. |
| 269245474, 271873957, 272543647 | Were perfect 100/100. |
| 272352194 | Was perfect agreement. |
| 270245082, 271912426, 272053386 | Added on top of leftover name/neutral/validate noise. |

Keep the four correct Nos: 265908165, 273290320, 273900663, 270583790.

---

## Call highlights

- **269768640** — star. Four stricts (H 100 / AI 57) → perfect. Contact N/A, unit, open-ended, closing.
- **269829113** — ownership voicemail-forward landed (45 → 86). Still closing strict + new validate lenient.
- **273900663** — open-ended leftover correctly No. Now perfect.
- **273100258** — 100/100 control → 76. Open-ended + closing (conference to maintenance).
- **265908165** — open-ended leftover cleared; dropped CSO mobile DQ; validate still lenient.
- **273290320** — contact N/A and open-ended leftover cleared. New ownership miss. Score delta 0 because ownership −5 and validate +5 cancel.

---

## Run 4.0 recs — same 20 calls

1. **Open-ended (P0)** — “Already explained → No” must not veto a later diagnostic probe or “what is this regarding?” Restore the 10 Yeses. Keep the four Nos above. Keep 269768640 / 269829113 / 274349154 as Yes.
2. **Secure-info** — Restore on-site staff personal/direct numbers (265908165 CSO mobile 813-992-9423). Caller phone readback is still not a DQ.
3. **Ownership** — Contact N/A on a lockout is not an ownership fail. Keep 273290320 Yes for explaining protocol.
4. **Closing** — Conference / committed next step is enough (273100258, 269829113 truncated voicemail). Don’t over-require a wrap-up.

Leave hold, FHA, contact, and unit alone. Validate and name are leftover noise, not why this run missed 90%. Reverse the 10 open-ended stricts and this set is **16 disagreements / 92%** before any other work.
