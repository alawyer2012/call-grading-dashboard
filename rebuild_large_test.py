#!/usr/bin/env python3
"""
Append (or refresh) the Residents → Large Test dashboard run from the
July/August 490-call comparison tabs.

Does NOT rebuild Runs 1–8. Hand-written recs / rootCause on those runs stay put.

Source: ~/Downloads/20 Call Resident Comparison (10).xlsx
  - Manual Grades JulyAugust Reside
  - AI JulyAugust Resident Simuluat
  - Graded stats (Yes-rate check only)
"""

import openpyxl
import re
from datetime import date

XLSX_PATH = "/Users/alawyer/Downloads/20 Call Resident Comparison (10).xlsx"
HTML_PATH = "/Users/alawyer/Entrata PM/Dashboard/call-grading/index.html"

RUN_ID = 90
MANUAL_TAB = "Manual Grades JulyAugust Reside"
AI_TAB = "AI JulyAugust Resident Simuluat"

# Human tab column order is NOT the same as the AI tab.
HUMAN_COL_MAP = {
    3: "hold_permission",
    4: "validate_concern",
    5: "greeting",
    6: "name_usage",
    7: "contact_info",
    8: "unit_number",
    9: "reason_for_call",
    10: "acknowledged",
    11: "next_steps",
    12: "final_closing_q",  # human-only — not in AI tab
    13: "fha",
    14: "secure_info",
}

AI_COL_MAP = {
    2: "greeting",
    3: "name_usage",
    4: "contact_info",
    5: "unit_number",
    6: "reason_for_call",
    7: "acknowledged",
    8: "next_steps",
    9: "neutral_language",  # AI-only — not in human tab
    10: "hold_permission",
    11: "validate_concern",
    12: "fha",
    13: "secure_info",
}

# Overlapping scored questions (37 pts). Neutral (5) and final-closing are
# one-sided, so they are excluded from agreement and from the weighted score.
WEIGHTS = {
    "greeting": 4,
    "name_usage": 3,
    "contact_info": 4,
    "unit_number": 4,
    "reason_for_call": 7,
    "acknowledged": 5,
    "next_steps": 3,
    "hold_permission": 2,
    "validate_concern": 5,
}
SCORED_KEYS = list(WEIGHTS.keys())
DQ_KEYS = ["fha", "secure_info"]
ALL_KEYS = SCORED_KEYS + DQ_KEYS
TOTAL_POINTS = sum(WEIGHTS.values())  # 37

Q_LABELS = {
    "greeting": "Greeting",
    "name_usage": "Name usage",
    "contact_info": "Contact info",
    "unit_number": "Unit number",
    "reason_for_call": "Reason for the call",
    "acknowledged": "Acknowledged/ownership",
    "next_steps": "Next steps",
    "hold_permission": "Hold permission",
    "validate_concern": "Validate concern",
    "fha": "FHA violation (DQ)",
    "secure_info": "Secure info (DQ)",
}

Q_FULL_LABELS = {
    "greeting": "Greeting (property name + intro)",
    "name_usage": "Asked for name + used it",
    "contact_info": "Contact info confirmed (if update requested)",
    "unit_number": "Unit / apartment number confirmed",
    "reason_for_call": "Reason for the call",
    "acknowledged": "Acknowledged caller / took ownership",
    "next_steps": "Next steps / role / what to expect",
    "hold_permission": "Asked permission before hold",
    "validate_concern": "Validated caller concern (if expressed)",
}

STRICT_REASONS = {
    "greeting": "Agent greeted with property name and intro. AI did not detect it.",
    "name_usage": "Human credited asked-for + used name. AI marked No.",
    "contact_info": "Agent confirmed contact info. AI did not credit it.",
    "unit_number": "Unit/apartment number was confirmed. AI missed it.",
    "reason_for_call": "Human credited capturing the reason for the call. AI marked No.",
    "acknowledged": "Agent acknowledged caller / took ownership. AI did not credit it.",
    "next_steps": "Human credited next steps / role / expectation. AI marked No.",
    "hold_permission": "Agent asked hold permission (or hold not needed). AI did not credit it.",
    "validate_concern": "Human credited concern validation. AI did not.",
}

LENIENT_REASONS = {
    "greeting": "AI credited greeting; human says it was insufficient.",
    "name_usage": "AI credited name usage; human says name was not properly used.",
    "contact_info": "AI credited contact confirmation; human disagrees.",
    "unit_number": "AI credited unit confirmation; human disagrees.",
    "reason_for_call": "AI credited capturing the reason; human marked No.",
    "acknowledged": "AI credited acknowledgment; human required more explicit ownership.",
    "next_steps": "AI credited next steps; human did not.",
    "hold_permission": "AI credited hold permission; human says it was not properly asked.",
    "validate_concern": "AI credited concern validation; human did not.",
    "fha": "AI flagged an FHA violation; human did not find one.",
    "secure_info": "AI flagged secure info disclosure; human did not find one.",
}


def parse_yes_no(val):
    if val is None or str(val).strip() == "":
        return None
    return 1 if str(val).strip().lower().startswith("y") else 0


def call_id_str(val):
    if val is None:
        return None
    try:
        return str(int(float(val)))
    except (TypeError, ValueError):
        s = str(val).strip()
        return s or None


def calc_score(answers):
    earned = sum(WEIGHTS[q] for q in SCORED_KEYS if answers.get(q) == 1)
    base = (earned / TOTAL_POINTS) * 100 if TOTAL_POINTS else 0
    dq_count = sum(1 for dq in DQ_KEYS if answers.get(dq) == 1)
    return round(base * max(0, 1 - 0.20 * dq_count), 2)


def js_escape(s):
    return (
        s.replace("\\", "\\\\")
        .replace("'", "\\'")
        .replace('"', '\\"')
        .replace("`", "\\`")
        .replace("\n", " ")
    )


def load_tab(ws, id_col, col_map, skip_dups_label):
    data = {}
    dups = 0
    for row in range(2, ws.max_row + 1):
        cid = call_id_str(ws.cell(row, id_col).value)
        if not cid:
            continue
        if cid in data:
            dups += 1
            continue
        answers = {k: parse_yes_no(ws.cell(row, c).value) for c, k in col_map.items()}
        data[cid] = answers
    print(f"  {skip_dups_label}: {len(data)} unique, {dups} duplicate rows skipped")
    return data


print("Reading spreadsheet...")
wb = openpyxl.load_workbook(XLSX_PATH, data_only=True)
print(f"  Tabs: {wb.sheetnames}")
human_raw = load_tab(wb[MANUAL_TAB], 1, HUMAN_COL_MAP, "Human")
ai_raw = load_tab(wb[AI_TAB], 1, AI_COL_MAP, "AI")
both = sorted(set(human_raw) & set(ai_raw))
print(f"  Matching IDs: {len(both)}  human-only={len(set(human_raw)-set(ai_raw))}  ai-only={len(set(ai_raw)-set(human_raw))}")

answer_data = {}
for cid in both:
    h = human_raw[cid]
    a = ai_raw[cid]
    answer_data[cid] = {q: [h.get(q), a.get(q)] for q in ALL_KEYS}

# ── metrics ──
total_agree = total_dis = total_s = total_l = total_comp = 0
dq_dis = 0
score_deltas = []
calls_data = []
q_acc = {
    q: {"agree": 0, "dis": 0, "strict": 0, "lenient": 0, "total": 0, "strict_ids": [], "lenient_ids": []}
    for q in ALL_KEYS
}

for cid in both:
    cd = answer_data[cid]
    h_score = calc_score({q: cd[q][0] for q in ALL_KEYS})
    a_score = calc_score({q: cd[q][1] for q in ALL_KEYS})
    score_deltas.append(abs(h_score - a_score))
    strict = lenient = 0
    details = []
    for q in ALL_KEYS:
        h_val, a_val = cd[q]
        if h_val is None or a_val is None:
            continue
        is_dq = q in DQ_KEYS
        q_acc[q]["total"] += 1
        if not is_dq:
            total_comp += 1
        if h_val == a_val:
            q_acc[q]["agree"] += 1
            if not is_dq:
                total_agree += 1
            continue
        q_acc[q]["dis"] += 1
        if is_dq:
            dq_dis += 1
        else:
            total_dis += 1
        w = WEIGHTS.get(q, "DQ")
        if h_val == 1 and a_val == 0:
            strict += 1
            q_acc[q]["strict"] += 1
            q_acc[q]["strict_ids"].append(cid)
            if not is_dq:
                total_s += 1
            details.append(
                {
                    "q": Q_LABELS[q],
                    "w": "DQ" if is_dq else w,
                    "ai": "No",
                    "h": "Yes",
                    "reason": STRICT_REASONS.get(q, "AI missed this behavior."),
                    "key": q,
                }
            )
        else:
            lenient += 1
            q_acc[q]["lenient"] += 1
            q_acc[q]["lenient_ids"].append(cid)
            if not is_dq:
                total_l += 1
            details.append(
                {
                    "q": Q_LABELS[q],
                    "w": "DQ" if is_dq else w,
                    "ai": "Yes",
                    "h": "No",
                    "reason": LENIENT_REASONS.get(q, "AI over-credited this behavior."),
                    "key": q,
                }
            )

    weight_lost = sum(d["w"] for d in details if d["w"] != "DQ")
    has_dq = any(d["w"] == "DQ" for d in details)
    n_dis = strict + lenient
    if n_dis == 0:
        note = "Perfect agreement on every overlapping question."
        rec = "No action needed — perfect agreement."
    else:
        bits = []
        if strict:
            bits.append(f"{strict} strict")
        if lenient:
            bits.append(f"{lenient} lenient")
        note = f"{n_dis} disagreement{'s' if n_dis != 1 else ''}: {', '.join(bits)}."
        rec_parts = []
        s_qs = sorted([d for d in details if d["ai"] == "No" and d["w"] != "DQ"], key=lambda x: -x["w"])[:3]
        l_qs = sorted([d for d in details if d["ai"] == "Yes" and d["w"] != "DQ"], key=lambda x: -x["w"])[:3]
        if s_qs:
            rec_parts.append("Strict: " + ", ".join(f"{d['q']} ({d['w']}pts)" for d in s_qs))
        if l_qs:
            rec_parts.append("Lenient: " + ", ".join(f"{d['q']} ({d['w']}pts)" for d in l_qs))
        rec = ". ".join(rec_parts) + "."

    calls_data.append(
        {
            "id": cid,
            "human": h_score,
            "ai": a_score,
            "strict": strict,
            "lenient": lenient,
            "weightLost": weight_lost,
            "disqualifier": has_dq,
            "note": note,
            "details": details,
            "recommendation": rec,
        }
    )

questions_data = []
for q in ALL_KEYS:
    acc = q_acc[q]
    if acc["dis"] == 0:
        # still include hold/FHA-class near-perfects so the table is complete
        direction = "mixed"
    elif acc["strict"] > acc["lenient"]:
        direction = "strict"
    elif acc["lenient"] > acc["strict"]:
        direction = "lenient"
    else:
        direction = "mixed"
    questions_data.append(
        {
            "short": q,
            "label": Q_FULL_LABELS.get(q, Q_LABELS[q]),
            "weight": WEIGHTS.get(q, 0),
            "agree": acc["agree"],
            "disagree": acc["dis"],
            "total": acc["total"],
            "strict": acc["strict"],
            "lenient": acc["lenient"],
            "dir": direction,
            "strict_ids": acc["strict_ids"],
            "lenient_ids": acc["lenient_ids"],
        }
    )
questions_data.sort(key=lambda q: (-q["disagree"], -q["weight"]))

agreement_pct = round((total_agree / total_comp) * 100, 1) if total_comp else 0
avg_delta = round(sum(score_deltas) / len(score_deltas), 1) if score_deltas else 0
perfect = sum(1 for c in calls_data if c["strict"] == 0 and c["lenient"] == 0)
mean_h = round(sum(c["human"] for c in calls_data) / len(calls_data), 1)
mean_a = round(sum(c["ai"] for c in calls_data) / len(calls_data), 1)
need_90 = max(0, int(0.9 * total_comp - total_agree + 0.999))

print("\n" + "=" * 70)
print("LARGE TEST METRICS")
print("=" * 70)
print(f"  Calls: {len(both)}")
print(f"  Agreement: {agreement_pct}% ({total_agree}/{total_comp})")
print(f"  Disagreements: {total_dis}  strict {total_s}  lenient {total_l}  DQ {dq_dis}")
print(f"  Avg |delta|: {avg_delta}%   mean H {mean_h}  mean AI {mean_a}")
print(f"  Perfect: {perfect}/{len(both)}")
print(f"  90% needs {need_90} more scored agreements")
for q in questions_data:
    print(
        f"  {q['short']:20} {q['agree']}/{q['total']} = {round(q['agree']/q['total']*100,1)}%  "
        f"S{q['strict']} L{q['lenient']}"
    )

# ── writeups ──
qmap = {q["short"]: q for q in questions_data}
name_q = qmap["name_usage"]
next_q = qmap["next_steps"]
reason_q = qmap["reason_for_call"]
val_q = qmap["validate_concern"]
greet_q = qmap["greeting"]
unit_q = qmap["unit_number"]
ack_q = qmap["acknowledged"]
contact_q = qmap["contact_info"]
hold_q = qmap["hold_permission"]

name_pct = round(name_q["agree"] / name_q["total"] * 100, 1)
next_pct = round(next_q["agree"] / next_q["total"] * 100, 1)
reason_pct = round(reason_q["agree"] / reason_q["total"] * 100, 1)
val_pct = round(val_q["agree"] / val_q["total"] * 100, 1)
strict_share = round(total_s / total_dis * 100) if total_dis else 0
name_w = name_q["weight"] * name_q["strict"]
reason_w = reason_q["weight"] * reason_q["strict"]
next_w = next_q["weight"] * next_q["strict"]
val_w = val_q["weight"] * val_q["lenient"]
top3_s = name_q["strict"] + next_q["strict"] + reason_q["strict"]

key_findings = f"""<p><strong>Large Test holds at {agreement_pct}% on 490 calls — the 20-call Run 8.0 result was not a small-n illusion.</strong> Spreadsheet tabs <strong>Manual Grades JulyAugust Reside</strong> vs <strong>AI JulyAugust Resident Simuluat</strong>. Scored agreement <strong>{agreement_pct}%</strong> ({total_agree}/{total_comp}) across 9 overlapping questions. Disagreements: <strong>{total_dis}</strong> ({total_s} strict / {total_l} lenient). Avg absolute score delta <strong>{avg_delta} pp</strong> on a 37-pt overlapping card. Perfect agreement: <strong>{perfect} of {len(both)}</strong> ({round(perfect/len(both)*100,1)}%). Mean AI {mean_a}% vs mean human {mean_h}%.</p>
<p><strong>The model is too strict at scale, not too loose.</strong> {strict_share}% of disagreements are AI=No / human=Yes ({total_s} of {total_dis}). Human is the reference, not the truth — but a {strict_share}/{100-strict_share} split that large is the model under-crediting agents. Run 8.0 on 20 calls was roughly even (10 strict / 11 lenient). The 490-call set exposes the real bias.</p>
<p><strong>Three questions are the work.</strong> Name usage <strong>{name_pct}%</strong> ({name_q["agree"]}/{name_q["total"]}, {name_q["strict"]}S/{name_q["lenient"]}L) — worst agreement, AI Yes-rate 69% vs human 91%. Next steps <strong>{next_pct}%</strong> ({next_q["agree"]}/{next_q["total"]}, {next_q["strict"]}S/{next_q["lenient"]}L) — AI 81% vs human 97%. Reason for the call <strong>{reason_pct}%</strong> ({reason_q["agree"]}/{reason_q["total"]}, {reason_q["strict"]}S/{reason_q["lenient"]}L) — AI 85% vs human 95%, and the highest weighted miss because it is 7 pts ({reason_w} strict-pts). Those three are {top3_s} of {total_s} strict errors ({round(top3_s/total_s*100) if total_s else 0}%).</p>
<p><strong>Validate concern is the only big lenient miss</strong> — {val_pct}% ({val_q["agree"]}/{val_q["total"]}, {val_q["strict"]}S/{val_q["lenient"]}L). AI Yes-rate 99% vs human 86%. The auto-Yes-when-no-concern rule is over-firing. Contact {round(contact_q["agree"]/contact_q["total"]*100,1)}%, hold {round(hold_q["agree"]/hold_q["total"]*100,1)}%, FHA 100%, secure-info 99.4% — leave them alone.</p>
<p><strong>Scorecard caveat:</strong> humans also graded a 10th item (final closing question, 92% Yes) that AI never answered. AI also graded neutral language (99.8% Yes) that humans never answered. Agreement below uses the 9 shared scored questions + 2 DQs (37 pts, not the 42-pt 20-call card). Graded-stats Yes-rates on the sheet match these Yes-rates. <strong>→ Open Recommendations</strong> — 90% is {need_90} more scored agreements. Recovering half of name strict + half of next-steps strict clears it.</p>"""

root_strict = f"""<p><strong>Name usage ({name_q["strict"]} strict / {name_q["lenient"]} lenient, {name_w} weighted pts):</strong> Lowest agreement on the card ({name_pct}%). AI misses 26% of the calls humans marked Yes. Pattern from the 20-call leftovers, now at volume: volunteered names, one-time confirmatory use (“Thank you, Mia”), and “can you confirm your first name?” plus a later use. Do not require a sales-style ask + reuse pair.</p>
<p><strong>Next steps ({next_q["strict"]} strict / {next_q["lenient"]} lenient, {next_w} weighted pts):</strong> This is the old closing problem with the new wording (“explain next steps, their role, or what the caller should expect”). Humans credit callback / note-to-property / answering-service-will-have-onsite-follow-up. AI still wants a more formal recap. {next_q["strict"]} false Nos vs only {next_q["lenient"]} false Yeses — loosen, do not add negative examples first.</p>
<p><strong>Reason for the call ({reason_q["strict"]} strict / {reason_q["lenient"]} lenient, {reason_w} weighted pts):</strong> Highest weighted miss. Unlike Run 8.0, both sides graded the same question (not human-open-ended vs AI-reason). Remaining gap is still AI too strict: acting on the caller’s opening issue (package, notice, inspection, work order) should count even without a recap sentence. Hold the {reason_q["lenient"]} false Yeses — do not copy AI 7’s over-credit.</p>
<p><strong>Second-tier strict cluster:</strong> Greeting {greet_q["strict"]}S/0L, unit {unit_q["strict"]}S/0L, ownership {ack_q["strict"]}S/{ack_q["lenient"]}L. Real misses, but smaller than the top three. Do not retune these in the same paste as name / next steps / reason — one lever at a time or we will not know what moved.</p>"""

root_lenient = f"""<p><strong>Validate concern ({val_q["lenient"]} lenient / {val_q["strict"]} strict, {val_w} weighted over-credit pts):</strong> The only question where AI is systematically generous. Protocol says auto-Yes when no concern was expressed. Humans still marked No on {val_q["total"] - (val_q["agree"] - 0) - val_q["strict"]} of these calls (human Yes-rate 85.5%). Either the caller raised a real concern and the agent never validated it, or humans are not applying the auto-Yes. Treat this as “do not auto-Yes through frustration / urgency / a stated problem.” Specific validating phrase required when concern language is present.</p>
<p><strong>Everything else lenient is noise.</strong> Contact {contact_q["lenient"]}L, reason {reason_q["lenient"]}L, next steps {next_q["lenient"]}L, name {name_q["lenient"]}L, secure-info 3, hold 2. Do not write protocols for these. Tightening them will cost Yeses we need on the strict side.</p>
<p>Human is the reference. On validate, humans may be the ones being harsh. On name / next steps / reason, the Yes-rate gaps (22 / 16 / 10 pp) are too large to pin on grader noise.</p>"""

recs = [
    {
        "num": 1,
        "title": f"Name usage — {name_pct}% is the largest gap ({name_q['strict']} strict)",
        "severity": "critical",
        "severityLabel": f"{name_q['agree']}/{name_q['total']} · {name_q['strict']} strict / {name_q['lenient']} lenient · AI Yes 69% vs human 91%",
        "owner": "AI Engineering",
        "ownerClass": "info",
        "problem": f"""<p>Name is the worst question on 490 calls. Agreement {name_pct}%. AI misses <strong>26%</strong> of human-Yes names ({name_q['strict']}/{name_q['agree'] + name_q['strict']}). False Yeses are only {name_q['lenient']} — this is not a loosen-vs-tighten toss-up. Recovering all {name_q['strict']} stricts would land at {round((total_agree + name_q['strict']) / total_comp * 100, 1)}%. Recovering half ({name_q['strict'] // 2}) gets to {round((total_agree + name_q['strict'] // 2) / total_comp * 100, 1)}%, one point shy of 90% by itself.</p>
<p>Same pattern we already saw on the 20-call set (16/20), now undeniable at n=490: one-time confirmatory use and volunteered names.</p>""",
        "protocols": [
            {
                "label": "Name usage — credit one-time use + volunteered names",
                "current": "Both required: (1) agent asked for the name or caller offered it, AND (2) agent used it at least once.",
                "recommended": (
                    "Keep both conditions, but mark YES when either is clearly met in a resident service call:<br><br>"
                    "1. Caller volunteers a name and the agent uses it once (“Thanks, Mia” / “Okay, John”). That is asked-or-offered + used.<br>"
                    "2. Agent confirms (“Can I get your first name?” / “Can you confirm your name?”) and later uses it once. Do not require a third reuse.<br>"
                    "3. Phonetic / transcript-garbled names still count if the agent is clearly addressing the caller by name.<br><br>"
                    f"Do <strong style=\"color:var(--red);\">not</strong> require sales-style repeated name use. Gate: name ≥ 85% agreement on this 490, strict &lt; 70. Leave the {name_q['lenient']} false Yeses unless they move with the loosen."
                ),
            }
        ],
    },
    {
        "num": 2,
        "title": f"Next steps — {next_pct}% , {next_q['strict']} false Nos (the closing problem at scale)",
        "severity": "critical",
        "severityLabel": f"{next_q['agree']}/{next_q['total']} · {next_q['strict']} strict / {next_q['lenient']} lenient · AI Yes 81% vs human 97%",
        "owner": "AI Engineering",
        "ownerClass": "info",
        "problem": f"""<p>Human Yes-rate on next steps is 96.9%. AI is 81.0%. Almost every miss is a false No ({next_q['strict']} vs {next_q['lenient']} false Yes). This is Run 8.0 closing (17/20) multiplied by 25. Resident calls are message-taking: arrange callback, leave a note, explain “I am the answering service / the property will call you.” That is the next step.</p>
<p>Half of these {next_q['strict']} plus half of name strict is {name_q['strict'] // 2 + next_q['strict'] // 2} agreements → {round((total_agree + name_q['strict'] // 2 + next_q['strict'] // 2) / total_comp * 100, 1)}%, which clears 90% without touching reason.</p>""",
        "protocols": [
            {
                "label": "Next steps — callback / note / role counts",
                "current": "Did the agent clearly explain at least one of: the next steps, their role, or what the caller should expect moving forward?",
                "recommended": "Mark YES if the agent does any one of these:<br><br>1. States a follow-up (“I will have the property call you back,” “I am leaving a note,” “someone from the office will reach out”).<br>2. Explains their role (“I am with the after-hours answering service; on-site will follow up”).<br>3. Tells the caller what happens next on the issue they called about.<br><br>A cooperative ending plus a stated follow-up is enough. “Anything else?” is <strong style=\"color:var(--red);\">not</strong> required. Truncated / voicemail-forward still counts. Gate: next steps ≥ 90% on this 490, strict &lt; 40.",
            }
        ],
    },
    {
        "num": 3,
        "title": f"Reason for the call — {reason_pct}%, highest weighted miss ({reason_w} pts)",
        "severity": "critical",
        "severityLabel": f"{reason_q['agree']}/{reason_q['total']} · {reason_q['strict']} strict / {reason_q['lenient']} lenient · 7-pt slot",
        "owner": "AI Engineering",
        "ownerClass": "info",
        "problem": f"""<p>Unlike Run 8.0, both sides graded “Did the agent capture the reason for the call?” — not human-open-ended vs AI-reason. Agreement is better than that 15/20 (75%) apples-to-oranges number, but still {reason_pct}% with {reason_q['strict']} false Nos. At 7 pts this is the most expensive strict question on the card ({reason_w} weighted pts vs name {name_w} vs next steps {next_w}).</p>
<p>Keep the new question. Do not revert to 2+ open-ended. Do not ship the AI 7 over-credit pattern ({reason_q['lenient']} false Yeses already exist; do not add more).</p>""",
        "protocols": [
            {
                "label": "Reason for the call — acting on the opening issue counts",
                "current": "Did the agent capture the reason for the call?",
                "recommended": (
                    "Mark YES if the agent restates, confirms, or <strong style=\"color:var(--red);\">handles</strong> the reason the caller phoned. "
                    "The caller’s opening statement plus the agent proceeding on that issue (package, work order, notice, inspection, callback, payment, maintenance) is capture.<br><br>"
                    "Do not require a recap sentence (“So you are calling about…”). Do not mark No just because the reason was obvious from the first turn.<br><br>"
                    f"Keep No when the agent never engages the actual issue. Gate: reason ≥ 90% on this 490, strict &lt; 30, lenient still ≤ {reason_q['lenient']}."
                ),
            }
        ],
    },
    {
        "num": 4,
        "title": f"Validate concern — only major lenient miss ({val_q['lenient']} false Yeses)",
        "severity": "warning",
        "severityLabel": f"{val_q['agree']}/{val_q['total']} · {val_q['strict']} strict / {val_q['lenient']} lenient · AI Yes 99% vs human 86%",
        "owner": "AI Engineering",
        "ownerClass": "info",
        "problem": f"""<p>Validate is the mirror image of name/next/reason. AI Yes-rate 98.6% vs human 85.5%. {val_q['lenient']} of {total_l} lenient errors ({round(val_q['lenient']/total_l*100) if total_l else 0}%) live here. The auto-Yes when no concern was expressed is eating calls where humans heard frustration, urgency, or a problem and wanted a validating phrase.</p>
<p>Do not retune this in the same paste as recs 1–3 if you need a clean read on those. Second paste. Target: cut lenient ~in half without creating a new strict pile.</p>""",
        "protocols": [
            {
                "label": "Validate concern — auto-Yes only when no concern language",
                "current": "If the caller expressed a concern, agent must use at least one specific validating phrase. If no concerns were expressed, always answer Yes.",
                "recommended": "Keep auto-Yes for purely transactional calls with no concern / frustration / urgency language.<br><br>Mark NO when the caller states a problem, delay, missed callback, unsafe condition, billing issue, or clear frustration and the agent does not acknowledge it with a validating phrase (“I understand,” “I am sorry you are dealing with that,” “that makes sense”). Proceeding to a callback without acknowledgment is not validation.<br><br>Gate: validate agreement ≥ 90%, lenient &lt; 35, strict still ≤ 10.",
            }
        ],
    },
    {
        "num": 5,
        "title": "Do not retune hold, contact, FHA, greeting, unit, or ownership in the next paste",
        "severity": "success",
        "severityLabel": f"Hold {round(hold_q['agree']/hold_q['total']*100,1)}% · contact {round(contact_q['agree']/contact_q['total']*100,1)}% · FHA 100% · greeting/unit/ack are second-tier strict",
        "owner": "Austin + AI Engineering",
        "ownerClass": "info",
        "problem": f"""<p>Hold, FHA, and contact are done. Greeting ({greet_q['strict']}S), unit ({unit_q['strict']}S), and ownership ({ack_q['strict']}S) are real but smaller. Mixing them into the name/next/reason paste will make the next 490-call read uninterpretable.</p>
<p><strong>90% math:</strong> {total_agree}/{total_comp} today. Need {need_90} more scored agreements (3969/{total_comp}). Name+next-steps half-recovery is {name_q['strict']//2 + next_q['strict']//2} → {round((total_agree + name_q['strict']//2 + next_q['strict']//2)/total_comp*100,1)}%. All name strict recovered is {round((total_agree + name_q['strict'])/total_comp*100,1)}% by itself.</p>
<p>Re-run on this same 490 after recs 1–3. Same human tab. New AI tab. Success = ≥90% scored agreement, name ≥85%, next steps ≥90%, reason ≥90%, validate not worse than {val_pct}%.</p>""",
        "protocols": [
            {
                "label": "Large Test execution checklist",
                "current": f"Large Test: {agreement_pct}% on {len(both)} calls, {total_dis} disagrees ({total_s}S/{total_l}L), {avg_delta}% avg delta. Name {name_pct}%. Next steps {next_pct}%. Reason {reason_pct}%. Validate {val_pct}%.",
                "recommended": "One paste, three questions: name (loosen one-time/volunteered), next steps (callback/note/role = Yes), reason (acting on the opening issue = Yes). Do not revert reason to open-ended. Do not import AI 7. Do not retune validate until the next cycle. Same 490 IDs. Success = 90%+, name ≥ 85%, next steps ≥ 90%, reason ≥ 90%.",
            }
        ],
    },
]

# ── JS emit ──
def emit_recs(recs_list):
    chunks = []
    for r in recs_list:
        proto_js = []
        for p in r.get("protocols") or []:
            proto_js.append(
                "          {\n"
                f'            label: "{js_escape(p["label"])}",\n'
                f'            current: "{js_escape(p["current"])}",\n'
                f'            recommended: `{p["recommended"]}`\n'
                "          }"
            )
        proto_block = ",\n".join(proto_js)
        chunks.append(
            f"""      {{
        num: {r['num']},
        title: "{js_escape(r['title'])}",
        severity: "{r['severity']}",
        severityLabel: "{js_escape(r['severityLabel'])}",
        owner: "{r['owner']}",
        ownerClass: "{r['ownerClass']}",
        problem: `{r['problem']}`,
        protocols: [
{proto_block}
        ]
      }}"""
        )
    return "    recommendations: [\n" + ",\n".join(chunks) + "\n    ]"


def emit_questions(qs):
    lines = []
    for q in qs:
        s_ids = ",".join(f'"{i}"' for i in q["strict_ids"])
        l_ids = ",".join(f'"{i}"' for i in q["lenient_ids"])
        lines.append(
            f'      {{ short: "{q["short"]}", label: "{js_escape(q["label"])}", weight: {q["weight"]}, '
            f'agree: {q["agree"]}, disagree: {q["disagree"]}, total: {q["total"]}, '
            f'lenient: {q["lenient"]}, strict: {q["strict"]}, dir: "{q["dir"]}",\n'
            f"        evidence: [], strictCalls: [{s_ids}], lenientCalls: [{l_ids}] }}"
        )
    return "    questions: [\n" + ",\n".join(lines) + "\n    ]"


def emit_calls(calls):
    lines = []
    for c in calls:
        details_parts = []
        for d in c["details"]:
            w_js = f'"{d["w"]}"' if d["w"] == "DQ" else str(d["w"])
            details_parts.append(
                f'          {{ q: "{js_escape(d["q"])}", w: {w_js}, ai: "{d["ai"]}", h: "{d["h"]}", reason: "{js_escape(d["reason"])}" }}'
            )
        details_str = ",\n".join(details_parts)
        if details_str:
            details_str = "\n" + details_str + "\n        "
        lines.append(
            f'      {{ id: "{c["id"]}", human: {c["human"]}, ai: {c["ai"]}, strict: {c["strict"]}, lenient: {c["lenient"]}, weightLost: {c["weightLost"]}, disqualifier: {"true" if c["disqualifier"] else "false"},\n'
            f"        note: '{js_escape(c['note'])}',\n"
            f"        details: [{details_str}],\n"
            f"        recommendation: '{js_escape(c['recommendation'])}' }}"
        )
    return "    calls: [\n" + ",\n".join(lines) + "\n    ]"


run_js = f"""  {{
    id: {RUN_ID},
    label: "Large Test",
    date: "September 1, 2026",
    largeEval: true,
    totalPoints: {TOTAL_POINTS},
    scoredQuestions: {len(SCORED_KEYS)},
    criteriaCount: {len(ALL_KEYS)},
    description: "July/August resident simulation — {len(both)} matching human vs AI grades on the overlapping 9-question card (reason-for-call scorecard, not the 20-call protocol series).",
    changes: "New 490-call eval, not a protocol run. Compared Manual Grades JulyAugust Reside vs AI JulyAugust Resident Simuluat. Shared questions only (37 pts). Human-only final-closing and AI-only neutral language excluded.",
    meta: {{
      agreement: {agreement_pct},
      totalDisagreements: {total_dis},
      avgDelta: {avg_delta},
      strictErrors: {total_s},
      lenientErrors: {total_l},
      target: 90
    }},
    keyFindings: `{key_findings}`,
{emit_questions(questions_data)},
{emit_calls(calls_data)},
    recTargetDisagreements: {int(total_comp * 0.10)},
{emit_recs(recs)},
    rootCause: {{
      overall: "Large Test is {len(both)} July/August resident calls at {agreement_pct}% scored agreement — {need_90} agreements short of 90%. Not comparable to the 20-call protocol series. Dominant bias is AI too strict ({strict_share}% of disagrees). Name, next steps, and reason-for-call are the work. Validate is the only large lenient miss.",
      what_worked: "Hold {round(hold_q['agree']/hold_q['total']*100,1)}%, FHA 100%, secure-info 99.4%, contact {round(contact_q['agree']/contact_q['total']*100,1)}%. {perfect} of {len(both)} calls are perfect. Mean scores are both high (human {mean_h}, AI {mean_a}) — this is not a broken grader, it is three questions.",
      what_didnt: "Name {name_pct}% ({name_q['strict']}S), next steps {next_pct}% ({next_q['strict']}S), reason {reason_pct}% ({reason_q['strict']}S, {reason_w} weighted pts). Validate {val_pct}% the other way ({val_q['lenient']}L). Greeting {greet_q['strict']}S, unit {unit_q['strict']}S, ownership {ack_q['strict']}S as a second-tier cluster.",
      path_to_90: "Need {need_90} more scored agreements ({int(0.9*total_comp)}/{total_comp}). Half of name strict + half of next-steps strict = {name_q['strict']//2 + next_q['strict']//2} → {round((total_agree + name_q['strict']//2 + next_q['strict']//2)/total_comp*100,1)}%. All name strict recovered is {round((total_agree + name_q['strict'])/total_comp*100,1)}% by itself. Re-run this same 490 after recs 1–3.",
      strictDetail: `{root_strict}`,
      lenientDetail: `{root_lenient}`
    }},
    matrixQuestions: [
      {{ short: "Greeting", key: "greeting" }},
      {{ short: "Name", key: "name_usage" }},
      {{ short: "Contact info", key: "contact_info" }},
      {{ short: "Unit #", key: "unit_number" }},
      {{ short: "Reason for call", key: "reason_for_call" }},
      {{ short: "Acknowledged", key: "acknowledged" }},
      {{ short: "Next steps", key: "next_steps" }},
      {{ short: "Hold perm.", key: "hold_permission" }},
      {{ short: "Validate concern", key: "validate_concern" }},
      {{ short: "FHA (DQ)", key: "fha" }},
      {{ short: "Secure info (DQ)", key: "secure_info" }}
    ]
  }}"""

ad_lines = []
for cid in both:
    cd = answer_data[cid]
    parts = []
    for q in ALL_KEYS:
        h, a = cd[q]
        h_str = str(h) if h is not None else "null"
        a_str = str(a) if a is not None else "null"
        parts.append(f"{q}:[{h_str},{a_str}]")
    ad_lines.append(f'    "{cid}": {{{",".join(parts)}}}')
answers_js = f"  {RUN_ID}: {{\n" + ",\n".join(ad_lines) + "\n  }"

RUN_START = "// ═══ LARGE TEST START ═══"
RUN_END = "// ═══ LARGE TEST END ═══"
ANS_START = "// ═══ LARGE TEST ANSWERS START ═══"
ANS_END = "// ═══ LARGE TEST ANSWERS END ═══"

run_block = f"{RUN_START}\n{run_js}\n  {RUN_END}"
ans_block = f"{ANS_START}\n{answers_js}\n  {ANS_END}"


def replace_between(html, start, end, new_block):
    if start in html and end in html:
        i0 = html.index(start)
        i1 = html.index(end) + len(end)
        return html[:i0] + new_block + html[i1:], True
    return html, False


print("\nPatching HTML...")
with open(HTML_PATH, "r") as f:
    html = f.read()

html, ok = replace_between(html, RUN_START, RUN_END, run_block)
if ok:
    print("  Replaced existing Large Test run block")
else:
    needle = "\n];\n\nconst residentMatrixQuestions = ["
    if needle not in html:
        raise SystemExit("Could not find residentRuns closing / residentMatrixQuestions")
    html = html.replace(
        needle,
        ",\n\n" + run_block + "\n];\n\nconst residentMatrixQuestions = [",
        1,
    )
    print("  Inserted Large Test run into residentRuns")

html, ok = replace_between(html, ANS_START, ANS_END, ans_block)
if ok:
    print("  Replaced existing Large Test answer block")
else:
    needle = "\n};\n\nconst residentRuns = ["
    if needle not in html:
        raise SystemExit("Could not find residentRunAnswerData closing")
    html = html.replace(
        needle,
        ",\n  " + ans_block + "\n};\n\nconst residentRuns = [",
        1,
    )
    print("  Inserted Large Test answers into residentRunAnswerData")

# Subtitle: mention Large Test
old_sub = "Run 1.0: 12 calls \\u00b7 Run 2.0–5.0 + 8.0: 20 calls \\u00b7 12 scoring criteria"
new_sub = "Run 1.0: 12 · Run 2.0–8.0: 20 · Large Test: 490 calls \\u00b7 overlapping 9 scored + 2 DQs"
if old_sub in html:
    html = html.replace(old_sub, new_sub)
    print("  Updated resident leadConfig subtitle")
elif "Large Test: 490" in html:
    print("  Subtitle already mentions Large Test")
else:
    print("  WARNING: could not update subtitle (pattern changed)")

with open(HTML_PATH, "w") as f:
    f.write(html)

print(f"\nWrote {HTML_PATH}")
print(f"  Large Test: {agreement_pct}% agreement, {total_dis} disagreements, {avg_delta}% avg delta, {len(both)} calls")
print(f"  Date stamp {date.today().isoformat()}")
