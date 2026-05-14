# Field Extraction Guide — CRM Meeting Report

This guide covers how to handle messy, incomplete, or ambiguous meeting notes. Consult it whenever the notes are unclear or a field is missing.

---

## Date

| Situation | What to do |
|---|---|
| Exact date given ("May 7", "April 29, 2026") | Use as-is; add year only if it makes the report unambiguous |
| Relative date ("yesterday", "last Tuesday") | Convert to a real date using today's date as the reference point |
| Vague ("last week", "recently") | Write "Approximate: [your best estimate]" and note the uncertainty |
| Not mentioned at all | Write "Not specified" |

Never leave the date field blank.

---

## Location

| Situation | What to do |
|---|---|
| Specific address or office name | Use as stated |
| "Their office" / "client site" | Write "[Company name] office" |
| "Zoom", "Teams", "Google Meet", "call", "video" | Write the platform name (e.g., "Zoom") |
| "Phone call" | Write "Phone call" |
| Not mentioned | Write "Not specified" |

---

## Participants

### Missing names
- If a participant is referenced by role only ("their PM", "the CFO"), create a row with the title filled in and the name left blank — don't omit the person entirely.
- If a name is partially noted (first name only, misspelled, or abbreviated), use what's there. Do not guess the full name.

### Missing titles
- Leave the Title cell blank. Do not invent a title.
- Exception: if the notes make the role obvious in context ("she runs their engineering org"), you may write a brief descriptive title in italics to signal it's inferred: *Engineering Lead (inferred)*

### Missing company / which side they're on
- If it's unclear whether someone is from the customer side or your side, place them at the bottom of the table and leave the Company cell blank.
- Do not assign your company name to someone unless the notes explicitly place them on your side.

### Row order
- Customer / prospect attendees first
- Your-side attendees second
- Unknown affiliation at the bottom

### Duplicate names
- If the same person appears to be listed twice (e.g., once with a title, once without), merge them into a single row using the most complete information.

---

## Key Topics

### How much detail to include
- Aim for 3–6 bullets. Fewer is fine if the meeting was short or narrowly focused.
- Each bullet should convey one distinct idea, decision, concern, or action.
- If the notes are very sparse (2–3 sentences total), include all substantive points even if it means only 1–2 bullets.

### What counts as a topic
Include:
- Decisions made or agreed upon
- Key concerns or objections raised
- Demos, presentations, or materials reviewed
- Questions asked that shaped the discussion
- Explicit next steps or follow-up commitments

Exclude:
- Small talk or pleasantries
- Logistical details (parking, lunch, room setup)
- Repetition of what's already captured in the participants table

### Action items and next steps
- If a specific follow-up was agreed upon (with an owner and/or deadline), make it the last bullet and bold the "Next step:" prefix.
- If multiple follow-ups were mentioned, list the most important one as the final bullet; fold others into the relevant topic bullets.

### Tone
- Neutral and factual. Report what was said, not your assessment of it.
- Avoid: "great discussion", "exciting opportunity", "the client loved it"
- Prefer: "the client responded positively", "there was strong interest in X"

---

## Handling Very Incomplete Notes

If the notes are so sparse that key sections cannot be filled in meaningfully:

1. Fill in what you can from the available information.
2. Use "Not specified" for missing required fields (date, location).
3. Keep the Topics section short and accurate — do not pad it.
4. Do not fabricate details to make the report look more complete.

The goal is an accurate, honest summary of what the notes contain — not a polished document that overstates what happened.
