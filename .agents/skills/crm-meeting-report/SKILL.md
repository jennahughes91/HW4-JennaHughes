---
name: crm-meeting-report
description: Creates a short, quality CRM meeting report in a standardized format. Use when notes from customer meetings are available that have meeting participant details and key discussion topics. 
---

# CRM Meeting Report Skill

## When to Use This Skill

Use this skill when there are notes available for a meeting with a customer or customers. The notes must include enough detail to identify who attended and what was discussed.

**Do not use this skill when:**
- There are no meeting notes available
- The meeting was not a business meeting with a customer (e.g., internal team meetings, personal appointments)
- The name of the customer or customer organization is not captured anywhere in the meeting notes

## Expected Inputs

Meeting notes from a customer meeting that include:
- Meeting participant names
- Meeting participant titles
- Meeting participant employers / companies
- Meeting date
- Meeting location
- Details of what was discussed

## Step-by-Step Instructions

Follow these steps in order when generating a report:

1. **Identify the meeting date** from the notes
2. **Identify the meeting location** from the notes
3. **Identify all meeting participants**, capturing each person's name, title, and company
4. **Identify the key topics discussed** during the meeting
5. **Generate the CRM Meeting Report** using all of the above data in the format specified below

## Expected Output Format

The report must follow this structure exactly:

```
CRM Meeting Report

Date:      [meeting date]
Location:  [meeting location]

Meeting Participants
| Name | Title | Company |
|------|-------|---------|
| ...  | ...   | ...     |

Key Topics Discussed
• [Topic one]
• [Topic two]
• [Topic three]
```

- **Title:** "CRM Meeting Report" at the top of the document
- **Date line:** The date of the meeting
- **Location line:** Where the meeting took place
- **Meeting Participants table:** A table with three columns — Name, Title, and Company — listing every attendee
- **Key Topics Discussed:** A bulleted list summarizing what was discussed in the meeting

## Important Limitations

- **Do not generate or invent any data that is not present in the meeting notes.** If a field (date, location, participant title, etc.) is not captured in the notes, mark it as not specified rather than making something up.
- **Only generate the report if the meeting was a customer meeting.** If the notes describe an internal meeting, a personal call, or any context where no customer is present, do not produce a report — instead inform the user that this skill applies to customer meetings only.

## How to Generate the .docx File

Use the bundled Python script at `scripts/generate_report.py`. The `python-docx` library is required (version 1.x+).

Prepare a JSON object from the extracted data:

```json
{
  "date": "May 8, 2026",
  "location": "Zoom",
  "participants": [
    { "name": "Sarah Chen", "title": "VP of Sales", "company": "Acme Corp" },
    { "name": "James Ruiz", "title": "Account Executive", "company": "Your Company" }
  ],
  "topics": [
    "Discussed renewal timeline — budget confirmed for Q3.",
    "Raised concerns about onboarding speed; follow-up scheduled with implementation team.",
    "Demoed the new analytics dashboard; positive reception."
  ],
  "output_path": "/path/to/crm-meeting-report-YYYY-MM-DD.docx"
}
```

Then run:

```bash
python3 <path-to-skill>/scripts/generate_report.py '<JSON_DATA>'
```

Confirm the file was created, then share it with the user as a `computer://` link.

## Reference Files

Two reference files are available in `references/` — load them when needed, not by default:

- **`references/report-template.md`** — A fully worked example of a finished report with formatting notes. Consult this if you are uncertain about layout, section order, level of detail, or tone.
- **`references/field-extraction-guide.md`** — Rules for handling messy or incomplete notes: missing names, ambiguous dates, unknown company affiliation, action items, and more. Consult this whenever the notes are unclear, a field is missing, or a participant's details are only partially known.
