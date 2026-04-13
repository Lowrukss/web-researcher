---
name: web-researcher
description: |
  Use this skill to research any topic on the web and produce a comprehensive, well-structured knowledge report.
  Trigger when the user says "research X", "find out about X", "I need to know everything about X",
  "gather information on X", "deep dive into X", "what do you know about X" (when web lookup is needed),
  "look up X", "investigate X", or asks for a thorough overview of any subject.
  Also trigger when the user implies they want more than a quick answer — e.g. "give me a full picture of X",
  "I want to understand X deeply", "compile what's known about X".
---

## Dependencies
- `python-docx` — for generating `.docx` output: `pip install python-docx`
- `fpdf2` — for generating `.pdf` output: `pip install fpdf2`
- Claude's built-in `WebSearch` and `WebFetch` tools for research.

---

## Fixed Workflow

### Step 1 — Clarify Scope (only if genuinely ambiguous)
If the topic is broad or the user's intent is unclear, ask **one** focused question before starting:
- "Any particular angle — technical, historical, business, news?"
- Skip this step if intent is already clear.

---

### Step 2 — Form Search Queries
Before searching, plan 4–6 distinct queries that cover different angles:
- **Broad overview**: `"[topic] overview"` or `"what is [topic]"`
- **Recent developments**: `"[topic] 2024 2025"`
- **Technical / in-depth**: `"[topic] how it works"` or `"[topic] explained"`
- **Criticism / limitations**: `"[topic] problems"` or `"[topic] criticism"`
- **Applications / use cases**: `"[topic] use cases"` or `"[topic] examples"`
- **Expert opinions / research**: `"[topic] research"` or `"[topic] study"`

Adapt query angles to the topic type (person, technology, event, concept, company, etc.).

---

### Step 3 — Execute Searches
Run all planned queries using `WebSearch`. For each result:
- Note the source name, URL, and headline
- Flag the 4–6 most promising URLs for deep reading (prefer authoritative sources: .edu, .gov, Wikipedia, major publications, official sites)

---

### Step 4 — Fetch & Read Key Sources
Use `WebFetch` to read each flagged URL. For each page:
- Extract key facts, data points, definitions, and quotes
- Note the publication date (for recency)
- Mark anything surprising, contested, or requiring follow-up

---

### Step 5 — Identify Gaps & Follow-Up
After the first round, check:
- Are there subtopics mentioned but not yet covered?
- Are there conflicting claims that need a third source?
- Is there missing recent data?

Run 1–3 additional targeted searches to fill gaps.

---

### Step 6 — Synthesize & Write Report

Compose the full report content using this structure:

```
# [Topic]

## TL;DR
2–3 sentence summary of the most important things to know.

## Overview
What it is, where it comes from, why it matters. (3–5 paragraphs)

## Key Facts & Data
Bullet list of the most concrete, verifiable facts: numbers, dates, names, stats.

## [Subtopic 1]
## [Subtopic 2]
## [Subtopic 3]
(Use as many subtopic sections as needed — name them based on what emerged from research)

## Current State & Recent Developments
What's happening now, latest news, trends as of [date].

## Controversies / Limitations / Open Questions
What's debated, unknown, or problematic. Skip if genuinely not applicable.

## Sources
- [Source Name](URL) — one-line note on what it contributed
```

**Quality bar for the report:**
- Minimum 600 words; aim for thorough over brief
- Every major claim should trace back to a fetched source
- Use concrete numbers and examples, not vague generalities
- Do not pad with filler — if a section has nothing meaningful, omit it

---

### Step 7 — Export to File

Ask the user which format they prefer if they haven't specified: **PDF** or **DOCX**.

Then generate the file using a Python script via the `Bash` tool.

**For DOCX** (using `python-docx`):
```python
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document()

# Title
title = doc.add_heading('[Topic]', level=0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER

# Add each section: doc.add_heading('Section Name', level=1) then doc.add_paragraph('content')
# For bullet lists: doc.add_paragraph('item', style='List Bullet')
# For source links: add as plain text (DOCX doesn't render hyperlinks visually by default)

doc.save('[topic-slug]-research.docx')
print("Saved: [topic-slug]-research.docx")
```

**For PDF** (using `fpdf2`):
```python
from fpdf import FPDF

class ReportPDF(FPDF):
    def header(self):
        self.set_font('Helvetica', 'B', 10)
        self.cell(0, 8, '[Topic] — Research Report', align='C', new_x='LMARGIN', new_y='NEXT')
        self.ln(2)

    def footer(self):
        self.set_y(-12)
        self.set_font('Helvetica', 'I', 8)
        self.cell(0, 8, f'Page {self.page_no()}', align='C')

pdf = ReportPDF()
pdf.add_page()
pdf.set_auto_page_break(auto=True, margin=15)

# Title
pdf.set_font('Helvetica', 'B', 20)
pdf.cell(0, 12, '[Topic]', align='C', new_x='LMARGIN', new_y='NEXT')
pdf.ln(4)

# Section headings: pdf.set_font('Helvetica', 'B', 13) then pdf.multi_cell(0, 7, 'Heading')
# Body text: pdf.set_font('Helvetica', '', 11) then pdf.multi_cell(0, 6, 'paragraph text')
# Bullet items: prepend '• ' to each item in multi_cell

pdf.output('[topic-slug]-research.pdf')
print("Saved: [topic-slug]-research.pdf")
```

**Naming convention:** use a lowercase kebab-case slug of the topic, e.g. `quantum-computing-research.pdf`.

**Save location:** current working directory unless the user specifies otherwise.

After saving, confirm the file path to the user.
