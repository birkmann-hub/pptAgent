# pptAgent

A canonical-model consulting presentation generator.
Given a few parameters the agent builds a fully styled **.pptx** file in
corporate design — ready to open in PowerPoint.

---

## Features

| Capability | Detail |
|---|---|
| **Canonical slide model** | 12 reusable slide types, each described in its own YAML file under `slides/` |
| **Corporate design** | Consistent colour palette, typography and layout via `src/pptAgent/corporate_design.py` |
| **Adaptive structure** | Number and type of slides are selected automatically based on meeting duration |
| **AI content** | When an `OPENAI_API_KEY` is set the agent calls GPT to fill in slide text; otherwise professional placeholder content is used |
| **PowerPoint output** | Output is a standard `.pptx` file (python-pptx) — editable in PowerPoint / LibreOffice |

---

## Canonical Slide Types

Each slide type is defined in `slides/<name>.yaml` with a description,
layout specification, element list and an AI prompt hint.

| Slide | File | When used |
|---|---|---|
| Cover / Title | `slides/cover.yaml` | First slide, always |
| Agenda | `slides/agenda.yaml` | Always |
| Executive Summary | `slides/executive_summary.yaml` | >= 30 min |
| Section Divider | `slides/section_divider.yaml` | Always |
| Context & Background | `slides/context_background.yaml` | Always |
| Problem Statement | `slides/problem_statement.yaml` | >= 60 min |
| Methodology | `slides/methodology.yaml` | >= 90 min |
| Key Findings | `slides/findings.yaml` | Always |
| Recommendations | `slides/recommendations.yaml` | Always |
| Roadmap | `slides/roadmap.yaml` | >= 120 min |
| Next Steps | `slides/next_steps.yaml` | Always |
| Closing | `slides/closing.yaml` | Last slide, always |

---

## Project Structure

```
pptAgent/
├── main.py                    # CLI entry point
├── requirements.txt
├── pytest.ini
├── .env.example               # API key template
├── slides/                    # Canonical slide type definitions (YAML)
│   ├── cover.yaml
│   ├── agenda.yaml
│   ├── executive_summary.yaml
│   ├── context_background.yaml
│   ├── problem_statement.yaml
│   ├── methodology.yaml
│   ├── findings.yaml
│   ├── recommendations.yaml
│   ├── roadmap.yaml
│   ├── next_steps.yaml
│   ├── section_divider.yaml
│   └── closing.yaml
├── src/pptAgent/              # Python package
│   ├── __init__.py
│   ├── models.py              # Pydantic input/output models
│   ├── corporate_design.py    # Colours, fonts, layout constants
│   ├── agent.py               # Slide selection + content generation
│   ├── slide_builder.py       # Renders each slide type onto python-pptx
│   └── generator.py           # Assembles the final .pptx file
└── tests/
    ├── test_models.py
    ├── test_agent.py
    └── test_generator.py
```

---

## Quick Start

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. (Optional) Set your OpenAI API key

```bash
cp .env.example .env
# edit .env and add your OPENAI_API_KEY
```

Without a key, a complete presentation is still generated using professional
placeholder content.

### 3. Generate a presentation

```bash
python main.py \
  --topic "Digital Transformation Strategy 2025" \
  --target-group "C-Suite executives" \
  --sector financial_services \
  --project-type digital_transformation \
  --duration 90 \
  --presenter "Jane Doe" \
  --email "jane.doe@company.com" \
  --client "Acme Corp" \
  --output presentation.pptx
```

Run `python main.py --help` for the full list of options.

---

## Parameters

| Parameter | Required | Description |
|---|---|---|
| `--topic` | Yes | Main topic / title |
| `--target-group` | Yes | Audience (role, seniority, background) |
| `--sector` | Yes | Industry sector (see choices) |
| `--project-type` | Yes | Type of consulting engagement |
| `--duration` | Yes | Meeting/workshop length in minutes |
| `--presenter` | | Lead presenter name |
| `--email` | | Contact e-mail on closing slide |
| `--client` | | Client company name |
| `--date` | | Presentation date (YYYY-MM-DD) |
| `--context` | | Additional context for the AI |
| `--output` | | Output file path (default: `presentation.pptx`) |
| `--language` | | Language code, e.g. `en`, `de` |

### Supported sectors

`financial_services` · `healthcare` · `technology` · `manufacturing` · `retail` · `energy` · `public_sector` · `automotive` · `consumer_goods` · `other`

### Supported project types

`strategy` · `digital_transformation` · `operations` · `organizational` · `mergers_acquisitions` · `market_entry` · `cost_reduction` · `innovation` · `risk_compliance` · `other`

---

## Running tests

```bash
pip install pytest
python -m pytest tests/ -v
```

---

## Corporate Design

The colour palette and typography are defined in
`src/pptAgent/corporate_design.py`:

| Token | Hex | Usage |
|---|---|---|
| `DARK_BLUE` | `#003366` | Header bands, badges, primary fill |
| `MEDIUM_BLUE` | `#0066CC` | Timeline bars, impact tags |
| `LIGHT_BLUE` | `#66AADD` | Subtitles on dark backgrounds |
| `ACCENT_ORANGE` | `#FF6600` | Call-outs, section numbers, highlight boxes |
| `LIGHT_GRAY` | `#F0F2F5` | Alternating table rows, chart placeholders |
| `DARK_TEXT` | `#1A1A2E` | Body text |

Font family: **Calibri** (standard PowerPoint font).
