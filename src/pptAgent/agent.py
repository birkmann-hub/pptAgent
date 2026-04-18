"""AI agent that generates slide content using an OpenAI-compatible LLM.

The agent accepts a :class:`PresentationRequest`, decides which slides to
include (based on the duration), generates the textual content for each slide
and returns a :class:`PresentationPlan` that can be fed to
:class:`PresentationGenerator`.

When no API key is available the agent falls back to sensible placeholder
content so that the generator can still produce a presentation skeleton.
"""

from __future__ import annotations

import json
import logging
import os
from datetime import date

from .models import (
    PresentationPlan,
    PresentationRequest,
    ProjectType,
    SectorType,
    SlideContent,
    SlideType,
)

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Slide selection logic
# ---------------------------------------------------------------------------

def _select_slide_types(duration_minutes: int) -> list[SlideType]:
    """Return an ordered list of slide types appropriate for the meeting length."""
    # Core slides that always appear
    core: list[SlideType] = [
        SlideType.COVER,
        SlideType.AGENDA,
        SlideType.SECTION_DIVIDER,
        SlideType.CONTEXT_BACKGROUND,
        SlideType.FINDINGS,
        SlideType.RECOMMENDATIONS,
        SlideType.NEXT_STEPS,
        SlideType.CLOSING,
    ]

    if duration_minutes >= 30:
        core.insert(2, SlideType.EXECUTIVE_SUMMARY)

    if duration_minutes >= 60:
        # Add problem statement section
        idx = core.index(SlideType.CONTEXT_BACKGROUND)
        core.insert(idx + 1, SlideType.PROBLEM_STATEMENT)

    if duration_minutes >= 90:
        # Add methodology
        idx = core.index(SlideType.FINDINGS)
        core.insert(idx, SlideType.METHODOLOGY)

    if duration_minutes >= 120:
        # Add a second findings slide and a roadmap
        idx = core.index(SlideType.RECOMMENDATIONS)
        core.insert(idx, SlideType.FINDINGS)
        core.insert(idx + 2, SlideType.ROADMAP)

    return core


# ---------------------------------------------------------------------------
# Fallback content builder (used when no LLM is available)
# ---------------------------------------------------------------------------

def _build_fallback_content(
    slide_type: SlideType,
    request: PresentationRequest,
    slide_index: int,
) -> SlideContent:
    """Return a SlideContent populated with sensible placeholder text."""
    today = request.presentation_date or date.today()
    date_str = today.strftime("%d.%m.%Y")

    if slide_type == SlideType.COVER:
        return SlideContent(
            slide_type=slide_type,
            title=request.topic,
            subtitle=f"{request.client_name or 'Client'} | {request.project_type.value.replace('_', ' ').title()}",
            body_text=f"{request.sector.value.replace('_', ' ').title()} | {request.project_type.value.replace('_', ' ').title()} | {date_str}",
            footnote=request.presenter_name or "",
        )

    if slide_type == SlideType.AGENDA:
        items = [
            "Context & Background",
            "Problem Statement",
            "Key Findings",
            "Recommendations",
            "Next Steps",
        ]
        return SlideContent(
            slide_type=slide_type,
            title="Agenda",
            bullet_points=items,
            footnote=f"Total duration: {request.duration_minutes} min",
        )

    if slide_type == SlideType.EXECUTIVE_SUMMARY:
        return SlideContent(
            slide_type=slide_type,
            title="Executive Summary",
            bullet_points=[
                f"The {request.sector.value.replace('_', ' ')} sector faces significant transformation pressure",
                "Current approach requires strategic re-alignment to capture full value",
                "Three high-impact initiatives will drive measurable results within 12 months",
            ],
            highlight_box=f"Bottom line: Immediate action on {request.topic} will deliver competitive advantage.",
        )

    if slide_type == SlideType.CONTEXT_BACKGROUND:
        return SlideContent(
            slide_type=slide_type,
            title="Context & Background",
            left_heading="Industry Trends",
            left_bullets=[
                "Rapid digitalisation reshaping competitive dynamics",
                "Regulatory requirements increasing complexity",
                "Customer expectations shifting towards personalised experiences",
            ],
            right_heading="Company Background",
            right_bullets=[
                f"Engagement scope: {request.topic}",
                f"Project type: {request.project_type.value.replace('_', ' ').title()}",
                f"Target audience: {request.target_group}",
            ],
            footnote="Sources: Industry reports, client documentation",
        )

    if slide_type == SlideType.PROBLEM_STATEMENT:
        return SlideContent(
            slide_type=slide_type,
            title="Problem Statement",
            highlight_box=f"How can we effectively address {request.topic} to create sustainable competitive advantage?",
            bullet_points=[
                "Driver 1: Growing gap between current capabilities and market requirements",
                "Driver 2: Inefficient processes leading to increased cost and reduced speed",
                "Driver 3: Limited visibility into key performance indicators",
            ],
            body_text="If left unaddressed, this leads to: Revenue loss, talent attrition, and reduced market share.",
        )

    if slide_type == SlideType.METHODOLOGY:
        return SlideContent(
            slide_type=slide_type,
            title="Our Approach",
            process_steps=[
                {"label": "Phase 1", "heading": "Discovery", "description": "Stakeholder interviews\nData collection & review"},
                {"label": "Phase 2", "heading": "Analysis", "description": "Data analysis\nBenchmarking & gap assessment"},
                {"label": "Phase 3", "heading": "Synthesis", "description": "Insight generation\nHypothesis testing"},
                {"label": "Phase 4", "heading": "Recommendations", "description": "Solution design\nBusiness case development", "active": True},
            ],
            footnote=f"Project scope: {request.duration_minutes} min workshop",
        )

    if slide_type == SlideType.FINDINGS:
        finding_n = slide_index
        return SlideContent(
            slide_type=slide_type,
            title="Key Findings",
            section_number=f"{finding_n:02d}",
            highlight_box=f"[Finding {finding_n}] Insert insight-driven headline here",
            bullet_points=[
                "Supporting data point or evidence 1",
                "Supporting data point or evidence 2",
                "Supporting data point or evidence 3",
            ],
            right_heading="[Chart / Data Visual]",
        )

    if slide_type == SlideType.RECOMMENDATIONS:
        return SlideContent(
            slide_type=slide_type,
            title="Our Recommendations",
            recommendation_cards=[
                {"priority": "01", "heading": "Establish governance framework", "description": "Define clear ownership, decision rights and accountability across functions.", "impact": "High", "effort": "Low"},
                {"priority": "02", "heading": "Launch quick-win initiatives", "description": "Implement 2-3 high-impact, low-effort improvements within 30 days.", "impact": "High", "effort": "Low"},
                {"priority": "03", "heading": "Build capability roadmap", "description": "Develop a 12-month plan for skill development and process improvement.", "impact": "Medium", "effort": "Medium"},
            ],
            footnote="Prioritisation based on impact/effort analysis",
        )

    if slide_type == SlideType.ROADMAP:
        return SlideContent(
            slide_type=slide_type,
            title="Implementation Roadmap",
            timeline_columns=["Q1", "Q2", "Q3", "Q4"],
            swim_lanes=[
                {"workstream": "Governance", "bars": [{"start": 1, "span": 1, "label": "Setup"}]},
                {"workstream": "Quick Wins", "bars": [{"start": 1, "span": 2, "label": "Implementation"}]},
                {"workstream": "Capability", "bars": [{"start": 2, "span": 3, "label": "Build & Deploy"}]},
            ],
        )

    if slide_type == SlideType.NEXT_STEPS:
        return SlideContent(
            slide_type=slide_type,
            title="Next Steps",
            table_rows=[
                ["1", "Schedule project kick-off meeting", "Project Lead", f"{date_str}", "Open"],
                ["2", "Provide requested data & documentation", "Client", f"{date_str}", "Open"],
                ["3", "Set up project governance structure", "PMO", f"{date_str}", "Open"],
                ["4", "Align on communication plan", "Comms Lead", f"{date_str}", "Open"],
            ],
            footnote="Status: Open | In Progress | Done",
        )

    if slide_type == SlideType.SECTION_DIVIDER:
        return SlideContent(
            slide_type=slide_type,
            section_number="01",
            title="Context & Background",
            subtitle="Setting the stage for today's discussion",
        )

    if slide_type == SlideType.CLOSING:
        return SlideContent(
            slide_type=slide_type,
            title="Thank You",
            subtitle="We look forward to your questions and next steps.",
            body_text=f"{request.presenter_name} | {request.contact_email}" if request.contact_email else request.presenter_name,
            footnote=request.client_name or "",
        )

    return SlideContent(slide_type=slide_type, title=slide_type.value.replace("_", " ").title())


# ---------------------------------------------------------------------------
# LLM-based content generation
# ---------------------------------------------------------------------------

_SYSTEM_PROMPT = """\
You are an expert management consultant and presentation designer.
Your task is to generate professional, data-driven slide content for a consulting
presentation. Always follow the pyramid principle: lead with the conclusion,
support with evidence.

Respond ONLY with a valid JSON object following the schema provided.
Do not add markdown, explanations or any text outside the JSON object.
"""


def _build_llm_prompt(slide_type: SlideType, request: PresentationRequest) -> str:
    schema = {
        "slide_type": slide_type.value,
        "title": "string",
        "subtitle": "string (optional)",
        "bullet_points": ["string (max 6 items)"],
        "highlight_box": "string (max 120 chars, optional)",
        "left_heading": "string (optional)",
        "left_bullets": ["string (optional)"],
        "right_heading": "string (optional)",
        "right_bullets": ["string (optional)"],
        "section_number": "string like '01' (optional)",
        "table_rows": [["string cell (optional)"]],
        "process_steps": [{"label": "Phase N", "heading": "Noun", "description": "2 lines", "active": False}],
        "timeline_columns": ["string (optional)"],
        "swim_lanes": [{"workstream": "string", "bars": [{"start": 1, "span": 2, "label": "string"}]}],
        "recommendation_cards": [{"priority": "01", "heading": "string", "description": "string", "impact": "High|Medium|Low", "effort": "High|Medium|Low"}],
        "footnote": "string (optional)",
    }
    context = {
        "topic": request.topic,
        "target_group": request.target_group,
        "sector": request.sector.value,
        "project_type": request.project_type.value,
        "duration_minutes": request.duration_minutes,
        "client_name": request.client_name,
        "presenter_name": request.presenter_name,
        "date": str(request.presentation_date),
        "additional_context": request.additional_context,
        "language": request.language,
    }
    return (
        f"Generate content for a '{slide_type.value}' slide.\n\n"
        f"Presentation context:\n{json.dumps(context, indent=2)}\n\n"
        f"Respond with a JSON object conforming to this schema:\n{json.dumps(schema, indent=2)}"
    )


def _call_openai(slide_type: SlideType, request: PresentationRequest) -> SlideContent | None:
    """Call OpenAI API and parse the JSON response into a SlideContent."""
    try:
        import openai  # imported lazily
    except ImportError:
        logger.debug("openai package not available; using fallback content")
        return None

    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        logger.debug("OPENAI_API_KEY not set; using fallback content")
        return None

    try:
        client = openai.OpenAI(api_key=api_key)
        response = client.chat.completions.create(
            model=os.environ.get("OPENAI_MODEL", "gpt-4o-mini"),
            messages=[
                {"role": "system", "content": _SYSTEM_PROMPT},
                {"role": "user", "content": _build_llm_prompt(slide_type, request)},
            ],
            temperature=0.7,
            response_format={"type": "json_object"},
        )
        raw = response.choices[0].message.content
        data = json.loads(raw)
        data["slide_type"] = slide_type.value
        return SlideContent(**{k: v for k, v in data.items() if k in SlideContent.model_fields})
    except Exception as exc:  # noqa: BLE001
        logger.warning("LLM call failed for slide '%s': %s", slide_type.value, exc)
        return None


# ---------------------------------------------------------------------------
# Public agent class
# ---------------------------------------------------------------------------

class PresentationAgent:
    """Orchestrates slide selection and content generation.

    Uses the OpenAI API when ``OPENAI_API_KEY`` is set; otherwise generates
    placeholder content so the pipeline always produces a complete file.
    """

    def run(self, request: PresentationRequest) -> PresentationPlan:
        """Generate a :class:`PresentationPlan` for the given request."""
        slide_types = _select_slide_types(request.duration_minutes)
        slides: list[SlideContent] = []

        finding_index = 0
        for slide_type in slide_types:
            if slide_type == SlideType.FINDINGS:
                finding_index += 1
                content = _call_openai(slide_type, request) or _build_fallback_content(slide_type, request, finding_index)
            else:
                content = _call_openai(slide_type, request) or _build_fallback_content(slide_type, request, 0)
            slides.append(content)

        return PresentationPlan(request=request, slides=slides)
