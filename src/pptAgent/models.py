"""Pydantic models for presentation requests and slide content."""

from __future__ import annotations

from datetime import date
from enum import Enum
from typing import Any

from pydantic import BaseModel, Field


# ---------------------------------------------------------------------------
# Enumerations
# ---------------------------------------------------------------------------

class SectorType(str, Enum):
    FINANCIAL_SERVICES = "financial_services"
    HEALTHCARE = "healthcare"
    TECHNOLOGY = "technology"
    MANUFACTURING = "manufacturing"
    RETAIL = "retail"
    ENERGY = "energy"
    PUBLIC_SECTOR = "public_sector"
    AUTOMOTIVE = "automotive"
    CONSUMER_GOODS = "consumer_goods"
    OTHER = "other"


class ProjectType(str, Enum):
    STRATEGY = "strategy"
    DIGITAL_TRANSFORMATION = "digital_transformation"
    OPERATIONS = "operations"
    ORGANIZATIONAL = "organizational"
    MERGERS_ACQUISITIONS = "mergers_acquisitions"
    MARKET_ENTRY = "market_entry"
    COST_REDUCTION = "cost_reduction"
    INNOVATION = "innovation"
    RISK_COMPLIANCE = "risk_compliance"
    OTHER = "other"


class SlideType(str, Enum):
    COVER = "cover"
    AGENDA = "agenda"
    EXECUTIVE_SUMMARY = "executive_summary"
    CONTEXT_BACKGROUND = "context_background"
    PROBLEM_STATEMENT = "problem_statement"
    METHODOLOGY = "methodology"
    FINDINGS = "findings"
    RECOMMENDATIONS = "recommendations"
    ROADMAP = "roadmap"
    NEXT_STEPS = "next_steps"
    SECTION_DIVIDER = "section_divider"
    CLOSING = "closing"


# ---------------------------------------------------------------------------
# Input model
# ---------------------------------------------------------------------------

class PresentationRequest(BaseModel):
    """All parameters required to generate a consulting presentation."""

    topic: str = Field(
        ...,
        description="The main topic or title of the presentation.",
        examples=["Digital Transformation Strategy 2025"],
    )
    target_group: str = Field(
        ...,
        description="Audience description (role, seniority, expertise level).",
        examples=["C-Suite executives with limited technical background"],
    )
    sector: SectorType = Field(
        ...,
        description="Industry sector of the client.",
    )
    project_type: ProjectType = Field(
        ...,
        description="Type of consulting engagement.",
    )
    duration_minutes: int = Field(
        ...,
        ge=15,
        le=480,
        description="Planned duration of the workshop or meeting in minutes.",
    )
    presenter_name: str = Field(
        default="",
        description="Full name of the lead presenter.",
    )
    contact_email: str = Field(
        default="",
        description="Contact e-mail address shown on the closing slide.",
    )
    client_name: str = Field(
        default="",
        description="Name of the client or company the presentation is for.",
    )
    presentation_date: date = Field(
        default_factory=date.today,
        description="Date of the presentation.",
    )
    additional_context: str = Field(
        default="",
        description="Any additional context the agent should consider.",
    )
    output_path: str = Field(
        default="presentation.pptx",
        description="File path where the generated .pptx will be saved.",
    )
    language: str = Field(
        default="en",
        description="Language code for the generated content (e.g. 'en', 'de').",
    )


# ---------------------------------------------------------------------------
# Slide content model
# ---------------------------------------------------------------------------

class SlideContent(BaseModel):
    """Populated content for a single slide, ready for rendering."""

    slide_type: SlideType
    title: str = ""
    subtitle: str = ""
    body_text: str = ""
    bullet_points: list[str] = Field(default_factory=list)
    highlight_box: str = ""
    left_heading: str = ""
    left_bullets: list[str] = Field(default_factory=list)
    right_heading: str = ""
    right_bullets: list[str] = Field(default_factory=list)
    table_rows: list[list[str]] = Field(default_factory=list)
    process_steps: list[dict[str, Any]] = Field(default_factory=list)
    timeline_columns: list[str] = Field(default_factory=list)
    swim_lanes: list[dict[str, Any]] = Field(default_factory=list)
    recommendation_cards: list[dict[str, Any]] = Field(default_factory=list)
    section_number: str = ""
    footnote: str = ""
    # Raw extra data the agent may add
    extra: dict[str, Any] = Field(default_factory=dict)


# ---------------------------------------------------------------------------
# Full presentation model
# ---------------------------------------------------------------------------

class PresentationPlan(BaseModel):
    """Ordered list of slides that make up the final presentation."""

    request: PresentationRequest
    slides: list[SlideContent] = Field(default_factory=list)
