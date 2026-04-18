"""Tests for the presentation models."""

import pytest
from datetime import date

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from pptAgent.models import (
    PresentationRequest,
    PresentationPlan,
    SlideContent,
    SlideType,
    SectorType,
    ProjectType,
)


def make_request(**kwargs) -> PresentationRequest:
    defaults = dict(
        topic="Digital Transformation Strategy",
        target_group="C-Suite executives",
        sector=SectorType.FINANCIAL_SERVICES,
        project_type=ProjectType.DIGITAL_TRANSFORMATION,
        duration_minutes=90,
        presenter_name="Jane Doe",
        contact_email="jane@example.com",
        client_name="Acme Corp",
        presentation_date=date(2025, 6, 1),
    )
    defaults.update(kwargs)
    return PresentationRequest(**defaults)


class TestPresentationRequest:
    def test_valid_request(self):
        req = make_request()
        assert req.topic == "Digital Transformation Strategy"
        assert req.sector == SectorType.FINANCIAL_SERVICES
        assert req.duration_minutes == 90

    def test_duration_min_boundary(self):
        req = make_request(duration_minutes=15)
        assert req.duration_minutes == 15

    def test_duration_below_min_raises(self):
        with pytest.raises(Exception):
            make_request(duration_minutes=10)

    def test_duration_max_boundary(self):
        req = make_request(duration_minutes=480)
        assert req.duration_minutes == 480

    def test_duration_above_max_raises(self):
        with pytest.raises(Exception):
            make_request(duration_minutes=481)

    def test_default_output_path(self):
        req = make_request()
        assert req.output_path == "presentation.pptx"

    def test_custom_output_path(self):
        req = make_request(output_path="/tmp/out.pptx")
        assert req.output_path == "/tmp/out.pptx"

    def test_all_sectors(self):
        for sector in SectorType:
            req = make_request(sector=sector)
            assert req.sector == sector

    def test_all_project_types(self):
        for pt in ProjectType:
            req = make_request(project_type=pt)
            assert req.project_type == pt


class TestSlideContent:
    def test_default_fields(self):
        sc = SlideContent(slide_type=SlideType.COVER)
        assert sc.bullet_points == []
        assert sc.table_rows == []
        assert sc.extra == {}

    def test_bullet_points(self):
        sc = SlideContent(
            slide_type=SlideType.AGENDA,
            bullet_points=["Item 1", "Item 2"],
        )
        assert len(sc.bullet_points) == 2

    def test_recommendation_cards(self):
        sc = SlideContent(
            slide_type=SlideType.RECOMMENDATIONS,
            recommendation_cards=[
                {"priority": "01", "heading": "Do X", "description": "Desc", "impact": "High", "effort": "Low"}
            ],
        )
        assert sc.recommendation_cards[0]["priority"] == "01"


class TestPresentationPlan:
    def test_empty_slides(self):
        req = make_request()
        plan = PresentationPlan(request=req)
        assert plan.slides == []

    def test_with_slides(self):
        req = make_request()
        slides = [
            SlideContent(slide_type=SlideType.COVER, title="Test"),
            SlideContent(slide_type=SlideType.CLOSING),
        ]
        plan = PresentationPlan(request=req, slides=slides)
        assert len(plan.slides) == 2
        assert plan.slides[0].title == "Test"
