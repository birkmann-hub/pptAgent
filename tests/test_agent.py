"""Tests for the PresentationAgent (fallback / no-LLM mode)."""

import pytest
from datetime import date
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from pptAgent.models import SectorType, ProjectType, SlideType
from pptAgent.agent import PresentationAgent, _select_slide_types
from tests.test_models import make_request


class TestSlideSelection:
    def test_short_meeting_has_core_slides(self):
        types = _select_slide_types(20)
        assert SlideType.COVER in types
        assert SlideType.CLOSING in types
        assert SlideType.AGENDA in types

    def test_executive_summary_at_30min(self):
        types = _select_slide_types(30)
        assert SlideType.EXECUTIVE_SUMMARY in types

    def test_no_executive_summary_below_30min(self):
        types = _select_slide_types(25)
        assert SlideType.EXECUTIVE_SUMMARY not in types

    def test_problem_statement_at_60min(self):
        types = _select_slide_types(60)
        assert SlideType.PROBLEM_STATEMENT in types

    def test_methodology_at_90min(self):
        types = _select_slide_types(90)
        assert SlideType.METHODOLOGY in types

    def test_roadmap_at_120min(self):
        types = _select_slide_types(120)
        assert SlideType.ROADMAP in types

    def test_cover_is_first(self):
        types = _select_slide_types(90)
        assert types[0] == SlideType.COVER

    def test_closing_is_last(self):
        types = _select_slide_types(90)
        assert types[-1] == SlideType.CLOSING

    def test_more_slides_for_longer_meeting(self):
        short = _select_slide_types(30)
        long = _select_slide_types(120)
        assert len(long) > len(short)


class TestPresentationAgent:
    def test_run_returns_plan(self):
        agent = PresentationAgent()
        req = make_request(duration_minutes=60)
        plan = agent.run(req)
        assert plan.request is req
        assert len(plan.slides) > 0

    def test_plan_starts_with_cover(self):
        agent = PresentationAgent()
        plan = agent.run(make_request(duration_minutes=60))
        assert plan.slides[0].slide_type == SlideType.COVER

    def test_plan_ends_with_closing(self):
        agent = PresentationAgent()
        plan = agent.run(make_request(duration_minutes=60))
        assert plan.slides[-1].slide_type == SlideType.CLOSING

    def test_cover_has_topic_as_title(self):
        agent = PresentationAgent()
        req = make_request(topic="My Topic")
        plan = agent.run(req)
        cover = next(s for s in plan.slides if s.slide_type == SlideType.COVER)
        assert "My Topic" in cover.title

    def test_90min_includes_methodology(self):
        agent = PresentationAgent()
        plan = agent.run(make_request(duration_minutes=90))
        types = [s.slide_type for s in plan.slides]
        assert SlideType.METHODOLOGY in types

    def test_30min_includes_executive_summary(self):
        agent = PresentationAgent()
        plan = agent.run(make_request(duration_minutes=30))
        types = [s.slide_type for s in plan.slides]
        assert SlideType.EXECUTIVE_SUMMARY in types

    def test_all_slides_have_slide_type(self):
        agent = PresentationAgent()
        plan = agent.run(make_request(duration_minutes=120))
        for slide in plan.slides:
            assert slide.slide_type is not None
