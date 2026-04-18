"""Assembles a python-pptx Presentation from a PresentationPlan and saves it."""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.util import Emu

from .models import PresentationPlan, SlideContent
from .slide_builder import build_slide
from .corporate_design import SLIDE_WIDTH_EMU, SLIDE_HEIGHT_EMU


class PresentationGenerator:
    """Generates a .pptx file from a :class:`PresentationPlan`."""

    def generate(self, plan: PresentationPlan) -> Path:
        """Build and save the presentation; returns the output path."""
        prs = Presentation()
        prs.slide_width = Emu(SLIDE_WIDTH_EMU)
        prs.slide_height = Emu(SLIDE_HEIGHT_EMU)

        for slide_number, slide_content in enumerate(plan.slides, start=1):
            build_slide(prs, slide_content, slide_number)

        output_path = Path(plan.request.output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        prs.save(str(output_path))
        return output_path
