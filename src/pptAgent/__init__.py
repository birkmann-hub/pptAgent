"""pptAgent – consulting presentation generator."""

from .models import PresentationRequest, SlideContent
from .generator import PresentationGenerator
from .agent import PresentationAgent

__all__ = [
    "PresentationRequest",
    "SlideContent",
    "PresentationGenerator",
    "PresentationAgent",
]
