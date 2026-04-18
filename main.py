#!/usr/bin/env python3
"""Command-line entry point for pptAgent.

Usage
-----
    python main.py --topic "Digital Transformation Strategy" \
                   --target-group "C-Suite, financial background" \
                   --sector financial_services \
                   --project-type digital_transformation \
                   --duration 90 \
                   --presenter "Jane Doe" \
                   --email jane.doe@company.com \
                   --client "Acme Corp" \
                   --output presentation.pptx

Run ``python main.py --help`` for the full list of options.
"""

from __future__ import annotations

import argparse
import logging
import sys
from datetime import date
from pathlib import Path

# Allow running from the repo root without installing the package
sys.path.insert(0, str(Path(__file__).parent / "src"))

from pptAgent import PresentationAgent, PresentationGenerator, PresentationRequest
from pptAgent.models import ProjectType, SectorType


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="pptAgent",
        description="Generate a consulting presentation as a PowerPoint file.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    p.add_argument("--topic", required=True, help="Main topic of the presentation")
    p.add_argument("--target-group", required=True, dest="target_group",
                   help="Audience description (role, seniority, background)")
    p.add_argument(
        "--sector", required=True,
        choices=[s.value for s in SectorType],
        help="Industry sector of the client",
    )
    p.add_argument(
        "--project-type", required=True, dest="project_type",
        choices=[p.value for p in ProjectType],
        help="Type of consulting engagement",
    )
    p.add_argument(
        "--duration", required=True, type=int, dest="duration_minutes",
        help="Planned workshop/meeting duration in minutes",
    )
    p.add_argument("--presenter", default="", dest="presenter_name",
                   help="Lead presenter full name")
    p.add_argument("--email", default="", dest="contact_email",
                   help="Contact e-mail shown on the closing slide")
    p.add_argument("--client", default="", dest="client_name",
                   help="Client or company name")
    p.add_argument("--date", default=str(date.today()), dest="presentation_date",
                   help="Presentation date (YYYY-MM-DD)")
    p.add_argument("--context", default="", dest="additional_context",
                   help="Any additional context for the agent")
    p.add_argument("--output", default="presentation.pptx", dest="output_path",
                   help="Output file path (.pptx)")
    p.add_argument("--language", default="en",
                   help="Language code for generated content (e.g. 'en', 'de')")
    p.add_argument("--verbose", "-v", action="store_true",
                   help="Enable verbose logging")
    return p


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    request = PresentationRequest(
        topic=args.topic,
        target_group=args.target_group,
        sector=SectorType(args.sector),
        project_type=ProjectType(args.project_type),
        duration_minutes=args.duration_minutes,
        presenter_name=args.presenter_name,
        contact_email=args.contact_email,
        client_name=args.client_name,
        presentation_date=date.fromisoformat(args.presentation_date),
        additional_context=args.additional_context,
        output_path=args.output_path,
        language=args.language,
    )

    logging.info("Generating presentation: %s", request.topic)
    agent = PresentationAgent()
    plan = agent.run(request)
    logging.info("Slide plan: %d slides", len(plan.slides))

    generator = PresentationGenerator()
    output_path = generator.generate(plan)
    logging.info("Presentation saved to: %s", output_path)

    print(f"✅  Presentation saved to: {output_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
