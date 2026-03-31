#!/usr/bin/env python3
"""
Vendor Command Center - Entry Point

An autonomous agent loop that proactively manages vendor relationships.
Processes emails, handles ad-hoc requests from Telegram/Teams, and
proactively scans vendors when the inbox is clear.

Usage:
    python main.py                   # Run the full loop
    python main.py --vendor "Acme"   # Process a single vendor
    python main.py --scan-only       # Run proactive scan only
    python main.py --dry-run         # Analyze but don't execute actions
"""
import argparse
import logging
import sys

from config import settings
from core.loop import CommandCenter


def setup_logging():
    logging.basicConfig(
        level=getattr(logging, settings.LOG_LEVEL),
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def main():
    parser = argparse.ArgumentParser(description="Vendor Command Center")
    parser.add_argument("--vendor", type=str, help="Process a single vendor by name")
    parser.add_argument("--scan-only", action="store_true", help="Run proactive scan only")
    parser.add_argument("--dry-run", action="store_true", help="Analyze but don't execute")
    args = parser.parse_args()

    setup_logging()
    logger = logging.getLogger("command-center")

    logger.info("=" * 60)
    logger.info("  VENDOR COMMAND CENTER")
    logger.info("  Proactive Vendor Relationship Management")
    logger.info("=" * 60)

    center = CommandCenter()

    try:
        center.initialize()

        if args.vendor:
            # Single vendor mode
            logger.info(f"Processing single vendor: {args.vendor}")
            result = center.process_single_vendor(args.vendor)
            if result:
                logger.info(f"Result: {result.summary()}")
            else:
                logger.warning(f"Could not process vendor: {args.vendor}")

        elif args.scan_only:
            # Proactive scan only
            logger.info("Running proactive scan...")
            center._run_proactive_scan()
            center._process_queue()

        else:
            # Full loop
            center.run()

    except KeyboardInterrupt:
        logger.info("\nShutting down...")
    finally:
        center.shutdown()


if __name__ == "__main__":
    main()
