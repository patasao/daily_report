"""Command-line entry point for the live integration pipeline."""
import argparse
from pipeline import run


def main() -> None:
    parser = argparse.ArgumentParser(description="Run the daily reporting pipeline for one PMS.")
    parser.add_argument("system", choices=("portal_pms", "cloud_pms"), help="Configured PMS integration")
    args = parser.parse_args()
    run(args.system)


if __name__ == "__main__":
    main()
