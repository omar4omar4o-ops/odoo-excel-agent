import sys
import argparse

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--run-background", action="store_true", help="Run the background agent instead of the UI")
    parser.add_argument("--config", default="", help="Path to config.json")
    args, unknown = parser.parse_known_args()

    if args.run_background:
        import odoo_excel_background
        sys.argv = [sys.argv[0]] + unknown
        if args.config:
            sys.argv.extend(["--config", args.config])
        sys.exit(odoo_excel_background.main(sys.argv[1:]))
    else:
        import odoo_excel_agent_ui
        sys.argv = [sys.argv[0]] + unknown
        if args.config:
            sys.argv.extend(["--config", args.config])
        sys.exit(odoo_excel_agent_ui.main(sys.argv[1:]))

if __name__ == "__main__":
    main()
