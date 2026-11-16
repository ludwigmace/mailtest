"""Console entry point for the ``emails`` command.

Running ``emails`` with no arguments opens the local TUI. Any arguments are
passed through to ``pull_emails.py``, so ``emails --list-stored`` and similar
commands keep working.
"""

import sys

from pull_emails import main as pull_emails_main


def main():
    if not sys.argv[1:]:
        sys.argv = [sys.argv[0], "--tui"]
    pull_emails_main()


if __name__ == "__main__":
    main()
