import sys
import unittest
from unittest.mock import patch

import emails_command


class EmailsCommandTests(unittest.TestCase):
    def test_no_arguments_launches_the_tui(self):
        captured = {}

        def fake_main():
            captured["argv"] = list(sys.argv)

        with (
            patch.object(sys, "argv", ["emails"]),
            patch.object(emails_command, "pull_emails_main", fake_main),
        ):
            emails_command.main()

        self.assertEqual(captured["argv"], ["emails", "--tui"])

    def test_arguments_are_passed_through_unchanged(self):
        captured = {}

        def fake_main():
            captured["argv"] = list(sys.argv)

        with (
            patch.object(sys, "argv", ["emails", "--list-stored"]),
            patch.object(emails_command, "pull_emails_main", fake_main),
        ):
            emails_command.main()

        self.assertEqual(captured["argv"], ["emails", "--list-stored"])


if __name__ == "__main__":
    unittest.main()
