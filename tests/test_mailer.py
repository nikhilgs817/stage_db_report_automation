"""Notification tests use an in-memory SMTP fake; no email leaves the machine."""

from contextlib import redirect_stderr, redirect_stdout
import io
import json
from pathlib import Path
import smtplib
import sys
import tempfile
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import report_mailer as mailer


class MailerTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.root = Path(self.temporary.name)
        self.state = self.root / ".run"
        self.output = self.root / "output"
        self.state.mkdir()
        self.output.mkdir()
        self.commit = "a" * 40
        self.request = {"script": "monthly/customer_report.sql", "commit": self.commit, "recipients": ["stakeholder@example.com"]}
        self.result = {"status": "success", "script": self.request["script"], "commit": self.commit, "rows": 3, "worksheets": 1, "report_file": "customer_report.xlsx"}
        self.write_request()
        self.write_result()
        (self.output / "customer_report.xlsx").write_bytes(b"workbook-test-data")
        self.env = {
            "SMTP_HOST": "relay.example.com", "SMTP_PORT": "587", "SMTP_SECURITY": "starttls",
            "MAIL_FROM": "reports@example.com", "SMTP_TIMEOUT_SECONDS": "30", "MAX_EMAIL_BYTES": "1048576",
            "FAILURE_RECIPIENTS": "operators@example.com", "STAKEHOLDERS": "stakeholder@example.com",
            "JOB_NAME": "STAGE DB REPORT", "BUILD_NUMBER": "1", "BUILD_URL": "https://jenkins.example.com/job/reports/1/",
        }

    def write_request(self):
        (self.state / "request.json").write_text(json.dumps(self.request), encoding="utf-8")

    def write_result(self):
        (self.output / "result.json").write_text(json.dumps(self.result), encoding="utf-8")

    def invoke(self, mode):
        output = io.StringIO()
        with redirect_stdout(output), redirect_stderr(output):
            result = mailer.main([mode, "--state-dir", str(self.state), "--output-dir", str(self.output)], self.env)
        return result, output.getvalue()

    def test_success_attaches_selected_sql_basename_and_uses_tls(self):
        with patch.object(mailer.smtplib, "SMTP") as factory:
            client = factory.return_value.__enter__.return_value
            client.send_message.return_value = {}
            code, _ = self.invoke("success")
            message = client.send_message.call_args.args[0]
            self.assertEqual(code, 0)
            self.assertEqual([p.get_filename() for p in message.iter_attachments()], ["customer_report.xlsx"])
            self.assertEqual(message["To"], "stakeholder@example.com")
            self.assertEqual(client.send_message.call_args.kwargs["to_addrs"], ["stakeholder@example.com"])
            client.starttls.assert_called_once()
            client.login.assert_not_called()
        record = json.loads((self.output / "notification.success.json").read_text())
        self.assertEqual(record["status"], "accepted_by_relay")

    def test_failure_has_no_attachment_and_includes_operator_and_sanitized_error(self):
        self.result.update(status="failure", error="DB2_PASSWORD=secret SELECT account FROM private_table")
        self.write_result()
        self.env["DB2_PASSWORD"] = "secret"
        self.env["FAILURE_STAGE"] = "Generate report"
        settings = mailer.load_settings(self.env)
        message, recipients = mailer.build_message("failure", self.state, self.output, self.env, settings)
        self.assertEqual(list(message.iter_attachments()), [])
        self.assertEqual(recipients, ["stakeholder@example.com", "operators@example.com"])
        body = message.get_content()
        self.assertIn("Generate report", body)
        self.assertNotIn("secret", body)
        self.assertNotIn("private_table", body)

    def test_invalid_stakeholders_fall_back_to_operators_when_request_missing(self):
        (self.state / "request.json").unlink()
        self.env["STAKEHOLDERS"] = "valid@example.com\r\nBcc: other@example.com"
        settings = mailer.load_settings(self.env)
        _, recipients = mailer.build_message("failure", self.state, self.output, self.env, settings)
        self.assertEqual(recipients, ["operators@example.com"])

    def test_partial_recipient_refusal_fails_without_retry(self):
        with patch.object(mailer.smtplib, "SMTP") as factory:
            client = factory.return_value.__enter__.return_value
            client.send_message.return_value = {"stakeholder@example.com": (550, b"policy rejection")}
            code, text = self.invoke("success")
            self.assertEqual(code, 1)
            self.assertIn("some recipients may already", text)
            client.send_message.assert_called_once()
        self.assertEqual(json.loads((self.output / "notification.json").read_text())["status"], "failed")

    def test_smtp_error_does_not_echo_server_text_or_password(self):
        with patch.object(mailer.smtplib, "SMTP") as factory:
            client = factory.return_value.__enter__.return_value
            client.send_message.side_effect = smtplib.SMTPDataError(554, b"secret body SQL SELECT account")
            code, text = self.invoke("success")
            self.assertEqual(code, 1)
            self.assertIn("SMTPDataError (SMTP 554)", text)
            self.assertNotIn("secret body", text)
            self.assertNotIn("SELECT", text)

    def test_mime_encoded_size_limit_fails_before_connecting(self):
        # Raw attachment fits; base64 plus headers exceed the configured cap.
        (self.output / "customer_report.xlsx").write_bytes(b"x" * 2000)
        self.env["MAX_EMAIL_BYTES"] = "2500"
        with patch.object(mailer.smtplib, "SMTP") as factory:
            code, text = self.invoke("success")
            self.assertEqual(code, 1)
            self.assertIn("after MIME encoding", text)
            factory.assert_not_called()

    def test_wrong_revision_result_is_not_emailed(self):
        self.result["commit"] = "b" * 40
        self.write_result()
        with patch.object(mailer.smtplib, "SMTP") as factory:
            code, _ = self.invoke("success")
            self.assertEqual(code, 1)
            factory.assert_not_called()

    def test_wrong_report_filename_is_not_emailed(self):
        self.result["report_file"] = "other.xlsx"
        self.write_result()
        with patch.object(mailer.smtplib, "SMTP") as factory:
            self.assertEqual(self.invoke("success")[0], 1)
            factory.assert_not_called()

    def test_validate_checks_settings_without_sending_or_writing_a_notification(self):
        with patch.object(mailer.smtplib, "SMTP") as factory:
            self.assertEqual(self.invoke("validate")[0], 0)
            factory.assert_not_called()
        self.assertFalse((self.output / "notification.json").exists())
        self.env.pop("SMTP_HOST")
        self.assertEqual(self.invoke("validate")[0], 1)

    def test_validate_rejects_stakeholder_header_injection_and_undotted_domain(self):
        for address in ("person@example.com\r\nBcc: other@example.com", "person@internal", "not a mailbox"):
            with self.subTest(address=address), patch.object(mailer.smtplib, "SMTP") as factory:
                self.env["STAKEHOLDERS"] = address
                self.assertEqual(self.invoke("validate")[0], 1)
                factory.assert_not_called()

    def test_plain_mode_rejects_credentials(self):
        self.env.update(SMTP_SECURITY="plain", SMTP_USER="smtp-user", SMTP_PASSWORD="smtp-password")
        with self.assertRaisesRegex(mailer.MailerError, "authentication requires"):
            mailer.load_settings(self.env)

    def test_failure_status_keeps_original_success_notification_failure_evidence(self):
        with patch.object(mailer.smtplib, "SMTP") as factory:
            client = factory.return_value.__enter__.return_value
            client.send_message.side_effect = [smtplib.SMTPServerDisconnected("lost"), {}]
            self.assertEqual(self.invoke("success")[0], 1)
            self.assertEqual(self.invoke("failure")[0], 0)
        self.assertEqual(json.loads((self.output / "notification.success.json").read_text())["status"], "failed")
        self.assertEqual(json.loads((self.output / "notification.failure.json").read_text())["status"], "accepted_by_relay")
        failure_message = client.send_message.call_args.args[0]
        self.assertIn("SMTPServerDisconnected", failure_message.get_content())

    def test_request_path_cannot_select_unrelated_attachment(self):
        self.request["script"] = "../other.sql"
        self.write_request()
        with patch.object(mailer.smtplib, "SMTP") as factory:
            self.assertEqual(self.invoke("success")[0], 1)
            factory.assert_not_called()


if __name__ == "__main__":
    unittest.main()
