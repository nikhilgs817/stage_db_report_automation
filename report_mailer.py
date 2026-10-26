"""Send report notifications through an approved SMTP relay.

All site settings come from Jenkins environment parameters. Passwords are supplied
by Jenkins Credentials Binding. A successful send means the relay accepted the
message, not that every recipient's mailbox has delivered it.
"""

from __future__ import annotations

import argparse
from datetime import datetime, timezone
from email.message import EmailMessage
from email.policy import SMTP as SMTP_POLICY
import json
import os
from pathlib import Path, PurePosixPath
import re
import smtplib
import ssl
import sys
import tempfile


class MailerError(Exception):
    """An operator-readable failure that does not include mail or secret data."""


_ADDRESS = re.compile(r"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9](?:[A-Za-z0-9.-]*[A-Za-z0-9])?\.[A-Za-z]{2,63}\Z")
_COMMIT = re.compile(r"[0-9a-fA-F]{40}\Z")


def parse_addresses(value: str) -> list[str]:
    """Accept plain mailbox addresses, not display names or header fragments."""
    if not isinstance(value, str) or any(ord(c) < 32 or ord(c) == 127 for c in value):
        raise MailerError("Email recipients contain a control character.")
    parts = [part.strip() for part in re.split(r"[,;]", value)]
    if not parts or any(not part or not _ADDRESS.fullmatch(part) for part in parts):
        raise MailerError("Provide comma- or semicolon-separated email addresses.")
    if len(parts) > 200:
        raise MailerError("Too many email recipients; use a distribution list.")
    return list(dict.fromkeys(parts))


def _required(env: dict, name: str) -> str:
    value = env.get(name, "").strip()
    if not value or any(ord(c) < 32 or ord(c) == 127 for c in value):
        raise MailerError(f"Configure {name} with a nonempty single-line value.")
    return value


def _integer(env: dict, name: str, minimum: int, maximum: int) -> int:
    try:
        value = int(_required(env, name))
    except ValueError:
        raise MailerError(f"{name} must be an integer.") from None
    if not minimum <= value <= maximum:
        raise MailerError(f"{name} must be between {minimum} and {maximum}.")
    return value


def load_settings(env: dict) -> dict:
    security = _required(env, "SMTP_SECURITY").lower()
    if security not in {"starttls", "ssl", "plain"}:
        raise MailerError("SMTP_SECURITY must be starttls, ssl, or plain.")
    sender = parse_addresses(_required(env, "MAIL_FROM"))
    if len(sender) != 1:
        raise MailerError("MAIL_FROM must contain one email address.")
    user, password = env.get("SMTP_USER", ""), env.get("SMTP_PASSWORD", "")
    if bool(user) != bool(password):
        raise MailerError("Supply both SMTP_USER and SMTP_PASSWORD, or neither.")
    if security == "plain" and user:
        raise MailerError("SMTP authentication requires starttls or ssl.")
    ca_file = env.get("SMTP_CA_FILE", "").strip()
    if ca_file and not Path(ca_file).is_file():
        raise MailerError("SMTP_CA_FILE does not exist.")
    return {
        "host": _required(env, "SMTP_HOST"),
        "port": _integer(env, "SMTP_PORT", 1, 65535),
        "security": security,
        "sender": sender[0],
        "timeout": _integer(env, "SMTP_TIMEOUT_SECONDS", 1, 300),
        "max_bytes": _integer(env, "MAX_EMAIL_BYTES", 1024, 100 * 1024 * 1024),
        "ca_file": ca_file or None,
        "user": user,
        "password": password,
        "operators": parse_addresses(_required(env, "FAILURE_RECIPIENTS")),
    }


def sanitize_error(value: object, env: dict) -> str:
    """Bound metadata and redact common credentials, JDBC URLs and SQL text."""
    text = str(value or "No additional error details were recorded.")
    for name in ("SMTP_PASSWORD", "DB2_PASSWORD", "SMTP_USER", "DB2_USER"):
        secret = env.get(name, "")
        if secret:
            text = text.replace(secret, "[redacted]")
    text = re.sub(r"jdbc:[^\s]+", "[JDBC URL redacted]", text, flags=re.IGNORECASE)
    text = re.sub(r"(?i)(password|passwd|pwd)\s*[=:]\s*[^\s;,]+", r"\1=[redacted]", text)
    # Drivers may echo an entire query. Preserve diagnostics before that query.
    text = re.split(r"\b(?:SELECT\s+|WITH\s+\w+\s+AS\s*\(|INSERT\s+INTO\s+|UPDATE\s+\w+\s+SET\s+|DELETE\s+FROM\s+)", text, maxsplit=1, flags=re.IGNORECASE)[0]
    return "".join(c if c in "\n\t" or ord(c) >= 32 else " " for c in text)[:2000].strip() or "See the Jenkins build log for diagnostics."


def _read_json(path: Path) -> dict:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        raise MailerError(f"Missing or invalid {path.name}.") from None
    if not isinstance(value, dict):
        raise MailerError(f"Invalid {path.name} object.")
    return value


def _request_recipients(request: dict) -> list[str]:
    recipients = request.get("recipients")
    if not isinstance(recipients, list) or any(not isinstance(v, str) for v in recipients):
        raise MailerError("Request contains invalid email recipients.")
    return parse_addresses(",".join(recipients))


def _script_name(request: dict) -> str:
    script = request.get("script", "")
    if not isinstance(script, str) or not script or "\\" in script or "|" in script:
        raise MailerError("Request contains an invalid SQL script path.")
    path = PurePosixPath(script)
    if path.is_absolute() or any(p in {"", ".", ".."} for p in script.split("/")):
        raise MailerError("Request contains an invalid SQL script path.")
    if any(ord(c) < 32 or ord(c) == 127 or c in ':<>"?*' for c in script):
        raise MailerError("Request contains an invalid SQL script path.")
    if path.suffix.lower() != ".sql" or not _COMMIT.fullmatch(str(request.get("commit", ""))):
        raise MailerError("Request must identify a SQL script and Git revision.")
    return path.stem + ".xlsx"


def build_message(mode: str, state_dir: Path, output_dir: Path, env: dict, settings: dict) -> tuple[EmailMessage, list[str]]:
    if mode not in {"success", "failure"}:
        raise MailerError("Unknown notification mode.")
    try:
        request = _read_json(state_dir / "request.json")
    except MailerError:
        if mode == "success":
            raise
        request = {}
    try:
        result = _read_json(output_dir / "result.json")
    except MailerError:
        if mode == "success":
            raise
        result = {}
    if mode == "success":
        recipients = _request_recipients(request)
        filename = _script_name(request)
        if result.get("status") != "success" or result.get("script") != request["script"] or result.get("commit") != request["commit"]:
            raise MailerError("Successful result does not match the selected SQL script and revision.")
        if result.get("report_file") != filename:
            raise MailerError("Successful result does not identify the expected workbook filename.")
    else:
        try:
            recipients = _request_recipients(request)
        except MailerError:
            try:
                recipients = parse_addresses(env.get("STAKEHOLDERS", ""))
            except MailerError:
                recipients = []
        recipients = list(dict.fromkeys(recipients + settings["operators"]))
    message = EmailMessage(policy=SMTP_POLICY)
    message["From"] = settings["sender"]
    message["To"] = ", ".join(recipients)
    job = sanitize_error(env.get("JOB_NAME", "STAGE DB REPORT"), env).replace("\n", " ").replace("\r", " ")[:150]
    build = sanitize_error(env.get("BUILD_NUMBER", "manual"), env).replace("\n", " ").replace("\r", " ")[:30]
    message["Subject"] = f"[{mode.upper()}] {job} #{build}"
    details = [f"Job: {job}", f"Build: {build}"]
    if request.get("script"):
        details.append("SQL script: " + sanitize_error(request["script"], env))
    if request.get("commit"):
        details.append("Git revision: " + sanitize_error(request["commit"], env))
    if env.get("BUILD_URL"):
        details.append("Jenkins build: " + sanitize_error(env["BUILD_URL"], env))
    if mode == "success":
        details += ["Report generation completed.", f"Rows: {result.get('rows', 'see report')}", f"Worksheets: {result.get('worksheets', 'see report')}"]
    else:
        details.append("Failed stage: " + sanitize_error(env.get("FAILURE_STAGE", "unknown"), env))
        if result.get("status") == "failure" and result.get("error"):
            details.append("Report error: " + sanitize_error(result["error"], env))
        if env.get("PIPELINE_ERROR"):
            details.append("Pipeline error: " + sanitize_error(env["PIPELINE_ERROR"], env))
        try:
            report_notification = _read_json(output_dir / "notification.success.json")
        except MailerError:
            report_notification = {}
        if report_notification.get("status") == "failed" and report_notification.get("error"):
            details.append("Report email error: " + sanitize_error(report_notification["error"], env))
        if result.get("status") != "failure" and not env.get("PIPELINE_ERROR"):
            details.append("No report error was recorded; review the Jenkins build log.")
        details.append("No report is attached to this failure notification.")
    message.set_content("\n".join(details))
    if mode == "success":
        target = output_dir / filename
        root = output_dir.resolve()
        if target.is_symlink() or target.resolve().parent != root or not target.is_file():
            raise MailerError("Expected generated workbook is missing or outside the output directory.")
        # The path is derived from the validated selection, never from an arbitrary
        # artifact field in result.json. Bound the read before constructing MIME.
        if target.stat().st_size > settings["max_bytes"]:
            raise MailerError("Workbook exceeds MAX_EMAIL_BYTES before MIME encoding.")
        message.add_attachment(target.read_bytes(), maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=filename)
    if len(message.as_bytes()) > settings["max_bytes"]:
        raise MailerError("Email exceeds MAX_EMAIL_BYTES after MIME encoding.")
    return message, recipients


def send_message(message: EmailMessage, recipients: list[str], settings: dict) -> None:
    """One attempt only: a retry after an uncertain/partial send can duplicate mail."""
    try:
        if settings["security"] == "ssl":
            context = ssl.create_default_context(cafile=settings["ca_file"])
            client = smtplib.SMTP_SSL(settings["host"], settings["port"], timeout=settings["timeout"], context=context)
        else:
            client = smtplib.SMTP(settings["host"], settings["port"], timeout=settings["timeout"])
        with client as client:
            client.ehlo()
            if settings["security"] == "starttls":
                context = ssl.create_default_context(cafile=settings["ca_file"])
                client.starttls(context=context)
                client.ehlo()
            if settings["user"]:
                client.login(settings["user"], settings["password"])
            refused = client.send_message(message, from_addr=settings["sender"], to_addrs=recipients)
            if refused:
                raise MailerError(f"SMTP refused {len(refused)} recipient(s); some recipients may already have received the message. No automatic retry was attempted.")
    except MailerError:
        raise
    except Exception as exc:
        code = getattr(exc, "smtp_code", None)
        label = type(exc).__name__
        if isinstance(code, int):
            label += f" (SMTP {code})"
        raise MailerError(f"SMTP notification failed: {label}. Delivery may be uncertain; no automatic retry was attempted.") from None


def _record(output_dir: Path, record: dict) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(record, indent=2) + "\n"
    for name in ("notification.json", f"notification.{record['mode']}.json"):
        temporary = None
        try:
            with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=output_dir, prefix=".mail-", suffix=".json", delete=False) as stream:
                stream.write(payload)
                temporary = Path(stream.name)
            temporary.replace(output_dir / name)
        finally:
            if temporary is not None:
                temporary.unlink(missing_ok=True)


def main(argv: list[str] | None = None, env: dict | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("mode", choices=("validate", "success", "failure"))
    parser.add_argument("--state-dir", type=Path, default=Path(".run"))
    parser.add_argument("--output-dir", type=Path, default=Path("output"))
    args = parser.parse_args(argv)
    environment = dict(os.environ if env is None else env)
    if args.mode == "validate":
        try:
            load_settings(environment)
            parse_addresses(_required(environment, "STAKEHOLDERS"))
        except Exception as exc:
            error = str(exc) if isinstance(exc, MailerError) else f"Mail configuration validation failed: {type(exc).__name__}."
            print(sanitize_error(error, environment), file=sys.stderr)
            return 1
        print("Mail configuration validated; no message was sent.")
        return 0
    record = {"mode": args.mode, "at_utc": datetime.now(timezone.utc).isoformat()}
    try:
        settings = load_settings(environment)
        message, recipients = build_message(args.mode, args.state_dir, args.output_dir, environment, settings)
        send_message(message, recipients, settings)
        record.update(status="accepted_by_relay", recipient_count=len(recipients))
        code = 0
    except Exception as exc:
        error = str(exc) if isinstance(exc, MailerError) else f"Notification failed: {type(exc).__name__}."
        record.update(status="failed", error=sanitize_error(error, environment))
        code = 1
    try:
        _record(args.output_dir, record)
    except OSError:
        print("Notification status could not be written.", file=sys.stderr)
        return 1
    print("Notification accepted by SMTP relay." if code == 0 else record["error"], file=sys.stdout if code == 0 else sys.stderr)
    return code


if __name__ == "__main__":
    raise SystemExit(main())
