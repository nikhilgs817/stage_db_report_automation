"""V1 Db2 report backend. Credentials are read from Jenkins-bound environment variables.

SQL filtering and JDBC readOnly are defense in depth. The DBA-controlled SELECT-only
account, reviewed Git repository, and Jenkins permissions form the security boundary.
"""
from __future__ import annotations

import argparse
import contextlib
import datetime as dt
from decimal import Decimal
import json
import math
import os
from pathlib import Path, PurePosixPath
import re
import stat
import subprocess
import sys
import tempfile
from typing import Any


class ReportError(Exception):
    """A controlled diagnostic safe to include in a notification."""


SHA_RE = re.compile(r"[0-9a-fA-F]{40}\Z")
EMAIL_RE = re.compile(r"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9](?:[A-Za-z0-9.-]*[A-Za-z0-9])?\Z")
INVALID_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]")
WINDOWS_RESERVED = re.compile(r"(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\..*)?\Z", re.I)
RUNTIME_DEFAULTS = {
    "jvm_path": None,
    "jvm_options": ["-Xmx512m"],
    "login_timeout_seconds": 30,
    "network_timeout_seconds": 120,
    "query_timeout_seconds": 900,
    "fetch_size": 500,
    "max_sql_bytes": 1048576,
    "max_rows_per_sheet": 1048575,
    "max_report_bytes": 20971520,
}
PROPERTY_ALLOWLIST = {
    "currentSchema", "currentSQLID", "currentPackageSet", "currentPackagePath",
    "clientProgramName", "clientWorkstation", "clientAccountingInformation",
    "securityMechanism",
}


def safe_script_path(value: str) -> str:
    if not isinstance(value, str) or not value or len(value) > 500:
        raise ReportError("Select a valid repository-relative SQL file.")
    if any(ord(c) < 32 or ord(c) == 127 for c in value) or any(c in value for c in '\\|<>:"?*'):
        raise ReportError("SQL paths must use forward slashes and contain no control or reserved characters.")
    parts = value.split("/")
    if any(not p or p in (".", "..") or p.endswith((" ", ".")) or WINDOWS_RESERVED.fullmatch(p) for p in parts):
        raise ReportError("SQL paths must be safe, relative paths without traversal or reserved Windows names.")
    path = PurePosixPath(value)
    if path.is_absolute() or path.suffix.lower() != ".sql":
        raise ReportError("The selected file must be a repository-relative .sql file.")
    return value


def parse_selection(value: str) -> tuple[str, str]:
    if not isinstance(value, str) or value.count("|") != 1:
        raise ReportError("SCRIPT_SELECTION must come from the SQL catalog (path|40-character Git commit).")
    path, commit = value.split("|")
    safe_script_path(path)
    if not SHA_RE.fullmatch(commit):
        raise ReportError("The selected Git commit must be a full 40-character SHA-1 revision.")
    return path, commit.lower()


def parse_recipients(value: str) -> list[str]:
    if not isinstance(value, str) or not value.strip() or any(c in value for c in "\r\n\x00"):
        raise ReportError("STAKEHOLDERS must contain comma- or semicolon-separated email addresses without line breaks.")
    recipients = [address.strip() for address in re.split(r"[,;]", value)]
    if len(recipients) > 100 or any(not EMAIL_RE.fullmatch(a) or len(a) > 254 or ".." in a for a in recipients):
        raise ReportError("STAKEHOLDERS contains an invalid email address; use addresses only, without display names.")
    return list(dict.fromkeys(recipients))


def atomic_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    handle, temporary = tempfile.mkstemp(prefix=path.name + ".", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(handle, "w", encoding="utf-8", newline="\n") as stream:
            json.dump(data, stream, indent=2, ensure_ascii=True)
            stream.write("\n")
        os.replace(temporary, path)
    finally:
        with contextlib.suppress(FileNotFoundError):
            os.unlink(temporary)


def git(repo_dir: Path, *arguments: str) -> bytes:
    try:
        result = subprocess.run(["git", "-C", str(repo_dir), *arguments], check=False,
                                stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60,
                                shell=False)
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise ReportError("Git could not inspect the checkout; check Git installation and repository access.") from exc
    if result.returncode:
        raise ReportError("Git could not inspect the checkout; verify the selected revision and repository configuration.")
    return result.stdout


def head_commit(repo_dir: Path) -> str:
    commit = git(repo_dir, "rev-parse", "--verify", "HEAD").decode("ascii").strip().lower()
    if not SHA_RE.fullmatch(commit):
        raise ReportError("V1 requires a repository using 40-character SHA-1 Git revisions.")
    return commit


def checked_file(repo_dir: Path, relative_path: str) -> Path:
    safe_script_path(relative_path)
    root = repo_dir.resolve(strict=True)
    candidate = root
    for part in PurePosixPath(relative_path).parts:
        candidate = candidate / part
        try:
            information = candidate.lstat()
        except OSError as exc:
            raise ReportError("The selected SQL file is missing from the checkout.") from exc
        if stat.S_ISLNK(information.st_mode) or getattr(information, "st_file_attributes", 0) & 0x400:
            raise ReportError("SQL files and their parent directories must not be symlinks or Windows reparse points.")
    if not candidate.is_file() or not candidate.resolve().is_relative_to(root):
        raise ReportError("The selected SQL file is not a regular file inside the report checkout.")
    return candidate


def tracked_sql(repo_dir: Path) -> list[str]:
    files = []
    # Tree modes exclude symlinks (120000) and submodules (160000).
    for record in git(repo_dir, "ls-tree", "-r", "-z", "HEAD").split(b"\0"):
        if not record:
            continue
        header, name = record.split(b"\t", 1)
        mode, kind, _object_id = header.split(b" ", 2)
        if kind != b"blob" or mode not in (b"100644", b"100755"):
            continue
        try:
            path = name.decode("utf-8")
            safe_script_path(path)
            checked_file(repo_dir, path)
        except (UnicodeError, ReportError):
            continue
        files.append(path)
    return sorted(files, key=lambda p: (p.casefold(), p))


def read_sql_at_commit(repo_dir: Path, path: str, commit: str, max_bytes: int = 1048576) -> str:
    """Check the pinned revision and clean worktree, then read immutable Git bytes."""
    path, commit = parse_selection(f"{path}|{commit}")
    if head_commit(repo_dir) != commit:
        raise ReportError("The checked-out report revision differs from SCRIPT_SELECTION; refresh the catalog and retry.")
    if path not in tracked_sql(repo_dir):
        raise ReportError("The selected SQL file is not a safe, committed file at the selected revision.")
    selected_file = checked_file(repo_dir, path)
    if selected_file.stat().st_size > max_bytes:
        raise ReportError("The selected SQL file exceeds runtime.max_sql_bytes.")
    script_bytes = selected_file.read_bytes()
    committed = git(repo_dir, "show", f"{commit}:{path}")
    if len(committed) > max_bytes:
        raise ReportError("The committed SQL file exceeds runtime.max_sql_bytes.")
    if script_bytes.replace(b"\r\n", b"\n") != committed.replace(b"\r\n", b"\n"):
        raise ReportError("The selected SQL file has uncommitted changes; use a clean report checkout.")
    try:
        return committed.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        raise ReportError("SQL files must be stored as UTF-8 text (UTF-8 BOM is accepted).") from exc


def split_read_queries(sql: str) -> list[str]:
    import sqlparse
    from sqlparse import tokens as tokens

    forbidden = {
        "INSERT", "UPDATE", "DELETE", "MERGE", "CREATE", "ALTER", "DROP", "TRUNCATE",
        "CALL", "EXEC", "EXECUTE", "GRANT", "REVOKE", "COMMIT", "ROLLBACK", "SET",
        "CONNECT", "DISCONNECT", "TERMINATE", "EXPORT", "IMPORT", "LOAD", "UNLOAD",
        "BEGIN", "DECLARE", "INTO", "REPLACE", "RENAME", "COMMENT", "REFRESH",
    }
    statements = []
    for statement in sqlparse.parse(sql):
        significant = [t for t in statement.flatten()
                       if not t.is_whitespace and t.ttype not in tokens.Comment]
        if not significant or all(t.value == ";" for t in significant):
            continue
        if statement.get_type() != "SELECT" or significant[0].normalized not in ("SELECT", "WITH", "WITH RECURSIVE"):
            raise ReportError("V1 SQL files may contain SELECT queries or WITH ... SELECT queries only.")
        for token in significant:
            if token.ttype in tokens.Literal.String or token.ttype in tokens.Name:
                continue
            if token.ttype in tokens.Keyword and any(word in forbidden for word in token.normalized.split()):
                raise ReportError("A SQL statement contains a command outside the V1 read-only query policy.")
        # Remove the SQL-file delimiter even when comments follow it. JDBC gets
        # one statement per call and must not receive the client delimiter.
        terminal = significant[-1] if significant[-1].value == ";" else None
        statement_sql = "".join(token.value for token in statement.flatten() if token is not terminal).strip()
        statements.append(statement_sql)
    if not statements:
        raise ReportError("The selected SQL file contains no report queries.")
    if len(statements) > 100:
        raise ReportError("V1 supports at most 100 result queries per SQL file.")
    return statements


def excel_value(value: Any) -> Any:
    """Preserve large numbers as text; reject lossy Excel cells instead of truncating."""
    if value is None or isinstance(value, bool):
        return value
    if isinstance(value, Decimal):
        if not value.is_finite():
            raise ReportError("A query returned a non-finite decimal, which Excel cannot represent.")
        outside_numeric_range = value != 0 and (value.adjusted() > 307 or value.adjusted() < -307)
        return str(value) if len(value.as_tuple().digits) > 15 or outside_numeric_range else value
    if isinstance(value, int):
        return str(value) if len(str(abs(value))) > 15 else value
    if isinstance(value, float):
        if not math.isfinite(value):
            raise ReportError("A query returned a non-finite number, which Excel cannot represent.")
        return value
    if isinstance(value, (dt.date, dt.datetime, dt.time)):
        if isinstance(value, (dt.datetime, dt.time)) and value.tzinfo is not None:
            return value.isoformat()
        return value
    if isinstance(value, bytes):
        value = value.hex()
    if not isinstance(value, str):
        value = str(value)
    if INVALID_XML.search(value):
        raise ReportError("A query returned text containing characters unsupported by Excel XML.")
    if len(value.encode("utf-16-le")) // 2 > 32767:
        raise ReportError("A query returned text longer than Excel's 32,767-character cell limit.")
    return value


def _literal_cell(sheet: Any, value: Any, *, header: bool = False) -> Any:
    from openpyxl.cell import WriteOnlyCell
    from openpyxl.styles import Font

    value = excel_value(value)
    cell = WriteOnlyCell(sheet, value=value)
    if isinstance(value, str):
        # Force an OOXML string cell, including strings starting =, +, -, @.
        # This retains the original value and prevents formula interpretation.
        cell.data_type = "s"
    if header:
        cell.font = Font(bold=True)
    return cell


def _jdbc_value(result_set: Any, index: int, jdbc_type: int) -> Any:
    text = result_set.getString(index)
    if text is None:
        return None
    text = str(text)
    if jdbc_type in (-7, 16):  # BIT, BOOLEAN
        return bool(result_set.getBoolean(index))
    if jdbc_type in (-6, 5, 4, -5):  # TINYINT, SMALLINT, INTEGER, BIGINT
        return int(text)
    if jdbc_type in (2, 3):  # NUMERIC, DECIMAL: never convert through Java double.
        return Decimal(text)
    if jdbc_type in (6, 7, 8):  # FLOAT, REAL, DOUBLE
        return float(text)
    # Dates/timestamps (including z/OS fractional seconds), LOBs and binary values
    # remain JDBC text representations; no Excel timezone or precision coercion.
    return text


def export_queries(connection: Any, queries: list[str], target_path: Path,
                   runtime: dict[str, Any]) -> dict[str, int]:
    from openpyxl import Workbook

    settings = {**RUNTIME_DEFAULTS, **runtime}
    target_path = Path(target_path)
    target_path.parent.mkdir(parents=True, exist_ok=True)
    handle, temporary_name = tempfile.mkstemp(prefix=target_path.stem + ".", suffix=".xlsx", dir=target_path.parent)
    os.close(handle)
    temporary = Path(temporary_name)
    workbook = Workbook(write_only=True)
    row_count = 0
    completed = False
    try:
        for query_index, query in enumerate(queries, 1):
            statement = connection.createStatement()
            result_set = None
            try:
                statement.setQueryTimeout(settings["query_timeout_seconds"])
                statement.setFetchSize(settings["fetch_size"])
                result_set = statement.executeQuery(query)
                metadata = result_set.getMetaData()
                column_count = int(metadata.getColumnCount())
                if column_count < 1 or column_count > 16384:
                    raise ReportError("A result exceeds Excel's 16,384-column limit or contains no columns.")
                sheet = workbook.create_sheet(f"Result_{query_index}")
                sheet.freeze_panes = "A2"
                sheet.append([_literal_cell(sheet, str(metadata.getColumnLabel(i)), header=True)
                              for i in range(1, column_count + 1)])
                types = [int(metadata.getColumnType(i)) for i in range(1, column_count + 1)]
                sheet_rows = 0
                while result_set.next():
                    sheet_rows += 1
                    if sheet_rows > settings["max_rows_per_sheet"] or sheet_rows > 1048575:
                        raise ReportError("A result exceeds the configured data-row limit or Excel's 1,048,575 data-row limit.")
                    values = [_literal_cell(sheet, _jdbc_value(result_set, i, types[i - 1]))
                              for i in range(1, column_count + 1)]
                    sheet.append(values)
                    row_count += 1
                sheet.auto_filter.ref = f"A1:{_excel_column(column_count)}{sheet_rows + 1}"
            finally:
                try:
                    if result_set is not None:
                        result_set.close()
                finally:
                    statement.close()
        workbook.save(temporary)
        if temporary.stat().st_size > settings["max_report_bytes"]:
            raise ReportError("The generated workbook exceeds max_report_bytes; narrow the query or arrange an approved delivery route.")
        os.replace(temporary, target_path)
        completed = True
        return {"rows": row_count, "worksheets": len(queries)}
    finally:
        if not completed:
            # Write-only worksheet writers use temporary XML files. Close and
            # remove those on errors so partial data is not left in TEMP.
            for sheet in workbook.worksheets:
                with contextlib.suppress(Exception):
                    sheet.close()
                writer = getattr(sheet, "_writer", None)
                if writer is not None:
                    with contextlib.suppress(Exception):
                        writer.cleanup()
        workbook.close()
        with contextlib.suppress(FileNotFoundError):
            temporary.unlink()


def _excel_column(number: int) -> str:
    letters = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _positive_int(value: Any, key: str, maximum: int = 2147483647) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or not 1 <= value <= maximum:
        raise ReportError(f"Configuration {key} must be a positive integer no greater than {maximum}.")
    return value


def load_config() -> dict[str, Any]:
    config_path = Path(os.environ.get("REPORT_CONFIG", ""))
    if not config_path.is_absolute() or not config_path.is_file():
        raise ReportError("REPORT_CONFIG must name an existing absolute path to an administrator-managed JSON file.")
    try:
        config = json.loads(config_path.read_text(encoding="utf-8-sig"))
        database = config["database"]
        runtime = {**RUNTIME_DEFAULTS, **config.get("runtime", {})}
        if not isinstance(database, dict) or not isinstance(runtime, dict):
            raise TypeError
    except (ValueError, KeyError, TypeError, OSError) as exc:
        raise ReportError("REPORT_CONFIG is not a valid report configuration JSON document.") from exc
    for key in ("host", "name"):
        if not isinstance(database.get(key), str) or not re.fullmatch(r"[A-Za-z0-9_.-]+", database[key]):
            raise ReportError(f"Configuration database.{key} must contain only letters, digits, underscores, dots or hyphens.")
    _positive_int(database.get("port"), "database.port", 65535)
    jars = database.get("jdbc_jars")
    if not isinstance(jars, list) or not jars or any(not isinstance(p, str) or not Path(p).is_absolute() or not Path(p).is_file() for p in jars):
        raise ReportError("Configuration database.jdbc_jars must list existing absolute IBM driver/license JAR paths.")
    cert = database.get("ssl_cert_location")
    if not isinstance(cert, str) or not Path(cert).is_absolute() or not Path(cert).is_file():
        raise ReportError("Configuration database.ssl_cert_location must name an existing absolute TLS certificate file.")
    properties = database.get("properties", {})
    if not isinstance(properties, dict) or any(key not in PROPERTY_ALLOWLIST or not isinstance(value, (str, int))
                                                or isinstance(value, bool) for key, value in properties.items()):
        raise ReportError("database.properties contains an unsupported property or value; consult the configuration guide.")
    for key in RUNTIME_DEFAULTS:
        if key not in ("jvm_path", "jvm_options"):
            _positive_int(runtime[key], f"runtime.{key}")
    if runtime["max_rows_per_sheet"] > 1048575:
        raise ReportError("runtime.max_rows_per_sheet cannot exceed Excel's data-row limit.")
    if runtime["jvm_path"] is not None and (not isinstance(runtime["jvm_path"], str) or not Path(runtime["jvm_path"]).is_file()):
        raise ReportError("runtime.jvm_path must be null or an existing JVM library file (jvm.dll on Windows).")
    if not isinstance(runtime["jvm_options"], list) or any(not isinstance(option, str) for option in runtime["jvm_options"]):
        raise ReportError("runtime.jvm_options must be a JSON list of JVM option strings.")
    return {"database": database, "runtime": runtime}


@contextlib.contextmanager
def db2_connection(config: dict[str, Any]):
    import jpype

    username = os.environ.get("DB2_USER")
    password = os.environ.get("DB2_PASSWORD")
    if not username or not password:
        raise ReportError("DB2_USER and DB2_PASSWORD must be bound from a Jenkins username/password credential.")
    database, runtime = config["database"], config["runtime"]
    if not jpype.isJVMStarted():
        jpype.startJVM(*runtime["jvm_options"], jvmpath=runtime["jvm_path"],
                       classpath=database["jdbc_jars"], convertStrings=True)
    driver = jpype.JClass("com.ibm.db2.jcc.DB2Driver")()
    driver_manager = jpype.JClass("java.sql.DriverManager")
    driver_manager.setLoginTimeout(runtime["login_timeout_seconds"])
    properties = jpype.JClass("java.util.Properties")()
    settings = {**database.get("properties", {}), "user": username, "password": password,
                "sslConnection": "true", "sslCertLocation": database["ssl_cert_location"],
                "sslClientHostnameValidation": "BASIC",
                "loginTimeout": runtime["login_timeout_seconds"],
                "blockingReadConnectionTimeout": runtime["network_timeout_seconds"]}
    for key, value in settings.items():
        properties.setProperty(key, str(value))
    url = f"jdbc:db2://{database['host']}:{database['port']}/{database['name']}"
    connection = driver.connect(url, properties)
    if connection is None:
        raise ReportError("IBM JCC did not accept the configured JDBC URL.")
    primary_error = None
    try:
        connection.setAutoCommit(False)
        connection.setReadOnly(True)
        yield connection
    except BaseException as exc:
        primary_error = exc
        raise
    finally:
        try:
            connection.rollback()
        except Exception:
            if primary_error is None:
                raise
        finally:
            connection.close()


def safe_error(error: BaseException) -> str:
    if isinstance(error, ReportError):
        message = str(error)
    elif hasattr(error, "getSQLState") and hasattr(error, "getErrorCode"):
        # JDBC messages can echo SQL and returned values. Send stable diagnostics
        # only; the DBA can look up the SQLCODE/SQLSTATE without leaking data.
        try:
            state = str(error.getSQLState())
            code = int(error.getErrorCode())
            state = state if re.fullmatch(r"[A-Za-z0-9]{5}", state) else "unknown"
            message = f"Db2 operation failed: SQLSTATE={state}, SQLCODE={code}. Ask the DBA to inspect the server diagnostics."
        except Exception:
            message = "Db2 operation failed; ask the DBA to inspect the server diagnostics."
    elif isinstance(error, ModuleNotFoundError):
        message = "A required Python package is missing; install the approved requirements in the agent's virtual environment."
    else:
        message = f"Report processing failed ({type(error).__name__}); check the approved agent runtime and configuration."
    for name in ("DB2_PASSWORD", "DB2_USER", "SMTP_PASSWORD", "SMTP_USER"):
        secret = os.environ.get(name, "")
        if secret:
            message = message.replace(secret, "[redacted]")
    return re.sub(r"[\x00-\x1f\x7f]", " ", message)[:1000]


def prepare_command(args: argparse.Namespace) -> None:
    path, commit = parse_selection(os.environ.get("SCRIPT_SELECTION", ""))
    recipients = parse_recipients(os.environ.get("STAKEHOLDERS", ""))
    state = Path(args.state_dir)
    state.mkdir(parents=True, exist_ok=True)
    atomic_json(state / "request.json", {"version": 1, "script": path, "commit": commit, "recipients": recipients})
    (state / "commit.txt").write_text(commit, encoding="ascii")
    (state / "recipients.txt").write_text(",".join(recipients), encoding="ascii")
    print("Report selection and recipients validated.")


def catalog_command(args: argparse.Namespace) -> None:
    repo = Path(args.repo_dir)
    commit = head_commit(repo)
    paths = tracked_sql(repo)
    if not paths:
        raise ReportError("No safe, committed .sql files were found in the report repository.")
    atomic_json(Path(args.output), {"version": 1, "commit": commit,
                                  "choices": [f"{path}|{commit}" for path in paths]})
    print(f"Published {len(paths)} SQL script choices at revision {commit}.")


def run_command(args: argparse.Namespace) -> None:
    output = Path(args.output_dir)
    output.mkdir(parents=True, exist_ok=True)
    result: dict[str, Any] = {"status": "failure", "script": None, "commit": None,
                              "rows": 0, "worksheets": 0, "report_file": None, "error": None}
    report_path = None
    try:
        request = json.loads((Path(args.state_dir) / "request.json").read_text(encoding="utf-8"))
        path, commit = parse_selection(f"{request['script']}|{request['commit']}")
        parse_recipients(",".join(request["recipients"]))
        result.update(script=path, commit=commit)
        repo = Path(args.repo_dir)
        if head_commit(repo) != commit:
            raise ReportError("The checked-out report revision differs from SCRIPT_SELECTION; refresh the catalog and retry.")
        config = load_config()
        sql = read_sql_at_commit(repo, path, commit, config["runtime"]["max_sql_bytes"])
        queries = split_read_queries(sql)
        report_path = output / (PurePosixPath(path).stem + ".xlsx")
        with db2_connection(config) as connection:
            statistics = export_queries(connection, queries, report_path, config["runtime"])
        result.update(status="success", report_file=report_path.name, error=None, **statistics)
        atomic_json(output / "result.json", result)
        print(f"Report completed: {statistics['rows']} data rows across {statistics['worksheets']} worksheet(s).")
    except BaseException as exc:
        if report_path is not None:
            with contextlib.suppress(FileNotFoundError):
                report_path.unlink()
        result["error"] = safe_error(exc)
        atomic_json(output / "result.json", result)
        raise


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    prepare = commands.add_parser("prepare", help="Validate the UI selection and recipients before checkout")
    prepare.add_argument("--state-dir", default=".run")
    prepare.set_defaults(handler=prepare_command)
    catalog = commands.add_parser("catalog", help="Publish SQL choices from an existing Git checkout")
    catalog.add_argument("--repo-dir", default=".reports")
    catalog.add_argument("--output", default="sql-catalog.json")
    catalog.set_defaults(handler=catalog_command)
    run = commands.add_parser("run", help="Run the selected SQL file and atomically write XLSX/result.json")
    run.add_argument("--repo-dir", default=".reports")
    run.add_argument("--state-dir", default=".run")
    run.add_argument("--output-dir", default="output")
    run.set_defaults(handler=run_command)
    args = parser.parse_args(argv)
    try:
        args.handler(args)
        return 0
    except KeyboardInterrupt:
        print("Report processing interrupted.", file=sys.stderr)
        return 130
    except Exception as exc:
        print(safe_error(exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
