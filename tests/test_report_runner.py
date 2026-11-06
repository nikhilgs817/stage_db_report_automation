"""V1 contract tests. No live DB2, Jenkins, or SMTP endpoint is contacted."""

from __future__ import annotations

import argparse
import contextlib
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from decimal import Decimal
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import report_runner as runner


SHA = "a" * 40
ROOT = Path(__file__).resolve().parents[1]


class ValidationTests(unittest.TestCase):
    def test_selection_preserves_relative_path_and_revision(self):
        self.assertEqual(
            runner.parse_selection(f"monthly/revenue.sql|{SHA}"),
            ("monthly/revenue.sql", SHA),
        )

    def test_selection_rejects_escape_and_ambiguous_windows_paths(self):
        paths = [
            "../report.sql", "/report.sql", "C:/report.sql",
            "folder/../report.sql", "folder\\report.sql", "report.txt",
            "folder//report.sql", "report.sql\n", "report.sql:payload",
            "CON.sql", "aux.sql", "folder./report.sql",
        ]
        for path in paths:
            with self.subTest(path=repr(path)), self.assertRaises(Exception):
                runner.parse_selection(f"{path}|{SHA}")

    def test_selection_requires_one_complete_revision(self):
        for selection in ["report.sql", "report.sql|main", "report.sql|abc123", f"report.sql|{SHA}|other"]:
            with self.subTest(selection=selection), self.assertRaises(Exception):
                runner.parse_selection(selection)

    def test_recipient_delimiters(self):
        self.assertEqual(
            runner.parse_recipients(" owner@example.com; reviewer@example.com, lead@example.org "),
            ["owner@example.com", "reviewer@example.com", "lead@example.org"],
        )

    def test_recipient_header_injection_and_invalid_addresses(self):
        for value in ["", "owner@example.com\r\nBcc: other@example.com", "not-an-address", "owner@example.com\x00"]:
            with self.subTest(value=repr(value)), self.assertRaises(Exception):
                runner.parse_recipients(value)


class SqlValidationTests(unittest.TestCase):
    def test_comments_strings_ctes_and_multiple_selects(self):
        sql = """-- report comment includes DELETE
        WITH x AS (SELECT ';INSERT' AS label FROM SYSIBM.SYSDUMMY1)
        SELECT label FROM x;
        /* comment includes DROP TABLE */ SELECT 'a;b' FROM SYSIBM.SYSDUMMY1;
        """
        queries = runner.split_read_queries(sql)
        self.assertEqual(len(queries), 2)
        self.assertIn("WITH", queries[0].upper())

    def test_rejects_write_statements_and_db2_data_change_table_references(self):
        statements = [
            "DELETE FROM report_table",
            "SELECT * FROM report_table; DROP TABLE report_table;",
            "WITH x AS (SELECT 1 FROM SYSIBM.SYSDUMMY1) INSERT INTO t SELECT * FROM x",
            "SELECT * FROM FINAL TABLE (INSERT INTO t VALUES (1))",
            "SELECT * FROM OLD TABLE (DELETE FROM t WHERE id=1)",
            "SELECT * FROM report_table FOR UPDATE",
            "CALL schema.procedure()",
            "CONNECT TO STAGE",
            "EXPORT TO report.csv OF DEL SELECT * FROM t",
            "SET CURRENT SCHEMA = APP",
        ]
        for sql in statements:
            with self.subTest(sql=sql), self.assertRaises(Exception):
                runner.split_read_queries(sql)

    def test_rejects_empty_or_comment_only_scripts(self):
        for sql in ["", "  ", "-- explanation only", "/* comment */;"]:
            with self.subTest(sql=sql), self.assertRaises(Exception):
                runner.split_read_queries(sql)

    def test_quoted_write_words_are_data(self):
        self.assertEqual(len(runner.split_read_queries('SELECT \'DELETE\' AS "UPDATE" FROM SYSIBM.SYSDUMMY1')), 1)


class Metadata:
    def __init__(self, labels, types):
        self.labels, self.types = labels, types

    def getColumnCount(self):
        return len(self.labels)

    def getColumnLabel(self, column):
        return self.labels[column - 1]

    def getColumnType(self, column):
        return self.types[column - 1]


class ResultSet:
    def __init__(self, labels, types, rows):
        self.metadata = Metadata(labels, types)
        self.rows = rows
        self.index = -1
        self.closed = False
        self.last_value = None

    def getMetaData(self):
        return self.metadata

    def next(self):
        self.index += 1
        return self.index < len(self.rows)

    def getString(self, column):
        self.last_value = self.rows[self.index][column - 1]
        return None if self.last_value is None else str(self.last_value)

    def getBoolean(self, column):
        self.last_value = self.rows[self.index][column - 1]
        return bool(self.last_value)

    def wasNull(self):
        return self.last_value is None

    def close(self):
        self.closed = True


class Statement:
    def __init__(self, result):
        self.result = result
        self.closed = False
        self.timeout = None
        self.fetch_size = None

    def setQueryTimeout(self, timeout):
        self.timeout = timeout

    def setFetchSize(self, size):
        self.fetch_size = size

    def executeQuery(self, sql):
        self.sql = sql
        return self.result

    def close(self):
        self.closed = True


class Connection:
    def __init__(self, results):
        self.results = iter(results)
        self.statements = []

    def createStatement(self):
        statement = Statement(next(self.results))
        self.statements.append(statement)
        return statement


class WorkbookTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.target = Path(self.tmp.name) / "report.xlsx"
        self.runtime = {
            "query_timeout_seconds": 30, "fetch_size": 20,
            "max_rows_per_sheet": 100, "max_report_bytes": 1_000_000,
        }

    def load(self):
        from openpyxl import load_workbook
        workbook = load_workbook(self.target, read_only=True, data_only=False)
        self.addCleanup(workbook.close)
        return workbook

    def test_each_select_gets_a_sheet_and_resources_close(self):
        results = [ResultSet(["ID"], [4], [(1,), (2,)]), ResultSet(["NAME"], [12], [])]
        connection = Connection(results)
        summary = runner.export_queries(connection, ["SELECT 1", "SELECT 2"], self.target, self.runtime)
        self.assertEqual(summary["rows"], 2)
        self.assertEqual(summary["worksheets"], 2)
        workbook = self.load()
        self.assertEqual(len(workbook.worksheets), 2)
        self.assertEqual(list(workbook.worksheets[0].values), [("ID",), (1,), (2,)])
        self.assertEqual(list(workbook.worksheets[1].values), [("NAME",)])
        self.assertTrue(all(result.closed for result in results))
        self.assertTrue(all(statement.closed for statement in connection.statements))
        self.assertTrue(all(statement.timeout == 30 for statement in connection.statements))
        self.assertTrue(all(statement.fetch_size == 20 for statement in connection.statements))

    def test_formula_like_headers_and_data_are_not_formulas(self):
        values = ["=HYPERLINK(\"https://example.invalid\")", "+SUM(A1)", "-1+2", "@SUM(A1)"]
        result = ResultSet(["=danger"], [12], [(value,) for value in values])
        runner.export_queries(Connection([result]), ["SELECT data"], self.target, self.runtime)
        cells = list(self.load().worksheets[0].iter_rows())
        self.assertEqual(len(cells), 5)
        for cell, original in zip([row[0] for row in cells], ["=danger"] + values):
            with self.subTest(value=original):
                self.assertNotEqual(cell.data_type, "f")
                self.assertTrue(str(cell.value).endswith(original))

    def test_high_precision_numbers_and_null_remain_unambiguous(self):
        result = ResultSet(
            ["DECIMAL", "BIGINT", "NULL_VALUE", "SMALL_NUMBER"],
            [3, -5, 12, 4],
            [("1234567890123456.12", "9223372036854775807", None, "42")],
        )
        runner.export_queries(Connection([result]), ["SELECT numbers"], self.target, self.runtime)
        row = list(self.load().worksheets[0].values)[1]
        self.assertEqual(row, ("1234567890123456.12", "9223372036854775807", None, 42))

    def test_excel_text_and_row_limits_fail_without_completed_workbook(self):
        cases = [
            (ResultSet(["TEXT"], [12], [("x" * 32768,)]), self.runtime),
            (ResultSet(["ID"], [4], [(1,), (2,)]), dict(self.runtime, max_rows_per_sheet=1)),
        ]
        for result, runtime in cases:
            with self.subTest(runtime=runtime):
                connection = Connection([result])
                with self.assertRaises(Exception):
                    runner.export_queries(connection, ["SELECT data"], self.target, runtime)
                self.assertFalse(self.target.exists())
                self.assertTrue(result.closed)
                self.assertTrue(connection.statements[0].closed)

    def test_decimal_conversion_preserves_more_than_excel_precision(self):
        self.assertEqual(runner.excel_value(Decimal("1234567890123456.1234")), "1234567890123456.1234")

    def test_decimal_outside_excel_range_is_exported_as_text(self):
        for value in [Decimal("1E999"), Decimal("1E-999")]:
            with self.subTest(value=value):
                self.assertEqual(runner.excel_value(value), str(value))

    def test_statement_closes_when_result_set_close_fails(self):
        class CloseFailureResult(ResultSet):
            def close(self):
                self.closed = True
                raise RuntimeError("result cleanup failed")

        result = CloseFailureResult(["ID"], [4], [(1,)])
        connection = Connection([result])
        with self.assertRaisesRegex(RuntimeError, "result cleanup failed"):
            runner.export_queries(connection, ["SELECT 1"], self.target, self.runtime)
        self.assertTrue(connection.statements[0].closed)
        self.assertFalse(self.target.exists())

    def test_report_size_limit_prevents_completed_workbook(self):
        result = ResultSet(["ID"], [4], [(1,)])
        with self.assertRaises(runner.ReportError):
            runner.export_queries(Connection([result]), ["SELECT 1"], self.target, dict(self.runtime, max_report_bytes=1))
        self.assertFalse(self.target.exists())


class GitCheckoutTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.repo = Path(self.tmp.name) / "repo"
        self.repo.mkdir()
        self.git("init")
        self.git("config", "user.email", "test@example.invalid")
        self.git("config", "user.name", "V1 tests")
        (self.repo / "monthly").mkdir()
        (self.repo / "monthly" / "report.sql").write_text("SELECT 1 FROM SYSIBM.SYSDUMMY1;", encoding="utf-8")
        (self.repo / "README.md").write_text("not a report", encoding="utf-8")
        self.git("add", ".")
        self.git("commit", "-m", "test reports")
        self.sha = self.git("rev-parse", "HEAD").strip()

    def git(self, *args):
        return subprocess.run(["git", "-C", str(self.repo), *args], check=True, capture_output=True, text=True).stdout

    def test_tracked_sql_lists_only_committed_sql_files(self):
        (self.repo / "untracked.sql").write_text("SELECT 2", encoding="utf-8")
        self.assertEqual(runner.tracked_sql(self.repo), ["monthly/report.sql"])
        self.assertEqual(runner.head_commit(self.repo), self.sha)

    def test_catalog_command_is_removed(self):
        result = subprocess.run(
            [sys.executable, str(ROOT / "report_runner.py"), "catalog", "--repo-dir", str(self.repo)],
            capture_output=True, text=True,
        )
        self.assertNotEqual(result.returncode, 0)

    def test_tracked_sql_excludes_git_symlink_mode_even_for_regular_worktree_file(self):
        # Windows Git can materialize a symlink as an ordinary text file. Its
        # committed tree mode must still exclude it from selectable reports.
        link = self.repo / "shortcut.sql"
        link.write_text("monthly/report.sql", encoding="utf-8")
        blob = self.git("hash-object", "-w", "shortcut.sql").strip()
        self.git("update-index", "--add", "--cacheinfo", f"120000,{blob},shortcut.sql")
        self.git("commit", "-m", "add tree-mode symlink")
        self.assertEqual(runner.tracked_sql(self.repo), ["monthly/report.sql"])

    def prepare_request(self, commit=None):
        state = Path(self.tmp.name) / "state"
        with patch.dict(os.environ, {
            "SCRIPT_SELECTION": f"monthly/report.sql|{commit or self.sha}",
            "STAKEHOLDERS": "owner@example.com;reviewer@example.com",
        }):
            runner.prepare_command(argparse.Namespace(state_dir=state))
        self.assertEqual((state / "commit.txt").read_text(encoding="ascii"), commit or self.sha)
        return state

    def run_args(self, state):
        return argparse.Namespace(repo_dir=self.repo, state_dir=state, output_dir=Path(self.tmp.name) / "output")

    def test_pinned_run_writes_named_workbook_and_success_result(self):
        state = self.prepare_request()
        connection = Connection([ResultSet(["ID"], [4], [(1,)])])
        args = self.run_args(state)
        config = {"database": {}, "runtime": dict(runner.RUNTIME_DEFAULTS)}
        with patch.object(runner, "load_config", return_value=config), patch.object(
            runner, "db2_connection", return_value=contextlib.nullcontext(connection)
        ):
            runner.run_command(args)
        result = json.loads((args.output_dir / "result.json").read_text(encoding="utf-8"))
        self.assertEqual(result["status"], "success")
        self.assertEqual(result["report_file"], "report.xlsx")
        self.assertEqual(result["commit"], self.sha)
        self.assertEqual(result["rows"], 1)
        self.assertTrue((args.output_dir / "report.xlsx").is_file())

    def test_wrong_revision_fails_before_connection_and_records_failure(self):
        state = self.prepare_request(SHA)
        args = self.run_args(state)
        config = {"database": {}, "runtime": dict(runner.RUNTIME_DEFAULTS)}
        with patch.object(runner, "load_config", return_value=config), patch.object(runner, "db2_connection") as connect:
            with self.assertRaises(runner.ReportError):
                runner.run_command(args)
        connect.assert_not_called()
        result = json.loads((args.output_dir / "result.json").read_text(encoding="utf-8"))
        self.assertEqual(result["status"], "failure")
        self.assertIn("revision", result["error"].lower())
        self.assertFalse(any(args.output_dir.glob("*.xlsx")))

    def test_uncommitted_sql_is_rejected_before_connection(self):
        state = self.prepare_request()
        (self.repo / "monthly" / "report.sql").write_text("SELECT 999 FROM SYSIBM.SYSDUMMY1", encoding="utf-8")
        args = self.run_args(state)
        config = {"database": {}, "runtime": dict(runner.RUNTIME_DEFAULTS)}
        with patch.object(runner, "load_config", return_value=config), patch.object(runner, "db2_connection") as connect:
            with self.assertRaises(runner.ReportError):
                runner.run_command(args)
        connect.assert_not_called()
        result = json.loads((args.output_dir / "result.json").read_text(encoding="utf-8"))
        self.assertIn("uncommitted", result["error"].lower())
        self.assertEqual(result["status"], "failure")


class ErrorHandlingTests(unittest.TestCase):
    def test_jdbc_failure_reports_codes_without_sql_rows_or_credentials(self):
        class JdbcFailure(Exception):
            def getSQLState(self):
                return "42501"

            def getErrorCode(self):
                return -551

        with patch.dict(os.environ, {"DB2_USER": "service-account", "DB2_PASSWORD": "secret-password"}):
            message = runner.safe_error(JdbcFailure("SELECT private_column; service-account secret-password sensitive-row"))
        self.assertIn("42501", message)
        self.assertIn("-551", message)
        for value in ["SELECT", "private_column", "service-account", "secret-password", "sensitive-row"]:
            self.assertNotIn(value, message)

    def test_controlled_errors_are_redacted_and_single_line(self):
        with patch.dict(os.environ, {"DB2_PASSWORD": "secret-password"}):
            message = runner.safe_error(runner.ReportError("connection secret-password\r\nfailed"))
        self.assertNotIn("secret-password", message)
        self.assertNotIn("\r", message)
        self.assertNotIn("\n", message)


class JdbcConnectionTests(unittest.TestCase):
    def test_read_only_transaction_rolls_back_and_closes_on_query_failure(self):
        from types import SimpleNamespace

        class JdbcConnection:
            autocommit = None
            read_only = None
            rolled_back = False
            closed = False

            def setAutoCommit(self, value):
                self.autocommit = value

            def setReadOnly(self, value):
                self.read_only = value

            def rollback(self):
                self.rolled_back = True

            def close(self):
                self.closed = True

        class Properties(dict):
            def setProperty(self, key, value):
                self[key] = value

        connection = JdbcConnection()
        captured = {}

        class Driver:
            def getMajorVersion(self):
                return 4

            def getMinorVersion(self):
                return 29

            def connect(self, url, properties):
                captured.update(url=url, properties=dict(properties))
                return connection

        manager = SimpleNamespace(setLoginTimeout=lambda value: captured.update(login_timeout=value))
        classes = {"com.ibm.db2.jcc.DB2Driver": Driver, "java.sql.DriverManager": manager, "java.util.Properties": Properties}
        jpype = SimpleNamespace(isJVMStarted=lambda: True, JClass=lambda name: classes[name])
        config = {
            "database": {"host": "stage.example.invalid", "port": 50001, "name": "STAGELOC", "ssl_cert_location": "C:/approved/server.pem"},
            "runtime": dict(runner.RUNTIME_DEFAULTS),
        }
        with patch.dict(sys.modules, {"jpype": jpype}), patch.dict(os.environ, {"DB2_USER": "svc-report", "DB2_PASSWORD": "secret"}):
            with self.assertRaisesRegex(RuntimeError, "query failed"):
                with runner.db2_connection(config) as actual:
                    self.assertIs(actual, connection)
                    raise RuntimeError("query failed")
        self.assertIs(connection.autocommit, False)
        self.assertIs(connection.read_only, True)
        self.assertTrue(connection.rolled_back)
        self.assertTrue(connection.closed)
        self.assertEqual(captured["properties"]["sslConnection"], "true")
        self.assertEqual(captured["properties"]["sslClientHostnameValidation"], "BASIC")
        self.assertEqual(captured["properties"]["user"], "svc-report")
        self.assertEqual(captured["properties"]["password"], "secret")

    def test_older_jcc_is_rejected_before_connecting(self):
        from types import SimpleNamespace

        for major, minor in ((3, 72), (4, 0), (4, 28)):
            with self.subTest(version=(major, minor)):
                captured = {"connected": False}

                class Driver:
                    def getMajorVersion(self):
                        return major

                    def getMinorVersion(self):
                        return minor

                    def connect(self, url, properties):
                        captured["connected"] = True
                        raise AssertionError("An unsupported driver must never connect")

                jpype = SimpleNamespace(isJVMStarted=lambda: True, JClass=lambda name: Driver)
                config = {"database": {}, "runtime": dict(runner.RUNTIME_DEFAULTS)}
                with patch.dict(sys.modules, {"jpype": jpype}), patch.dict(
                    os.environ, {"DB2_USER": "svc-report", "DB2_PASSWORD": "secret"}
                ):
                    with self.assertRaisesRegex(runner.ReportError, "4.29 or later.*hostname validation"):
                        with runner.db2_connection(config):
                            self.fail("Unsupported JCC unexpectedly opened a connection")
                self.assertFalse(captured["connected"])


if __name__ == "__main__":
    unittest.main()
