# V1 validation — 26 October 2026

## Completed locally

- **39 tests passed** on Windows with Python 3.12.10.
- Pinned dependencies installed in an isolated virtual environment; `pip check` reported no broken requirements.
- Python source and tests compiled successfully.
- Real temporary Git repositories verified committed-file listing, exact revision selection, dirty-file rejection, and symlink tree-mode exclusion.
- Fake JDBC connections verified read-only transaction setup, rollback/cleanup, result streaming, and SQLSTATE/SQLCODE-only error diagnostics.
- The runtime rejects JCC versions older than 4.29 before connecting, preventing silent loss of TLS hostname validation on an older driver.
- Generated real XLSX workbooks were reopened to verify multiple result sheets, empty results, literal formula-like headers/values, NULLs, large integer/decimal preservation, and limit failures.
- Mocked SMTP verified TLS setup, exact attachment filename, recipient validation, partial recipient refusal, notification status, MIME size limits, and failure notifications without attachments.
- Integration review checked UI-only parameter declarations, credential binding boundaries, pinned Git checkout, timeout coverage, and Active Choices artifact permissions.

The local test command from the repository root is:

```powershell
& .\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

## Required in your environment

No live DB2 query or real email has been attempted. Jenkins pipeline execution and Active Choices script approvals have not been tested against your controller. The framework needs your approved JCC JAR/license arrangement, Java runtime, stage DDF endpoint/certificate, Jenkins credentials, Git access, and SMTP relay settings.

Use the first-run acceptance steps in README.md. Confirm package/driver compatibility, effective service-account read permissions, TLS hostname validation, representative report types and sizes, failure notification, and retention before sending actual reports.

Build status and report-generation status differ deliberately: an XLSX may be generated successfully while SMTP delivery fails, which fails the Jenkins build. Relay acceptance does not prove mailbox delivery. Jenkins monitoring must cover bootstrap checkout failure, unavailable agents/controllers, cleanup failures that remove the mailer, and a failed mail relay.

## Jenkinsfile consolidation — 6 November 2026

Catalog refresh and report generation now share one `Jenkinsfile`, selected by the UI-only `RUN_MODE` parameter. The dedicated catalog and report jobs retain separate parameter sets. The two production Python files are unchanged. Review confirms that catalog mode does not call database or email operations. Jenkins controller execution remains an environment acceptance check; no live Jenkins execution is claimed for this change.

## Catalog job removed — 6 November 2026

The separate catalog job, the `report_runner.py catalog` command and the `RUN_MODE` switch are removed; this supersedes the consolidation above. The `SCRIPT_SELECTION` Active Choices script now reads the committed `.sql` files and the approved branch's current commit from the GitLab API, so the dropdown cannot go stale. Choices keep the `path|commit` format, so revision pinning, Git verification and the report run are unchanged.

- **40 tests passed** locally (Python 3.11, Linux). New tests check that `tracked_sql` lists only committed `.sql` files and that the `catalog` command no longer exists.
- The dropdown script was run with Groovy 2.4.21 (the version Jenkins bundles) against a mock GitLab API, using stub Jenkins credential classes. It returned sorted, paginated `path|commit` choices, dropped a path containing `|`, and showed disabled messages for a missing credential, an HTTP 401 and a refused connection.
- Not tested: a real Jenkins controller, Script Approval, or a real GitLab instance.
