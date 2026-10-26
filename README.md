# STAGE DB REPORT — V1

Run a reviewed SQL file from Git against stage **Db2 for z/OS**, create a real `.xlsx` workbook with the SQL file's base name, and email it to the stakeholders supplied on **Build with Parameters**. The agent is Windows. Python handles selection validation, query execution, Excel generation, mail, and result records. Jenkins handles checkout, credentials, orchestration, and build status.

No parameter declarations, credentials, repository addresses, database addresses, recipients, or schedules are embedded in either Jenkinsfile. Configure parameter definitions and defaults manually in the Jenkins job UI. The database profile is an administrator-managed file on the agent, selected by a UI parameter.

## V1 flow

1. A small catalog job checks out the SQL repository and publishes the committed `.sql` file choices.
2. Open the report job and select **Build with Parameters**.
3. Choose `SCRIPT_SELECTION`, for example `monthly/customer_report.sql|<full Git revision>`, and enter `STAKEHOLDERS`.
4. The job validates the request and mail settings, fetches that exact Git revision, and verifies the selected SQL file against Git.
5. Python connects over TLS using the Jenkins-bound DB2 service credential and runs the SELECT queries.
6. `customer_report.xlsx` is generated. Each SELECT result gets a worksheet; an empty result has its column headers.
7. Python sends one email with that workbook. Database, validation, checkout, Excel, or mail errors fail the build and trigger a failure email without an attachment.

SMTP success means the relay accepted the message for every recipient. It cannot prove final mailbox delivery. No automatic report or email retry is enabled, to avoid duplicate reporting after an uncertain send.

## What to install on the Windows agent

**Neither the Db2 database server nor IBM Data Studio is required.** Use:

- An approved, supported **64-bit Python 3.10 or later**; this framework was locally tested with Python 3.12.
- An approved **64-bit Java runtime** compatible with your IBM JCC driver. Python and the JVM must have matching architecture. An existing suitable agent Java installation can be used.
- IBM **Data Server Driver for JDBC and SQLJ**, specifically `db2jcc4.jar`. Use an organization-approved driver version **4.29 or later**, compatible with your z/OS server. This framework explicitly requests TLS hostname validation, introduced in JCC 4.29, and rejects older drivers before connecting.
- The appropriate **Db2 Connect entitlement/license** for z/OS. Your DBA/licensing team must determine whether the installation uses server activation, a licensed gateway, or a client license JAR. Add a license JAR to the classpath only when your setup requires it; the sample license filename is a placeholder.
- The approved database TLS certificate file, and the packages in `requirements.lock.txt`: JPype1, openpyxl, sqlparse, and their pinned dependencies.
- Git for Windows and Windows PowerShell. Microsoft Excel, Office automation, VS Code, and a VS Code database plugin are not required.
- Agent network access to Git, the stage z/OS DDF TLS endpoint, and the approved SMTP relay. The controller needs its normal Jenkins/SCM connectivity; database traffic originates from the agent.

IBM Data Studio's documented **Db2 Client Command Line Processor** integration invokes a separately installed Db2 Client. IBM also supplies CLPPlus in the Data Server Driver Package. Those are viable separately configured alternatives; the implemented V1 executor uses JDBC directly to obtain structured rows for Excel. [IBM Data Studio](https://www.ibm.com/support/pages/whats-new-and-changed-ibm-data-studio-41x), [IBM driver packages](https://www.ibm.com/support/pages/getting-started-ibm-data-server-drivers), [JDBC installation](https://www.ibm.com/docs/en/db2/12.1.x?topic=java-installing-updating-upgrading-data-server-driver-jdbc-sqlj-db2-linux-unix-windows-systems), [JCC TLS properties](https://www.ibm.com/docs/en/db2-for-zos/13.0.0?topic=information-data-server-driver-jdbc-sqlj-configuration-properties).

### Install packages once, outside Jenkins workspaces

Use the internal package repository or an approved offline wheel bundle. Do not install packages during report builds. **The virtual environment, JARs, certificates, and database profile must be outside the Jenkins workspace** because each build cleans its workspace.

Example installation paths below are illustrative; replace them with approved paths. On a connected staging machine with the same Windows architecture/Python version, prepare wheels through your approved software process:

```powershell
python -m pip download --only-binary=:all: --dest C:\Transfer\report-wheels -r .\requirements.lock.txt
```

On the restricted agent, after transferring the reviewed files:

```powershell
python -m venv C:\ApprovedTools\StageDbReport\venv
& C:\ApprovedTools\StageDbReport\venv\Scripts\python.exe -m pip install --no-index --find-links C:\Transfer\report-wheels -r .\requirements.lock.txt
```

Restart the Jenkins agent service if Java, PATH, or JAVA_HOME changed. Verify paths and connectivity as the **Windows account running the agent service**, which may differ from your interactive login.

## DBA and Jenkins administrator setup

### 1. Provide the DB2 service identity

The Windows agent service account does not automatically become a DB2 identity. V1 binds a dedicated **Jenkins username/password credential** to `DB2_USER` and `DB2_PASSWORD`; that credential must identify the approved DB2 service account. The framework does not use a developer's Data Studio saved password.

Ask the DBA for the stage DDF hostname, TLS port, **Db2 LOCATION name** for the JDBC database component, certificate, approved JCC driver/license arrangement, and any required `currentSchema`, `currentSQLID`, or package settings. Do not assume that the workstation's local catalog alias is the LOCATION name.

The DBA must grant only the required read access to reporting tables/views and any needed connection/package privileges. Do not grant table writes, broad administration, or execute access to side-effecting routines. Review inherited/PUBLIC privileges as well. JDBC `setReadOnly(true)`, rollback, and SQL filtering are additional safeguards; database authorization is the enforcement boundary.

Copy `config.example.json` to an approved protected location such as `C:\ApprovedTools\StageDbReport\config.local.json`. Replace every example database value and JAR/certificate path. Remove the placeholder license JAR entry if your entitlement does not require a client license JAR. Configure `JAVA_HOME`, or set `runtime.jvm_path` to the approved `jvm.dll`. Keep passwords out of this file and out of Git. Restrict write access to administrators and read access to the agent account.

The connection is `jdbc:db2://HOST:TLS_PORT/LOCATION`. TLS certificate and hostname verification are enabled. `database.properties` accepts a small allowlist of package/schema/client properties, not arbitrary credential or TLS overrides.

### 2. Configure Jenkins

Use a standard **Pipeline** job with **Pipeline script from SCM**, not a multibranch job that regenerates parameter definitions. This repository is the framework repository, with `Jenkinsfile`, `report_runner.py`, and `report_mailer.py` at its root. The SQL repository can be a separate Git repository and is never executed as Python or Groovy.

Required plugins, at approved versions compatible with your Jenkins controller:

- Pipeline, including its Windows PowerShell step.
- Git and Git Client.
- Credentials and Credentials Binding.
- Timestamper (`timestamps`).
- **Active Choices** (`uno-choice`) for the dynamic script dropdown.

Email Extension and Pipeline Utility Steps are not required. Python sends mail from the agent. Store Git read credentials and the DB2 username/password credential in a folder accessible to these jobs. For HTTPS Git, use the appropriate Jenkins username/password credential with an approved token; for SSH, use an SSH private key credential. Never put credentials in the repository URL.

Restrict job Configure permissions, framework repository changes, database profiles, and credential administration. Build access must be limited to trusted report operators: users can supply recipient addresses, configuration paths, repository settings, and credential IDs. Restrict report-repository changes to reviewed SQL. Configure build/log/artifact retention in the UI according to your data policy. Concurrent builds get isolated Jenkins workspaces; set executor/concurrency limits appropriate for stage DB load.

### 3. Create the SQL catalog job

Example name: `Reporting/STAGE-DB-REPORT-CATALOG`. In Pipeline script from SCM, point **Script Path** to `Jenkinsfile.catalog`.

Enable **This project is parameterized** and manually create these **String parameters** with your environment's defaults:

- `AGENT_LABEL`: label for the prepared Windows agent.
- `PYTHON_EXE`: absolute virtual-environment `python.exe` path, outside the workspace. Supply the executable path only, with no surrounding quotation marks or extra arguments.
- `JOB_TIMEOUT_MINUTES`: whole-job execution timeout, for example `10` for catalog refresh.
- `REPORT_REPO_URL`: approved SQL Git repository URL, without embedded credentials.
- `REPORT_GIT_REF`: full approved branch ref, for example `refs/heads/main`.
- `GIT_CREDENTIALS_ID`: Jenkins credential ID with read access to that repository.

Run the catalog job once. Confirm its archived `sql-catalog.json` lists the expected files and revision. Refresh this job after approved SQL changes. Optional periodic or webhook refresh can be configured by the Jenkins administrator in the job UI; no trigger is embedded in the Jenkinsfile. Monitor catalog failures separately. The report form intentionally uses the last successful published catalog, so a later refresh failure leaves the previous pinned choices available.

### 4. Create the report job and parameters

Example name: `Reporting/STAGE-DB-REPORT`. In Pipeline script from SCM, point **Script Path** to `Jenkinsfile`.

Manually create these parameters in the job UI, with the operator inputs first:

- **String** `CATALOG_JOB`: full catalog job name, such as `Reporting/STAGE-DB-REPORT-CATALOG`.
- **Active Choices Reactive Parameter** `SCRIPT_SELECTION`: Single Select; enable filtering; **Referenced parameters** must be `CATALOG_JOB`. Paste the exact contents of `jenkins/active_choices.groovy` in its Groovy script field. The fallback script can be `return ['Catalog unavailable:disabled']`.
- **String** `STAKEHOLDERS`: report recipients as comma- or semicolon-separated full email addresses. Distribution lists are supported. Do not use display names.
- **String** `AGENT_LABEL`, `PYTHON_EXE`, `JOB_TIMEOUT_MINUTES`, `REPORT_REPO_URL`, `GIT_CREDENTIALS_ID`: same meanings as above. Use a report timeout suitable for the largest approved query, for example `30` minutes. Keep the repository URL aligned with the catalog job.
- **String** `REPORT_CONFIG`: absolute path to the protected agent database profile JSON file.
- **String** `DB_CREDENTIALS_ID`: ID of the dedicated DB2 service username/password credential. This is an ID, never the password.
- **String** `SMTP_HOST`, `SMTP_PORT`: approved relay hostname and port.
- **Choice** `SMTP_SECURITY`: `starttls`, `ssl`, or `plain`; use the mode required by your relay. Authenticated SMTP requires TLS. Use `plain` only for an explicitly approved internal unauthenticated relay.
- **String** `MAIL_FROM`: one approved sender email address.
- **String** `SMTP_CREDENTIALS_ID`: optional Jenkins username/password credential ID; leave blank for a relay that does not require authentication.
- **String** `SMTP_TIMEOUT_SECONDS`: connection/socket timeout, for example `30` (allowed 1–300).
- **String** `SMTP_CA_FILE`: optional PEM CA bundle for a private mail PKI. Leave blank to use Python's default trust roots. Certificate validation remains enabled for TLS modes.
- **String** `MAX_EMAIL_BYTES`: maximum complete MIME email size, for example `30000000`. Set below the relay's accepted message limit. Base64 encoding increases attachment size by roughly one third, so configure the database profile's `max_report_bytes` accordingly.
- **String** `FAILURE_RECIPIENTS`: required operator/support distribution list. Failure messages include this list in addition to valid stakeholder recipients. Invalid stakeholder input falls back to the operator list.
- **Boolean** `ARCHIVE_REPORT`: default `false`. When enabled, Jenkins also retains the workbook as a build artifact. Apply appropriate artifact access and retention controls first.

The Active Choices script reads a published Jenkins metadata artifact using Jenkins APIs. It checks the viewer's read permission on the catalog job and the optional per-build artifact permission when enabled. Your Jenkins administrator must review the script and approve the required methods in Script Security, or approve the exact script under your organization's policy. An unapproved script may show the disabled fallback. The framework does not disable global script security.

Active Choices relies on the interactive build form. The standard Git Parameter plugin lists branches/tags/revisions rather than arbitrary repository files, which is why this framework uses a separate file catalog. For future scheduled/API runs, pass an explicit validated `SCRIPT_SELECTION` and recipients; do not depend on browser dropdown evaluation. [Active Choices](https://plugins.jenkins.io/uno-choice/), [Git Parameter](https://plugins.jenkins.io/git-parameter/), [Git checkout](https://www.jenkins.io/doc/pipeline/steps/git/).

## Supported report SQL and output

- One or more UTF-8 `SELECT` or `WITH ... SELECT` statements, separated by semicolons. UTF-8 BOM is accepted.
- No Data Studio/CLP commands (`CONNECT`, `EXPORT`, terminator changes), SQL substitution prompts, DDL/DML, procedures, temporary-table preparation, or session `SET` commands in V1. Configure approved schema/package properties in the database profile instead. Client terminator directives must be removed; use normal semicolon separators.
- SELECT-only filtering is deliberately conservative and is not a complete SQL security parser. Read-only DBA grants and reviewed queries are mandatory. Reports using unsupported syntax need review before extending the policy.
- Each query creates `Result_1`, `Result_2`, etc. The workbook filename is the selected SQL basename with `.xlsx`, preserving its name. A script `monthly/customer_report.sql` produces `customer_report.xlsx`.
- Strings, including formula-like headers and values, are stored as literal Excel text. Integers/decimals beyond Excel's 15-digit precision are stored as text. Dates/timestamps are preserved in JDBC text form, including z/OS fractional seconds. NULL produces an empty cell.
- Rows stream to a write-only workbook. Excel row/column/text limits and configured workbook/email size limits fail the job without emailing a partial workbook. LOB/binary columns should be explicitly projected or encoded in SQL; V1 does not promise portable binary/LOB export.
- JDBC login/network-read/query timeouts are configured independently of the Jenkins overall timeout. Set them within the Jenkins timeout. Cancellation is best effort; the DBA should also apply appropriate server-side workload limits.

Metadata is archived as `output/result.json`, `output/notification.success.json`, and `output/notification.failure.json` when present. SQLSTATE/SQLCODE are retained for DB errors; raw SQL, row values, and credentials are excluded. `result.json` reports generation status; notification records report SMTP status. A generated workbook with a failed mail send still means the Jenkins job failed. If mail delivery fails after some recipients accepted it, no automatic resend is attempted.

Workspaces are cleaned after each build. With `ARCHIVE_REPORT=false`, the delivered attachment is the retained report copy; there is no Jenkins workbook artifact. Framework SCM checkout failure, an unavailable agent/controller, or an unavailable mail relay can prevent failure email. Use your existing Jenkins monitoring for those infrastructure failures. An archive/cleanup error can occur after email acceptance; check notification metadata before manually rerunning.

## First stage-environment acceptance run

1. Have the DBA validate the service identity's effective read-only grants and the JCC driver/license/certificate combination.
2. Put `examples/sample_report.sql` in the approved SQL repository, commit it, and refresh the catalog.
3. Build the report job with only a test recipient list. Confirm the chosen revision, service identity, expected stage LOCATION, workbook name, headers/data, and SMTP acceptance.
4. Run a known empty-result query and a reviewed multi-query script; confirm headers and separate worksheets.
5. Use an approved test setup to verify SQL authorization failure, query timeout, Git checkout failure, and SMTP rejection. Confirm failure status and no report attachment in failure notifications.
6. Validate representative real report types, precision, volume, email limits, agent restart behavior, and artifact retention before enabling stakeholder distribution.

Local tests do not replace these checks: no live stage DB, corporate Git, Jenkins controller, certificate/driver installation, or mail relay was supplied in this workspace.

## V2 scheduling from SharePoint

V1 is manually triggered. Daily, weekly, and monthly reports all use the same executor; no frequency logic is embedded in the SQL runner.

Reading due dates from the SharePoint spreadsheet is feasible with an approved Microsoft Graph/SharePoint integration and service identity. A future dispatcher can poll at an agreed interval, read due rows, and invoke this job with explicit selection and recipients. This is a separate approval/integration step, not part of V1.

Plan fields such as `ReportId`, SQL path, approved Git branch/revision policy, recipients, due date/time, timezone, frequency, and enabled status. Store a durable execution key such as `(ReportId, scheduled due instant)` and generation/delivery state so polling or retries do not send duplicates. Define timezone/DST behavior, overdue catch-up, holidays, recurrence, and handling of spreadsheet edits before implementation. A SharePoint List can be considered if concurrent editing and update tracking become important. If Graph access is unavailable, UI-configured Jenkins schedules with explicit per-report inputs are a fallback.

## Developer verification

From this directory, after installing the locked dependencies:

```powershell
& .\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

Tests use temporary local Git repositories, fake JDBC objects, real workbook readback, and mocked SMTP. They do not connect to DB2 or send email. See `VALIDATION.md` for the actual local results and deployment checks still outstanding.
