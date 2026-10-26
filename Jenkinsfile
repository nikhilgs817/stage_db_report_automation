// Configure every parameter in the job UI; see README.md. No properties() here.
node(params.AGENT_LABEL) {
    timestamps {
        def failedStage = 'Framework checkout'
        def python = { String command ->
            powershell(script: """
                \$ErrorActionPreference = 'Stop'
                & \$env:PYTHON_EXE ${command}
                exit \$LASTEXITCODE
            """)
        }
        def notify = { String mode ->
            def bindings = []
            if (params.SMTP_CREDENTIALS_ID?.trim()) {
                bindings << usernamePassword(credentialsId: params.SMTP_CREDENTIALS_ID,
                    usernameVariable: 'SMTP_USER', passwordVariable: 'SMTP_PASSWORD')
            }
            withCredentials(bindings) {
                python("'report_mailer.py' ${mode} --state-dir '.run' --output-dir 'output'")
            }
        }
        try {
            timeout(time: params.JOB_TIMEOUT_MINUTES.toInteger(), unit: 'MINUTES') {
                stage('Framework checkout') {
                    deleteDir()
                    checkout scm
                }
                stage('Validate request') {
                    failedStage = 'Validate request'
                    python("'report_runner.py' prepare --state-dir '.run'")
                    notify('validate')
                }
                stage('Fetch selected SQL revision') {
                    failedStage = 'Fetch selected SQL revision'
                    def revision = readFile('.run/commit.txt').trim()
                    dir('.reports') {
                        checkout scmGit(branches: [[name: revision]],
                            userRemoteConfigs: [[url: params.REPORT_REPO_URL,
                                credentialsId: params.GIT_CREDENTIALS_ID]])
                    }
                }
                stage('Generate Excel report') {
                    failedStage = 'Generate Excel report'
                    withCredentials([usernamePassword(credentialsId: params.DB_CREDENTIALS_ID,
                        usernameVariable: 'DB2_USER', passwordVariable: 'DB2_PASSWORD')]) {
                        python("'report_runner.py' run --repo-dir '.reports' --state-dir '.run' --output-dir 'output'")
                    }
                }
                stage('Email report') {
                    failedStage = 'Email report'
                    notify('success')
                }
                stage('Retain build records') {
                    failedStage = 'Retain build records'
                    archiveArtifacts artifacts: 'output/*.json', allowEmptyArchive: false
                    if (params.ARCHIVE_REPORT == true) {
                        archiveArtifacts artifacts: 'output/*.xlsx', allowEmptyArchive: false
                    }
                }
                stage('Clean workspace') {
                    failedStage = 'Clean workspace'
                    deleteDir()
                }
            }
        } catch (Exception failure) {
            currentBuild.result = 'FAILURE'
            withEnv(["FAILURE_STAGE=${failedStage}", "PIPELINE_ERROR=${failure.message ?: 'Pipeline execution failed'}"]) {
                try {
                    if (fileExists('report_mailer.py')) {
                        notify('failure')
                    } else {
                        echo 'Framework unavailable: failure email could not be attempted. See Jenkins monitoring.'
                    }
                } catch (Exception notificationFailure) {
                    echo 'Failure email could not be sent. Build remains failed; see notification status and Jenkins monitoring.'
                }
            }
            try {
                archiveArtifacts artifacts: 'output/*.json', allowEmptyArchive: true
                if (params.ARCHIVE_REPORT == true) {
                    archiveArtifacts artifacts: 'output/*.xlsx', allowEmptyArchive: true
                }
            } catch (Exception archiveFailure) {
                echo 'Failure records could not be archived; see Jenkins monitoring.'
            }
            try {
                deleteDir()
            } catch (Exception cleanupFailure) {
                echo 'Workspace cleanup failed; ask the Jenkins administrator to remove retained report files.'
            }
            throw failure
        }
    }
}
