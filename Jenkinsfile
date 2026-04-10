pipeline {
    agent any

    environment {
        REPORTS_DIR = "C:\\Autoreports\\SanityCheck\\reports"
        WORKSPACE_DIR = "${env.WORKSPACE}"
        PYLINT_THRESHOLD = "5"
        SAMPLE_DATA = "C:\\Autoreports\\SanityCheck\\sample_data"
    }

    stages {

        stage('Checkout') {
            steps {
                git url: 'https://github.com/Ons26bm/Sanity-Check-.git'
            }
        }

        stage('Setup') {
            steps {
                echo "📁 Initialisation des dossiers..."
                bat "if not exist \"%REPORTS_DIR%\" mkdir \"%REPORTS_DIR%\""
                bat "if exist \"%WORKSPACE_DIR%\\reports\" rmdir /s /q \"%WORKSPACE_DIR%\\reports\""
                bat "mkdir \"%WORKSPACE_DIR%\\reports\""
            }
        }

        stage('Install Tools') {
            steps {
                echo "⚙️ Installation outils Python..."
                bat 'python -m pip install --user black pylint bandit pytest pip-audit'
            }
        }

        stage('Python Syntax Check') {
            steps {
                echo "🔍 Vérification syntaxe Python..."
                bat """
                for %%f in (*.py) do (
                    python -m py_compile %%f 2>> "%REPORTS_DIR%\\syntax_errors.txt"
                )
                """
            }
        }

        stage('Code Format Fix (Black)') {
            steps {
                echo "🎨 Formatage code avec Black..."
                bat 'python -m black .'
            }
        }

        stage('Run Pylint') {
            steps {
                catchError(buildResult: 'SUCCESS', stageResult: 'UNSTABLE') {
                    bat """
                    docker run --rm -v "%WORKSPACE_DIR%:/workspace" -w /workspace sanity-python:latest ^
                    pylint *.py --fail-under=%PYLINT_THRESHOLD% --output-format=json --disable=R0801 ^
                    1>"%REPORTS_DIR%\\pylint_report.json" 2>&1
                    """
                }
            }
        }

        stage('Security Scan (Bandit)') {
            steps {
                catchError(buildResult: 'SUCCESS', stageResult: 'UNSTABLE') {
                    bat 'python -m bandit -r . -f json -o "%REPORTS_DIR%\\bandit_report.json"'
                }

                bat '''
                if exist "convert_bandit_to_sonar.py" (
                    echo ✅ Conversion Bandit → Sonar
                    python convert_bandit_to_sonar.py "%REPORTS_DIR%\\bandit_report.json" "%REPORTS_DIR%\\bandit_report_sonar.json"
                ) else (
                    echo ⚠️ Script manquant
                )
                '''
            }
        }

        stage('Dependency Scan (pip-audit)') {
            steps {
                catchError(buildResult: 'SUCCESS', stageResult: 'UNSTABLE') {
                    bat """
                    python -m pip_audit --format json > "%REPORTS_DIR%\\pip_audit_report.json"
                    """
                }

                bat "type \"%REPORTS_DIR%\\pip_audit_report.json\""
            }
        }

        stage('AI Report Summary') {
            steps {
                script {

                    def pylintRaw = readFile("${REPORTS_DIR}\\pylint_report.json")
                    def banditRaw = readFile("${REPORTS_DIR}\\bandit_report.json")
                    def auditRaw = readFile("${REPORTS_DIR}\\pip_audit_report.json")

                    def prompt = """
Tu es un expert en qualité de code Python.

PYLINT:
${pylintRaw.take(1000)}

BANDIT:
${banditRaw.take(1000)}

PIP-AUDIT:
${auditRaw.take(1000)}

Donne un résumé clair en français :
1. Problèmes critiques
2. Problèmes mineurs
3. Temps de correction
4. Solutions
"""

                    writeFile file: "request.json", text: """
{
  "model": "claude-sonnet-4-20250514",
  "max_tokens": 1024,
  "messages": [
    {
      "role": "user",
      "content": "${prompt.replace('"','\\"')}"
    }
  ]
}
"""

                    withCredentials([string(credentialsId: 'anthropic-key', variable: 'ANTHROPIC_API_KEY')]) {

                        def response = bat(
                            returnStdout: true,
                            script: """
curl -s -X POST https://api.anthropic.com/v1/messages ^
-H "x-api-key: %ANTHROPIC_API_KEY%" ^
-H "anthropic-version: 2023-06-01" ^
-H "content-type: application/json" ^
--data @request.json
"""
                        ).trim()

                        echo "Claude response: ${response}"
                        env.AI_SUMMARY = response
                    }
                }
            }
        }

        stage('Run Sanity Check on Sample Data') {
            steps {
                catchError(buildResult: 'SUCCESS', stageResult: 'UNSTABLE') {
                    bat """
                    for %%f in (*.py) do (
                        python %%f "%SAMPLE_DATA%" > "%REPORTS_DIR%\\%%~nf_results.txt" 2>&1
                    )
                    """
                }
            }
        }

        stage('SonarQube Analysis') {
            steps {
                withSonarQubeEnv('SonarQubeServer') {
                    bat """
                    sonar-scanner ^
                        -Dsonar.projectKey=SanityCheck ^
                        -Dsonar.sources=. ^
                        -Dsonar.python.version=3.12 ^
                        -Dsonar.exclusions=reports/* ^
                        -Dsonar.python.pylint.reportPaths=%REPORTS_DIR%\\pylint_report.json
                    """
                }
            }
        }

        stage('Quality Gate') {
            steps {
                catchError(buildResult: 'UNSTABLE', stageResult: 'FAILURE') {
                    script {
                        def response = bat(returnStdout: true, script: """
                        curl -s "http://localhost:9000/api/qualitygates/project_status?projectKey=SanityCheck"
                        """).trim()

                        def json = new groovy.json.JsonSlurper().parseText(response)

                        if (json.projectStatus.status == 'ERROR') {
                            error "❌ Quality Gate failed!"
                        } else {
                            echo "✅ Quality Gate passed"
                        }
                    }
                }
            }
        }

        stage('Generate HTML Report') {
            steps {
                script {
                    echo "📄 Génération rapport HTML..."
                }
            }
        }
    }

    post {
        always {
            emailext (
                subject: "Sanity Check - ${currentBuild.currentResult}",
                body: """Build: ${currentBuild.displayName}
Résultat: ${currentBuild.currentResult}""",
                attachmentsPattern: "reports/sanity_check_report.html",
                to: "pw39f@ningen-group.com"
            )
        }
    }
}
