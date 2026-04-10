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
                echo "📊 Analyse qualité du code avec Pylint..."
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
                echo "🔐 Scan de sécurité avec Bandit..."
                catchError(buildResult: 'SUCCESS', stageResult: 'UNSTABLE') {
                    bat 'python -m bandit -r . -f json -o "%REPORTS_DIR%\\bandit_report.json"'
                }
            }
        }

        stage('Dependency Scan (pip-audit)') {
            steps {
                echo "🔍 Analyse dépendances..."
                catchError(buildResult: 'SUCCESS', stageResult: 'UNSTABLE') {
                    bat """
                    python -m pip_audit --format json > "%REPORTS_DIR%\\pip_audit_report.json"
                    """
                }
            }
        }

        stage('AI Report Summary') {
            steps {
                script {

                    def pylintRaw  = readFile("${REPORTS_DIR}\\pylint_report.json")
                    def banditRaw  = readFile("${REPORTS_DIR}\\bandit_report.json")
                    def auditRaw   = readFile("${REPORTS_DIR}\\pip_audit_report.json")

                    def prompt = """
Tu es un expert en qualité de code Python.

PYLINT:
${pylintRaw.take(1000)}

BANDIT:
${banditRaw.take(1000)}

PIP-AUDIT:
${auditRaw.take(1000)}

Donne un résumé clair en français avec :
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

                        echo "Claude raw response: ${response}"
                        env.AI_SUMMARY = response
                    }
                }
            }
        }

        stage('Run Sanity Check on Sample Data') {
            steps {
                echo "🧪 Exécution scripts sur données échantillons..."
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
                echo "📈 Analyse SonarQube..."
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
                    withCredentials([usernamePassword(credentialsId: 'sonar-creds',
                                                      usernameVariable: 'SONAR_USER',
                                                      passwordVariable: 'SONAR_PASS')]) {
                        script {
                            def response = bat(returnStdout: true, script: """
                                @echo off
                                curl -s -u %SONAR_USER%:%SONAR_PASS% "http://localhost:9000/api/qualitygates/project_status?projectKey=SanityCheck"
                            """).trim()

                            def jsonStart = response.indexOf('{')
                            def jsonText = response.substring(jsonStart)
                            def json = new groovy.json.JsonSlurper().parseText(jsonText)

                            if (json.projectStatus.status == 'ERROR') {
                                error "❌ Quality Gate failed!"
                            } else {
                                echo "✅ Quality Gate passed"
                            }
                        }
                    }
                }
            }
        }

        stage('Generate HTML Report') {
            steps {
                catchError(buildResult: 'SUCCESS', stageResult: 'UNSTABLE') {
                    script {

                        echo "📄 Génération rapport HTML..."

                        def syntaxContent = fileExists("${REPORTS_DIR}\\syntax_errors.txt") ?
                                readFile("${REPORTS_DIR}\\syntax_errors.txt").trim() : ""

                        def syntaxStatus = syntaxContent ? "Erreurs detectees" : "OK"
                        def syntaxClass = syntaxContent ? "fail" : "ok"

                        def html = """
                        <html><body>
                        <h2>Sanity Check Report</h2>

                        <p>Syntaxe: ${syntaxStatus}</p>
                        <p>Build: ${currentBuild.currentResult}</p>

                        </body></html>
                        """

                        writeFile file: "${REPORTS_DIR}\\sanity_check_report.html", text: html
                        bat "copy \"${REPORTS_DIR}\\sanity_check_report.html\" \"${WORKSPACE_DIR}\\reports\\sanity_check_report.html\""

                        echo "Rapport HTML généré"
                    }
                }
            }
        }
    }

    post {
        always {
            emailext (
                subject: "Sanity Check - ${currentBuild.currentResult}",
                body: "Build terminé: ${currentBuild.displayName}",
                attachmentsPattern: "reports/sanity_check_report.html",
                to: "pw39f@ningen-group.com"
            )
        }
    }
}
