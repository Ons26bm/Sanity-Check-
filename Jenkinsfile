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
                echo "🔐 Scan sécurité avec Bandit..."
                catchError(buildResult: 'SUCCESS', stageResult: 'UNSTABLE') {
                    bat 'python -m bandit -r . -f json -o "%REPORTS_DIR%\\bandit_report.json"'
                }
            }
        }

        stage('Dependency Scan (pip-audit)') {
            steps {
                echo "📦 Analyse dépendances avec pip-audit..."
                bat 'python -m pip_audit --json > "%REPORTS_DIR%\\pip_audit_report.json"'
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
                echo "📈 Analyse avec SonarQube..."
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
                echo "🚦 Vérification Quality Gate SonarQube..."
                withCredentials([usernamePassword(credentialsId: 'sonar-creds', 
                                                  usernameVariable: 'SONAR_USER', 
                                                  passwordVariable: 'SONAR_PASS')]) {
                    script {
                        def response = bat(returnStdout: true, script: """
                            curl -s -u %SONAR_USER%:%SONAR_PASS% "http://localhost:9000/api/qualitygates/project_status?projectKey=SanityCheck"
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
                echo "📄 Génération rapport HTML consolidé..."
                script {
                    def html = """
                        <html><body>
                        <h2>Sanity Check Résumé</h2>
                        <ul>
                            <li>✅ Syntaxe : ${fileExists("${REPORTS_DIR}/syntax_errors.txt") ? "❌ erreurs" : "✔️ OK"}</li>
                            <li>📊 Pylint : <a href="pylint_report.json">Voir JSON</a></li>
                            <li>🔐 Bandit : <a href="bandit_report.json">Voir JSON</a></li>
                            <li>📦 pip-audit : <a href="pip_audit_report.json">Voir JSON</a></li>
                            <li>🧪 Exécution scripts : fichiers *_results.txt</li>
                        </ul>
                        </body></html>
                    """
                    writeFile file: "${REPORTS_DIR}/sanity_check_report.html", text: html
                }
            }
        }

        stage('Email Report') {
            steps {
                echo "📧 Envoi email au data analyst..."
                emailext (
                    subject: "Sanity Check - Résumé Scripts Python",
                    body: "Le rapport HTML est disponible ici: ${REPORTS_DIR}\\sanity_check_report.html",
                    to: "pw39f@ningen-group.com"
                )
            }
        }
    }

    post {
        always {
            echo "🔔 Pipeline terminé avec status = ${currentBuild.currentResult}"
        }
    }
}
