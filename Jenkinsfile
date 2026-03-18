pipeline {
    agent any

    environment {
        REPORTS_DIR = "C:\\Autoreports\\SanityCheck\\reports"
        WORKSPACE_DIR = "${env.WORKSPACE}"
        PYLINT_THRESHOLD = "6"
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
                echo "⚙️ Installation des outils Python..."
                bat 'python -m pip install --user black pylint bandit pytest'
            }
        }

        stage('Python Syntax Check') {
            steps {
                echo "🔍 Vérification de la syntaxe Python..."
                bat """
                for %%f in (*.py) do (
                    python -m py_compile %%f 2>> "%REPORTS_DIR%\\syntax_errors.txt"
                )
                """
            }
        }

        stage('Code Format Check (Black)') {
            steps {
                echo "🎨 Vérification du format du code (Black)..."

                catchError(buildResult: 'SUCCESS', stageResult: 'UNSTABLE') {
                    bat 'python -m black --check . 1>"%REPORTS_DIR%\\black_report.txt" 2>&1'
                }
            }
        }

        stage('Run Pylint') {
            steps {
                echo "📊 Analyse qualité du code avec Pylint..."

                catchError(buildResult: 'SUCCESS', stageResult: 'UNSTABLE') {
                    bat """
                    docker run --rm -v "%WORKSPACE_DIR%:/workspace" -w /workspace sanity-python:latest ^
                    pylint *.py --fail-under=%PYLINT_THRESHOLD% --output-format=json ^
                    1>"%REPORTS_DIR%\\pylint_report.json" 2>&1
                    """
                }
            }
        }

        stage('Run Tests (Optional)') {
            steps {
                echo "🧪 Exécution des tests (si existants)..."

                catchError(buildResult: 'SUCCESS', stageResult: 'UNSTABLE') {
                    bat 'pytest --junitxml=%REPORTS_DIR%\\pytest_report.xml'
                }
            }
        }

        stage('Security Scan (Bandit)') {
            steps {
                echo "🔐 Scan de sécurité avec Bandit..."

                catchError(buildResult: 'SUCCESS', stageResult: 'UNSTABLE') {
                    bat 'python -m bandit -r . -f json -o "%REPORTS_DIR%\\bandit_report.json"'
                }

                bat '''
                if exist "convert_bandit_to_sonar.py" (
                    echo ✅ Conversion Bandit → Sonar
                    python convert_bandit_to_sonar.py "%REPORTS_DIR%\\bandit_report.json" "%REPORTS_DIR%\\bandit_report_sonar.json"
                ) else (
                    echo ⚠️ Script de conversion manquant
                )
                '''
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
                        -Dsonar.python.pylint.reportPaths=%REPORTS_DIR%\\pylint_report.json ^
                        -Dsonar.externalIssuesReportPaths=%REPORTS_DIR%\\bandit_report_sonar.json ^
                        -Dsonar.qualitygate.wait=true
                    """
                }
            }
        }

        stage('Summary') {
            steps {
                echo "📋 ===== Résumé Sanity Check ====="
                echo "✔️ Syntaxe vérifiée"
                echo "🎨 Format Black (voir rapport)"
                echo "📊 Pylint (seuil = ${PYLINT_THRESHOLD})"
                echo "🧪 Tests exécutés (si présents)"
                echo "🔐 Sécurité analysée (Bandit)"
                echo "📈 Résultats disponibles dans SonarQube"
            }
        }
    }

    post {
        always {
            echo "📦 Archivage des rapports..."

            archiveArtifacts artifacts: 'reports/*', fingerprint: true

            echo "📁 Pipeline terminé. Vérifiez les rapports dans %REPORTS_DIR%"
        }

        unstable {
            echo "⚠️ Pipeline terminé avec des avertissements (code améliorable)"
        }

        failure {
            echo "❌ Pipeline échoué. Vérifier les erreurs critiques."
        }
    }
}
