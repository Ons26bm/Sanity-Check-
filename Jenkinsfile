pipeline {
    agent any

    environment {
        REPORTS_DIR = "C:\\Autoreports\\SanityCheck\\reports"
        WORKSPACE_DIR = "${env.WORKSPACE}"
    }

    stages {
        stage('Checkout') {
            steps {
                git url: 'https://github.com/Ons26bm/Sanity-Check-.git'
            }
        }

        stage('Setup') {
            steps {
                // Créer le dossier de rapports si nécessaire
                bat "if not exist \"%REPORTS_DIR%\" mkdir \"%REPORTS_DIR%\""
                // Supprimer les anciens rapports
                bat "if exist \"%WORKSPACE_DIR%\\reports\" rmdir /s /q \"%WORKSPACE_DIR%\\reports\""
                // Créer dossier reports dans le workspace
                bat "mkdir \"%WORKSPACE_DIR%\\reports\""
            }
        }

        stage('Python Syntax Check') {
            steps {
                bat """
                for %%f in (*.py) do python -m py_compile %%f 2>> "%REPORTS_DIR%\\syntax_errors.txt" || exit 0
                """
            }
        }

        stage('Code Format Check (Black)') {
            steps {
                bat 'python -m pip install --user black'
                bat 'python -m black --check . 1>"%REPORTS_DIR%\\black_report.txt" 2>&1 || exit 0'
            }
        }

        stage('Run Pylint') {
            steps {
                bat """
                docker run --rm -v "%WORKSPACE_DIR%:/workspace" -w /workspace sanity-python:latest ^
                pylint *.py --output-format=json 1>"%REPORTS_DIR%\\pylint_report.json" 2>&1 || exit 0
                """
            }
        }

        stage('Security Scan (Bandit)') {
            steps {
                echo 'Scan de sécurité avec Bandit...'
                bat 'python -m pip install --user bandit'
                bat 'python -m bandit -r . -f json -o "%REPORTS_DIR%\\bandit_report.json"'

                // Vérifier que le script de conversion existe
                bat 'if exist "convert_bandit_to_sonar.py" (echo Script trouvé) else (echo Script manquant & exit 1)'

                // Conversion du rapport Bandit vers format SonarQube
                bat 'python convert_bandit_to_sonar.py "%REPORTS_DIR%\\bandit_report.json" "%REPORTS_DIR%\\bandit_report_sonar.json" || exit 0'
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
                        -Dsonar.python.pylint.reportPaths=%REPORTS_DIR%\\pylint_report.json ^
                        -Dsonar.externalIssuesReportPaths=%REPORTS_DIR%\\bandit_report_sonar.json
                    """
                }
            }
        }
    }

    post {
        always {
            echo "Pipeline terminé. Vérifier les rapports dans %REPORTS_DIR%"
        }
        failure {
            echo "Pipeline échoué. Vérifier les logs et rapports."
        }
    }
}
