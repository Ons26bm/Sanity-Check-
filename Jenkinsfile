pipeline {
    agent any

    environment {
        REPORT_DIR = "C:\\Autoreports\\SanityCheck\\reports"
        WORKSPACE_DIR = "${env.WORKSPACE}"
    }

    triggers {
        pollSCM('H/2 * * * *')
    }

    stages {
        stage('Git Checkout') {
            steps {
                git branch: 'master',
                    url: 'https://github.com/Ons26bm/Sanity-Check-.git'
            }
        }

        stage('Setup') {
            steps {
                echo 'Création des dossiers nécessaires...'
                bat "if not exist \"${REPORT_DIR}\" mkdir \"${REPORT_DIR}\""
                bat """
                if exist \"${WORKSPACE_DIR}\\reports\" (
                    echo Tentative de suppression du dossier reports...
                    rmdir /s /q \"${WORKSPACE_DIR}\\reports\" || echo Impossible de supprimer reports
                )
                mkdir \"${WORKSPACE_DIR}\\reports\"
                """
            }
        }

        stage('Python Syntax Check') {
            steps {
                echo 'Vérification de la syntaxe Python...'
                bat """
                for %%f in (*.py) do python -m py_compile %%f 2>> "${REPORT_DIR}\\syntax_errors.txt" || exit 0
                """
            }
        }

        stage('Code Format Check (Black)') {
            steps {
                echo 'Vérification du format du code avec Black...'
                bat """
                python -m pip install --user black
                python -m black --check . > "${REPORT_DIR}\\black_report.txt" 2>&1 || exit 0
                """
            }
        }

        stage('Run Pylint') {
            steps {
                echo 'Exécution de Pylint pour l’analyse statique du code...'
                bat """
                docker run --rm ^
                -v "${WORKSPACE_DIR}:/workspace" ^
                -w /workspace ^
                sanity-python:latest ^
                pylint *.py --output-format=json > "${REPORT_DIR}\\pylint_report.json" 2>&1 || exit 0
                """
            }
        }

        stage('Security Scan (Bandit)') {
            steps {
                echo 'Scan de sécurité avec Bandit...'
                bat """
                python -m pip install --user bandit
                python -m bandit -r . -f txt -o "${REPORT_DIR}\\bandit_report.txt" || exit 0
                python "${WORKSPACE_DIR}\\convert_bandit_to_sonar.py" "${REPORT_DIR}\\bandit_report.txt" "${REPORT_DIR}\\bandit_report.json"
                if not exist "${REPORT_DIR}\\bandit_report.json" (
                    echo ERREUR : bandit_report.json introuvable!
                    exit 1
                )
                """
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
                    -Dsonar.python.pylint.reportPaths=reports/pylint_report.json ^
                    -Dsonar.externalIssuesReportPaths=${REPORT_DIR}/bandit_report.json
                    """
                }
            }
        }

        stage('Sanity Check Script') {
            steps {
                echo 'Exécution du script principal bot_NJERI.py...'
                bat "\"C:\\Users\\PW39F\\AppData\\Local\\Programs\\Python\\Python312\\python.exe\" bot_NJERI.py > \"${REPORT_DIR}\\bot_njeri_log.txt\" 2>&1 || exit 0"
            }
        }
    }

    post {
        success {
            echo 'Pipeline terminé avec succès ! Tous les rapports sont dans C:\\Autoreports\\SanityCheck\\reports'
        }
        failure {
            echo 'Pipeline échoué, vérifier les logs et rapports.'
        }
    }
}
