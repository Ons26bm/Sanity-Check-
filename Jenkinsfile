pipeline {
    agent any

    environment {
        REPORT_DIR = "C:\\Autoreports\\SanityCheck"
        WORKSPACE_DIR = "C:\\ProgramData\\Jenkins\\.jenkins\\workspace\\SanityCheckScripts"
    }

    stages {
        // Stage 1 : Checkout Git
        stage('Git Checkout') {
            steps {
                git branch: 'master',
                    url: 'https://github.com/Ons26bm/Sanity-Check-.git'
            }
        }

        // Stage 2 : Setup
        stage('Setup') {
            steps {
                echo 'Création des dossiers nécessaires...'
                bat "if not exist ${env.REPORT_DIR} mkdir ${env.REPORT_DIR}"
                bat "if not exist ${env.WORKSPACE_DIR}\\reports mkdir ${env.WORKSPACE_DIR}\\reports"
            }
        }

        // Stage 3 : Check Python
        stage('Check Python') {
            steps {
                bat "where python"
                bat "echo %PATH%"
            }
        }

        // Stage 4 : Run tests in Docker (pytest + pylint + coverage)
        stage('Run tests in Docker') {
            steps {
                echo 'Exécution des tests pytest et pylint avec coverage...'
                bat """
                docker run --rm ^
                -v "${env.WORKSPACE_DIR}:/workspace" ^
                -w "/workspace" ^
                sanity-python:latest ^
                bash -c "pytest --maxfail=1 --disable-warnings -q --cov=. --cov-report=xml:/workspace/reports/coverage.xml > /workspace/reports/pytest_report.txt 2>&1; pylint *.py > /workspace/reports/pylint_report.txt 2>&1 || true"
                """
            }
        }

        // Stage 5 : SonarQube Analysis
        stage('SonarQube Analysis') {
            steps {
                withSonarQubeEnv('SonarQubeServer') {
                    bat """
                    sonar-scanner ^
                    -Dsonar.projectKey=SanityCheck ^
                    -Dsonar.sources=./ ^
                    -Dsonar.python.coverage.reportPaths=${env.WORKSPACE_DIR}\\reports\\coverage.xml ^
                    -Dsonar.python.pylint.reportPaths=${env.WORKSPACE_DIR}\\reports\\pylint_report.txt
                    """
                }
            }
        }

        // Stage 6 : Exécution du script principal (sanity check)
        stage('Sanity Check Script') {
            steps {
                echo 'Exécution du script principal bot_NJERI.py...'
                bat "\"C:\\Users\\PW39F\\AppData\\Local\\Programs\\Python\\Python312\\python.exe\" bot_NJERI.py > ${env.REPORT_DIR}\\bot_njeri_log.txt 2>&1 || exit 0"
            }
        }
    }

    post {
        success {
            echo 'Pipeline terminé avec succès !'
        }
        failure {
            echo 'Pipeline échoué, vérifier les logs.'
        }
    }
}
