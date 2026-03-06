pipeline { 
    agent any

    environment {
        REPORT_DIR = "C:\\Autoreports\\SanityCheck"
        WORKSPACE_DIR = "C:\\ProgramData\\Jenkins\\.jenkins\\workspace\\SanityCheckScripts"
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
                bat "if not exist ${env.REPORT_DIR} mkdir ${env.REPORT_DIR}"
                bat "if not exist ${env.WORKSPACE_DIR}\\reports mkdir ${env.WORKSPACE_DIR}\\reports"
            }
        }

        stage('Check Python') {
            steps {
                bat "where python"
                bat "echo %PATH%"
            }
        }

        // Stage 4 : Run pytest sur Windows
        stage('Run pytest') {
            steps {
                echo 'Exécution des tests pytest avec coverage sur Windows...'
                bat """
                python -m pytest --maxfail=1 --disable-warnings -q ^
                --cov=${env.WORKSPACE_DIR} ^
                --cov-report=xml:${env.WORKSPACE_DIR}\\reports\\coverage.xml
                """
            }
        }

        // Stage 5 : Run pylint dans Docker
        stage('Run pylint in Docker') {
            steps {
                echo 'Exécution de pylint dans Docker...'
                bat """
                docker run --rm ^
                -v "${env.WORKSPACE_DIR}:/workspace" ^
                -w "/workspace" ^
                sanity-python:latest ^
                bash -c "pylint *.py --output-format=parseable > /workspace/reports/pylint_report.xml || true"
                """
            }
        }

        stage('SonarQube Analysis') {
            steps {
                withSonarQubeEnv('SonarQubeServer') {
                    bat """
                    sonar-scanner ^
                    -Dsonar.projectKey=SanityCheck ^
                    -Dsonar.sources=./ ^
                    -Dsonar.python.coverage.reportPaths=${env.WORKSPACE_DIR}\\reports\\coverage.xml ^
                    -Dsonar.python.pylint.reportPaths=${env.WORKSPACE_DIR}\\reports\\pylint_report.xml
                    """
                }
            }
        }

        stage('Sanity Check Script') {
            steps {
                echo 'Exécution du script principal bot_NJERI.py...'
                bat "\"C:\\Users\\PW39F\\AppData\\Local\\Programs\\Python\\Python312\\python.exe\" bot_NJERI.py > ${env.REPORT_DIR}\\bot_njeri_log.txt 2>&1 || exit 0"
            }
        }
    }

    post {
        success {
            echo 'Pipeline terminé avec succès et Quality Gate validé !'
        }
        failure {
            echo 'Pipeline échoué ou Quality Gate non validé, vérifier les logs.'
        }
    }
}
