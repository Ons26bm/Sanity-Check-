pipeline {
     agent any

    environment {
        REPORT_DIR = "C:\\Autoreports\\SanityCheck"
    }

    stages {
        // Stage 1 : Checkout Git
        stage('Git Checkout') {
            steps {
                git branch: 'master',
                    url: 'https://github.com/Ons26bm/Sanity-Check-.git'
            }
        }

        // Stage 2 : Préparation / Setup
        stage('Setup') {
            steps {
                echo 'Création du dossier de rapport si nécessaire...'
                bat "if not exist ${env.REPORT_DIR} mkdir ${env.REPORT_DIR}"
            }
        }
stage('Check Python') {
    steps {
        bat "where python"
        bat "echo %PATH%"
    }
}
stage('Run tests in Docker') {
    steps {
        echo 'Exécution des tests pytest et pylint dans Docker...'

        // On crée d'abord un dossier "reports" dans le workspace pour stocker les rapports
        bat "if not exist C:\\ProgramData\\Jenkins\\.jenkins\\workspace\\SanityCheckScripts\\reports mkdir C:\\ProgramData\\Jenkins\\.jenkins\\workspace\\SanityCheckScripts\\reports"

        // Lancement du conteneur Docker avec le workspace monté et exécution des tests
        bat """
        docker run --rm ^
        -v "C:\\ProgramData\\Jenkins\\.jenkins\\workspace\\SanityCheckScripts:/workspace" ^
        -w "/workspace" ^
        sanity-python:latest ^
        bash -c "pytest --maxfail=1 --disable-warnings -q > /workspace/reports/pytest_report.txt 2>&1; pylint *.py > /workspace/reports/pylint_report.txt 2>&1 || true"
        """
    }
}

        // Stage 5 : SonarQube Analysis
        stage('SonarQube Analysis') {
            steps {
                withSonarQubeEnv('SonarQubeServer') {
                    bat "sonar-scanner -Dsonar.projectKey=SanityCheck -Dsonar.sources=./"
                }
            }
        }

        // Stage 6 : Exécution du script principal (sanity check)
        stage('Sanity Check Script') {
            steps {
                echo 'Exécution du script principal bot_NJERI.py...'
                bat "C:\\Users\\PW39F\\AppData\\Local\\Programs\\Python\\Python312\\python.exe\" bot_NJERI.py > ${env.REPORT_DIR}\\bot_njeri_log.txt 2>&1 || exit 0"
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
