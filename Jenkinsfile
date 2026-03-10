pipeline {
    agent any

    environment {
        REPORT_DIR = "C:\\Autoreports\\SanityCheck\\reports"
       WORKSPACE_DIR = "${env.WORKSPACE}"
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
        // Création du dossier de rapports si inexistant
        bat "if not exist \"${env.REPORT_DIR}\" mkdir \"${env.REPORT_DIR}\""

        // Suppression sécurisée du dossier reports
        bat """
        if exist \"${env.WORKSPACE_DIR}\\reports\" (
            echo Tentative de suppression du dossier reports...
            rmdir /s /q \"${env.WORKSPACE_DIR}\\reports\" || echo Impossible de supprimer reports, peut être utilisé par un autre processus
        )
        mkdir \"${env.WORKSPACE_DIR}\\reports\\"
        """
    }
}

        // Stage 3 : Check Python
        stage('Check Python') {
            steps {
                bat "where python"
                bat "echo %PATH%"
            }
        }

        // Stage 4 : Run Pylint in Docker
stage('Run Pylint in Docker') {
    steps {
        echo 'Exécution de Pylint pour l’analyse statique du code...'
        bat """
        docker run --rm ^
        -v "C:\\ProgramData\\Jenkins\\.jenkins\\workspace\\SanityCheckScripts:/workspace" ^
        -w /workspace ^
        sanity-python:latest ^
        pylint *.py --output-format=parseable > \"${env.WORKSPACE_DIR}\\reports\\pylint_report.json\" 2>&1 || exit 0
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
-Dsonar.sources=. ^
-Dsonar.python.version=3.12 ^
-Dsonar.exclusions=reports/* ^
-Dsonar.python.pylint.reportPaths=reports/pylint_report.json
            """
        }
    }
}
  stage('Quality Gate') {
    steps {
        timeout(time: 5, unit: 'MINUTES') {
            waitForQualityGate abortPipeline: true
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
            echo 'Pipeline terminé avec succès et Quality Gate validé !'
        }
        failure {
            echo 'Pipeline échoué ou Quality Gate non validé, vérifier les logs.'
        }
    }
}
