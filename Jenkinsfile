pipeline {
    agent any

  environment {
    PYTHON_PATH = "C:\\Users\\PW39F\\AppData\\Local\\Programs\\Python\\Python312\\python.exe"
    REPORT_DIR = "C:\\Autoreports\\TestReports"
}

    stages {
        // Stage 1 : Checkout Git
        stage('Git Checkout') {
            steps {
                git branch: 'master',
                    url: 'https://github.com/Ons26bm/Sanity-Check-.git'
            }
        }

        // Stage 2 : Build / Setup
        stage('Build / Setup') {
            steps {
                echo 'Installation des dépendances Python (si nécessaire)...'
                bat "if not exist ${env.REPORT_DIR} mkdir ${env.REPORT_DIR}"
            }
        }

      stage('SonarQube Analysis') {
       steps {
        withSonarQubeEnv('SonarQubeServer') {
            bat "C:\\Users\\PW39F\\Downloads\\sonar-scanner-cli-8.0.1.6346-windows-x64\\sonar-scanner-8.0.1.6346-windows-x64\\bin\\sonar-scanner.bat -Dsonar.projectKey=TestAutoreports -Dsonar.sources=./"
        }
    }
}

        // // Stage 4 : Test / Sanity Check
        stage('Test / Sanity Check') {
            steps {
                echo 'Exécution du fichier de test...'
                bat "${env.PYTHON_PATH} bot_NJERI.py > ${env.REPORT_DIR}\\bot_njeri_log.txt 2>&1"
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
