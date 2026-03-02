pipeline {
    agent any

    environment {
        PYTHON_PATH = "C:\Users\PW39F\AppData\Local\Programs\Python\Python312\python.exe"   // chemin vers Python
        REPORT_DIR = "C:\Users\PW39F\SharePointConnection\bot_NJERI.py"
    }

    stages {
        // Stage 1 : Checkout Git
        stage('Git Checkout') {
            steps {
                git branch: 'main',
                    url: 'https://github.com/Ons26bm/Sanity-Check-.git'
            }
        }

        // Stage 2 : Build / Setup
        stage('Build / Setup') {
            steps {
                echo 'Installation des dépendances Python (si nécessaire)...'
                bat "${env.PYTHON_PATH} -m pip install -r requirements.txt"
                bat "if not exist ${env.REPORT_DIR} mkdir ${env.REPORT_DIR}"
            }
        }

        // Stage 3 : SonarQube Analyse
        stage('SonarQube Analysis') {
            steps {
                withSonarQubeEnv('SonarQubeServer') {
                    bat "sonar-scanner -Dsonar.projectKey=TestAutoreports -Dsonar.sources=./"
                }
            }
        }

        // Stage 4 : Test / Sanity Check
        stage('Test / Sanity Check') {
            steps {
                echo 'Exécution du fichier de test...'
                bat "${env.PYTHON_PATH} test_script.py > ${env.REPORT_DIR}\\test_script_log.txt 2>&1"
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
