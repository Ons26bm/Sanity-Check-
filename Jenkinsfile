pipeline {
    // On utilise l'image Docker que tu as créée

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
        stage('Run tests in Docker') {
    steps {
        echo 'Lancer les tests avec pytest et pylint dans Docker...'
        bat """
        docker run --rm ^
        -v "C:\\ProgramData\\Jenkins\\.jenkins\\workspace\\SanityCheckScripts:C:/workspace" ^
        -w "C:/workspace" ^
        sanity-python:latest ^
        cmd.exe /c "pytest && pylint *.py"
        """
    }
}

        // Stage 3 : Lint avec pylint
        stage('Lint') {
            steps {
                echo 'Exécution de pylint...'
                bat "pylint . > ${env.REPORT_DIR}\\pylint_report.txt || exit 0"
            }
        }

        // Stage 4 : Test avec pytest
        stage('Tests') {
            steps {
                echo 'Exécution des tests pytest...'
                bat "pytest --maxfail=1 --disable-warnings -q > ${env.REPORT_DIR}\\pytest_report.txt || exit 0"
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
                bat "python bot_NJERI.py > ${env.REPORT_DIR}\\bot_njeri_log.txt 2>&1 || exit 0"
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
