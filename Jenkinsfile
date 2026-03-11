pipeline {
    agent any

    environment {
        REPORT_DIR = "C:\\Autoreports\\SanityCheck\\reports"
        WORKSPACE_DIR = "${env.WORKSPACE}"
    }

    triggers {
        // Déclenchement automatique à chaque push (polling toutes les 2 minutes)
        pollSCM('H/2 * * * *')
    }

    stages {
        // Stage 1 : Checkout Git
        stage('Git Checkout') {
            steps {
                git branch: 'master',
                    url: 'https://github.com/Ons26bm/Sanity-Check-.git'
            }
        }

        // Stage 2 : Setup dossiers
        stage('Setup') {
            steps {
                echo 'Création des dossiers nécessaires...'
                bat "if not exist \"${env.REPORT_DIR}\" mkdir \"${env.REPORT_DIR}\""
                bat """
                if exist \"${env.WORKSPACE_DIR}\\reports\" (
                    echo Tentative de suppression du dossier reports...
                    rmdir /s /q \"${env.WORKSPACE_DIR}\\reports\" || echo Impossible de supprimer reports, peut être utilisé par un autre processus
                )
                mkdir \"${env.WORKSPACE_DIR}\\reports\"
                """
            }
        }

        // Stage 3 : Vérification Syntaxe Python
        stage('Python Syntax Check') {
            steps {
                echo 'Vérification de la syntaxe Python...'
                // Boucle sur chaque fichier .py pour Windows
                bat """
                for %%f in (*.py) do python -m py_compile %%f 2>> "${env.REPORT_DIR}\\syntax_errors.txt"
                """
            }
        }

        // Stage 4 : Vérification Format Code (Black)
      stage('Code Format Check (Black)') {
    steps {
        echo 'Vérification du format du code avec Black...'
        bat """
        python -m pip install --user black
        python -m black --check . > ${env.REPORT_DIR}\\black_report.txt 2>&1 || exit 0
        """
    }
}

        // Stage 5 : Analyse statique Pylint dans Docker
        stage('Run Pylint') {
            steps {
                echo 'Exécution de Pylint pour l’analyse statique du code...'
                bat """
                docker run --rm ^
                -v "${env.WORKSPACE_DIR}:/workspace" ^
                -w /workspace ^
                sanity-python:latest ^
                pylint *.py --output-format=json > "${env.WORKSPACE_DIR}\\reports\\pylint_report.json" 2>&1 || exit 0
                """
            }
        }

        // Stage 6 : Security Scan (Bandit)
           stage('Security Scan (Bandit)') {
            steps {
                echo 'Scan de sécurité avec Bandit...'
                bat """
                python -m pip install --user bandit
                python -m bandit -r . -f txt -o ${env.REPORT_DIR}\\bandit_report.txt
                """
            }
        }

        // Stage 7 : SonarQube Analysis (Dashboard)
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

        // Stage 8 : Exécution du script principal
        stage('Sanity Check Script') {
            steps {
                echo 'Exécution du script principal bot_NJERI.py...'
                bat "\"C:\\Users\\PW39F\\AppData\\Local\\Programs\\Python\\Python312\\python.exe\" bot_NJERI.py > ${env.REPORT_DIR}\\bot_njeri_log.txt 2>&1 || exit 0"
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
