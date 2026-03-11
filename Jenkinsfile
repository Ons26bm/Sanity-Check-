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
                bat "if not exist \"${REPORTS_DIR}\" mkdir \"${REPORTS_DIR}\""
                bat "if exist \"${WORKSPACE_DIR}\\reports\" rmdir /s /q \"${WORKSPACE_DIR}\\reports\""
                bat "mkdir \"${WORKSPACE_DIR}\\reports\""
            }
        }

    stage('Python Syntax Check') {
    steps {
        bat """
        for %%f in (*.py) do python -m py_compile %%f 2>> "C:\\Autoreports\\SanityCheck\\reports\\syntax_errors.txt" || exit 0
        """
    }
}

        stage('Code Format Check (Black)') {
            steps {
                bat 'python -m pip install --user black'
                bat 'python -m black --check . 1>"${REPORTS_DIR}\\black_report.txt" 2>&1 || exit 0'
            }
        }

        stage('Run Pylint') {
            steps {
                bat """
                docker run --rm -v "${WORKSPACE_DIR}:/workspace" -w /workspace sanity-python:latest \
                pylint *.py --output-format=json 1>"${REPORTS_DIR}\\pylint_report.json" 2>&1 || exit 0
                """
            }
        }

stage('Security Scan (Bandit)') {
    steps {
        echo 'Scan de sécurité avec Bandit...'
        // Installer Bandit si nécessaire
        bat 'python -m pip install --user bandit'

        // Générer le rapport JSON standard Bandit
        bat 'python -m bandit -r . -f json -o "C:\\Autoreports\\SanityCheck\\reports\\bandit_report.json"'

        // Vérifier que le fichier de conversion existe dans le workspace
        bat 'if exist "convert_bandit_to_sonar.py" (echo Script trouvé) else (echo Script manquant & exit 1)'

        // Exécuter le script de conversion pour SonarQube
        bat 'python convert_bandit_to_sonar.py "C:\\Autoreports\\SanityCheck\\reports\\bandit_report.json" "C:\\Autoreports\\SanityCheck\\reports\\bandit_report_sonar.json"'
    }
}

stage('SonarQube Analysis') {
    steps {
        withSonarQubeEnv('SonarQubeServer') {
            echo 'Analyse SonarQube...'
            bat '''
            sonar-scanner ^
                -Dsonar.projectKey=SanityCheck ^
                -Dsonar.sources=. ^
                -Dsonar.python.version=3.12 ^
                -Dsonar.exclusions=reports/* ^
                -Dsonar.python.pylint.reportPaths=reports/pylint_report.json ^
                -Dsonar.externalIssuesReportPaths=reports/bandit_report_sonar.json
            '''
        }
    }
}
    }

    post {
        always {
            echo "Pipeline terminé. Vérifier les rapports dans ${REPORTS_DIR}"
        }
        failure {
            echo "Pipeline échoué. Vérifier les logs et rapports."
        }
    }
}
