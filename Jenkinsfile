pipeline {
    agent any

    environment {
        REPORTS_DIR = "C:\\Autoreports\\SanityCheck\\reports"
        WORKSPACE_DIR = "${env.WORKSPACE}"
        PYLINT_THRESHOLD = "5"
        SAMPLE_DATA = "C:\\Autoreports\\SanityCheck\\sample_data"
    }

    stages {
        stage('Checkout') {
            steps {
                git url: 'https://github.com/Ons26bm/Sanity-Check-.git'
            }
        }

        stage('Setup') {
            steps {
                echo "📁 Initialisation des dossiers..."
                bat "if not exist \"%REPORTS_DIR%\" mkdir \"%REPORTS_DIR%\""
                bat "if exist \"%WORKSPACE_DIR%\\reports\" rmdir /s /q \"%WORKSPACE_DIR%\\reports\""
                bat "mkdir \"%WORKSPACE_DIR%\\reports\""
            }
        }

        stage('Install Tools') {
            steps {
                echo "⚙️ Installation outils Python..."
                bat 'python -m pip install --user black pylint bandit pytest pip-audit'
            }
        }

        stage('Python Syntax Check') {
            steps {
                echo "🔍 Vérification syntaxe Python..."
                bat """
                for %%f in (*.py) do (
                    python -m py_compile %%f 2>> "%REPORTS_DIR%\\syntax_errors.txt"
                )
                """
            }
        }

        stage('Code Format Fix (Black)') {
            steps {
                echo "🎨 Formatage code avec Black..."
                bat 'python -m black .'
            }
        }

        stage('Run Pylint') {
            steps {
                echo "📊 Analyse qualité du code avec Pylint..."
                catchError(buildResult: 'SUCCESS', stageResult: 'UNSTABLE') {
                    bat """
                    docker run --rm -v "%WORKSPACE_DIR%:/workspace" -w /workspace sanity-python:latest ^
                    pylint *.py --fail-under=%PYLINT_THRESHOLD% --output-format=json --disable=R0801 ^
                    1>"%REPORTS_DIR%\\pylint_report.json" 2>&1
                    """
                }
            }
        }

        stage('Security Scan (Bandit)') {
            steps {
                echo "🔐 Scan de sécurité avec Bandit..."
                catchError(buildResult: 'SUCCESS', stageResult: 'UNSTABLE') {
                    bat 'python -m bandit -r . -f json -o "%REPORTS_DIR%\\bandit_report.json"'
                }
                bat '''
                if exist "convert_bandit_to_sonar.py" (
                    echo ✅ Conversion Bandit → Sonar
                    python convert_bandit_to_sonar.py "%REPORTS_DIR%\\bandit_report.json" "%REPORTS_DIR%\\bandit_report_sonar.json"
                ) else (
                    echo ⚠️ Script de conversion manquant
                )
                '''
            }
        }

        stage('Dependency Scan (pip-audit)') {
            steps {
                echo "🔍 Analyse des dépendances Python avec pip-audit..."
                catchError(buildResult: 'SUCCESS', stageResult: 'UNSTABLE') {
                    bat """
                    python -m pip_audit --format json > "%REPORTS_DIR%\\pip_audit_report.json"
                    """
                }
                bat """
                type "%REPORTS_DIR%\\pip_audit_report.json"
                """
            }
        }

        stage('AI Report Summary') {
            steps {
                script {
                    def pylintRaw = readFile("${REPORTS_DIR}\\pylint_report.json")
                    def banditRaw = readFile("${REPORTS_DIR}\\bandit_report.json")
                    def auditRaw  = readFile("${REPORTS_DIR}\\pip_audit_report.json")

                    def prompt = """
Tu es un expert en qualité de code Python.

PYLINT:
${pylintRaw.take(1000)}

BANDIT:
${banditRaw.take(1000)}

PIP-AUDIT:
${auditRaw.take(1000)}

Donne un résumé clair en français avec :
1. Problèmes critiques
2. Problèmes mineurs
3. Temps de correction
4. Solutions
"""

                    writeFile file: "request.json", text: """
{
  "model": "claude-sonnet-4-20250514",
  "max_tokens": 1024,
  "messages": [
    {
      "role": "user",
      "content": "${prompt.replace('"', '\\"')}"
    }
  ]
}
"""

                    withCredentials([string(credentialsId: 'anthropic-key', variable: 'ANTHROPIC_API_KEY')]) {
                        def response = bat(
                            returnStdout: true,
                            script: """
curl -s -X POST https://api.anthropic.com/v1/messages ^
-H "x-api-key: %ANTHROPIC_API_KEY%" ^
-H "anthropic-version: 2023-06-01" ^
-H "content-type: application/json" ^
--data @request.json
"""
                        ).trim()

                        echo "Claude raw response: ${response}"
                        env.AI_SUMMARY = response
                    }
                }
            }
        }

        stage('Run Sanity Check on Sample Data') {
            steps {
                echo "🧪 Exécution scripts sur données échantillons..."
                catchError(buildResult: 'SUCCESS', stageResult: 'UNSTABLE') {
                    bat """
                    for %%f in (*.py) do (
                        python %%f "%SAMPLE_DATA%" > "%REPORTS_DIR%\\%%~nf_results.txt" 2>&1
                    )
                    """
                }
            }
        }

        stage('SonarQube Analysis') {
            steps {
                echo "📈 Analyse avec SonarQube..."
                withSonarQubeEnv('SonarQubeServer') {
                    bat """
                    sonar-scanner ^
                        -Dsonar.projectKey=SanityCheck ^
                        -Dsonar.sources=. ^
                        -Dsonar.python.version=3.12 ^
                        -Dsonar.exclusions=reports/* ^
                        -Dsonar.python.pylint.reportPaths=%REPORTS_DIR%\\pylint_report.json
                    """
                }
            }
        }

        stage('Quality Gate') {
            steps {
                catchError(buildResult: 'UNSTABLE', stageResult: 'FAILURE') {
                    withCredentials([usernamePassword(credentialsId: 'sonar-creds',
                                                      usernameVariable: 'SONAR_USER',
                                                      passwordVariable: 'SONAR_PASS')]) {
                        script {
                            def response = bat(returnStdout: true, script: """
                                @echo off
                                curl -s -u %SONAR_USER%:%SONAR_PASS% "http://localhost:9000/api/qualitygates/project_status?projectKey=SanityCheck"
                            """).trim()

                            def jsonStart = response.indexOf('{')
                            def jsonText  = response.substring(jsonStart)
                            def json      = new groovy.json.JsonSlurper().parseText(jsonText)

                            if (json.projectStatus.status == 'ERROR') {
                                error "❌ Quality Gate failed!"
                            } else {
                                echo "✅ Quality Gate passed"
                            }
                        }
                    }
                }
            }
        }

        stage('Generate HTML Report') {
            steps {
                catchError(buildResult: 'SUCCESS', stageResult: 'UNSTABLE') {
                    script {
                        echo "Generation rapport HTML consolide..."

                        // ── 1. SYNTAXE ──────────────────────────────────────
                        def syntaxContent = ""
                        if (fileExists("${REPORTS_DIR}\\syntax_errors.txt")) {
                            syntaxContent = readFile("${REPORTS_DIR}\\syntax_errors.txt").trim()
                        }
                        def syntaxStatus = syntaxContent.isEmpty() ? "OK" : "Erreurs detectees"
                        def syntaxClass  = syntaxContent.isEmpty() ? "ok" : "fail"
                        def syntaxDetail = syntaxContent.isEmpty() ? "" : """
                            <pre class='detail fail-bg'>${syntaxContent}</pre>"""

                        // ── 2. PYLINT ───────────────────────────────────────
                        def pylintSummary = "Rapport non disponible"
                        def pylintClass   = "warn"
                        def pylintDetail  = ""
                        if (fileExists("${REPORTS_DIR}\\pylint_report.json")) {
                            try {
                                def pylintRaw  = readFile("${REPORTS_DIR}\\pylint_report.json")
                                def pylintJson = new groovy.json.JsonSlurper().parseText(pylintRaw)

                                def errors      = pylintJson.count { it.type == "error"      }
                                def warnings    = pylintJson.count { it.type == "warning"    }
                                def conventions = pylintJson.count { it.type == "convention" }
                                def scoreEntry  = pylintJson.find  { it.type == "statement"  }
                                def score       = scoreEntry ? scoreEntry.score : "N/A"

                                pylintSummary = "Score : ${score}/10 | Erreurs : ${errors} | Warnings : ${warnings} | Conventions : ${conventions}"
                                pylintClass   = errors > 0 ? "fail" : (warnings > 0 ? "warn" : "ok")

                                def topErrors = pylintJson.findAll { it.type == "error" }.take(10)
                                if (topErrors) {
                                    def rows = topErrors.collect { msg ->
                                        "<tr><td>${msg.path ?: ''}</td><td>${msg.line ?: ''}</td><td>${msg.message ?: ''}</td></tr>"
                                    }.join("\n")
                                    pylintDetail = """
                                    <table class='detail-table'>
                                        <tr><th>Fichier</th><th>Ligne</th><th>Message</th></tr>
                                        ${rows}
                                    </table>"""
                                }
                            } catch (e) {
                                pylintSummary = "Erreur lecture rapport : ${e.message}"
                                pylintClass   = "warn"
                            }
                        }

                        // ── 3. BANDIT ───────────────────────────────────────
                        def banditSummary = "Rapport non disponible"
                        def banditClass   = "warn"
                        def banditDetail  = ""
                        if (fileExists("${REPORTS_DIR}\\bandit_report.json")) {
                            try {
                                def banditRaw  = readFile("${REPORTS_DIR}\\bandit_report.json")
                                def banditJson = new groovy.json.JsonSlurper().parseText(banditRaw)

                                def high   = banditJson.results.count { it.issue_severity == "HIGH"   }
                                def medium = banditJson.results.count { it.issue_severity == "MEDIUM" }
                                def low    = banditJson.results.count { it.issue_severity == "LOW"    }

                                banditSummary = "HIGH : ${high} | MEDIUM : ${medium} | LOW : ${low}"
                                banditClass   = high > 0 ? "fail" : (medium > 0 ? "warn" : "ok")

                                def topIssues = banditJson.results
                                    .findAll { it.issue_severity in ["HIGH", "MEDIUM"] }
                                    .take(10)
                                if (topIssues) {
                                    def rows = topIssues.collect { issue ->
                                        def sevClass = issue.issue_severity == "HIGH" ? "fail" : "warn"
                                        "<tr><td class='${sevClass}'>${issue.issue_severity}</td><td>${issue.filename ?: ''}</td><td>${issue.line_number ?: ''}</td><td>${issue.issue_text ?: ''}</td></tr>"
                                    }.join("\n")
                                    banditDetail = """
                                    <table class='detail-table'>
                                        <tr><th>Severite</th><th>Fichier</th><th>Ligne</th><th>Probleme</th></tr>
                                        ${rows}
                                    </table>"""
                                }
                            } catch (e) {
                                banditSummary = "Erreur lecture rapport : ${e.message}"
                                banditClass   = "warn"
                            }
                        }

                        // ── 4. PIP-AUDIT ────────────────────────────────────
                        def auditSummary = "Rapport non disponible"
                        def auditClass   = "warn"
                        def auditDetail  = ""
                        if (fileExists("${REPORTS_DIR}\\pip_audit_report.json")) {
                            try {
                                def auditRaw  = readFile("${REPORTS_DIR}\\pip_audit_report.json")
                                def auditJson = new groovy.json.JsonSlurper().parseText(auditRaw)

                                def allVulns   = auditJson.dependencies.findAll { it.vulns && it.vulns.size() > 0 }
                                def totalVulns = allVulns.sum { it.vulns.size() } ?: 0

                                auditSummary = totalVulns == 0
                                    ? "Aucune vulnerabilite detectee"
                                    : "${totalVulns} vulnerabilite(s) sur ${allVulns.size()} paquet(s)"
                                auditClass = totalVulns == 0 ? "ok" : "fail"

                                if (allVulns) {
                                    def rows = allVulns.collect { dep ->
                                        dep.vulns.collect { v ->
                                            "<tr><td>${dep.name} ${dep.version}</td><td>${v.id ?: ''}</td><td>${v.description ?: ''}</td><td>${v.fix_versions?.join(', ') ?: 'N/A'}</td></tr>"
                                        }.join("\n")
                                    }.join("\n")
                                    auditDetail = """
                                    <table class='detail-table'>
                                        <tr><th>Paquet</th><th>CVE</th><th>Description</th><th>Version corrigee</th></tr>
                                        ${rows}
                                    </table>"""
                                }
                            } catch (e) {
                                auditSummary = "Erreur lecture rapport : ${e.message}"
                                auditClass   = "warn"
                            }
                        }

                        // ── 5. SECTION IA ────────────────────────────────────
                        def aiSection = ""
                        if (env.AI_SUMMARY) {
                            try {
                                def aiJson = new groovy.json.JsonSlurper().parseText(env.AI_SUMMARY)
                                def aiText = aiJson.content[0].text.replace("\n", "<br>")
                                aiSection = """
                                <div class='card ai-card'>
                                    <span class='label'>Analyse IA — Résumé et priorités</span>
                                    <div class='ai-content'>${aiText}</div>
                                </div>"""
                            } catch (e) {
                                aiSection = "<div class='card warn'>Résumé IA non disponible</div>"
                            }
                        }

                        // ── 6. CONSTRUCTION HTML ────────────────────────────
                        def html = """<!DOCTYPE html>
<html>
<head>
  <meta charset='UTF-8'>
  <title>Sanity Check Report</title>
  <style>
    body        { font-family: Arial, sans-serif; padding: 20px; background: #f9f9f9; }
    h2          { color: #333; border-bottom: 2px solid #ccc; padding-bottom: 8px; }
    .card       { background: white; border-radius: 6px; padding: 16px; margin-bottom: 14px;
                  box-shadow: 0 1px 4px rgba(0,0,0,0.1); }
    .label      { font-weight: bold; font-size: 15px; }
    .ok         { color: #2e7d32; }
    .fail       { color: #c62828; }
    .warn       { color: #e65100; }
    .fail-bg    { background: #fff3f3; border-left: 4px solid #c62828; padding: 8px; }
    pre.detail  { font-size: 12px; overflow-x: auto; }
    .detail-table     { border-collapse: collapse; width: 100%; margin-top: 10px; font-size: 13px; }
    .detail-table th  { background: #eeeeee; padding: 6px 10px; text-align: left; }
    .detail-table td  { padding: 5px 10px; border-bottom: 1px solid #e0e0e0; }
    .build-info { font-size: 13px; color: #666; margin-bottom: 20px; }
  </style>
</head>
<body>

<h2>Sanity Check - Rapport de qualite</h2>
<p class='build-info'>
  Build : ${currentBuild.displayName} &nbsp;|&nbsp;
  Date  : ${new Date().format('dd/MM/yyyy HH:mm')} &nbsp;|&nbsp;
  Statut global : <strong>${currentBuild.currentResult}</strong>
</p>

<div class='card'>
  <span class='label'>Syntaxe Python</span> &nbsp;
  <span class='${syntaxClass}'>${syntaxStatus}</span>
  ${syntaxDetail}
</div>

<div class='card'>
  <span class='label'>Qualite du code (Pylint)</span> &nbsp;
  <span class='${pylintClass}'>${pylintSummary}</span>
  ${pylintDetail}
</div>

<div class='card'>
  <span class='label'>Securite (Bandit)</span> &nbsp;
  <span class='${banditClass}'>${banditSummary}</span>
  ${banditDetail}
</div>

<div class='card'>
  <span class='label'>Dependances (pip-audit)</span> &nbsp;
  <span class='${auditClass}'>${auditSummary}</span>
  ${auditDetail}
</div>

${aiSection}

</body>
</html>"""

                        // ── 7. ECRITURE ──────────────────────────────────────
                        writeFile file: "${REPORTS_DIR}\\sanity_check_report.html",
                                  text: html,
                                  encoding: "UTF-8"

                        bat "copy \"${REPORTS_DIR}\\sanity_check_report.html\" \"${WORKSPACE_DIR}\\reports\\sanity_check_report.html\""
                        echo "Rapport HTML genere avec succes"
                    }
                }
            }
        }

    } // fin stages

    post {
        always {
            script {
                def reportFile   = "C:/Autoreports/SanityCheck/reports/sanity_check_report.html"
                def reportExists = fileExists(reportFile)
                echo "📄 Rapport existe : ${reportExists}"
            }
            emailext(
                subject: "Sanity Check - Résultat: ${currentBuild.currentResult}",
                body: """Le pipeline est terminé.
Build: ${currentBuild.displayName}
Résultat: ${currentBuild.currentResult}
Voir rapport en PIECE JOINTE.""",
                attachmentsPattern: "reports/sanity_check_report.html",
                to: "pw39f@ningen-group.com"
            )
        }
    }

} // fin pipeline
