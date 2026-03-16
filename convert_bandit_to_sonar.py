import json
import sys
import os

# Mapping Bandit -> SonarQube
SEVERITY_MAP = {
    "HIGH": "CRITICAL",
    "MEDIUM": "MAJOR",
    "LOW": "MINOR",
    "UNDEFINED": "INFO"
}

def convert_bandit_to_sonar(input_file, output_file):
    with open(input_file, 'r', encoding='utf-8') as f:
        bandit_data = json.load(f)

    sonar_issues = {"issues": []}

    for issue in bandit_data.get("results", []):
        bandit_severity = issue.get("issue_severity", "UNDEFINED").upper()
        sonar_severity = SEVERITY_MAP.get(bandit_severity, "INFO")

        filename = issue.get("filename", "")
        line_number = issue.get("line_number", 1)

        # Lecture de la ligne réelle dans le fichier pour ajuster les colonnes
        start_col = issue.get("col_offset", 1)
        end_col = issue.get("end_col_offset", start_col)
        try:
            with open(filename, "r", encoding="utf-8") as f:
                lines = f.readlines()
            line_index = line_number - 1
            if 0 <= line_index < len(lines):
                line_length = len(lines[line_index].rstrip("\n"))
                start_col = min(max(start_col, 1), line_length)
                end_col = min(max(end_col, start_col), line_length)
            else:
                start_col = 1
                end_col = start_col
        except Exception:
            # fallback si le fichier n'existe pas ou problème de lecture
            start_col = 1
            end_col = 1

        # Préparer l'issue au format SonarQube
        sonar_issue = {
            "engineId": "bandit",
            "ruleId": issue.get("test_id", "unknown"),
            "severity": sonar_severity,
            "type": "VULNERABILITY",
            "primaryLocation": {
                "message": issue.get("issue_text", "No message provided"),
                "filePath": os.path.relpath(filename, start=os.getcwd()).replace("\\", "/"),
                "textRange": {
                    "startLine": line_number,
                    "endLine": line_number,
                    "startColumn": start_col,
                    "endColumn": end_col
                }
            }
        }
        sonar_issues["issues"].append(sonar_issue)

    # Sauvegarder le JSON au format SonarQube
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(sonar_issues, f, indent=2)

    print(f"Conversion terminée : {len(sonar_issues['issues'])} issues générées pour SonarQube")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python convert_bandit_to_sonar.py <bandit_report.json> <output_sonar.json>")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]

    convert_bandit_to_sonar(input_path, output_path)
