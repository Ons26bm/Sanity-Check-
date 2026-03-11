import json
import sys
import os

def convert_bandit_to_sonar(input_file, output_file):
    with open(input_file, 'r', encoding='utf-8') as f:
        bandit_data = json.load(f)

    sonar_issues = {"issues": []}

    for issue in bandit_data.get("results", []):
        # Préparer l'issue au format SonarQube
        sonar_issue = {
            "engineId": "bandit",
            "ruleId": issue.get("test_id", "unknown"),
            "severity": issue.get("issue_severity", "MAJOR").upper(),
            "type": "VULNERABILITY",
            "primaryLocation": {
                "message": issue.get("issue_text", "No message provided"),
                "filePath": os.path.relpath(issue.get("filename", ""), start=os.getcwd()).replace("\\", "/"),
                "textRange": {
                    "startLine": issue.get("line_number", 1),
                    "endLine": issue.get("line_number", 1),
                    "startColumn": issue.get("col_offset", 0),
                    "endColumn": issue.get("end_col_offset", 0)
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
