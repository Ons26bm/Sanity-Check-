# convert_bandit_json_to_sonar.py
import json
import sys

bandit_json_file = sys.argv[1]   # ex: reports/bandit_report.json
sonar_json_file = sys.argv[2]    # ex: reports/bandit_report_sonar.json

with open(bandit_json_file) as f:
    bandit_data = json.load(f)

issues = []

for issue in bandit_data.get("results", []):
    sonar_issue = {
        "engineId": "bandit",
        "ruleId": issue.get("test_id", "UNKNOWN"),
        "primaryLocation": {
            "filePath": issue.get("filename", "").replace("\\", "/"),
            "textRange": {
                "startLine": issue.get("line_number", 0),
                "endLine": issue.get("line_number", 0)
            }
        },
        "type": "VULNERABILITY",
        "severity": issue.get("issue_severity", "MINOR").upper(),
        "message": issue.get("issue_text", "")
    }
    issues.append(sonar_issue)

with open(sonar_json_file, "w") as f:
    json.dump({"issues": issues}, f, indent=2)

print(f"Conversion terminée : {len(issues)} issues générées pour SonarQube")
