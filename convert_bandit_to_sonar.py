# convert_bandit_to_sonar.py
import json
import sys

txt_file = sys.argv[1]
json_file = sys.argv[2]

issues = []

with open(txt_file) as f:
    lines = f.read().splitlines()

i = 0
while i < len(lines):
    if lines[i].startswith(">> Issue:"):
        issue = {}
        issue['message'] = lines[i].split("]")[-1].strip()
        issue['severity'] = lines[i+1].split(":")[-1].strip().lower()
        loc_line = lines[i+4]
        file_path = loc_line.split(":")[0].replace(".\\","")
        line_number = int(loc_line.split(":")[1])
        issue['component'] = file_path
        issue['line'] = line_number
        issues.append(issue)
        i += 6
    else:
        i += 1

with open(json_file, "w") as f:
    json.dump({"issues": issues}, f, indent=2)
