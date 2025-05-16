import os
import requests
import datetime
import openpyxl

TOKEN = os.environ['GH_TOKEN']
OWNER = "actions"
REPO = "runner-images"
START_DATE = "2025-05-01T00:00:00Z"
TODAY_DATE = datetime.datetime.utcnow().isoformat() + "Z"
PER_PAGE = 100

TARGET_LABELS = {"OS: macOS", "OS: Ubuntu", "OS: Windows", "OS: Ubuntu24"}

headers = {
    "Authorization": f"token {TOKEN}",
    "Accept": "application/vnd.github.v3+json"
}

def get_issues(state):
    issues = []
    page = 1
    while True:
        url = f"https://api.github.com/repos/{OWNER}/{REPO}/issues"
        params = {
            "state": state,
            "since": START_DATE,
            "per_page": PER_PAGE,
            "page": page
        }
        response = requests.get(url, headers=headers, params=params, timeout=10)
        response.raise_for_status()
        data = response.json()
        if not data:
            break
        # Filter issues created between 1 May and today, exclude PRs
        for issue in data:
            created_at = issue.get("created_at")
            if created_at and START_DATE <= created_at <= TODAY_DATE and "pull_request" not in issue:
                issues.append(issue)
        page += 1
    return issues

def issues_to_excel_grouped(issues, filename="issues_grouped.xlsx"):
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # Prepare groups
    grouped_issues = {label: [] for label in TARGET_LABELS}
    grouped_issues["Other"] = []

    for issue in issues:
        labels = {label_obj["name"] for label_obj in issue.get("labels", [])}
        matched_labels = TARGET_LABELS.intersection(labels)
        if matched_labels:
            for label in matched_labels:
                grouped_issues[label].append(issue)
        else:
            grouped_issues["Other"].append(issue)

    # Write each group to its own sheet
    for label, issues_list in grouped_issues.items():
        ws = wb.create_sheet(title=label.replace(":", "").replace(" ", "_"))
        headers = ["Number", "Title", "State", "Created At", "Closed At", "Labels"]
        ws.append(headers)
        for issue in issues_list:
            ws.append([
                issue["number"],
                issue["title"],
                issue["state"],
                issue["created_at"],
                issue.get("closed_at", ""),
                ", ".join([lbl["name"] for lbl in issue.get("labels", [])])
            ])

    wb.save(filename)

def issues_to_excel_all(issues, filename="all_issues.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "All Issues"

    headers = ["Number", "Title", "State", "Created At", "Closed At", "Labels"]
    ws.append(headers)

    for issue in issues:
        ws.append([
            issue["number"],
            issue["title"],
            issue["state"],
            issue["created_at"],
            issue.get("closed_at", ""),
            ", ".join([lbl["name"] for lbl in issue.get("labels", [])])
        ])

    wb.save(filename)

if __name__ == "__main__":
    open_issues = get_issues("open")
    closed_issues = get_issues("closed")
    all_issues = open_issues + closed_issues

    # Save grouped issues into one Excel workbook (multiple sheets)
    issues_to_excel_grouped(all_issues, filename="issues_grouped.xlsx")

    # Save all issues (no grouping) into a separate Excel workbook
    issues_to_excel_all(all_issues, filename="all_issues.xlsx")
