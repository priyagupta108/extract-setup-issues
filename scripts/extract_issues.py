import requests
import datetime
import openpyxl

TOKEN = os.environ['GH_TOKEN']
OWNER = "actions"
REPO = "runner-images"
SINCE_DATE = (datetime.datetime.utcnow() - datetime.timedelta(days=90)).isoformat() + "Z"
PER_PAGE = 100

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
            "since": SINCE_DATE,
            "per_page": PER_PAGE,
            "page": page
        }
        response = requests.get(url, headers=headers, params=params)
        data = response.json()
        if not data:
            break
        issues.extend(data)
        page += 1
    return [i for i in issues if "pull_request" not in i]

def export_to_excel(issues, filename="extracted_issues.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Issues"

    headers = ["Number", "Title", "State", "Created At", "Updated At", "URL"]
    ws.append(headers)

    for issue in issues:
        ws.append([
            issue["number"],
            issue["title"],
            issue["state"],
            issue["created_at"],
            issue["updated_at"],
            issue["html_url"]
        ])

    wb.save(filename)

if __name__ == "__main__":
    open_issues = get_issues("open")
    closed_issues = get_issues("closed")
    all_issues = open_issues + closed_issues
    export_to_excel(all_issues)
