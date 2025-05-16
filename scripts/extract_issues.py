import os
import requests
import datetime
import time
import pandas as pd

# Set up
TOKEN = os.environ['GH_TOKEN']
OWNER = "actions"
REPO = "runner-images"
LABELS_TO_INCLUDE = {"OS: macOS", "OS: Ubuntu", "OS: Windows", "OS: Ubuntu24"}

headers = {
    "Authorization": f"token {TOKEN}",
    "Accept": "application/vnd.github.v3+json"
}

def get_issues(state, since, until):
    issues = []
    page = 1
    per_page = 100

    while True:
        url = f"https://api.github.com/repos/{OWNER}/{REPO}/issues"
        params = {
            "state": state,
            "since": since,
            "per_page": per_page,
            "page": page
        }

        try:
            response = requests.get(url, headers=headers, params=params, timeout=10)
            response.raise_for_status()
        except requests.exceptions.RequestException as e:
            print(f"Request failed on page {page}: {e}")
            break

        data = response.json()
        if not data:
            break

        for issue in data:
            if "pull_request" in issue:
                continue

            created_at = datetime.datetime.strptime(issue["created_at"], "%Y-%m-%dT%H:%M:%SZ")
            if not (since_dt <= created_at <= until_dt):
                continue

            label_names = [label["name"] for label in issue.get("labels", [])]

            issues.append({
                "number": issue["number"],
                "title": issue["title"],
                "state": issue["state"],
                "created_at": issue["created_at"],
                "closed_at": issue.get("closed_at"),
                "labels": label_names
            })

        print(f"Fetched page {page} with {len(data)} items")
        page += 1
        time.sleep(1)  # Rate limiting

    return issues

# Date range
since_dt = datetime.datetime(2025, 5, 1)
until_dt = datetime.datetime.utcnow()
since_iso = since_dt.isoformat() + "Z"

# Get open and closed issues
open_issues = get_issues("open", since_iso, until_dt)
closed_issues = get_issues("closed", since_iso, until_dt)

# Combine issues
all_issues = open_issues + closed_issues

# Group issues by label
grouped_issues = {label: [] for label in LABELS_TO_INCLUDE}
unlabeled_issues = []

for issue in all_issues:
    matching_labels = set(issue["labels"]).intersection(LABELS_TO_INCLUDE)
    if matching_labels:
        for label in matching_labels:
            grouped_issues[label].append({
                "number": issue["number"],
                "title": issue["title"],
                "state": issue["state"],
                "created_at": issue["created_at"],
                "closed_at": issue["closed_at"],
                "labels": ", ".join(issue["labels"])
            })
    else:
        unlabeled_issues.append({
            "number": issue["number"],
            "title": issue["title"],
            "state": issue["state"],
            "created_at": issue["created_at"],
            "closed_at": issue["closed_at"],
            "labels": ", ".join(issue["labels"])
        })

# Write each group to a CSV
for label, issues in grouped_issues.items():
    df = pd.DataFrame(issues)
    safe_label = label.replace(" ", "_").replace(":", "")
    df.to_csv(f"{safe_label}_issues.csv", index=False)
    print(f"Saved {len(issues)} issues to {safe_label}_issues.csv")

# Write unlabeled or non-target-labeled issues to separate file
if unlabeled_issues:
    df_unlabeled = pd.DataFrame(unlabeled_issues)
    df_unlabeled.to_csv("other_issues.csv", index=False)
    print(f"Saved {len(unlabeled_issues)} issues to other_issues.csv")
