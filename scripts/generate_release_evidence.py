import os
import logging
import requests

from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
# -----------------------------------------------------------------------------
# Logging
# -----------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)

logger = logging.getLogger(__name__)

# -----------------------------------------------------------------------------
# GitHub Configuration
# -----------------------------------------------------------------------------
GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")
GITHUB_API_URL = os.getenv("GITHUB_API_URL", "https://api.github.com")
GITHUB_SERVER_URL = os.getenv("GITHUB_SERVER_URL", "https://github.com")

GITHUB_REPOSITORY = os.getenv("GITHUB_REPOSITORY")
GITHUB_RUN_ID = os.getenv("GITHUB_RUN_ID")
GITHUB_SHA = os.getenv("GITHUB_SHA")
GITHUB_REF_NAME = os.getenv("GITHUB_REF_NAME")
GITHUB_ACTOR = os.getenv("GITHUB_ACTOR")
GITHUB_WORKFLOW = os.getenv("GITHUB_WORKFLOW")
GITHUB_RUN_NUMBER = os.getenv("GITHUB_RUN_NUMBER")

# -----------------------------------------------------------------------------
# GitHub REST API Helper
# -----------------------------------------------------------------------------
def github_get(endpoint: str, params: dict | None = None) -> dict:
    """
    Executes a GET request against the GitHub REST API.
    """

    url = f"{GITHUB_API_URL}{endpoint}"

    headers = {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "ReleaseEvidenceGenerator",
    }

    logger.info("GET %s", url)

    response = requests.get(
        url=url,
        headers=headers,
        params=params,
        timeout=30,
    )

    response.raise_for_status()

    return response.json()
# Create reports directory
os.makedirs("reports", exist_ok=True)

wb = Workbook()

ws = wb.active
ws.title = "Release Summary"

ws["A1"] = "Release Evidence Report"

ws["A3"] = "Repository"
ws["B3"] = os.getenv("GITHUB_REPOSITORY")

ws["A4"] = "Branch"
ws["B4"] = os.getenv("GITHUB_REF_NAME")

ws["A5"] = "Commit SHA"
ws["B5"] = os.getenv("GITHUB_SHA")

ws["A6"] = "Workflow Run ID"
ws["B6"] = os.getenv("GITHUB_RUN_ID")

ws["A7"] = "Triggered By"
ws["B7"] = os.getenv("GITHUB_ACTOR")

ws["A8"] = "Generated On"
ws["B8"] = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")

# -----------------------------------------------------------------------------
# Repository Information
# -----------------------------------------------------------------------------
try:
    logger.info("Retrieving repository information...")

    repository = github_get(f"/repos/{GITHUB_REPOSITORY}")

    ws["A10"] = "Repository Name"
    ws["B10"] = repository.get("name")

    ws["A11"] = "Repository Owner"
    ws["B11"] = repository.get("owner", {}).get("login")

    ws["A12"] = "Repository Visibility"
    ws["B12"] = repository.get("visibility")

    ws["A13"] = "Default Branch"
    ws["B13"] = repository.get("default_branch")

    ws["A14"] = "Repository URL"
    ws["B14"] = repository.get("html_url")

    ws["A15"] = "Clone URL"
    ws["B15"] = repository.get("clone_url")

    ws["A16"] = "Created At"
    ws["B16"] = repository.get("created_at")

    ws["A17"] = "Last Updated"
    ws["B17"] = repository.get("updated_at")

    logger.info("Repository information collected successfully.")

except requests.exceptions.RequestException as ex:
    logger.error("Failed to retrieve repository information: %s", ex)

# -----------------------------------------------------------------------------
# Workflow Information
# -----------------------------------------------------------------------------
try:
    logger.info("Retrieving workflow run information...")

    workflow_run = github_get(
        f"/repos/{GITHUB_REPOSITORY}/actions/runs/{GITHUB_RUN_ID}"
    )

    ws["A19"] = "Workflow Name"
    ws["B19"] = workflow_run.get("name")

    ws["A20"] = "Run Number"
    ws["B20"] = workflow_run.get("run_number")

    ws["A21"] = "Run Attempt"
    ws["B21"] = workflow_run.get("run_attempt")

    ws["A22"] = "Workflow Event"
    ws["B22"] = workflow_run.get("event")

    ws["A23"] = "Workflow Status"
    ws["B23"] = workflow_run.get("status")

    ws["A24"] = "Workflow Conclusion"
    ws["B24"] = workflow_run.get("conclusion")

    ws["A25"] = "Workflow Created"
    ws["B25"] = workflow_run.get("created_at")

    ws["A26"] = "Workflow Updated"
    ws["B26"] = workflow_run.get("updated_at")

    ws["A27"] = "Workflow URL"
    ws["B27"] = workflow_run.get("html_url")

    logger.info("Workflow information collected successfully.")

except requests.exceptions.RequestException as ex:
    logger.error("Failed to retrieve workflow information: %s", ex)

# -----------------------------------------------------------------------------
# Commit Information
# -----------------------------------------------------------------------------
try:
    logger.info("Retrieving commit information...")

    commit = github_get(
        f"/repos/{GITHUB_REPOSITORY}/commits/{GITHUB_SHA}"
    )

    commit_info = commit.get("commit", {})
    author = commit_info.get("author", {})
    committer = commit_info.get("committer", {})

    ws["A29"] = "Commit Message"
    ws["B29"] = commit_info.get("message")

    ws["A30"] = "Author"
    ws["B30"] = author.get("name")

    ws["A31"] = "Author Email"
    ws["B31"] = author.get("email")

    ws["A32"] = "Author Date"
    ws["B32"] = author.get("date")

    ws["A33"] = "Committer"
    ws["B33"] = committer.get("name")

    ws["A34"] = "Committer Email"
    ws["B34"] = committer.get("email")

    ws["A35"] = "Commit Date"
    ws["B35"] = committer.get("date")

    ws["A36"] = "Commit URL"
    ws["B36"] = commit.get("html_url")

    logger.info("Commit information collected successfully.")

except requests.exceptions.RequestException as ex:
    logger.error("Failed to retrieve commit information: %s", ex)

# -----------------------------------------------------------------------------
# Build Information
# -----------------------------------------------------------------------------
try:
    logger.info("Collecting build information...")

    workflow_run = github_get(
        f"/repos/{GITHUB_REPOSITORY}/actions/runs/{GITHUB_RUN_ID}"
    )

    ws["A38"] = "Build Information"
    ws["B38"] = ""

    ws["A39"] = "Build ID"
    ws["B39"] = workflow_run.get("id")

    ws["A40"] = "Build Number"
    ws["B40"] = workflow_run.get("run_number")

    ws["A41"] = "Build Attempt"
    ws["B41"] = workflow_run.get("run_attempt")

    ws["A42"] = "Head Branch"
    ws["B42"] = workflow_run.get("head_branch")

    ws["A43"] = "Head SHA"
    ws["B43"] = workflow_run.get("head_sha")

    ws["A44"] = "Event"
    ws["B44"] = workflow_run.get("event")

    ws["A45"] = "Status"
    ws["B45"] = workflow_run.get("status")

    ws["A46"] = "Conclusion"
    ws["B46"] = workflow_run.get("conclusion")

    ws["A47"] = "Created At"
    ws["B47"] = workflow_run.get("created_at")

    ws["A48"] = "Updated At"
    ws["B48"] = workflow_run.get("updated_at")

    ws["A49"] = "HTML URL"
    ws["B49"] = workflow_run.get("html_url")

    ws["A50"] = "Actor"
    ws["B50"] = workflow_run.get("actor", {}).get("login")

    logger.info("Build information collected successfully.")

except requests.exceptions.RequestException as ex:
    logger.error("Failed to retrieve build information: %s", ex)

# -----------------------------------------------------------------------------
# Pull Request Information
# -----------------------------------------------------------------------------
try:
    logger.info("Retrieving pull request information...")

    pull_requests = github_get(
        f"/repos/{GITHUB_REPOSITORY}/commits/{GITHUB_SHA}/pulls"
    )

    if pull_requests:

        pull_request = pull_requests[0]

        ws["A52"] = "Pull Request Information"
        ws["B52"] = ""

        ws["A53"] = "Pull Request Number"
        ws["B53"] = pull_request.get("number")

        ws["A54"] = "Pull Request Title"
        ws["B54"] = pull_request.get("title")

        ws["A55"] = "Pull Request State"
        ws["B55"] = pull_request.get("state")

        ws["A56"] = "Merged"
        ws["B56"] = pull_request.get("merged")

        ws["A57"] = "Created By"
        ws["B57"] = pull_request.get("user", {}).get("login")

        ws["A58"] = "Created At"
        ws["B58"] = pull_request.get("created_at")

        ws["A59"] = "Merged At"
        ws["B59"] = pull_request.get("merged_at")

        ws["A60"] = "Base Branch"
        ws["B60"] = pull_request.get("base", {}).get("ref")

        ws["A61"] = "Source Branch"
        ws["B61"] = pull_request.get("head", {}).get("ref")

        ws["A62"] = "Pull Request URL"
        ws["B62"] = pull_request.get("html_url")

        logger.info("Pull request information collected successfully.")

    else:

        ws["A52"] = "Pull Request Information"
        ws["B52"] = "No Pull Request associated with this commit."

        logger.info("No pull request found for commit.")

except requests.exceptions.RequestException as ex:
    logger.error("Failed to retrieve pull request information: %s", ex)

# -----------------------------------------------------------------------------
# Changed Files
# -----------------------------------------------------------------------------
try:
    logger.info("Retrieving changed files...")

    commit = github_get(
        f"/repos/{GITHUB_REPOSITORY}/commits/{GITHUB_SHA}"
    )

    files = commit.get("files", [])

    start_row = 64

    ws[f"A{start_row}"] = "Changed Files"

    header_row = start_row + 1

    ws[f"A{header_row}"] = "File Name"
    ws[f"B{header_row}"] = "Status"
    ws[f"C{header_row}"] = "Additions"
    ws[f"D{header_row}"] = "Deletions"
    ws[f"E{header_row}"] = "Changes"

    current_row = header_row + 1

    if files:

        for file in files:

            ws[f"A{current_row}"] = file.get("filename")
            ws[f"B{current_row}"] = file.get("status")
            ws[f"C{current_row}"] = file.get("additions")
            ws[f"D{current_row}"] = file.get("deletions")
            ws[f"E{current_row}"] = file.get("changes")

            current_row += 1

        logger.info("Changed files collected successfully.")

    else:

        ws[f"A{current_row}"] = "No changed files found."

except requests.exceptions.RequestException as ex:
    logger.error("Failed to retrieve changed files: %s", ex)
output = "reports/ReleaseEvidence.xlsx"

wb.save(output)

logger.info("Evidence report generated: %s", output)
