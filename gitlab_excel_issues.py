"""
Generate gitlab issues for tabular excel sheets
"""
import sys
import argparse
import pandas as pd
import gitlab
import os
from dotenv import load_dotenv


class GitLabConfig:
    def __init__(self, url, api_key, project_id):
        self.url = url
        self.api_key = api_key
        self.project_id = project_id


def print_progress_bar(
    iteration, total, prefix="", suffix="", decimals=1, length=50, fill="█"
):
    """
    Call in a loop to create a terminal progress bar
    @params:
        iteration   - Required : current iteration (Int)
        total       - Required : total iterations (Int)
        prefix      - Optional : prefix string (Str)
        suffix      - Optional : suffix string (Str)
        decimals    - Optional : number of decimals in percent complete (Int)
        length      - Optional : character length of the bar (Int)
        fill        - Optional : bar fill character (Str)
    """
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filled_length = int(length * iteration // total)
    bar = fill * filled_length + "-" * (length - filled_length)
    sys.stdout.write("\r%s |%s| %s%% %s" % (prefix, bar, percent, suffix))
    sys.stdout.flush()

    if iteration == total:
        print()


def parse_args():
    parser = argparse.ArgumentParser(
        description="Generate GitLab issues from an Excel document."
    )
    parser.add_argument(
        "--xls", type=str, required=True, help="Path to the input XLS file."
    )
    parser.add_argument(
        "--sheet", type=str, required=True, help="Name of the active worksheet."
    )
    parser.add_argument(
        "--list", action="store_true", help="List column items (writes to console)."
    )
    parser.add_argument(
        "--gitlab-issue-label",
        type=str,
        required=True,
        help="Column name for issue label.",
    )
    parser.add_argument(
        "--gitlab-issue-tag", type=str, help="(Optional) Column name for issue tag."
    )
    parser.add_argument(
        "--gitlab-issue-due",
        type=str,
        help="(Optional) Column name for issue due date.",
    )
    parser.add_argument(
        "--gitlab-issue-description",
        type=str,
        help="(Optional) Column name for issue description.",
    )
    parser.add_argument(
        "--gitlab-dryrun",
        action="store_true",
        help="(Optional) Output issues without submitting to GitLab.",
    )

    parser.add_argument(
        "--gitlab-tag",
        action="append",
        help="(Optional) Custom tag to apply to all generated issues.",
    )

    return parser.parse_args()


def get_gitlab_config():
    load_dotenv()
    url = os.getenv("GITLAB_URL")
    api_key = os.getenv("GITLAB_API_KEY")
    project_id = os.getenv("GITLAB_PROJECT_ID")

    return GitLabConfig(url, api_key, project_id)


def create_gitlab_issues(df, config, issue_options, custom_tags=None, dryrun=False):
    gl = gitlab.Gitlab(config.url, private_token=config.api_key)
    project = None if dryrun else gl.projects.get(config.project_id)

    total_rows = len(df)
    for index, row in df.iterrows():
        issue_data = {
            "title": row[issue_options["label"]],
            "labels": row[issue_options["tag"]] if issue_options.get("tag") else None,
            "due_date": row[issue_options["due"]] if issue_options.get("due") else None,
            "description": row[issue_options["description"]]
            if issue_options.get("description")
            else None,
        }

        if custom_tags:
            if issue_data["labels"]:
                issue_data["labels"] += f",{','.join(custom_tags)}"
            else:
                issue_data["labels"] = ",".join(custom_tags)

        # Remove keys with None values
        issue_data = {k: v for k, v in issue_data.items() if v is not None}

        if dryrun:
            print("Dry run issue data:")
            print(issue_data)
        else:
            if index < 1:
                print("Generating GitLab issues:")

            project.issues.create(issue_data)
            print_progress_bar(
                index + 1, total_rows, prefix="Progress:", suffix="Complete", length=50
            )

    if not dryrun:
        issues_url = f"{config.url}/{project.path_with_namespace}/-/issues"
        print(f"\nIssues list URL: {issues_url}")


def read_excel_file(xls_path, sheet_name):
    df = pd.read_excel(xls_path, sheet_name=sheet_name)
    return df


def list_columns(df):
    print("Columns in the DataFrame:")
    for col in df.columns:
        print(col)


def main():
    args = parse_args()

    xls_path = args.xls
    sheet_name = args.sheet
    list_columns_only = args.list

    df = read_excel_file(xls_path, sheet_name)

    if list_columns_only:
        list_columns(df)
    else:
        config = get_gitlab_config()
        issue_options = {
            "label": args.gitlab_issue_label,
            "tag": args.gitlab_issue_tag,
            "due": args.gitlab_issue_due,
            "description": args.gitlab_issue_description,
        }
        create_gitlab_issues(
            df,
            config,
            issue_options=issue_options,
            custom_tags=args.gitlab_tag,
            dryrun=args.gitlab_dryrun,
        )


if __name__ == "__main__":
    sys.exit(main())
