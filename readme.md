# GitLab Excel Issues

Generate GitLab issues from an Excel document. This Python command-line utility reads issues from an Excel file and creates corresponding GitLab issues. It supports specifying various fields, such as labels, tags, due dates, and descriptions, using column names from the Excel sheet.

## Installation

### Prerequisites

- Python 3.8 or higher
- [Poetry](https://python-poetry.org/)


### Steps

1. Install dependencies using Poetry:
2. Create a `.env` file in the project root with your GitLab credentials:
	```
	GITLAB_URL=https://your.gitlab.instance.com
	GITLAB_API_KEY=your_private_token
	GITLAB_PROJECT_ID=your_project_id
	```

Replace the values with your own GitLab instance URL, private token, and project ID.


## Command-line Arguments

- `--xls [FILE]`: Path to the input XLS file.
- `--sheet [NAME]`: Name of the active worksheet.
- `--list`: List column items (writes to console).
- `--gitlab-tag [TAG]`: Apply tag to every generated issue.
- `--gitlab-issue-label [COLNAME]`: Column name for issue label.
- `--gitlab-issue-tag [COLNAME]`: (Optional) Column name for issue tag.
- `--gitlab-issue-due [COLNAME]`: (Optional) Column name for issue due date.
- `--gitlab-issue-description [COLNAME]`: (Optional) Column name for issue description.
- `--gitlab-dryrun`: (Optional) Output issues without submitting to GitLab.

## Example Usage

```bash
# List columns in the Excel file
python gitlab_excel_issues.py --xls issues.xlsx --sheet Sheet1 --list

# Create GitLab issues from the Excel file
python gitlab_excel_issues.py --xls issues.xlsx --sheet Sheet1 --gitlab-issue-label Title --gitlab-issue-tag Tag --gitlab-issue-due DueDate --gitlab-issue-description Description

# Perform a dry run without submitting issues to GitLab
python gitlab_excel_issues.py --xls issues.xlsx --sheet Sheet1 --gitlab-issue-label Title --gitlab-issue-tag Tag --gitlab-issue-due DueDate --gitlab-issue-description Description --gitlab-dryrun
```

