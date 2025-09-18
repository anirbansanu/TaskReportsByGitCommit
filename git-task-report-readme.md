# Git Task Management Report Generator

A Python script that generates task management Excel reports by analyzing Git commit history from one or more project repositories.

## Features

- **Dual Input Modes**: Command-line arguments or JSON configuration file
- **Multiple Repository Support**: Process multiple Git repositories simultaneously
- **Flexible Date Filtering**: Filter commits by date range and author
- **Smart Commit Merging**: Merge commits by date with line break concatenation
- **Excel Report Generation**: Professional formatted Excel reports with multiple sheets
- **Repository Statistics**: Detailed analysis and statistics for each repository
- **Error Handling**: Robust validation and error reporting
- **Progress Tracking**: Real-time progress updates during processing

## Installation

### Prerequisites

- Python 3.6 or higher
- Git installed and accessible from command line

### Required Python packages

```bash
pip install pandas openpyxl GitPython
```

## Usage

### Method 1: Command Line Arguments

```bash
# Basic usage
python git_task_report.py --repos /path/to/repo1 /path/to/repo2 --author "Anirban" --since 2025-07-01 --until 2025-09-15

# With custom filename
python git_task_report.py --repos ./project1 ./project2 --author "Anirban" --since 2025-07-01 --until 2025-09-15 --filename "custom_report.xlsx"

# With separate sheets for each repository
python git_task_report.py --repos ./project1 ./project2 --author "Anirban" --since 2025-07-01 --until 2025-09-15 --separate-sheets

# Process current directory only
python git_task_report.py --repos ./ --since 2025-08-01 --until 2025-09-18
```

### Method 2: JSON Configuration File

If no command-line arguments are provided, the script will look for `git_task_config.json` in the current directory.

**Sample configuration file:**

```json
{
  "repos": [
    "./project1", 
    "./project2",
    "/path/to/another/repo"
  ],
  "author": "Anirban",
  "since": "2025-07-01",
  "until": "2025-09-15",
  "filename": "task_report_custom.xlsx",
  "separate_sheets": true
}
```

**Run with config file:**

```bash
python git_task_report.py
```

## Command Line Arguments

| Argument | Description | Example |
|----------|-------------|---------|
| `--repos` | One or more repository paths | `--repos ./repo1 ./repo2` |
| `--author` | Filter commits by author name | `--author "Anirban"` |
| `--since` | Start date (YYYY-MM-DD format) | `--since 2025-07-01` |
| `--until` | End date (YYYY-MM-DD format) | `--until 2025-09-15` |
| `--filename` | Custom output filename | `--filename "my_report.xlsx"` |
| `--separate-sheets` | Create separate sheet per repo | `--separate-sheets` |

## Output Format

### Excel Report Structure

The generated Excel file contains:

#### Task Sheets
Each task sheet includes the following columns:

| Column | Description | Example |
|--------|-------------|---------|
| **Task Name** | Commit message(s) | `Fix user authentication bug` |
| **Task Priority** | Empty (for manual input) | ` ` |
| **Assign Date** | One day before commit date | `14-09-2025` |
| **Due Date** | Empty (for manual input) | ` ` |
| **Planned End Date** | One day before commit date | `14-09-2025` |
| **Actual End Date** | Commit date | `15-09-2025` |
| **Assignee** | Fixed value | `Arvind Sir` |

#### Summary Sheet
Contains detailed statistics for each repository:

- Total commits processed
- Date range covered
- Days with commits
- Average commits per day
- Maximum commits in a single day
- Number of unique authors

### Commit Processing Rules

1. **Merge Commit Filtering**: Commits starting with "Merge" are automatically excluded
2. **Date-based Merging**: Multiple commits on the same date are merged into a single task row
3. **Message Concatenation**: Multiple commit messages are joined with line breaks (not semicolons)
4. **Date Formatting**: All dates use DD-MM-YYYY format
5. **Sorting**: Tasks are sorted by Actual End Date in ascending order

## File Naming Convention

If no custom filename is provided:
- `task_report_YYYYMMDD_to_YYYYMMDD.xlsx` (when date range specified)
- `task_report_YYYYMMDD.xlsx` (when no date range specified)

## Examples

### Example 1: Single Repository Analysis

```bash
python git_task_report.py --repos ./my-project --author "Anirban" --since 2025-09-01 --until 2025-09-18
```

Output: `task_report_20250901_to_20250918.xlsx`

### Example 2: Multiple Repositories with Separate Sheets

```bash
python git_task_report.py --repos ./frontend ./backend ./api --separate-sheets --filename "project_analysis.xlsx"
```

Output: `project_analysis.xlsx` with sheets: `Tasks_frontend`, `Tasks_backend`, `Tasks_api`, `Summary`

### Example 3: Using Configuration File

Create `git_task_config.json`:
```json
{
  "repos": ["./web-app", "./mobile-app"],
  "author": "Anirban",
  "since": "2025-08-01",
  "until": "2025-09-18",
  "filename": "august_september_report.xlsx",
  "separate_sheets": false
}
```

Run:
```bash
python git_task_report.py
```

## Error Handling

The script handles various error conditions:

- **Invalid repository paths**: Warns and skips non-existent or non-Git directories
- **Missing Git**: Checks for Git installation and provides helpful error messages
- **Invalid date formats**: Validates date format and provides correction guidance
- **Permission issues**: Handles file access and permission errors gracefully
- **Missing dependencies**: Checks for required Python packages on startup

## Troubleshooting

### Common Issues

**1. "Not a git repository" error**
- Ensure the specified path contains a `.git` directory
- Verify you're pointing to the repository root, not a subdirectory

**2. "No commits found" message**
- Check date range (commits might be outside specified dates)
- Verify author name spelling and capitalization
- Ensure the repository has commits from the specified author

**3. Package import errors**
- Install missing packages: `pip install pandas openpyxl`
- For older Python versions, you might need: `pip install GitPython`

**4. Permission denied when saving Excel file**
- Ensure the output file isn't open in Excel
- Check write permissions in the target directory
- Try a different filename or directory

### Performance Considerations

- Processing large repositories (10,000+ commits) may take several minutes
- Multiple repositories are processed sequentially with progress updates
- Excel file generation time depends on the number of tasks (typically under 30 seconds)

## Advanced Usage

### Filtering Tips

- **Author names**: Use exact spelling as it appears in Git commits
- **Date ranges**: Use YYYY-MM-DD format; both dates are inclusive
- **Repository paths**: Can be absolute or relative paths

### Integration Ideas

- **CI/CD Integration**: Run weekly/monthly to generate team reports
- **Project Management**: Import Excel data into project management tools
- **Performance Tracking**: Analyze commit patterns and productivity trends
- **Team Reporting**: Generate individual or team-based task summaries

## Contributing

Feel free to enhance the script with additional features:
- Support for different date formats
- Additional Excel formatting options
- Integration with project management APIs
- Custom commit message filtering rules

## Version History

- **v1.0**: Initial release with core functionality
- **v1.1**: Added JSON configuration support and repository statistics
- **v1.2**: Enhanced error handling and progress reporting

## License

This script is provided as-is for educational and professional use.