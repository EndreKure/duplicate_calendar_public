# Calendar Duplicator
## Purpose
The Calendar Duplicator tool is designed to create duplicate appointments in calendars to block out time. It consists of several files and provides flexibility through configuration options.

## Files
* `config.toml`: Configuration file for specifying parameters.
* `requirements.in/requirements.txt`: Lists required Python packages.
* `environment.yml`: Conda environment setup file.
* `update_calendar.py`: Main script for updating calendars.

Note: This tool currently supports only non-recurring meetings.

## Getting Started
### Requirements
* Python 3.8 or higher is required.
* It is recommended to use Anaconda for managing environments.

### Setting up the Environment
Create the Conda environment by running the following command from the project folder:

```conda env create -f environment.yml ```

Activate the created environment:

```conda activate cal_duplicator```

Install the required Python packages:

```
python -m piptools compile requirements.in
pip install -r requirements.txt
```

### How to run the script
Open a command prompt window and navigate to the project folder.
Execute the following command to run the script:

```python update_calendar.py```

The script will perform the following actions:

* Delete all previously created placeholders (meetings with names matching `subject_delete`).
* Create new placeholders (using the format specified in `subject`) in the chosen calendars.

### Configuration Options
You can customize the tool's behavior by editing the `config.toml` file. Here are the available options:

* `target_emails`: List of calendars to be updated, specified as ["zz@xx.no", "zz@yy.no"].
* `subject`: Subject for meetings to be added to the calendar, along with the calendar name (e.g., f"||{subject}|{{calendar_name}||").
* `subject_delete`: Subject for meetings with the calendar name to be deleted, specified as f"||{subject}|{calendar_name}||".
* `body`: Default value is "Placeholder for meeting in other calendar".
* `start_time`: Format (YYYY-MM-DD). Default value is today's date.
* `duration_days`: Number of days in the horizon (integer). Default value is 7 (one week).
* `just_delete_placeholders`: Set to true to delete all previously created placeholders without updating with new meetings.