import win32com.client
import pandas as pd
import numpy as np
import datetime as dt
import re
import toml


def get_calendar_and_accounts(emails):
    """ Retrieves an updated view on the Outlook calendar"""

    # Create an Outlook application instance
    outlook = win32com.client.Dispatch("Outlook.Application")

    # Get the MAPI namespace
    namespace = outlook.GetNamespace("MAPI")

    # Get the collection of accounts configured in Outlook
    accounts = namespace.Folders
    accounts = [account for account in accounts if account.name in emails]

    calendar = get_meetings_for_all_accounts(accounts)
    return outlook, namespace, accounts, calendar


def get_meetings_for_all_accounts(accounts):
    """ Retrieve all meetings in all the Outlook calendar as a pandas DataFrame"""
    try:
        # Loop through each account
        output = pd.DataFrame()
        for account in accounts:
            print("account", account)
            # Access the Calendar folder for the current account
            calendar_folder = account.Folders("Calendar")

            # Now, you can work with the Calendar folder object for this account
            # For example, you can loop through the items in the calendar for this account
            for calendar_item in calendar_folder.Items:
                subject = calendar_item.Subject
                start = calendar_item.start
                end = calendar_item.end
                duration = calendar_item.duration

                temp_dict = {
                    "account": account,
                    "subject": subject,
                    "start": f"{start}",
                    "end": f"{end}"
                }
                df_dictionary = pd.DataFrame([temp_dict])
                output = pd.concat([output, df_dictionary], ignore_index=True)

                # Do something with each calendar item
                # print(calendar_item.Subject)

        calendar = output.copy()
        calendar["account_name"] = calendar["account"].apply(lambda x: x.name)
        for y in ["start", "end"]:
            calendar[f"{y}_time"] = calendar[f"{y}"].apply(
                lambda x: np.datetime64(dt.datetime.strptime(x, "%Y-%m-%d %H:%M:%S%z")))
        return calendar

    except Exception as e:
        print(f"An error occurred: {str(e)}")


def create_shadow_meetings(calendar, start_time, duration_days, account_name):
    """ Retrieve all meetings that should have a shadow meeting for a give account for a given timespan"""

    # Copy calendar instance
    shadow_meetings = calendar.copy()

    # Filter out only relevant time
    end_time = start_time + dt.timedelta(days=duration_days)
    selector1 = (shadow_meetings.account_name != account_name)
    selector2 = (shadow_meetings.start_time >= start_time)
    selector3 = (shadow_meetings.end_time <= end_time)
    shadow_meetings = shadow_meetings[selector1 & selector2 & selector3]

    return shadow_meetings


def create_new_meeting(target_account_email, accounts, start, end, subject, body):
    """ Creates a meeting instance """
    # Loop through each account
    for account in accounts:
        print("account", account)
        # Access the Calendar folder for the current account

        if account.name == target_account_email:
            calendar_folder = account.Folders("Calendar")

            appointment = calendar_folder.Items.Add()
            appointment.Start = start  # yyyy-MM-dd hh:mm
            appointment.End = end
            appointment.Subject = subject
            appointment.Body = body
            appointment.MeetingStatus = 1  # Indicates that this is a meeting in the respective outlook calendar
            appointment.Save()
            print("Meeting is created")


def update_calendars(target_emails, start_time, duration_days, body, sub_subject):
    """ updates the calendar with meetings"""
    outlook, namespace, accounts, calendar = get_calendar_and_accounts(target_emails)

    for target_account_email in target_emails:
        shadow_meetings = create_shadow_meetings(calendar, start_time, duration_days=duration_days,
                                                  account_name=target_account_email)

        for index, row in shadow_meetings.iterrows():
            cal_name = re.search(r'@(\w+)', row["account_name"]).group(1)
            subject = f"||{sub_subject}|{cal_name}||"
            print(subject)
            print(row["start_time"])
            create_new_meeting(target_account_email, accounts, row["start_time"].strftime("%Y-%m-%d %H:%M"),
                               row["end_time"].strftime("%Y-%m-%d %H:%M"), subject, body)


def delete_meetings(accounts, target_account_email, subject):
    """ Deletes all shadow meetings previously created.
    Returns FALSE if no meeting was deleted, TRUE if one or more meeting was deleted """

    meeting_deleted = False
    # Loop through each account
    for account in accounts:
        print("account", account)
        # Access the Calendar folder for the current account

        if account.name == target_account_email:
            calendar_folder = account.Folders("Calendar")

            for item in calendar_folder.Items:
                if item.MeetingStatus == 1 and item.Subject == subject:
                    print(f"Found meeting to delete: Subject - {item.Subject}, Start Time - {item.Start}")
                    item.Delete()  # Delete the meeting
                    meeting_deleted = True
                    print("Meeting deleted.")
    return meeting_deleted


def clean_calendar_of_old_shadow_meetings(target_emails, subject_delete):
    """ Remove all old shadow meetings"""

    # Updated calendar
    outlook, namespace, accounts, calendar = get_calendar_and_accounts(target_emails)

    # Iterate calendar
    meeting_deleted = True
    counter = 0  # just added as a fail-safe
    while (meeting_deleted and counter < 100):
        meeting_deleted = False
        for target_account_email in target_emails:
            emails_copy = target_emails.copy()
            emails_copy.remove(target_account_email)
            print(emails_copy)
            for cal_name in [re.search(r'@(\w+)', email).group(1) for email in emails_copy]:
                print(f"------{target_account_email}-----")
                subject = f"||{subject_delete}|{cal_name}||"
                print(subject)
                meeting_deleted = delete_meetings(accounts, target_account_email, subject=subject) or meeting_deleted
        counter += 1
        print(counter, meeting_deleted)


if __name__ == "__main__":
    config = toml.load("config.toml")

    # config values
    target_emails = config["target_emails"]
    subject = config["subject"]
    subject_delete = config["subject_delete"]
    body = config["body"]
    start_time = config["start_time"]
    duration_days = config["duration_days"]
    just_delete_placeholders = config["just_delete_placeholders"]
    print(target_emails, body, start_time, duration_days, subject_delete, subject, just_delete_placeholders)

    # Default value if values are not given
    if "".__eq__(start_time):
        start_time = dt.datetime.now().strftime("%Y-%m-%d")
    # Convert star_time to date_time object
    start_time = dt.datetime.strptime(start_time, "%Y-%m-%d")

    if "".__eq__(f"{duration_days}"):
        duration_days = 7

    if "".__eq__(body):
        body = "Placeholder for meeting in other calendar"

    print(target_emails, body, start_time, duration_days, subject, subject_delete, subject, just_delete_placeholders)
    try:
        # clear out old shadow agreements
        clean_calendar_of_old_shadow_meetings(target_emails, subject_delete)
    except Exception as e:
        print(f"An error occurred when deleting appointments: {str(e)}")
    if not just_delete_placeholders:
        try:
            # updates calendars with meetings
            update_calendars(target_emails, start_time, duration_days, body, subject)
        except Exception as e:
            print(f"An error occurred when creating appointments: {str(e)}")
