import pandas as pd
from datetime import datetime, timedelta
from pytz import timezone
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pathlib import Path

# -------------------------------------------------------------------
# CONFIGURATION
# -------------------------------------------------------------------
PROJECT_ROOT = Path(__file__).resolve().parent
FILE_PATH = PROJECT_ROOT / "license_data" / "licenses.xlsx"

frequency_mapping = {
    "6 MONTH": timedelta(days=180),
    "1 YEAR": timedelta(days=365),
    "2 YEAR": timedelta(days=730),
}

SENDER_EMAIL = ""
SENDER_PASSWORD = ""
SMTP_SERVER = ""
SMTP_PORT = 465

# -------------------------------------------------------------------
# EMAIL SENDER
# -------------------------------------------------------------------
def send_email(subject, body, to_emails, cc_emails=None):
    msg = MIMEMultipart()
    msg["From"] = SENDER_EMAIL
    msg["To"] = ", ".join(to_emails)
    msg["Subject"] = subject
    if cc_emails:
        msg["Cc"] = ", ".join(cc_emails)

    msg.attach(MIMEText(body, "html"))
    recipients = to_emails + (cc_emails if cc_emails else [])

    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.sendmail(SENDER_EMAIL, recipients, msg.as_string())

    print(f"✅ Email sent to: {', '.join(recipients)}")

# -------------------------------------------------------------------
# MAIN FUNCTION
# -------------------------------------------------------------------
def update_and_notify():

    if not FILE_PATH.exists():
        print(f"❌ Excel file not found: {FILE_PATH}")
        return

    # 🔐 Read everything as TEXT (CRITICAL)
    df = pd.read_excel(FILE_PATH, dtype=str)
    df.columns = df.columns.str.strip().str.upper()

    print("📋 Columns:", list(df.columns))

    today = datetime.now(timezone("Asia/Kolkata")).date()
    print(f"🕒 Today: {today}")

    # -------------------------------------------------------------------
    # HELPER DATE COLUMNS (LOGIC ONLY — NEVER SAVED)
    # -------------------------------------------------------------------
    df["_VALIDITY_DATE"] = pd.to_datetime(df["VALIDITY"], errors="coerce")

    if "REMINDER" in df.columns:
        df["_REMINDER_DATE"] = pd.to_datetime(df["REMINDER"], errors="coerce")
    else:
        df["_REMINDER_DATE"] = pd.NaT

    # -------------------------------------------------------------------
    # STEP 1: SEND REMINDERS (READ-ONLY)
    # -------------------------------------------------------------------
    remind_indices = []

    for i, row in df.iterrows():
        reminder_date = row["_REMINDER_DATE"]

        if pd.isna(reminder_date):
            continue

        days_left = (reminder_date.date() - today).days

        if days_left in (0, 1, 4, 5):
            remind_indices.append(i)

    if remind_indices:
        reminder_rows = df.loc[remind_indices].copy()

        # format helper columns for email view
        for col in reminder_rows.columns:
            if col.startswith("_"):
                reminder_rows.drop(columns=col, inplace=True)

        html_table = reminder_rows.to_html(index=False)

        body = f"""
        <html>
        <body>
            <p>Dear Sir,</p>
            <p>The following licenses/certificates require attention:</p>
            {html_table}
            <p>Please ensure timely renewal.</p>
            <p>Regards,<br>Sharada Industries</p>
        </body>
        </html>
        """

        send_email(
            "License / Certificate Reminder - Sharada Industries",
            body,
            ["pghadge@algoanalytics.com"]
        )
    else:
        print("✅ No reminders today.")

    # -------------------------------------------------------------------
    # STEP 2: UPDATE EXPIRED VALIDITY (ONLY VIA FREQUENCY)
    # -------------------------------------------------------------------
    updated = False

    if "FREQUENCY" in df.columns:
        for i, row in df.iterrows():

            validity_date = row["_VALIDITY_DATE"]
            freq_raw = row.get("FREQUENCY")

            # validity must be real date
            if pd.isna(validity_date):
                continue

            # frequency must exist
            if not freq_raw:
                continue

            freq = freq_raw.strip().upper()
            freq = freq.replace("MONTHS", "MONTH").replace("YEARS", "YEAR")

            if freq not in frequency_mapping:
                continue

            # must be expired
            if validity_date.date() >= today:
                continue

            new_validity = validity_date + frequency_mapping[freq]
            new_reminder = new_validity - timedelta(days=15)

            # 🔥 ONLY PLACE WHERE CHANGES HAPPEN
            df.at[i, "VALIDITY"] = new_validity.strftime("%d-%m-%Y")
            if "REMINDER" in df.columns:
                df.at[i, "REMINDER"] = new_reminder.strftime("%d-%m-%Y")

            updated = True
            print(f"🔁 Row {i} renewed → VALIDITY {new_validity.date()}")

    # -------------------------------------------------------------------
    # CLEANUP HELPER COLUMNS
    # -------------------------------------------------------------------
    df.drop(columns=[c for c in df.columns if c.startswith("_")], inplace=True)

    # -------------------------------------------------------------------
    # SAVE FILES
    # -------------------------------------------------------------------
    if updated:
        try:
            df.to_excel(FILE_PATH, index=False)
            print("✅ Excel updated successfully.")
        except PermissionError:
            print("⚠️ Please close Excel file and retry.")

    df.to_csv("licenses.csv", index=False)
    print("🏁 Process complete.")

# -------------------------------------------------------------------
if __name__ == "__main__":
    update_and_notify()
