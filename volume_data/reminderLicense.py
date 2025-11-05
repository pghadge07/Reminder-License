import pandas as pd
from datetime import datetime, timedelta
from pytz import timezone
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pathlib import Path

# -------------------------------------------------------------------
# CONFIGURATION (keep your existing values)
# -------------------------------------------------------------------
PROJECT_ROOT = Path(__file__).resolve().parent
FILE_PATH = PROJECT_ROOT / "license_data" / "licenses.xlsx"

frequency_mapping = {
    "6 MONTH": timedelta(days=180),
    "1 YEAR": timedelta(days=365),
    "2 YEAR": timedelta(days=730),
}

SENDER_EMAIL = "alerts@sharadaindustries.com"
SENDER_PASSWORD = "Sharada@321"
SMTP_SERVER = "sharadaindustries.mithiskyconnect.com"
SMTP_PORT = 465

# -------------------------------------------------------------------
# EMAIL SENDER (unchanged)
# -------------------------------------------------------------------
def send_email(subject, body, to_emails, cc_emails=None):
    msg = MIMEMultipart()
    msg["From"] = SENDER_EMAIL
    msg["To"] = ", ".join(to_emails)
    msg["Subject"] = subject
    if cc_emails:
        msg["Cc"] = ", ".join(cc_emails)

    msg.attach(MIMEText(body, "html"))
    all_recipients = to_emails + (cc_emails if cc_emails else [])

    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, all_recipients, msg.as_string())
        print(f"✅ Email sent to: {', '.join(all_recipients)}")
    except Exception as e:
        print(f"❌ Failed to send email: {e}")

# -------------------------------------------------------------------
# MAIN FUNCTION (updated to iterate over VALIDITY)
# -------------------------------------------------------------------
def update_and_notify():
    # Check if file exists
    if not FILE_PATH.exists():
        print(f"❌ Excel file not found at: {FILE_PATH}")
        return

    print(f"📂 Reading file from: {FILE_PATH}")
    df = pd.read_excel(FILE_PATH)

    # Normalize headers to uppercase (to avoid KeyErrors)
    df.columns = df.columns.str.strip().str.upper()
    print("📋 Columns found in Excel:", list(df.columns))

    required_cols = {"VALIDITY"}
    missing = required_cols - set(df.columns)
    if missing:
        print(f"⚠️ Missing columns in Excel: {missing}")
        print("Please check the header names in your Excel file.")
        return

    # Convert to datetime
    df["VALIDITY"] = pd.to_datetime(df["VALIDITY"], errors="coerce")
    if "REMINDER" in df.columns:
        df["REMINDER"] = pd.to_datetime(df["REMINDER"], errors="coerce")

    today = datetime.now(timezone("Asia/Kolkata")).date()
    print(f"🕒 Running reminders for: {today}")

    # -------------------------------
    # STEP 1: SEND REMINDERS
    # -------------------------------
    remind_indices = []
    for i, row in df.iterrows():
        validity = row.get("VALIDITY")
        if pd.isna(validity):
            continue
        validity_date = pd.Timestamp(validity).date()
        days_left = (validity_date - today).days

        if days_left in (14, 15, 4, 5):
            remind_indices.append(i)
            print(f"🔔 Will remind for row {i} — VALIDITY {validity_date} (in {days_left} days)")

    if remind_indices:
        remind_rows = df.loc[remind_indices].copy()
        for c in remind_rows.columns:
            if pd.api.types.is_datetime64_any_dtype(remind_rows[c]):
                remind_rows[c] = remind_rows[c].dt.strftime("%d-%m-%Y")

        html_table = remind_rows.to_html(index=False)
        subject = "License/Certificate Reminder - Sharada Industries"
        body = f"""
        <html>
        <body>
            <p>Dear Sir,</p>
            <p>This is an automated reminder for the following licenses/certificates which are approaching their validity date:</p>
            {html_table}
            <p>Please ensure renewal before the validity date.</p>
            <p>Regards,<br>Sharada Industries</p>
        </body>
        </html>
        """
        send_email(subject, body,
                   ["suryakant.mhetre@sharadaindustries.com"],
                   ["amrapali.ingle@sharadaindustries.com"])
    else:
        print("✅ No reminders triggered by VALIDITY windows today.")

    # -------------------------------
    # STEP 2: UPDATE EXPIRED ROWS
    # -------------------------------
    updated = False
    if "FREQUENCY" in df.columns:
        for i, row in df.iterrows():
            validity = row.get("VALIDITY")
            if pd.isna(validity):
                continue
            validity_date = pd.Timestamp(validity).date()
            if validity_date < today:
                freq_raw = row.get("FREQUENCY")
                if pd.notnull(freq_raw):
                    # Normalize frequency (remove plural "S")
                    freq = str(freq_raw).strip().upper().rstrip('S')
                    if freq in frequency_mapping:
                        new_validity = pd.Timestamp(validity) + frequency_mapping[freq]
                        df.at[i, "VALIDITY"] = new_validity
                        new_reminder_date = (new_validity - timedelta(days=15)).date()
                        if "REMINDER" in df.columns:
                            df.at[i, "REMINDER"] = pd.Timestamp(new_reminder_date)
                        updated = True
                        print(f"🔁 Row {i} validity expired; updated VALIDITY to {new_validity.date()} and REMINDER to {new_reminder_date}")
                    else:
                        print(f"⚠️ Row {i} has FREQUENCY '{freq_raw}' which is not in mapping — skipping update")
                else:
                    print(f"⚠️ Row {i} validity expired but no FREQUENCY provided — skipping update")
    else:
        print("ℹ️ No FREQUENCY column found; skipping validity updates for expired rows.")

    # -------------------------------
    # STEP 3: SAVE (sorted by S.N)
    # -------------------------------
    if updated:
        # Convert datetime columns to date-only
        for col in ["VALIDITY", "REMINDER"]:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors="coerce").dt.date

        # Sort by S.N if available, else fallback to VALIDITY
        if "S.N" in df.columns:
            df.sort_values(by="S.N", inplace=True, ignore_index=True)
        elif "VALIDITY" in df.columns:
            df.sort_values(by="VALIDITY", inplace=True, ignore_index=True)

        # Safe save (handle file lock)
        try:
            # Convert datetime columns to formatted string for Excel
            for col in ["VALIDITY", "REMINDER"]:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%d-%m-%Y")

            # Safe save (handle file lock)
            try:
                df.to_excel(FILE_PATH, index=False)
                print("✅ Updated validity and reminder dates saved to Excel (sorted by S.N, date-only format).")
            except PermissionError:
                print(f"⚠️ Excel file '{FILE_PATH}' is open. Please close it and rerun the script.")

            print("✅ Updated validity and reminder dates saved to Excel (sorted by S.N).")
        except PermissionError:
            print(f"⚠️ Excel file '{FILE_PATH}' is open. Please close it and rerun the script.")
    else:
        print("ℹ️ No updates made to Excel — no expired validity dates.")

    print("🏁 Reminder check complete.")


# -------------------------------------------------------------------
if __name__ == "__main__":
    update_and_notify()
