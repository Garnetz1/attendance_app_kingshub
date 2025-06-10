import uuid
from flask import Flask, render_template, request, redirect, url_for, flash
import pandas as pd
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'  # Needed for flashing messages

# File paths
CHURCH_DB = 'G:/My Drive/Attendance/attendance_app/churchDB.xlsx'
ATTENDANCE_FILE = 'G:/My Drive/Attendance/attendance_app/attendance_record.xlsx'

@app.route("/", methods=["GET", "POST"])
def attendance():
    # Load employee data for dropdown from churchDB.xlsx
    try:
        df = pd.read_excel(CHURCH_DB)
    except FileNotFoundError:
        df = pd.DataFrame(columns=['Name', 'Team', 'PCG'])  # Empty fallback
    
    names = df['Name'].tolist()
    team = ""
    pcg = ""

    if request.method == "POST":
        name = request.form["name"]
        matched = df[df["Name"].str.lower() == name.lower()]
        team = matched["Team"].values[0] if not matched.empty else ""
        pcg = matched["PCG"].values[0] if not matched.empty else ""

        service = request.form["service"]
        status = request.form["status"]
        date = datetime.now().strftime("%Y-%m-%d")

        new_data = pd.DataFrame([{
            "ID": str(uuid.uuid4()),  # Unique identifier
            "Name": name,
            "Team": team,
            "PCG": pcg,
            "Service": service,
            "Status": status,
            "Date": date
        }])

        try:
            existing = pd.read_excel(ATTENDANCE_FILE)
            # Remove duplicates for same Name, Team, PCG, Service, Date
            mask = ~(
                (existing["Name"].str.lower() == name.lower()) &
                (existing["Team"] == team) &
                (existing["PCG"] == pcg) &
                (existing["Service"] == service) &
                (existing["Date"] == date)
            )
            filtered = existing[mask]
            result = pd.concat([filtered, new_data], ignore_index=True)
        except FileNotFoundError:
            result = new_data

        result.to_excel(ATTENDANCE_FILE, index=False)

        # Clear team and pcg for next form load
        team = ""
        pcg = ""

    return render_template("attendance_form.html", names=names, team=team, pcg=pcg)

@app.route("/first-timer", methods=["GET", "POST"])
def first_timer():
    if request.method == "POST":
        name = request.form["name"].strip()
        email = request.form["email"].strip()
        phone = request.form["phone"].strip()
        pcg = request.form["pcg"]  # Should be "First Timer"
        service = request.form["service"]
        status = request.form["status"]
        date = datetime.now().strftime("%Y-%m-%d")

        # Load existing churchDB
        try:
            existing_church = pd.read_excel(CHURCH_DB)
        except FileNotFoundError:
            existing_church = pd.DataFrame(columns=['Name', 'Email', 'Phone', 'PCG', 'Date'])

        # Prevent duplicate registration in ChurchDB
        if not existing_church[existing_church['Name'].str.lower() == name.lower()].empty:
            flash(f"First Timer '{name}' is already registered!", "warning")
            return redirect(url_for('first_timer'))

        # Write to ChurchDB
        new_church_entry = pd.DataFrame([{
            "Name": name,
            "Email": email,
            "Phone": phone,
            "PCG": pcg,
            "Date": date
        }])
        updated_church = pd.concat([existing_church, new_church_entry], ignore_index=True)
        updated_church.to_excel(CHURCH_DB, index=False)

        # Write to Attendance
        attendance_entry = pd.DataFrame([{
            "ID": str(uuid.uuid4()),
            "Name": name,
            "Team": "",  # No team for first timers
            "PCG": pcg,
            "Service": service,
            "Status": status,
            "Date": date
        }])

        try:
            existing_attendance = pd.read_excel(ATTENDANCE_FILE)
            result = pd.concat([existing_attendance, attendance_entry], ignore_index=True)
        except FileNotFoundError:
            result = attendance_entry

        result.to_excel(ATTENDANCE_FILE, index=False)

        flash(f"First Timer '{name}' registered successfully!", "success")
        return redirect(url_for('first_timer'))

    return render_template("first_timer_form.html")

if __name__ == "__main__":
    app.run(debug=True)
