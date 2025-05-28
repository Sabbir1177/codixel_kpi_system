#-------when i found the api url for that time-----------
# import pandas as pd
# import requests
# from io import StringIO

# # ---------- STEP 1: Load Attendance Data from API ----------
# def load_attendance_data():
#     url = "https://your-api.com/attendance"  
#     response = requests.get(url)
#     csv_data = StringIO(response.text)
#     df = pd.read_csv(csv_data)
#     df["Hours Worked"] = pd.to_datetime(df["Check-out Time"], format="%H:%M") - pd.to_datetime(df["Check-in Time"], format="%H:%M")
#     df["Present"] = df["Present? (Y/N)"].apply(lambda x: 1 if x == "Y" else 0)
#     return df

# # ---------- STEP 2: Load Codixel Extra Task Data from API ----------
# def fetch_codixel_internal_data():
#     url = "https://your-api.com/extra-tasks"  
#     response = requests.get(url)
#     csv_data = StringIO(response.text)
#     df = pd.read_csv(csv_data)
#     return df

# # ---------- STEP 3: Load Involvement Data from API ----------
# def load_involvement_data():
#     url = "https://your-api.com/involvement" 
#     response = requests.get(url)
#     csv_data = StringIO(response.text)
#     df = pd.read_csv(csv_data)
#     return df


import pandas as pd

# ---------- STEP 1: Load Attendance Data ----------
def load_attendance_data():
    df = pd.read_csv("attendance.csv")
    df["Hours Worked"] = pd.to_datetime(df["Check-out Time"], format="%H:%M") - pd.to_datetime(df["Check-in Time"], format="%H:%M")
    df["Present"] = df["Present? (Y/N)"].apply(lambda x: 1 if x == "Y" else 0)
    return df

# ---------- STEP 2: Load Codixel Extra Task Data ----------
def fetch_codixel_internal_data():
    df = pd.read_csv("extra_activities.csv")
    return df

# ---------- STEP 3: Load Involvement Data ----------
def load_involvement_data():
    df = pd.read_csv("involvement.csv")
    return df

# ---------- STEP 4: KPI Calculation ----------
def calculate_kpi(attendance_df, extra_df, involvement_df):
    members = attendance_df["Name"].unique()
    result = []

    for name in members:
        present_days = attendance_df[attendance_df["Name"] == name]["Present"].sum()
        total_days = attendance_df[attendance_df["Name"] == name]["Date"].nunique()
        attendance_score = (present_days / total_days) * 100 if total_days else 0

        extra_time = extra_df[extra_df["Name"] == name]["Time Spent (hrs)"].sum()
        extra_score = min((extra_time / 10) * 100, 100)

        involvement_row = involvement_df[involvement_df["Name"] == name]
        if not involvement_row.empty:
            meeting = involvement_row["Internal Meeting Attended"].values[0]
            feedback = involvement_row["Gave Feedback"].values[0]
            helped = 10 if involvement_row["Helped Others (Y/N)"].values[0] == "Y" else 0
            involvement_score = meeting * 10 + feedback * 10 + helped
        else:
            involvement_score = 0

        final_score = (attendance_score * 0.4 +
                       extra_score * 0.3 +
                       involvement_score * 0.3)

        result.append({
            "Name": name,
            "Attendance Score": round(attendance_score, 2),
            "Extra Activity Score": round(extra_score, 2),
            "Involvement Score": round(involvement_score, 2),
            "Total KPI Score": round(final_score, 2)
        })

    return pd.DataFrame(result)

# ---------- STEP 5: Export Everything ----------
def export_to_excel(attendance_df, extra_df, involvement_df, kpi_df, output_path):
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        attendance_df.to_excel(writer, index=False, sheet_name="Attendance Tracker", startrow=2)
        extra_df.to_excel(writer, index=False, sheet_name="Extra Activities", startrow=2)
        involvement_df.to_excel(writer, index=False, sheet_name="Involvement Log", startrow=2)
        kpi_df.to_excel(writer, index=False, sheet_name="KPI Scoring", startrow=2)
        workbook = writer.book
        # Title format
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 16,
            'align': 'center',
            'valign': 'vcenter'
        })
        sheets = {
            "Attendance Tracker": writer.sheets["Attendance Tracker"],
            "Extra Activities": writer.sheets["Extra Activities"],
            "Involvement Log": writer.sheets["Involvement Log"],
            "KPI Scoring": writer.sheets["KPI Scoring"]
        }

           # Import Excel column utility
        from xlsxwriter.utility import xl_col_to_name

        # Write title to each sheet
        for sheet_name, worksheet in sheets.items():
            # Dynamically get column count for that sheet
            if sheet_name == "Attendance Tracker":
                col_count = attendance_df.shape[1]
            elif sheet_name == "Extra Activities":
                col_count = extra_df.shape[1]
            elif sheet_name == "Involvement Log":
                col_count = involvement_df.shape[1]
            elif sheet_name == "KPI Scoring":
                col_count = kpi_df.shape[1]
            else:
                col_count = 8  # fallback

            # Convert column index to Excel column letter (e.g. 7 → H)
            end_col_letter = xl_col_to_name(col_count - 1)

            # Merge the top row and add the title
            worksheet.merge_range(f'A1:{end_col_letter}1', 'Codixel KPI Report', title_format)



# ---------- RUN EVERYTHING ----------
def run_kpi_system():
    attendance = load_attendance_data()
    extra_data = fetch_codixel_internal_data()
    involvement = load_involvement_data()

    kpi = calculate_kpi(attendance, extra_data, involvement)
    export_to_excel(attendance, extra_data, involvement, kpi, "Codixel_KPI_Report.xlsx")
    print("✅ KPI Report Generated: Codixel_KPI_Report.xlsx")
  

# Run the script
run_kpi_system()
