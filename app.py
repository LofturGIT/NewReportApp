import os
from flask import Flask, request, render_template, send_file
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from datetime import datetime
import re

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
REPORTS_FOLDER = 'reports'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORTS_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def upload_files():
    if request.method == 'POST':
        # Get uploaded files
        pending_users_file = request.files.get('pending_users')
        course_status_files = request.files.getlist('course_status')

        if not pending_users_file or not course_status_files:
            return "Please upload both Pending Users and Course Status files!", 400

        # Save uploaded files
        pending_users_path = os.path.join(UPLOAD_FOLDER, pending_users_file.filename)
        pending_users_file.save(pending_users_path)

        course_status_paths = []
        for file in course_status_files:
            file_path = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(file_path)
            course_status_paths.append(file_path)

        # Process files
        report_path = process_files(pending_users_path, course_status_paths)

        # Send the generated report back to the user
        return send_file(report_path, as_attachment=True)

    return render_template('upload.html')

def process_files(pending_users_path, course_status_paths):
    # Load the pending users dataset
    pending_users_df = pd.read_csv(pending_users_path)
    pending_users_df['Email'] = pending_users_df['Email'].str.strip().str.lower()
    pending_users_df['Email Domain'] = pending_users_df['Email'].str.split('@').str[1]

    template_path = 'template2.xlsx'
    logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logo.png')
    combined_reports = []

    for course_status_path in course_status_paths:
        # Load the course status report
        course_status_df = pd.read_csv(course_status_path)
        course_status_df['Email'] = course_status_df['Email'].str.strip().str.lower()
        course_status_df['Email Domain'] = course_status_df['Email'].str.split('@').str[1]

        # Filter pending users by matching email domains
        matching_domains = set(course_status_df['Email Domain'].dropna().unique())
        filtered_pending_users = pending_users_df[pending_users_df['Email Domain'].isin(matching_domains)]

        pending_users = filtered_pending_users[['Email', 'Last invite sent at']].copy()
        pending_users['Full name'] = 'Pending User'
        pending_users['Course name'] = 'Pending Course'
        pending_users['Progress'] = 'Not started'
        pending_users['Score'] = '0%'
        pending_users['Course completion %'] = '0%'
        pending_users['Enrolled'] = '-'
        pending_users['Started'] = '-'
        pending_users['Last accessed'] = '-'
        pending_users['Completed'] = pending_users['Last invite sent at'].apply(
            lambda x: f"Invite last sent: {x}" if pd.notna(x) and x.strip() else "Unknown"
        )

        # Combine course status and pending users
        combined_df = pd.concat([course_status_df, pending_users], ignore_index=True)

        # Define 'Status' column
        combined_df['Status'] = combined_df.apply(
            lambda row: 'Passed' if row['Progress'] == 'Passed'
            else 'In Progress' if row['Progress'] == 'In Progress'
            else 'Not started',
            axis=1
        )

        # Update 'Completed' column logic for course data
        combined_df['Completed'] = combined_df.apply(
            lambda row: f"Completed: {row['Completed'].split(' ')[0]}" if row['Status'] == 'Passed' and pd.notna(row['Completed']) and row['Completed'] != '-'
            else f"Last accessed course: {row['Last accessed'].split(' ')[0]}" if row['Status'] == 'In Progress' and pd.notna(row['Last accessed']) and row['Last accessed'] != '-'
            else f"Enrolled on: {row['Enrolled'].split(' ')[0]}" if row['Status'] == 'Not started' and pd.notna(row['Enrolled']) and row['Enrolled'] != '-'
            else row['Completed'],  # Preserve "Invite last sent" for pending users
            axis=1
        )

        # Prepare final DataFrame
        final_df = combined_df[['Full name', 'Email', 'Course name', 'Status', 'Completed', 'Score']].copy()
        final_df.rename(columns={'Full name': 'User'}, inplace=True)

        # Save to Excel
        wb = load_workbook(template_path)
        ws = wb.active
        start_row = 13  # Starting row for pasting data
        start_col = 2   # Starting column for pasting data
        for r_idx, row in final_df.iterrows():
            for c_idx, value in enumerate(row):
                ws.cell(row=start_row + r_idx, column=start_col + c_idx, value=value)

        # Add logo to the worksheet
        img = Image(logo_path)
        img.anchor = "D4"  # Place the image in cell D4
        ws.add_image(img)

        # Save final report
        today_date = datetime.now().strftime('%Y-%m-%d')
        course_name = re.sub(r'[<>:"/\\|?*]', '_', combined_df['Course name'].iloc[0])
        output_file = os.path.join(REPORTS_FOLDER, f"New_Report_{course_name}_{today_date}.xlsx")
        wb.save(output_file)

        combined_reports.append(output_file)

    return combined_reports[0]

if __name__ == '__main__':
    print("Starting Flask App...")
    app.run(debug=True)
