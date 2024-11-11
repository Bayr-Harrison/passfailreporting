import os
import streamlit as st
import pandas as pd
import pg8000
from io import BytesIO
import zipfile

# Function to query database and generate coversheets in a zip file
def generate_coversheets_zip(curriculum, startdate, enddate):
    db_connection = pg8000.connect(
        database=os.environ["SUPABASE_DB_NAME"],
        user=os.environ["SUPABASE_USER"],
        password=os.environ["SUPABASE_PASSWORD"],
        host=os.environ["SUPABSE_HOST"],
        port=os.environ["SUPABASE_PORT"]
    )

    db_cursor = db_connection.cursor()

    # Format query to select data within specified date range and curriculum
    db_query = f"""SELECT student_list.name,                  
                    student_list.iatc_id, 
                    student_list.nat_id,
                    student_list.class,
                    student_list.curriculum,
                    exam_results.exam,
                    exam_results.result,
                    exam_results.date,
                    exam_results.type
                    FROM exam_results 
                    JOIN student_list ON exam_results.nat_id = student_list.nat_id
                    WHERE student_list.curriculum = '{curriculum}' 
                    AND exam_results.date >= '{startdate}' 
                    AND exam_results.date <= '{enddate}'
                """
    db_cursor.execute(db_query)
    output_data = db_cursor.fetchall()
    db_cursor.close()
    db_connection.close()

    # Convert output to DataFrame for Excel export
    col_names = ['Name', 'IATC ID', 'National ID', 'Class', 'Curriculum', 'Exam', 'Result', 'Date', 'Exam Type']
    df = pd.DataFrame(output_data, columns=col_names)

    # Create an in-memory ZIP file to store individual Excel files
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        # Save the entire DataFrame to one Excel file
        excel_buffer = BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="Pass_Fail_Report")
        
        # Save Excel file in the zip
        excel_buffer.seek(0)
        zip_file.writestr("Pass_Fail_Report.xlsx", excel_buffer.read())

    zip_buffer.seek(0)
    return zip_buffer

# Streamlit interface
st.title("Generate Theory Exam Pass/Fail Report by Faculty and Date Range")
st.write("Select a Faculty and date range to generate an Excel Pass Fail Report within the specified period")

# Curriculum selection
curriculum = st.selectbox("Select Faculty:", ["EASA", "GACA", "UAS"])

# Date input from user
startdate = st.date_input("Select Start Date:")
enddate = st.date_input("Select End Date:")

# Button to generate and download coversheets
if st.button("Generate Report"):
    if startdate > enddate:
        st.error("Error: End Date must be after Start Date.")
    else:
        try:
            st.write("Generating Report...")

            # Generate the zip file in memory
            zip_file = generate_coversheets_zip(curriculum, startdate, enddate)

            # Download button for the zip file
            st.download_button(
                label="Download Report",
                data=zip_file,
                file_name="Pass_Fail_Report.zip",
                mime="application/zip"
            )
            st.success("Pass Fail Report zip generated successfully!")
        except Exception as e:
            st.error(f"An error occurred: {e}")
