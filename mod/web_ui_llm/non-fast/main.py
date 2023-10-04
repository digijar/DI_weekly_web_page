from flask import Flask, render_template, request
import os
# Importing necessary libraries for handling Excel and BigQuery
import pandas as pd
from google.cloud import bigquery
from flask_cors import CORS
import re

# Replace 'your-project-id' with your actual Google Cloud Project ID
from dotenv import load_dotenv
load_dotenv('gcloud.env')
PROJECT_ID = os.getenv('PROJECT_ID')
DATASET_ID = os.getenv('DATASET_ID')


app = Flask(__name__)
CORS(app)

@app.route('/')
def index():
    return render_template('index.html')

def sanitize_column_name(column_name, existing_columns):
    # Remove any characters that are not allowed in BigQuery column names
    sanitized_name = re.sub(r'[^a-zA-Z0-9_]', '_', column_name)
    
    # Handle duplicate column names
    count = 2
    while sanitized_name in existing_columns:
        sanitized_name = f"{sanitized_name}_{count}"
        count += 1
    
    return sanitized_name

@app.route('/upload', methods=['POST'])
def upload():
    try:
        if 'file' not in request.files:
            return 'No file part', 400

        files = request.files.getlist('file')
        if len(files) == 0:
            return 'No file selected', 400

        table_id = request.form.get('tableIdInput')  # Get the user-inputted table ID

        if not table_id:
            return 'Table ID is required', 400

        # Initialize a list to store sanitized column names
        sanitized_column_names = []

        # Iterate through uploaded files and read data using pandas
        for file in files:
            # Convert the file to a binary file
            df = pd.read_excel(file, header=0)

            # Sanitize column names
            df.columns = [sanitize_column_name(col, sanitized_column_names) for col in df.columns]

            # Ensure that all columns are converted to string type
            df = df.astype(str)

            # Add sanitized column names to the list
            sanitized_column_names.extend(df.columns)

            # Upload the data to BigQuery
            client = bigquery.Client(project="testing-bigquery-vertexai")
            dataset_id = "web_UI"
            table_ref = client.dataset(dataset_id).table(table_id)
            job_config = bigquery.LoadJobConfig()
            job_config.autodetect = True
            job = client.load_table_from_dataframe(df, table_ref, job_config=job_config)
            job.result()  # Wait for the job to complete

        return 'File(s) uploaded successfully!'
    except Exception as e:
        return f'Error uploading file: {str(e)}', 500
    
@app.route('/model_testing')
def model_testing():
    return render_template('model_testing.html')

@app.route('/view_table')
def edit_row():
    return render_template('view_table.html')

@app.route('/preview')
def preview():
    return render_template('preview.html')

@app.route('/slicer')
def slicer():
    return render_template('slicer.html')

if __name__ == '__main__':
    app.run(debug=True)