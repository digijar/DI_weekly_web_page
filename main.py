from flask import Flask, render_template, request
import os
# Importing necessary libraries for handling Excel and BigQuery
import pandas as pd
from google.cloud import bigquery
from flask_cors import CORS

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

        # Iterate through uploaded files and read data using pandas
        for file in files:

            # Convert the file to a binary file

            df = pd.read_excel(file, header=1)

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

if __name__ == '__main__':
    app.run(debug=True)