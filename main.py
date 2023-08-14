from flask import Flask, render_template, request, jsonify
import os
import pandas as pd
from google.cloud import bigquery
from flask_cors import CORS
import json

# Replace 'your-project-id' with your actual Google Cloud Project ID
from dotenv import load_dotenv
load_dotenv('gcloud.env')
PROJECT_ID = os.getenv('PROJECT_ID')
DATASET_ID = os.getenv('DATASET_ID')


app = Flask(__name__)
CORS(app)

# Set up BigQuery client
# os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'testing-bigquery-vertexai-service-account.json'
client = bigquery.Client()

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
            table_ref = client.dataset("testing-bigquery-vertexai").table("web_UI")
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

@app.route('/opportunity/<string:id>', methods=["GET"])
def return_text(id):
    sql_query = """
    SELECT * FROM `testing-bigquery-vertexai.MergerMarket.MergerMarket_Cleaned_v2`
    WHERE Opportunity_ID = {id}
    """.format(id=id)

    query_job = client.query(sql_query)

    for row in query_job.result():
        opportunity = row['Opportunity']
        value = {"opportunity": opportunity}
        return jsonify(value)

@app.route('/classification/<string:id>', methods=["GET"])
def classify_text(id):
    job = """
    SELECT
    ml_generate_text_result['predictions'][0]['content'] AS predicted_classification,
    ml_generate_text_result['predictions'][0]['safetyAttributes']
      AS safety_attributes,
    * EXCEPT (ml_generate_text_result)
    FROM
    ML.GENERATE_TEXT(
      MODEL `bqml_tutorial.llm_model`,
      (
        SELECT
          CONCAT('Your task is to perform text classification on the following text, and return one of the following categories: ```Completed or Available```...',
          Opportunity) AS prompt,
          *
        FROM
          `testing-bigquery-vertexai.MergerMarket.MergerMarket_Cleaned_v2`
        WHERE Opportunity_ID = {id}
      ),
      STRUCT(
        0.2 AS temperature,
        10 AS max_output_tokens));""".format(id=id)

    result = client.query(job)
    for row in result:
        return row[0]

@app.route('/summarization/<string:id>', methods=["GET"])
def summarize_text(id):
    job = """
    SELECT
    ml_generate_text_result['predictions'][0]['content'] AS predicted_classification,
    ml_generate_text_result['predictions'][0]['safetyAttributes']
      AS safety_attributes,
    * EXCEPT (ml_generate_text_result)
    FROM
    ML.GENERATE_TEXT(
      MODEL `bqml_tutorial.llm_model`,
      (
        SELECT
          CONCAT('Summarize the following text in 100 words: ', Opportunity) AS prompt,
          *
        FROM
          `testing-bigquery-vertexai.MergerMarket.MergerMarket_Cleaned_v2`
        WHERE Opportunity_ID = {id}
      ),
      STRUCT(
        0.2 AS temperature,
        200 AS max_output_tokens));""".format(id=id)

    result = client.query(job)
    for row in result:
        return row[0]

if __name__ == '__main__':
    app.run(debug=True)