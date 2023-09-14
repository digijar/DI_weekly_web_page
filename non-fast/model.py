from flask import Flask, render_template
import os
from google.cloud import bigquery
import json
from flask_cors import CORS

app = Flask(__name__)

CORS(app)

os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'testing-bigquery-vertexai-service-account.json'
client = bigquery.Client()

@app.route('/opportunity/<string:id>', methods=["GET"])
def return_text(id):
    sql_query = """
    SELECT * FROM `testing-bigquery-vertexai.MergerMarket.MergerMarket_Cleaned_v2`
    WHERE Opportunity_ID = {id}
    """.format(id = id)

    query_job = client.query(sql_query)

    for row in query_job.result():
        opportunity = row['Opportunity']

        value = {
            "opportunity": opportunity
        }

        #returns bigquery data as json object, which will be queried in index.html
        return json.dumps(value)
        # return opportunity

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
          CONCAT('Your task is to perform text classification on the following text, and return one of the following categories: ```Completed or Available```. An opportunity classified as ‘Completed’ may include the following phrases or keywords: “Company X have already mandated, mandates company for their deal. Company X has raised, raises, closes fundraise, completes, gains, closes, inject, receives, of abc amount in series Y. Company X has sold, stake sold, exited, exists, their xyz% stake. Company X has acquired, already acquired, acquires, buys, completes acquisition, Company Y. Withdraws IPO, will be halted, launch of pre-structuring. Anti-takeover mechanism, takeover defense measures, hostile suitor, poison pill.” An opportunity classified as ‘Available’ may include the following phrases or keywords: “Plans to buy, to acquire, to be acquired by. Plans to sell, x% stake to be placed for sale, mulling sale, stake put up for sale, seeks exit. Want, seeks, to raise, invites investors, to receive, new funding. Seek pre IPO funding, plans to file, eyes IPO.”', Opportunity) AS prompt,
          *
        FROM
          `testing-bigquery-vertexai.MergerMarket.MergerMarket_Cleaned_v2`
        WHERE Opportunity_ID = {id}
      ),
      STRUCT(
        0.2 AS temperature,
        10 AS max_output_tokens));""".format(id = id)

    result = client.query(job)
    for row in result:
        print(row[0])
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
        200 AS max_output_tokens));""".format(id = id)

    result = client.query(job)
    for row in result:
        print(row[0])
        return row[0]  

if __name__ == '__main__':
    app.run(debug=True, port=5001)