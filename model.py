import os
from google.cloud import bigquery
import json

from fastapi import FastAPI
from fastapi.encoders import jsonable_encoder
from fastapi.responses import JSONResponse
import uvicorn

app = FastAPI()

os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'testing-bigquery-vertexai-service-account.json'
client = bigquery.Client()

@app.get("/")
async def root():
    return {"message":"testing"}

@app.get("/opportunity/{id}")
def return_text(id: str):
    sql_query = """SELECT * FROM `testing-bigquery-vertexai.MergerMarket.MergerMarket_Cleaned_v2`WHERE Opportunity_ID = {id}""".format(id = id)

    query_job = client.query(sql_query)

    for row in query_job.result():
        opportunity = row['Opportunity']

        value = {
            "opportunity": opportunity
        }

        return_json = jsonable_encoder(value)

        return JSONResponse(content=return_json)


@app.get("/classification/{id}")
async def classify_text(id: str):
    job = """
    SELECT
    ml_generate_text_result['predictions'][0]['content'] AS predicted_classification,
    ml_generate_text_result['predictions'][0]['safetyAttributes']
      AS safety_attributes,
    * EXCEPT (ml_generate_text_result)
    FROM
    ML.GENERATE_TEXT(
      MODEL `bqml_tutorial.llm_model`,
      (SELECT
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
        return JSONResponse(row[0])

@app.get("/summarization/{id}")
async def summarize_text(id: str):

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
    
        return JSONResponse(row[0])


if __name__ == '__main__':
    uvicorn.run("main:app", host='127.0.0.1', port=5001, reload=True)