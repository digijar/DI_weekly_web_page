import os
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.encoders import jsonable_encoder
from fastapi.responses import JSONResponse
from google.cloud import bigquery

app = FastAPI()

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/bq/{row_num}")
def retrieve_bigquery(row_num: int):
    # Set the path to your Google Cloud service account key
    json_path = "./testing_bigquery_vertexai_service_account.json"
    json_abs_path = os.path.abspath(json_path)
    os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = json_abs_path

    # Initialize the BigQuery client
    bq = bigquery.Client()

    # Construct the SQL query to retrieve data
    sql_query = """
    SELECT
        _Target AS company_name,
        Deal_Intelligence_info,
        Target_country AS dominant_country,
        Target_Region AS region,
        KPMG_View___Redacted AS next_step,
        Asset_pack_ AS other_info,
        Website AS link
    FROM
        `testing-bigquery-vertexai.web_UI.Rolling_08-09-23`
    WHERE
        Num = '{num}'
    """.format(num=row_num)

    # Run the query
    query_job = bq.query(sql_query)

    # Get the results
    # results = query_job.result()
    results = []
    for row in query_job:
        results.append(row)

    # Initialize dictionaries to store column summaries
    col_summary = {}

    # Define columns for which you want to generate summaries
    cols = ["Deal_Intelligence_info", "Business_Description"]

    for rscol in cols:
        job = """
        SELECT
            ml_generate_text_result['predictions'][0]['content'] AS generated_text
        FROM
            ML.GENERATE_TEXT(
                MODEL `bqml_tutorial.llm_model`,
                (
                SELECT
                    CONCAT('Summarize the following text in 50 words: ', {tgt_column}) AS prompt,
                    *
                FROM
                    `testing-bigquery-vertexai.web_UI.Rolling_08-09-23`
                WHERE Num = '{num}'
                ),
                STRUCT(
                0.2 AS temperature,
                200 AS max_output_tokens
                )
            );
        """.format(num=row_num, tgt_column=rscol)

        result = bq.query(job)

        # Extract and store column summaries
        for r in result:
            col_summary[rscol] = r["generated_text"]

    # # Split the "other_info" column into a list
    other_info = results[0]["other_info"].split(";")

    # Convert the results to a JSON response
    response_data = {
        "company_name": results[0]["company_name"],
        "deal_intelligence": results[0]["Deal_Intelligence_info"],
        "dominant_country": results[0]["dominant_country"],
        "region": results[0]["region"],
        "next_step": results[0]["next_step"],
        "other_info": other_info,
        "col_summary": col_summary,
        "link": results[0]["link"]
    }

    return JSONResponse(content=jsonable_encoder(response_data))

if __name__ == '__main__':
    import uvicorn
    uvicorn.run("retrieve_bigquery_service:app", host='127.0.0.1', port=5011, reload=True)