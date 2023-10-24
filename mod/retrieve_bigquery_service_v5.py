import requests
import os
from fastapi import FastAPI, Request
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

json_path = "./testing_bigquery_vertexai_service_account.json"
json_abs_path = os.path.abspath(json_path)
os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = json_abs_path
bq = bigquery.Client()

@app.get("/")
async def get_rolling_shortlist_data():
    sql_query = "SELECT * FROM `testing-bigquery-vertexai.templates.Rolling_Shortlist`"
    result = bq.query(sql_query)

    target_list = []
    for row in result:
        # print(row[' Target'])
        # print(row['Business Description'])
        target_list.append(row[' Target'])
        # return "succeeded"
    return JSONResponse(content = target_list)

@app.get("/bq/{row_num}")
async def retrieve_bigquery(row_num: float):
    # Construct the SQL query to retrieve data
    sql_query = """
    SELECT
        ` Target` AS company_name,
        `Deal Intelligence info`,
        `Target country` AS dominant_country,
        `Target Region` AS region,
        `KPMG View - Redacted` AS next_step,
        `Asset pack ` AS other_info,
        `Website` AS link,
        `Business Description` AS biz_desc
    FROM
        `testing-bigquery-vertexai.templates.Rolling_Shortlist`
    WHERE
        Num = {num}
    """.format(num=row_num)

    # Run the query
    query_job = bq.query(sql_query)

    # Get the results
    # results = query_job.result()
    results = []
    for row in query_job:
        results.append(row)

    print(results)
    # Initialize dictionaries to store column summaries
    col_summary = {}

    # Define columns for which you want to generate summaries
    cols = ["Deal Intelligence info", "Business Description"]

    for rscol in cols:
        # check if column is empty
        check_result = {}
        checking_job = "SELECT `" + rscol + f"` from `testing-bigquery-vertexai.templates.Rolling_Shortlist` WHERE Num = {row_num}"
        print(checking_job)
        # checking_job = """SELECT {tgt_column} from `testing-bigquery-vertexai.templates.Rolling_Shortlist` WHERE Num = {Num}""".format(tgt_column = rscol, Num = row_num)
        checking_query_job = bq.query(checking_job)
        for row in checking_query_job:
            check_result[rscol] = row[rscol]

        print('Checking ok!')
        # summarise for non-empty column
        if check_result[rscol] != "nan":
            job = """
            SELECT
                ml_generate_text_result['predictions'][0]['content'] AS generated_text
            FROM
                ML.GENERATE_TEXT(
                    MODEL `bqml_tutorial.llm_model`,
                    (
                    SELECT
                        CONCAT('Summarize the following text in 50 words: ', `{tgt_column}`) AS prompt,
                        *
                    FROM
                        `testing-bigquery-vertexai.templates.Rolling_Shortlist`
                    WHERE Num = {num}
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
        else:
            col_summary[rscol] = '" No text to summarise. "'

    print('ML done.')
    
    # # Split the "other_info" column into a list
    if results[0]["other_info"] != None:
        other_info = results[0]["other_info"].split(";")
    else:
        other_info = 'None'

    # Convert the results to a JSON response
    response_data = {
        "company_name": results[0]["company_name"],
        "deal_intelligence": results[0]["Deal Intelligence info"],
        "dominant_country": results[0]["dominant_country"],
        "region": results[0]["region"],
        "next_step": results[0]["next_step"],
        "other_info": other_info,
        "col_summary": col_summary,
        "link": results[0]["link"],
        "biz_desc": results[0]["biz_desc"]
    }

    return JSONResponse(content=jsonable_encoder(response_data))

@app.post("/update")
async def update_rolling_shortlist(request: Request):
    try:
        data = await request.json()
        num = data.get('num')
        target = data.get('target')
        json_obj = data.get('scraped_data')
        print(json_obj)

        update_sql_query = "UPDATE `testing-bigquery-vertexai.templates.Rolling_Shortlist` SET "

        for company, info in json_obj.items():
            if len(info) != 0:
                for column, data in info.items():
                    if column == "Revenue":
                        data_str = str(data)
                        update_sql_query += "`Revenue_USD_M` = " + "'" + data_str + "'"
                        update_sql_query += ", "
                    elif column == "EBITDA":
                        data_str = str(data)
                        update_sql_query += "`EBITDA_USD_M` = " + "'" + data_str + "'"
                        update_sql_query += ", "
                    elif column == "Valuation":
                        data_str = str(data)
                        update_sql_query += "`Valuation_USD_M` = " +"'" + data_str + "'"
                        update_sql_query += ", "
                    elif column == "Other Info":
                        update_sql_query += "`Other Info` = " + "'" + data + "'"
                        update_sql_query += ", "
                    elif column == "Asset Pack":
                        update_sql_query += "`Asset pack ` = " + "'" + data + "'"
                        update_sql_query += ", "
                    elif column == "Business Description":
                        update_sql_query += "`Business Description` = " + "'" + data + "'"
                        update_sql_query += ", "
                    elif column == "Target Region":
                        update_sql_query +=  "`Target Region` = " + "'" + data + "'"
                        update_sql_query += ", "
                    elif column == "Website":
                        update_sql_query += "`Website` = " + "'" + data + "'"
                        update_sql_query += ", "
                
                update_sql_query = update_sql_query[:-2]
                update_sql_query += " WHERE Num = {};".format(num)
            
            else:
                update_sql_query = "UPDATE `testing-bigquery-vertexai.templates.Rolling_Shortlist` SET `Other Info`='Webscraper did not find data' WHERE Num = {};".format(num)

        print(update_sql_query)
        bq.query(update_sql_query)
        
        success = True
    except Exception as e:
        print(e)
        success = False
    return {"success": success}

if __name__ == '__main__':
    import uvicorn
    uvicorn.run("retrieve_bigquery_service_v5:app", host='127.0.0.1', port=5011, reload=True)