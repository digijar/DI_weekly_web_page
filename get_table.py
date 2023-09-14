import os
from google.cloud import bigquery
import json

from fastapi import FastAPI
from fastapi.encoders import jsonable_encoder
from fastapi.responses import JSONResponse
import uvicorn
import time

app = FastAPI()

os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'testing-bigquery-vertexai-service-account.json'
client = bigquery.Client()

@app.get("/")
async def root():
    return {"message": "Testing"}

@app.get("/get_table")
def get_table():
    sql_query = """SELECT table_name from `testing-bigquery-vertexai.MergerMarket`.INFORMATION_SCHEMA.TABLES"""

    query_job = client.query(sql_query)

    table_dict = {}
    count = 0
    for row in query_job.result():
        print(row[0])

        table_dict['table' + str(count)] = row[0]
        count += 1

    json_compatible_table_data = jsonable_encoder(table_dict)
    print(json_compatible_table_data)
    return json_compatible_table_data

@app.get("/get_table/{table_name}")
def get_table_data(table_name: str):
    startTime = time.time()

    job = """SELECT * FROM `testing-bigquery-vertexai.MergerMarket.{table_name}`ORDER BY Opportunity_id;""".format(table_name = table_name)
    print(job)
    
    result = client.query(job)

    return_json = {}
    opportunity_id = 0

    # returns a json dump with Opportunity_ID as primary key

    for row in result.result():

        row_dict = {}

        row_dict["Vendors"] = row["Vendors"]
        row_dict["Bidders"] = row["Bidders"]
        row_dict["States"] = row["States"]
        row_dict["Sectors"] = row["Sectors"]
        row_dict["Intelligence_Type"] = row["Intelligence_Type"]
        row_dict["Geography"] = row["Geography"]
        row_dict["Stake_Value"] = row["Stake_Value"]
        row_dict["Intelligence_Size"] = row["Intelligence_Size"]
        row_dict["Opportunity"] = row["Opportunity"]
        row_dict["Type_of_transaction"] = row["Type_of_transaction"]
        row_dict["HS_sector_classification"] = row["HS_sector_classification"]
        row_dict["Targets"] = row["Targets"]
        row_dict["Intelligence_Grade"] = row["Intelligence_Grade"]
        row_dict["Source"] = row["Source"]
        row_dict["Value_Description"] = row["Value_Description"]
        row_dict["Dominant_Sector"] = row["Dominant_Sector"]
        row_dict["Others"] = row["Others"]
        row_dict["Competitors"] = row["Competitors"]
        row_dict["Heading"] = row["Heading"]
        row_dict["Sub_Sectors"] = row["Sub_Sectors"]
        row_dict["Dominant_Geography"] = row["Dominant_Geography"]
        row_dict["Value_USD_M"] = row["Value_USD_M"]
        row_dict["Date"] = row["Date"]
        row_dict["Short_BD"] = row["Short_BD"]
        row_dict["Issuers"] = row["Issuers"]
        row_dict["Topics"] = row["Topics"]
        row_dict["Lead_type"] = row["Lead_type"]
        row_dict["Opportunity_ID"] = row["Opportunity_ID"]

        opportunity_id = int(row["Opportunity_ID"])
        json_compatible_item_data = jsonable_encoder(row_dict)
        return_json[opportunity_id] = json_compatible_item_data
        return_json_compatible_data = jsonable_encoder(return_json) 

    endTime = time.time()
    elapsedTime = endTime - startTime
    print(elapsedTime)

    return JSONResponse(content=return_json_compatible_data)


if __name__ == "__main__":
    uvicorn.run("get_table_fastapi:app", port=8002, reload=True)