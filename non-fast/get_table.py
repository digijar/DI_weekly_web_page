from flask import Flask, render_template
import os
from google.cloud import bigquery
import json
from flask_cors import CORS
import pandas as pd

app = Flask(__name__)

CORS(app)

os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'testing-bigquery-vertexai-service-account.json'
client = bigquery.Client()

@app.route('/get_table', methods=["GET"])
def get_table():

    sql_query = """
    SELECT table_name from `testing-bigquery-vertexai.MergerMarket`.INFORMATION_SCHEMA.TABLES    
    """

    query_job = client.query(sql_query)

    table_dict = {}
    count = 0
    for row in query_job.result():
        print(row[0])

        table_dict['table' + str(count)] = row[0]
        count += 1
    
    return json.dumps(table_dict)

@app.route('/get_table/<string:table_name>', methods=["GET"])
def get_table_data(table_name):
    job = """
    SELECT * FROM `testing-bigquery-vertexai.templates.{table_name}`
    ORDER BY Opportunity_id;
    """.format(table_name = table_name)

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
        return_json[opportunity_id] = row_dict
        
    # return json.dumps(return_json)
    return return_json


if __name__ == '__main__':
    app.run(debug=True, port=5002)