from flask import Flask, render_template, jsonify, request
import os
from google.cloud import bigquery
import json
from flask_cors import CORS
import pandas as pd

app = Flask(__name__)

CORS(app)

os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'testing-bigquery-vertexai-service-account.json'
client = bigquery.Client()

@app.route('/get_table/<string:table_name>', methods=["GET"])
def get_table_data(table_name):
    job = """
    SELECT * FROM `testing-bigquery-vertexai.templates.{table_name}`
    """.format(table_name=table_name)

    result = client.query(job)

    return_json = []

    # Iterate through rows
    for row in result:
        row_dict = {}
        
        # Iterate through columns dynamically
        for column_name in row.keys():
            row_dict[column_name] = row[column_name]

        return_json.append(row_dict)

    return jsonify(return_json)

    # for row in result.result():

    #     row_dict = {}

    #     row_dict["Vendors"] = row["Vendors"]
    #     row_dict["Bidders"] = row["Bidders"]
    #     row_dict["States"] = row["States"]
    #     row_dict["Sectors"] = row["Sectors"]
    #     row_dict["Intelligence_Type"] = row["Intelligence Type"]
    #     row_dict["Geography"] = row["Geography"]
    #     row_dict["Stake_Value"] = row["Stake Value"]
    #     row_dict["Intelligence_Size"] = row["Intelligence Size"]
    #     row_dict["Opportunity"] = row["Opportunity"]
    #     row_dict["Type_of_transaction"] = row["Type of transaction"]
    #     row_dict["HS_sector_classification"] = row["HS sector classification"]
    #     row_dict["Targets"] = row["Targets"]
    #     row_dict["Intelligence_Grade"] = row["Intelligence Grade"]
    #     row_dict["Source"] = row["Source"]
    #     row_dict["Value_Description"] = row["Value Description"]
    #     row_dict["Dominant_Sector"] = row["Dominant Sector"]
    #     row_dict["Others"] = row["Others"]
    #     row_dict["Competitors"] = row["Competitors"]
    #     row_dict["Heading"] = row["Heading"]
    #     row_dict["Sub_Sectors"] = row["Sub Sectors"]
    #     row_dict["Dominant_Geography"] = row["Dominant Geography"]
    #     row_dict["Value_INR_M"] = row["Value INR_m"]
    #     row_dict["Date"] = row["Date"]
    #     row_dict["Short_BD"] = row["Short BD"]
    #     row_dict["Issuers"] = row["Issuers"]
    #     row_dict["Topics"] = row["Topics"]
    #     row_dict["Lead_type"] = row["Lead type"]
    #     row_dict["Opportunity_ID"] = row["Opportunity_ID"]

    #     opportunity_id = int(row["Opportunity_ID"])
    #     return_json[opportunity_id] = row_dict
        
    # # return json.dumps(return_json)
    # return return_json


@app.route('/update_table/<string:table_name>', methods=["POST"])
def update_table_data(table_name):
    data = request.get_json()
    success = True  # Assume success initially

    # for row_data in data:
    #         row_id = row_data.get("row_id")
    #         updated_column = row_data.get("column_name")
    #         updated_value = row_data.get("edited_value")
    
    # return jsonify({"success": success, "row_id": row_id, "updated_column": updated_column, "updated_value": updated_value})

    try:
        for row_data in data:
            row_id = row_data.get("row_id")
            updated_column = row_data.get("column_name")
            updated_value = row_data.get("edited_value")

            # Construct the UPDATE query dynamically
            update_query = f"""
            UPDATE `testing-bigquery-vertexai.templates.{table_name}`
            SET `{updated_column}` = '{updated_value}'
            WHERE Opportunity_ID = {row_id}
            """

            # Execute the query using BigQuery API or library of your choice
            # Set up your Google Cloud SDK credentials for BigQuery access
            client.query(update_query)
            
            # Simulate success/failure
            # Uncomment the above line and handle the response as per your library's response structure
            # For now, we'll assume the query was successful
            # client.query(update_query)
            
    except Exception as e:
        print(e)
        success = False
    
    return jsonify({"success": success})


if __name__ == '__main__':
    app.run(debug=True, port=5003)