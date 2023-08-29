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

@app.route('/get_mergermarket', methods=["GET"])
def get_mergermarket():
    table_name = "MergerMarket"
    job = """
    SELECT
        Opportunity_ID,
        Date,
        `Value INR_m`,
        `Value Description`,
        Heading,
        Opportunity,
        Targets,
        `Lead type`,
        `Type of transaction`,
        `HS sector classification`,
        `Short BD`,
        Source,
        `Intelligence Type`,
        `Intelligence Grade`,
        `Intelligence Size`,
        `Stake Value`,
        `Dominant Sector`,
        Sectors,
        `Sub Sectors`,
        `Dominant Geography`,
        Geography,
        States,
        Topics,
        Bidders,
        Vendors,
        Issuers,
        Competitors,
        Others,
        Completed
    FROM `testing-bigquery-vertexai.templates.{table_name}`
    ORDER BY Opportunity_ID;
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


@app.route('/update_mergermarket', methods=["POST"])
def update_mergermarket():
    table_name = "MergerMarket"
    data = request.get_json()
    success = True  # Assume success initially

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
            client.query(update_query)
            
    except Exception as e:
        print(e)
        success = False
    
    return jsonify({"success": success})


@app.route('/get_marketscan', methods=["GET"])
def get_marketscan():
    table_name = "Market Scan"
    job = """
    SELECT
        `Num`,
        `Picks`,
        `Captured in the week`,
        `Date`,
        `Lead type`,
        `Target HS industry classification`,
        `Dominant sector_as per MM`,
        `Sectors_as per MM`,
        `Sub Sectors_as per MM`,
        `HS Target region`,
        `Dominant country`,
        `Countries`,
        `Next Steps`,
        `Target`,
        `Target_Chinese`,
        `Deal intelligence`,
        `Short BD`,
        `Bidders`,
        `Sellers_or_Vendors`,
        `Type of transaction`,
        `Intelligence type_as per MM`,
        `Topics_as per MM`,
        `Value_USDm`,
        `Value description`,
        `Intelligence size`,
        `Intelligence size bucket`,
        `Stake value_percent`,
        `Held by ASPAC priority firm_Y_or_N`,
        `PE priority firm `,
        `Held since`,
        `Held for more than three years_Y_N_NA`,
        `KPMG credentials_PE_Company_Both_N`,
        `KPMG firm`,
        `Engagement partner`,
    FROM `testing-bigquery-vertexai.templates.{table_name}`
    ORDER BY Num;
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

@app.route('/update_marketscan', methods=["POST"])
def update_marketscan():
    table_name = "Market Scan"
    data = request.get_json()
    success = True  # Assume success initially

    try:
        for row_data in data:
            row_id = row_data.get("row_id")
            updated_column = row_data.get("column_name")
            updated_value = row_data.get("edited_value")

            # Construct the UPDATE query dynamically
            update_query = f"""
            UPDATE `testing-bigquery-vertexai.templates.{table_name}`
            SET `{updated_column}` = '{updated_value}'
            WHERE Num = {row_id}
            """
            client.query(update_query)
            
    except Exception as e:
        print(e)
        success = False
    
    return jsonify({"success": success})


@app.route('/get_rollingshortlist', methods=["GET"])
def get_rollingshortlist():
    table_name = "Rolling Shortlist"
    job = """
    SELECT
        `Num`,
        `Captured date`,
        `BP comments`,
        `Partner_or_Director`,
        `Target country`,
        `Sub sector`,
        `Source`,
        ` Target`,
        `Business Description`,
        `Financials`,
        `Revenue_USD_M`,
        `EBITDA_USD_M`,
        `Valuation_USD_M`,
        `Other Info`,
        `Deal Intelligence info`,
        `News Date`,
        `KPMG View - Redacted`,
        `Credentials`,
        `HS contact`,
        `Investment date`,
        `Geographic region`,
        `Asset pack `,
        `Target Region`,
        `Shareholders`,
        `Lead type`,
        `Target HS industry classification `,
        `Stake for sale_percent`,
        `Value_USD_M`,
        `Value description `,
        `Website`,
    FROM `testing-bigquery-vertexai.templates.{table_name}`
    ORDER BY Num;
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

@app.route('/update_rollingshortlist', methods=["POST"])
def update_rollingshortlist():
    table_name = "Rolling Shortlist"
    data = request.get_json()
    success = True  # Assume success initially

    try:
        for row_data in data:
            row_id = row_data.get("row_id")
            updated_column = row_data.get("column_name")
            updated_value = row_data.get("edited_value")

            # Construct the UPDATE query dynamically
            update_query = f"""
            UPDATE `testing-bigquery-vertexai.templates.{table_name}`
            SET `{updated_column}` = '{updated_value}'
            WHERE Num = {row_id}
            """
            client.query(update_query)
            
    except Exception as e:
        print(e)
        success = False
    
    return jsonify({"success": success})


if __name__ == '__main__':
    app.run(debug=True, port=5003)