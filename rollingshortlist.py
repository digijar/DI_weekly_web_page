from fastapi import FastAPI, Request, Response
import os
from google.cloud import bigquery
from json import JSONDecodeError
import json
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
from pydantic import BaseModel
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
import datetime
from datetime import date

app = FastAPI()

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'testing-bigquery-vertexai-service-account.json'
client = bigquery.Client()

@app.get("/download_rollingshortlist")
def download_RS():
    table_name = "Rolling_Shortlist"
    query = """
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

    # Run the BigQuery query
    job = client.query(query)

    # Fetch the query results
    result = job.result()

    # Extract data rows and header positions
    data_tuple = (
        [
            (
                row.Num,
                row['Captured date'],
                row['BP comments'],
                row['Partner_or_Director'],
                row['Target country'],
                row['Sub sector'],
                row['Source'],
                row[' Target'],
                row['Business Description'],
                row['Financials'],
                row['Revenue_USD_M'],
                row['EBITDA_USD_M'],
                row['Valuation_USD_M'],
                row['Other Info'],
                row['Deal Intelligence info'],
                row['News Date'],
                row['KPMG View - Redacted'],
                row['Credentials'],
                row['HS contact'],
                row['Investment date'],
                row['Geographic region'],
                row['Asset pack '],
                row['Target Region'],
                row['Shareholders'],
                row['Lead type'],
                row['Target HS industry classification '],
                row['Stake for sale_percent'],
                row['Value_USD_M'],
                row['Value description '],
                row['Website'],
            )
            for row in result
        ],
        {
            'Num': 0,
            'Captured date': 1,
            'BP comments': 2,
            'Partner_or_Director': 3,
            'Target country': 4,
            'Sub sector': 5,
            'Source': 6,
            ' Target': 7,
            'Business Description': 8,
            'Financials': 9,
            'Revenue_USD_M': 10,
            'EBITDA_USD_M': 11,
            'Valuation_USD_M': 12,
            'Other Info': 13,
            'Deal Intelligence info': 14,
            'News Date': 15,
            'KPMG View - Redacted': 16,
            'Credentials': 17,
            'HS contact': 18,
            'Investment date': 19,
            'Geographic region': 20,
            'Asset pack ': 21,
            'Target Region': 22,
            'Shareholders': 23,
            'Lead type': 24,
            'Target HS industry classification ': 25,
            'Stake for sale_percent': 26,
            'Value_USD_M': 27,
            'Value description ': 28,
            'Website': 29,
        }
    )



    # Extract row values and header positions
    row_values, header_positions = data_tuple

    # Create a new Workbook
    wb = Workbook()
    ws = wb.active

    # Append the header row to the worksheet based on the header positions
    header_row = [None] * len(header_positions)
    for header, position in header_positions.items():
        header_row[position] = header
    ws.append(header_row)

    # Make the header row bold
    for cell in ws[1]:
        cell.font = Font(bold=True)

    # Append the data rows to the worksheet
    for row_data in row_values:
        # Check if row_data has at least 2 elements before converting date to string
        if len(row_data) >= 2 and isinstance(row_data[1], datetime.date):
            row_data = list(row_data)  # Convert tuple to list to allow modification
            row_data[1] = row_data[1].strftime('%Y-%m-%d')  # Convert date to string

        # Replace None values with an empty string
        row_data = ['' if value is None else value for value in row_data]

        # Replace '\r' and '\n' characters with line breaks in strings
        row_data = [value.replace('\r', '\n') if isinstance(value, str) else value for value in row_data]

        # Append the preprocessed row_data to the worksheet
        ws.append(row_data)

    # Create an in-memory file-like object to save the workbook
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    # Set response headers for file download
    headers = {
        'Content-Disposition': f'attachment; filename={"Rolling Shortlist.xlsx"}',
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Headers": "*",
        "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
    }
    media_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

    return Response(
        content=output.read(),
        headers=headers,
        media_type=media_type
    )


@app.get('/get_rollingshortlist')
async def get_rollingshortlist():
    table_name = "Rolling_Shortlist"
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

    return return_json

@app.post('/update_rollingshortlist')
async def update_rollingshortlist(request: Request):
    table_name = "Rolling_Shortlist"

    try:
        data = await request.json()

    except JSONDecodeError:
        return 'Invalid JSON data.'
    
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
            WHERE `Num` = {row_id}
            """
            client.query(update_query)
            
    except Exception as e:
        print(e)
        success = False
    
    return {"success": success}

class RollingShortlistRow(BaseModel):
    Num: float
    Captured_date: float
    BP_comments: str
    Partner_or_Director: str
    Target_country: str
    Sub_sector: str
    Source: str
    Target: str
    Business_Description: str
    Financials: str
    Revenue_USD_M: str
    EBITDA_USD_M: str
    Valuation_USD_M: str
    Other_Info: str
    Deal_Intelligence_info: str
    News_Date: float
    KPMG_View_Redacted: str
    Credentials: str
    HS_contact: str
    Investment_date: str
    Geographic_region: str
    Asset_pack: str
    Target_Region: str
    Shareholders: str
    Lead_type: str
    Target_HS_industry_classification: str
    Stake_for_sale_percent: str
    Value_USD_M: str
    Value_description: str
    Website: str

@app.post('/add_rollingshortlist_row')
async def add_rollingshortlist_row(request: Request, row_data: RollingShortlistRow):
    table_name = "Rolling_Shortlist"

    try:
        # Initialize the BigQuery client
        client = bigquery.Client()

        # Construct the INSERT query dynamically
        insert_query = f"""
        INSERT INTO `testing-bigquery-vertexai.templates.{table_name}`
        (Num, `Captured date`, `BP comments`, `Partner_or_Director`, `Target country`, `Sub sector`, `Source`, ` Target`, `Business Description`, `Financials`, `Revenue_USD_M`, `EBITDA_USD_M`, `Valuation_USD_M`, `Other Info`, `Deal Intelligence info`, `News Date`, `KPMG View - Redacted`, `Credentials`, `HS contact`, `Investment date`, `Geographic region`, `Asset pack `, `Target Region`, `Shareholders`, `Lead type`, `Target HS industry classification `, `Stake for sale_percent`, `Value_USD_M`, `Value description `, `Website`) VALUES 
        ({row_data.Num}, {row_data.Captured_date}, '{row_data.BP_comments}', '{row_data.Partner_or_Director}', 
        '{row_data.Target_country}', '{row_data.Sub_sector}', '{row_data.Source}', '{row_data.Target}', 
        '{row_data.Business_Description}', '{row_data.Financials}', '{row_data.Revenue_USD_M}', 
        '{row_data.EBITDA_USD_M}', '{row_data.Valuation_USD_M}', '{row_data.Other_Info}', 
        '{row_data.Deal_Intelligence_info}', {row_data.News_Date}, '{row_data.KPMG_View_Redacted}', 
        '{row_data.Credentials}', '{row_data.HS_contact}', '{row_data.Investment_date}', 
        '{row_data.Geographic_region}', '{row_data.Asset_pack}', '{row_data.Target_Region}', 
        '{row_data.Shareholders}', '{row_data.Lead_type}', '{row_data.Target_HS_industry_classification}', 
        '{row_data.Stake_for_sale_percent}', '{row_data.Value_USD_M}', '{row_data.Value_description}', 
        '{row_data.Website}')
        """

        # Run the query
        query_job = client.query(insert_query)

        # Wait for the query to complete (optional)
        query_job.result()

        return {"success": True}

    except Exception as e:
        print(e)
        return {"success": False}

if __name__ == '__main__':
    import uvicorn
    uvicorn.run("preview:app", host='127.0.0.1', port=5004, reload=True)