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

@app.get("/download_marketscan")
def download_MS():
    table_name = "Market_Scan"
    query = """
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

    # Run the BigQuery query
    job = client.query(query)

    # Fetch the query results
    result = job.result()

    # Extract data rows and header positions
    data_tuple = (
        [
            (
                row.Num,
                row.Picks,
                row['Captured in the week'],
                row.Date,
                row['Lead type'],
                row['Target HS industry classification'],
                row['Dominant sector_as per MM'],
                row['Sectors_as per MM'],
                row['Sub Sectors_as per MM'],
                row['HS Target region'],
                row['Dominant country'],
                row.Countries,
                row['Next Steps'],
                row.Target,
                row['Target_Chinese'],
                row['Deal intelligence'],
                row['Short BD'],
                row.Bidders,
                row['Sellers_or_Vendors'],
                row['Type of transaction'],
                row['Intelligence type_as per MM'],
                row['Topics_as per MM'],
                row['Value_USDm'],
                row['Value description'],
                row['Intelligence size'],
                row['Intelligence size bucket'],
                row['Stake value_percent'],
                row['Held by ASPAC priority firm_Y_or_N'],
                row['PE priority firm '],
                row['Held since'],
                row['Held for more than three years_Y_N_NA'],
                row['KPMG credentials_PE_Company_Both_N'],
                row['KPMG firm'],
                row['Engagement partner']
            )
            for row in result
        ],
        {
            'Num': 0,
            'Picks': 1,
            'Captured in the week': 2,
            'Date': 3,
            'Lead type': 4,
            'Target HS industry classification': 5,
            'Dominant sector_as per MM': 6,
            'Sectors_as per MM': 7,
            'Sub Sectors_as per MM': 8,
            'HS Target region': 9,
            'Dominant country': 10,
            'Countries': 11,
            'Next Steps': 12,
            'Target': 13,
            'Target_Chinese': 14,
            'Deal intelligence': 15,
            'Short BD': 16,
            'Bidders': 17,
            'Sellers_or_Vendors': 18,
            'Type of transaction': 19,
            'Intelligence type_as per MM': 20,
            'Topics_as per MM': 21,
            'Value_USDm': 22,
            'Value description': 23,
            'Intelligence size': 24,
            'Intelligence size bucket': 25,
            'Stake value_percent': 26,
            'Held by ASPAC priority firm_Y_or_N': 27,
            'PE priority firm ': 28,
            'Held since': 29,
            'Held for more than three years_Y_N_NA': 30,
            'KPMG credentials_PE_Company_Both_N': 31,
            'KPMG firm': 32,
            'Engagement partner': 33
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
        'Content-Disposition': f'attachment; filename={"Market Scan.xlsx"}',
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


@app.get('/get_marketscan')
async def get_marketscan():
    table_name = "Market_Scan"
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

    return return_json

@app.post('/update_marketscan')
async def update_marketscan(request: Request):
    table_name = "Market_Scan"

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

            if updated_column == "Picks":
                if updated_value == "True":
                    update_query = (
                        f'UPDATE `testing-bigquery-vertexai.templates.{table_name}'
                        'SET Picks = TRUE '
                        f'WHERE Num = {row_id}')

            else: 
                update_query = f"""
                    UPDATE `testing-bigquery-vertexai.templates.{table_name}`
                    SET `{updated_column}` = `{updated_value}`
                    WHERE `Num` = {row_id}
                    """

            client.query(update_query)
            
    except Exception as e:
        print(e)
        success = False
    
    return {"success": success}

class MarketScanRow(BaseModel):
    Num: float
    Picks: bool
    Captured_in_the_week: str
    Date: float
    Lead_type: str
    Target_HS_industry_classification: str
    Dominant_sector_as_per_MM: str
    Sectors_as_per_MM: str
    Sub_Sectors_as_per_MM: str
    HS_Target_region: str
    Dominant_country: str
    Countries: str
    Next_Steps: str
    Target: str
    Target_Chinese: str
    Deal_intelligence: str
    Short_BD: str
    Bidders: str
    Sellers_or_Vendors: str
    Type_of_transaction: str
    Intelligence_type_as_per_MM: str
    Topics_as_per_MM: str
    Value_USDm: str
    Value_description: str
    Intelligence_size: str
    Intelligence_size_bucket: str
    Stake_value_percent: str
    Held_by_ASPAC_priority_firm_Y_or_N: str
    PE_priority_firm: str
    Held_since: str
    Held_for_more_than_three_years_Y_N_NA: str
    KPMG_credentials_PE_Company_Both_N: str
    KPMG_firm: str
    Engagement_partner: str

@app.post('/add_marketscan_row')
async def add_marketscan_row(request: Request, row_data: MarketScanRow):
    table_name = "Market_Scan"

    try:
        # Initialize the BigQuery client
        client = bigquery.Client()

        # Construct the INSERT query dynamically
        insert_query = f"""
        INSERT INTO `testing-bigquery-vertexai.templates.{table_name}`
        (Num, Picks, `Captured in the week`, Date, `Lead type`, `Target HS industry classification`,
        `Dominant sector_as per MM`, `Sectors_as per MM`, `Sub Sectors_as per MM`, `HS Target region`,
        `Dominant country`, Countries, `Next Steps`, Target, Target_Chinese, `Deal intelligence`,
        `Short BD`, Bidders, Sellers_or_Vendors, `Type of transaction`, `Intelligence type_as per MM`,
        `Topics_as per MM`, `Value_USDm`, `Value description`, `Intelligence size`, `Intelligence size bucket`,
        `Stake value_percent`, `Held by ASPAC priority firm_Y_or_N`, `PE priority firm `, `Held since`,
        `Held for more than three years_Y_N_NA`, `KPMG credentials_PE_Company_Both_N`, `KPMG firm`,
        `Engagement partner`)
        VALUES 
        ({row_data.Num}, {row_data.Picks}, '{row_data.Captured_in_the_week}', {row_data.Date}, 
        '{row_data.Lead_type}', '{row_data.Target_HS_industry_classification}', '{row_data.Dominant_sector_as_per_MM}',
        '{row_data.Sectors_as_per_MM}', '{row_data.Sub_Sectors_as_per_MM}', '{row_data.HS_Target_region}',
        '{row_data.Dominant_country}', '{row_data.Countries}', '{row_data.Next_Steps}', '{row_data.Target}',
        '{row_data.Target_Chinese}', '{row_data.Deal_intelligence}', '{row_data.Short_BD}', '{row_data.Bidders}',
        '{row_data.Sellers_or_Vendors}', '{row_data.Type_of_transaction}', '{row_data.Intelligence_type_as_per_MM}',
        '{row_data.Topics_as_per_MM}', '{row_data.Value_USDm}', '{row_data.Value_description}', 
        '{row_data.Intelligence_size}', '{row_data.Intelligence_size_bucket}', '{row_data.Stake_value_percent}',
        '{row_data.Held_by_ASPAC_priority_firm_Y_or_N}', '{row_data.PE_priority_firm}', '{row_data.Held_since}',
        '{row_data.Held_for_more_than_three_years_Y_N_NA}', '{row_data.KPMG_credentials_PE_Company_Both_N}',
        '{row_data.KPMG_firm}', '{row_data.Engagement_partner}')
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
    uvicorn.run("marketscan:app", host='127.0.0.1', port=5003, reload=True)