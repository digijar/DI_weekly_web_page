from fastapi import FastAPI, Request, Response
import os
from google.cloud import bigquery
from json import JSONDecodeError
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from io import BytesIO
from openpyxl import Workbook
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

@app.get("/download_mergermarket")
def download_MM():

    """
        The above Python function downloads data from a BigQuery table and saves it as an Excel file for the
        user to download.
        :return: The code is returning a Response object with the content of the Excel file, headers for
        file download, and the media type set to
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'.
    """

    table_name = "MergerMarket"
    query = """
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

    # Run the BigQuery query
    job = client.query(query)

    # Fetch the query results
    result = job.result()

    # Extract data rows and header positions
    data_tuple = (
        [(row.Opportunity_ID, row.Date, row['Value INR_m'], row['Value Description'], row.Heading, row.Opportunity,
        row.Targets, row['Lead type'], row['Type of transaction'], row['HS sector classification'], row['Short BD'], row.Source,
        row['Intelligence Type'], row['Intelligence Grade'], row['Intelligence Size'], row['Stake Value'], row['Dominant Sector'],
        row.Sectors, row['Sub Sectors'], row['Dominant Geography'], row.Geography, row.States, row.Topics, row.Bidders,
        row.Vendors, row.Issuers, row.Competitors, row.Others, row.Completed)
        for row in result],
        {
            'Opportunity_ID': 0,
            'Date': 1,
            'Value INR_m': 2,
            'Value Description': 3,
            'Heading': 4,
            'Opportunity': 5,
            'Targets': 6,
            'Lead type': 7,
            'Type of transaction': 8,
            'HS sector classification': 9,
            'Short BD': 10,
            'Source': 11,
            'Intelligence Type': 12,
            'Intelligence Grade': 13,
            'Intelligence Size': 14,
            'Stake Value': 15,
            'Dominant Sector': 16,
            'Sectors': 17,
            'Sub Sectors': 18,
            'Dominant Geography': 19,
            'Geography': 20,
            'States': 21,
            'Topics': 22,
            'Bidders': 23,
            'Vendors': 24,
            'Issuers': 25,
            'Competitors': 26,
            'Others': 27,
            'Completed': 28
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
        'Content-Disposition': f'attachment; filename={"MergerMarket.xlsx"}',
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

@app.get('/get_mergermarket')
async def get_mergermarket():
    """
    The function `get_mergermarket` retrieves data from a BigQuery table named "MergerMarket" and
    returns it as a list of dictionaries.
    :return: The function `get_mergermarket` returns a list of dictionaries. Each dictionary represents
    a row from the `MergerMarket` table and contains the values for each column in that row.
    """

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

    return return_json


@app.post('/update_mergermarket')
async def update_mergermarket(request: Request):
    """
    The function `update_mergermarket` updates rows in the "MergerMarket" table in a BigQuery dataset
    based on the provided JSON data.
    
    :param request: The `request` parameter is an instance of the `Request` class, which represents an
    HTTP request made to the server. It contains information about the request, such as the HTTP method,
    headers, and body. In this case, the `request` object is used to retrieve the JSON data sent
    :type request: Request
    :return: a JSON response with a key-value pair indicating the success status of the update
    operation. The key is "success" and the value is a boolean indicating whether the update operation
    was successful or not.
    """

    table_name = "MergerMarket"

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
            WHERE `Opportunity_ID` = {row_id}
            """
            client.query(update_query)
            
    except Exception as e:
        print(e)
        success = False
    
    return {"success": success}

class MergerMarketRow(BaseModel):
    Opportunity_ID: float
    Date: date
    Value_INR_m: str
    Value_Description: str
    Heading: str
    Opportunity: str
    Targets: str
    Lead_Type: str
    Type_of_transaction: str
    HS_sector_classification: str
    Short_BD: str
    Source: str
    Intelligence_Type: str
    Intelligence_Grade: str
    Intelligence_Size: str
    Stake_Value: str
    Dominant_Sector: str
    Sectors: str
    Sub_Sectors: str
    Dominant_Geography: str
    Geography: str
    States: str
    Topics: str
    Bidders: str
    Vendors: str
    Issuers: str
    Competitors: str
    Others: str
    Completed: str

@app.post('/add_mergermarket_row')
async def add_mergermarket_row(request: Request, row_data: MergerMarketRow):
    """
    The `add_mergermarket_row` function inserts a row of data into a BigQuery table named "MergerMarket"
    using the provided `MergerMarketRow` object.
    
    :param request: The `request` parameter is of type `Request` and represents the HTTP request made to
    the server. It can be used to access information about the request, such as headers, query
    parameters, and request body
    :type request: Request
    :param row_data: The `row_data` parameter is an instance of the `MergerMarketRow` class. It contains
    the data for a single row that needs to be inserted into the "MergerMarket" table in BigQuery
    :type row_data: MergerMarketRow
    :return: a dictionary with a key "success" and a boolean value indicating whether the operation was
    successful or not.
    """
    
    table_name = "MergerMarket"

    try:
        # Initialize the BigQuery client
        client = bigquery.Client()

        # Convert Date to ISO formatted string
        formatted_date = row_data.Date.isoformat()

        # Construct the INSERT query dynamically
        insert_query = f"""
        INSERT INTO `testing-bigquery-vertexai.templates.{table_name}`
        (Opportunity_ID, Date, `Value INR_m`, `Value Description`, Heading, Opportunity, Targets,
        `Lead type`, `Type of transaction`, `HS sector classification`, `Short BD`, Source,
        `Intelligence Type`, `Intelligence Grade`, `Intelligence Size`, `Stake Value`, `Dominant Sector`,
        Sectors, `Sub Sectors`, `Dominant Geography`, Geography, States, Topics, Bidders, Vendors,
        Issuers, Competitors, Others, Completed)
        VALUES 
        ({row_data.Opportunity_ID}, '{formatted_date}', '{row_data.Value_INR_m}', '{row_data.Value_Description}',
        '{row_data.Heading}', '{row_data.Opportunity}', '{row_data.Targets}', '{row_data.Lead_Type}',
        '{row_data.Type_of_transaction}', '{row_data.HS_sector_classification}', '{row_data.Short_BD}',
        '{row_data.Source}', '{row_data.Intelligence_Type}', '{row_data.Intelligence_Grade}',
        '{row_data.Intelligence_Size}', '{row_data.Stake_Value}', '{row_data.Dominant_Sector}',
        '{row_data.Sectors}', '{row_data.Sub_Sectors}', '{row_data.Dominant_Geography}', '{row_data.Geography}',
        '{row_data.States}', '{row_data.Topics}', '{row_data.Bidders}', '{row_data.Vendors}', '{row_data.Issuers}',
        '{row_data.Competitors}', '{row_data.Others}', '{row_data.Completed}')
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
    uvicorn.run("mergermarket:app", host='127.0.0.1', port=5002, reload=True)