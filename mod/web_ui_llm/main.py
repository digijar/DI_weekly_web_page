"""
    This is a FastAPI application that facilitates the routing of the web pages and the uploading of the excel files to BigQuery.
    :return: The code is returning a FastAPI application instance named "app".
"""

from fastapi import FastAPI, UploadFile, Form, HTTPException, Request
import os
import pandas as pd
from google.cloud import bigquery
from fastapi.middleware.cors import CORSMiddleware
import uvicorn
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
import re
from pydantic import BaseModel
from typing import List
from typing_extensions import Annotated
from json import JSONDecodeError
import json

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Configure templates directory
templates = Jinja2Templates(directory="templates")

@app.get('/', response_class=HTMLResponse)
async def index(request: Request):
    # Render the index.html template
    return templates.TemplateResponse("index.html" , {"request":request})

# Function to clean column names
def clean_column_name(column_name, existing_names):
    # Remove any characters that are not allowed in BigQuery column names
    cleaned_name = re.sub(r"[^a-zA-Z0-9_]+", "_", column_name)

    # Check for duplicates and add a suffix if necessary
    counter = 2
    while cleaned_name in existing_names:
        cleaned_name = f"{column_name}_{counter}"
        counter += 1

    return cleaned_name

@app.post('/upload')
async def upload(file: List[UploadFile], tableIdInput: Annotated[str, Form()]):
    """
    This function allows users to upload Excel files, clean the column names, and
    upload the data to Google BigQuery.
    
    :param column_name: The `column_name` parameter is a string that represents the name of a column in
    a dataset
    :param existing_names: The `existing_names` parameter is a list of column names that already exist
    in the dataset. It is used in the `clean_column_name` function to check for duplicates and add a
    suffix if necessary
    """

    try:
        if not file:
            raise HTTPException(status_code=400, detail='No file selected')

        if not tableIdInput:
            raise HTTPException(status_code=400, detail='Table ID is required')

        # Iterate through uploaded files and read data using pandas
        for file_single in file:
            content = await file_single.read()
            with pd.io.common.BytesIO(content) as buffer:
                df = pd.read_excel(buffer, header=0)

            # Clean column names using the clean_column_name function
            existing_column_names = df.columns.tolist()
            cleaned_column_names = [
                clean_column_name(column_name, existing_column_names)
                for column_name in df.columns
            ]
            df.columns = cleaned_column_names

            # Upload the data to BigQuery
            client = bigquery.Client(project="testing-bigquery-vertexai")
            dataset_id = "web_UI"
            table_ref = client.dataset(dataset_id).table(tableIdInput)
            job_config = bigquery.LoadJobConfig()
            job_config.autodetect = True
            job = client.load_table_from_dataframe(df, table_ref, job_config=job_config)
            job.result()  # Wait for the job to complete

        return 'File(s) uploaded successfully!'
    except Exception as e:
        raise HTTPException(status_code=500, detail=f'Error uploading file: {str(e)}')

@app.get("/model_testing", response_class=HTMLResponse)
def model_testing(request: Request):
    return templates.TemplateResponse("model_testing.html", {"request":request})

@app.get("/preview", response_class=HTMLResponse)
def preview(request: Request):
    return templates.TemplateResponse("preview.html", {"request":request})

@app.get("/mergermarket", response_class=HTMLResponse)
def view_mergermarket(request: Request):
    return templates.TemplateResponse("mergermarket.html", {"request":request})

@app.get("/marketscan", response_class=HTMLResponse)
def view_marketscan(request: Request):
    return templates.TemplateResponse("marketscan.html", {"request":request})

@app.get("/rollingshortlist", response_class=HTMLResponse)
def view_rollingshortlist(request: Request):
    return templates.TemplateResponse("rollingshortlist.html", {"request":request})


if __name__ == '__main__':
    uvicorn.run("main:app", host='127.0.0.1', port=5000, reload=True)