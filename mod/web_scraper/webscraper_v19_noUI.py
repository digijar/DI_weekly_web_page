import requests
from fastapi import FastAPI, Request, Response
from google.cloud import bigquery
from json import JSONDecodeError
import json
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

'''import webscraper libraries'''
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup
import time
from datetime import datetime
from urllib.parse import urlparse, parse_qs

username = 'smustu\elijah.khor.2021'
password = 'AyabeKyoto1!'

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--headless')
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument("--no-first-run")
chrome_options.add_argument("--no-default-browser-check")
chrome_options.add_argument('--start-maximized')

app = FastAPI()

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def get_orbis (company, final_dict, service_instance):
    '''This function takes in the name of a company, the final_dict dictionary and returns a dictionary of up to 6 items, of each account 
        and its value as the key and value respectively from the Orbis platform.

        The accounts scraped from Orbis are:
        - Operating revenue (Turnover)
        - Costs of goods sold
        - Gross profit
        - EBITDA
        - Enterprise value
        - Date (Date of records)

        These will be returned in a dictionary with the respective keys (also the same keys returned by CapitalIQ):
        - Revenue
        - COGS
        - Gross Profit
        - EBITDA
        - Valuation
        - Date

        Do note that not all 6 items may be returned, as the availability of information may vary from company to company.

        Inputs: name of company
        Output: dictionary of key:value - Account: Amount
    '''
    # columns mapping of Orbis account names - to standardize results across Orbis and Capital IQ
    orbis_mapping = {'Enterprise value':'Valuation', 'Operating revenue (Turnover)':'Revenue', 'Costs of goods sold':'COGS', 'Gross profit':'Gross Profit', 'EBITDA':'EBITDA'}
    

    url = 'http://libproxy.smu.edu.sg/login?url=https://orbis4.bvdinfo.com/ip'
    orbis = webdriver.Chrome(service=service_instance, options=chrome_options)
    orbis.get(url)

    un_locator = (By.ID, 'userNameInput')
    # wait maximum 5 seconds for username input to load
    un = WebDriverWait(orbis, 5).until(EC.presence_of_element_located(un_locator))
    un.send_keys(username)
    pwd = orbis.find_element(By.ID, 'passwordInput')
    pwd.send_keys(password)
    pwd.send_keys(Keys.RETURN)

    # default wait time - 5 seconds
    wait = WebDriverWait(orbis, 5)

    # Close any new popups that appear
    try:
        # check if popup overlay present
        popup_overlay_locator = (By.ID, 'walkme-overlay-all')
        popup_overlay = wait.until(EC.presence_of_element_located(popup_overlay_locator))
        # try block will continue if popup overlay is found (not timeout from waiting)
        close_btn_locator = (By.CSS_SELECTOR, 'div[class="walkme-x-button wm-close-button"]')
        close_btn = wait.until(EC.presence_of_element_located(close_btn_locator))
        close_btn.click()
    except:
        pass
    
    try:
        # wait for search box to be found, up till 5 seconds
        company_input_locator = (By.ID, 'search')
        company_input = wait.until(EC.presence_of_element_located(company_input_locator))
        company_input.send_keys(company)

        # wait for suggestions to load, up till 5 seconds
        suggestions_locator = (By.CLASS_NAME, 'suggestions')
        suggestions = wait.until(EC.presence_of_element_located(suggestions_locator))

        time.sleep(3)
        # wait for first suggestion to load, up till 5 seconds
        first_locator = (By.CSS_SELECTOR, 'a[role="link"]')
        first_three_suggestions = orbis.find_elements(By.CSS_SELECTOR, 'a[role="link"]')[:3]
        
        # click on first suggestion
        first_three_suggestions[0].click()
    except:
        # no suggestions found, close browser and return None
        orbis.quit()
        return final_dict
    
    # initalise bvd_id list and fill with top 3 suggestions
    bvd_list = []
    for x in first_three_suggestions:
        bvd_id = x.get_attribute('data-parambvdid')
        bvd_list.append(bvd_id)
    
    # initialise list to collect data of top 3 suggestions
    top_3_suggestions =[]
    
    # start with first bvd_id
    bvd_id = bvd_list[0]
    
    # define subfunction to collect data from each of the suggestions
    def orbis_info (bvd_id):
        # nonlocal orbis to use the same browser
        nonlocal orbis
        
        # initalise Other_info str with bvd_id
        asset_pack_str = "Orbis_BvdID:" + bvd_id + ";"

        # initalise raw list with bvd_Id in asset_pack to collect the data
        raw = [('Asset Pack', asset_pack_str)]
    
        # check if navigation sidebar is closed, if it is, click on it to open
        try:
            sidebar_locator = (By.CLASS_NAME, 'side-expanded--closed')
            sidebar = wait.until(EC.presence_of_element_located(sidebar_locator))
            sidebar.click()
        except:
            pass
    
        # Get Region - Find chapter search input
        search_locator = (By.CSS_SELECTOR,'input[name="titleInput"]')
        search = wait.until(EC.presence_of_element_located(search_locator))
        search.send_keys("Geographic footprint")

        # find search suggestions for geographic footprint
        geographic = orbis.find_element(By.CSS_SELECTOR,'li[class="search-result section-inBook candidate"]').find_element(By.CSS_SELECTOR,'a[title="Geographic footprint"]')
        geographic.click()
    
        # check if its in table form, if not, click to change to table form
        try:
            list_toggle_locator = (By.CSS_SELECTOR,'a[aria-label="Show list"]')
            list_toggle = wait.until(EC.presence_of_element_located(list_toggle_locator))
            list_toggle.click()
        except:
            pass
    
        time.sleep(2)
        # wait for countries table to load
        countries_locator = (By.CSS_SELECTOR,'table[class="ETBL ownership-table heatmap-table"]')
        countries_table = wait.until(EC.presence_of_element_located(countries_locator))

        soup = BeautifulSoup(orbis.page_source, 'lxml')
    
        try:
            countries = soup.find_all('td', {'class':'ownership-table__no-left-border ownership-table--left heatmap-table__element'})
            region_str = ''
            for x in countries:
                region_str += x.get_text(strip=True) + ';'
            if len(region_str) != 0:
                region_str = region_str[:-1]
                raw.append(('Target Region', region_str))
        except:
            pass
    
        # Find the header element using its attributes to expand
        header_locator = (By.CSS_SELECTOR,'a[title="Financials"]')
        header = wait.until(EC.presence_of_element_located(header_locator))

        # check if header expanded, if no then click on it
        if header.get_attribute('aria-expanded') == 'false':
            element = header.find_element(By.CLASS_NAME, 'menu__view-selection-item-icon')
            # expand header
            element.click()

        # find global standard format and click on it    
        element = orbis.find_element(By.CSS_SELECTOR,'a[title="Global standard format"]')
        element.click()
        time.sleep(5)
        # use beautiful soup to parse the page 
        soup = BeautifulSoup(orbis.page_source, 'lxml')

        # try to find the financial data table
        datatable = soup.find('table',{'class':'FinDataTable'})

        # if no table is found, there is no data - close browser and return None
        if datatable is None:
            return dict(raw)

        # find latest date and get it in the format of DD MMM YYYY e.g. 27 Jul 2023
        date_obj = soup.find('table',{'class':'FinDataTable'}).find('tr', {'class': 'finHead'}).find('p').get_text(strip=True).split('<br/>')[0].replace('USD', '').replace('th', '').strip()
        try:
            date = datetime.strptime(date_obj, '%d/%m/%Y').strftime('%d %b %Y')
        except:
            date = datetime.strptime(date_obj, '%Y').strftime('%Y')

        # find all rows in datatable
        table = datatable.find_all('tr', class_=['fin1', 'fin2', 'fin3'])

        # initialise list to collect the tuples of data
        raw.append(('Date', date))

        # initalise other info string
        other_info_str = ''

        # initalise other info accounts to be put into the string
        other_info_list = ['COGS', 'Gross Profit']

        na_values = ['-', '0', 'n.a.']

        # iterate through rows
        for x in table:
            # find all account headers, which are in the first cell position, index 0
            label_element = x.find_all('td')
            label = label_element[0].find('div').get_text(strip=True).replace("∟", "").strip()

            # check if it is one of the metrics mapped in the mapping dictionary
            if label in orbis_mapping:
                label = orbis_mapping[label]
            else:
                continue

            # get value of account and convert to float 
            # for Orbis, the value in 2nd cell position is the value from the latest recording date, index 1
            value = label_element[1].get_text(strip=True)
            if value in na_values:
                continue
            else:
                float_value = float(value.replace(',',''))

            if label not in other_info_list:
                # append the account label and value to list
                raw.append((label, float_value))
            else:
                other_info_str += label + ': ' + str(float_value) + ', '

        # append other info to list if its not empty
        if len(raw) >1:
            other_info_str += 'Units: USD Thousands (k)'

        if other_info_str != '':
            raw.append(('Other Info', other_info_str))
        # convert list to dictionary
        result = dict(raw)
        
        return result
    
    # get result of 1st suggestion
    r1 = orbis_info(bvd_id)
    top_3_suggestions.append(r1)
    
    print(top_3_suggestions)
    
    # define subfunction to search next 2 bvd_ids
    def new_bvdID_search(bvd_id):
        nonlocal orbis
        homepage = 'https://orbis-r1-bvdinfo-com.libproxy.smu.edu.sg/'
        # go to homepage
        orbis.get(homepage)
        # wait for search box to be found, up till 5 seconds
        company_input_locator = (By.ID, 'search')
        company_input = wait.until(EC.presence_of_element_located(company_input_locator))
        company_input.send_keys(bvd_id)

        # wait for suggestions to load, up till 5 seconds
        suggestions_locator = (By.CLASS_NAME, 'suggestions')
        suggestions = wait.until(EC.presence_of_element_located(suggestions_locator))

        time.sleep(3)
        # wait for first suggestion to load, up till 5 seconds
        first_locator = (By.CSS_SELECTOR, 'a[role="link"]')
        first = wait.until(EC.presence_of_element_located(first_locator))
        first.click()

        # checking if popup for replace search appear
        try:
            time.sleep(5)
            popup = orbis.find_element(By.CSS_SELECTOR,'div[class="popup popup__dialog new-search"]')
            buttons = popup.find_element(By.CSS_SELECTOR,'div[class="button popup__buttons"]')
            ok_button = buttons.find_element(By.CSS_SELECTOR,'a[class="button submit ok"]')
            ok_button.click()
        except:
            pass
    
    for i in bvd_list[1:]:
        new_bvdID_search(i)
        res = orbis_info(i)
        top_3_suggestions.append(res)
        print(res)
    
    print(top_3_suggestions)
    
    # select best result based on number of items and order of search results
    # initialise max_len
    max_len = 0
    for x in top_3_suggestions:
        if len(x) > max_len:
            max_len = len(x)
    
    for y in top_3_suggestions:
        if len(y) == max_len:
            result = y
            break
    
    final_dict.update(result)
    
    # close browser
    orbis.quit()
    
    return final_dict


def get_data_capital(table):
    '''This is a helper function to scrape data from Capital IQ specifically, as the format of the datatable is more complicated
        than Orbis. 
        The input is a list of <a role='clickThru'> elements from the datatable that contain the required metrics,
        which has been filtered prior to this function. The output will be a dictionary of up to 6 items, of each account 
        and its value as the key and value respectively.

        This function will be used twice, for the Income Statement page and the Capitalization page. The Income statement page
        contains most of the information below, except Total Enterprise Value (TEV)/Total Equity, which we will get from the
        Capitalization page.
        
        The accounts scraped from Capital IQ are:
        - Revenue OR Total Revenue
        - Cost Of Goods Sold
        - Gross Profit
        - EBITDA
        - Total Enterprise Value (TEV) OR Total Equity
        - Date (Date of records)

        Note that for Revenue/Total Revenue and Total Enterprise Value (TEV)/Total Equity, only one of each will be returned as
        some companies only have 1 or the other.

        These will be returned in a dictionary with the respective keys (also the same keys returned by Orbis):
        - Revenue
        - COGS
        - Gross Profit
        - EBITDA
        - Valuation
        - Date

        Similarly, not all 6 items may be returned, as the availability of information may vary from company to company.

        Input: list of <a> elements
        Output: dictionary of key:value - Account: Amount
    '''
    # columns mapping of Capital IQ account names - to standardize results across Orbis and Capital IQ
    capitaliq_mapping ={'Total Revenue':'Revenue', 'Revenue':'Revenue', 'Cost Of Goods Sold':'COGS', 'Gross Profit':'Gross Profit', 'EBITDA':'EBITDA', 'Total Enterprise Value (TEV)':'Valuation', 'Total Equity':'Valuation'}
    
    # initalise list to collect extracted data from <a> elements
    raw = []

    # loop through <a> elements and append split titles into the list
    for x in table:
        try:
            raw.append(x['title'].split('\n'))
        except:
            continue
    
    # add account header
    for x in raw:
        if x[0]:
            x[0] = 'Account: ' + x[0]

    # initialise list to collect data
    data = []

    for row in raw:
        data_dict = {}
        for pair in row:
            # split embedded kev:value pairs if exist (see 'Value, Currency, Millions' example above)
            record = pair.strip().split(', ')

            # single key:value pair
            if len(record)==1:
                try:
                    k, v = record[0].split(':')
                except:
                    continue
                else:
                    data_dict[k.strip()] = v.strip()

            # multiple key:value pairs
            else:
                for x in record:
                    if x.count(':')==0:
                        continue
                    else:
                        k, v = x.split(':')
                        data_dict[k.strip()] = v.strip()

        data.append(data_dict)
        
    key_list = []
    for row in data:
        key_list.extend(row.keys())

    key_list = set(key_list)
    
    for row in data:
        for key in key_list:
            if key not in row:
                row[key] = None
    
    other_info_str = ''
    other_info_list = ['COGS', 'Gross Profit']

    result = {}
    for x in data:
        if x["Account"] in capitaliq_mapping and capitaliq_mapping[x["Account"]] not in result:
            if capitaliq_mapping[x["Account"]] not in other_info_list:
                result[capitaliq_mapping[x["Account"]]] = float(x["Value"].replace(',', ''))
            else:
                other_info_str += capitaliq_mapping[x["Account"]] + ': ' + str(float(x["Value"].replace(',', ''))) + ', '
    try:
        latest_date = data[0]['Filing Date']
        latest_date_formatted = datetime.strptime(latest_date, "%b-%d-%Y").strftime("%d %b %Y")
        result['Date'] = latest_date_formatted
    except:
        pass
    
    # as long as there is more than just date - ie there are financials, add in 
    if len(result) >1:
        other_info_str += 'Units: USD Thousands (k)'
    
    if other_info_str != '':
        result['Other Info'] = other_info_str
    return result

def get_capitaliq (company, final_dict, service_instance):
    '''
        This function searches the name of a company and gathers the required financial information from Captial IQ. It will get a
        list of <a> elements containing the financial information of the latest record date and utilize the get_data_capital 
        helper function to parse and get the dictionary.

        The helper function get_data_capital will be used twice for the Income statement page and Capitalization Page to collect all
        the required information.

        The accounts scraped from Capital IQ are:
        - Revenue OR Total Revenue
        - Cost Of Goods Sold
        - Gross Profit
        - EBITDA
        - Total Enterprise Value (TEV) OR Total Equity
        - Date (Date of records)
        
        Note that for Revenue/Total Revenue and Total Enterprise Value (TEV)/Total Equity, only one of each will be returned as
        some companies only have 1 or the other.

        These will be returned in a dictionary with the respective keys (also the same keys returned by Orbis):
        - Revenue
        - COGS
        - Gross Profit
        - EBITDA
        - Valuation
        - Date

        Similarly, not all 6 items may be returned, as the availability of information may vary from company to company.

        Input: name of company
        Output: dictionary of key:value - Account: Amount
       
    '''    
    # Captial IQ
    # Login via SMU credentials
    url = 'https://login.spglobal.com/oamfed/sp/initiatesso?providerid=IDP_SMU&returnurl=https://www.capitaliq.com/CIQDotNet/saml-sso.aspx'
    capitaliq = webdriver.Chrome(service=service_instance, options=chrome_options)
    capitaliq.get(url)
    un_locator = (By.ID, 'userNameInput')
    un = WebDriverWait(capitaliq, 5).until(EC.presence_of_element_located(un_locator))
    un.send_keys(username)
    pwd = capitaliq.find_element(By.ID, 'passwordInput')
    pwd.send_keys(password)
    pwd.send_keys(Keys.RETURN)

    wait = WebDriverWait(capitaliq, 5)
    # search for company
    try:
        company_input_locator = (By.CLASS_NAME,'cSearchBoxBorderMiddle')
        company_input = wait.until(EC.presence_of_element_located(company_input_locator)).find_element(By.TAG_NAME,'input')
        company_input.send_keys(company)

        

        suggestions_locator = (By.CSS_SELECTOR,'div[class="acResults regularAutoCompleteSearch "]')
        suggestions = wait.until(EC.presence_of_element_located(suggestions_locator))

        time.sleep(5)
        first_3_suggestions = suggestions.find_elements(By.CSS_SELECTOR,'a[class="acResultLink"]')[:3]

        # initialise list of links
        links = []
        # append each link to list
        for x in first_3_suggestions:
            link = x.get_attribute('href')
            links.append(link)
    except:
        capitaliq.quit()
        return final_dict
    
    top_3_suggestions = []
    
    # define subfunction to get info
    
    def capiq_info(link):
        nonlocal capitaliq

        capitaliq.get(link)
        parsed_url = urlparse(link)
        query_params = parse_qs(parsed_url.query)

        company_id = query_params.get('companyId', [''])[0]
        if company_id =='':
            return {}

        asset_info = 'CapIQ_CompanyID:'+company_id+';'

        # initialise raw dict to store data
        raw = {}

        # Get numOfEmployees, yearFounded and Business Description, webSite

        # wait for presence of table, wait for it to load finish
        tables_locator = (By.CSS_SELECTOR, 'table[class="cTblListBody"]')
        tables = wait.until(EC.presence_of_element_located(tables_locator))

        time.sleep(3)
        soup = BeautifulSoup(capitaliq.page_source, 'lxml')
        info = soup.find_all('td', id=['numOfEmployees', 'yearFounded', 'webSite'])

        na_values = ['-', 'n.a.', '']

        for i in info:
            label = i.get('id')
            val = i.get_text(strip=True)
            if val not in na_values:
                if label == 'webSite':
                    raw['Website'] = val
                else:
                    asset_info+= label + ':' + val + ';'

        if asset_info != '':        
            raw['Asset Pack'] = asset_info[:-1]

        biz_desc_table = soup.find_all('table', {'class':'cTblListBody'})[1]
        biz_desc = biz_desc_table.find('span').get_text(strip=True)

        if biz_desc != '':
            raw['Business Description'] = biz_desc

        statement = ['IncomeStatement', 'Capitalization']
        url = 'https://www.capitaliq.com/CIQDotNet/Financial/{}.aspx?companyId={}'

        for i in statement:
            capitaliq.get(url.format(i, company_id))
            time.sleep(2)

            # change currency and units to USD and Thousands (k)
            # look for more options expander
            try:
                more_options_locator = (By.ID, '_pageHeader_ShowMoreLink')
                more_options = wait.until(EC.presence_of_element_located(more_options_locator))
                more_options.click()

                # change to USD
                currency_dropdown_locator = (By.CSS_SELECTOR, 'select[id="_pageHeader_fin_dropdown_currency"]')
                currency_dropdown = wait.until(EC.presence_of_element_located(currency_dropdown_locator))

                currency_select = Select(currency_dropdown)
                currency_select.select_by_visible_text('US Dollar')

                # change to Thousands (k)
                units_dropdown_locator = (By.CSS_SELECTOR, 'select[id="_pageHeader_fin_dropdown_units"]')
                units_dropdown = wait.until(EC.presence_of_element_located(units_dropdown_locator))

                units_select = Select(units_dropdown)
                units_select.select_by_visible_text('Thousands (k)')

                # submit filters
                submit_btn = capitaliq.find_elements(By.CSS_SELECTOR,'td[class=cTblFuncTxt]')[-2].find_element(By.CSS_SELECTOR, 'input[type="submit"]')
                submit_btn.click()
            except:
                continue

            time.sleep(5)
            soup = BeautifulSoup(capitaliq.page_source, 'lxml')
            datatable = soup.find('table',{'class':'FinancialGridView'})
            if datatable is None:
                print('datatable empty')
                continue

            row = datatable.find_all('tr')
            table = []
            for x in row:
                data = x.find_all('a',{'class':'clickThru'})
                if data:
                    table.append(data[-1])
            try:
                res = get_data_capital(table)
                raw.update(res)
            except:
                pass

        return raw
    
    for x in links:
        res = capiq_info(x)
        top_3_suggestions.append(res)
        
    # select best result based on number of items and order of search results
    # initialise max_len
    max_len = 0
    for x in top_3_suggestions:
        if len(x) > max_len:
            max_len = len(x)

    for y in top_3_suggestions:
        if len(y) == max_len:
            result = y
            break
    
    # extract out website and business description, which will be updated no matter what, Asset pack will be updated separately
    keys_to_extract = ["Website", "Business Description", "Asset Pack"]

    # only website and business description are updated immediately
    new_cols_dict = {key: result[key] for key in keys_to_extract[:2]}
    final_dict.update(new_cols_dict)

    # addon asset pack in final_dict
    if 'Asset Pack' in final_dict:
        final_dict['Asset Pack'] += result['Asset Pack']
    else:
        final_dict['Asset Pack'] = result['Asset Pack']

    # extract out remaining columns for date checking
    checking_dict = {key: result[key] for key in result if key not in keys_to_extract}

    cap_iq_date = checking_dict['Date']
    if 'Date' in final_dict:
        orbis_date = final_dict['Date']

        try:
            cap_iq_date = datetime.strptime(cap_iq_date, "%d %b %Y")
            orbis_date = datetime.strptime(orbis_date, "%d %b %Y")
            if cap_iq_date >orbis_date:
                final_dict.update(checking_dict)
        except:
            pass
    else:
        final_dict.update(checking_dict)
    capitaliq.quit()

    return final_dict

def get_company(company):
    service_instance = Service()
    company = company.strip()
    
    # initialise final_dict json
    final_dict = {}
    
    try:
        final_dict = get_orbis(company, final_dict, service_instance)
        print("================ final dict has orbis information =====================")
        print(final_dict)
        print(len(final_dict))
        raise orbis_exception
    except Exception as orbis_exception:
        try:
            final_dict = get_capitaliq(company, final_dict, service_instance)

            print("================ final dict has capital IQ information =====================")
            print(final_dict)
            print(len(final_dict))
            service_instance.stop()
            return ({company:final_dict})
        except:

            print("================ ran into exception case =====================")
            print(final_dict)
            print(len(final_dict))
            service_instance.stop()
            return ({company: final_dict})
        

microservice_url = "http://127.0.0.1:5011"

@app.get('/scrape/{row_num}')
async def webscraping(row_num:float):
    # Make an HTTP GET request to the endpoint with row_num in the path
    response = requests.get(f"{microservice_url}/bq/{row_num}")

    # Parse the JSON response
    response_data = response.json()

    # Extract the data you need, for example, the company_name
    company_name = response_data.get("company_name")
    print(company_name)
    print('Scraping...')
    # Scrape data
    scraped_data = get_company(company_name)

    # Package in JSON
    update_data = {
        "num": row_num,
        "target": company_name,
        "scraped_data": scraped_data
    }
    ## send back to retrieve to update
    update_response = requests.post(f"{microservice_url}/update", json=update_data)
    success = True
    if update_response.status_code == 200:
        # Check if the update was successful
        update_result = update_response.json()
        if update_result.get("success"):
            print("Data update successful.")
        else:
            print("Data update failed.")
            success = False
        return {"Data update successful?": success}
    else:
        print(f"Update request failed with status code: {update_response.status_code}")

if __name__ == '__main__':
    import uvicorn
    uvicorn.run("webscraper_v18_noUI:app", host='127.0.0.1', port=5009, reload=True)