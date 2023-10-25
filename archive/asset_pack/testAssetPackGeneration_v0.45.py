# # install comtypes, google-cloud-bigquery
# pip install comtypes --trusted-host=files.pythonhosted.org

import comtypes.client
# to get current date
from datetime import datetime
# connection to google cloud bigquery
import os
from google.cloud import bigquery
# text processing
import re



def retrieve_bigquery(row_num):
    # # Connection to Big Query - get authorisation from GCP to access Big Query
    json_path = "./testing_bigquery_vertexai_service_account.json"
    json_abs_path = os.path.abspath(json_path)
    os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = json_abs_path

    # Access Big Query
    bq = bigquery.Client()
    # sql_query - extract the whole row out based on given Num identifier
    sql_query = """SELECT * FROM `testing-bigquery-vertexai.web_UI.Rolling_08-09-23` WHERE Num = '{num}'""".format(num = row_num) 
    # run query
    query_job = bq.query(sql_query)
    # get result
    results = query_job.result()

    # assign results to temporary variables
    for row in results:
        company_name = row["_Target"]
        deal_intelligence = row["Deal_Intelligence_info"]
        dominant_country = row["Target_country"]
        # sbd = row["Short BD"]
        # bd = row["Business_Description"]
        region = row["Target_Region"]
        next_step = row["KPMG_View___Redacted"]
        others = row["Asset_pack_"]
        link = str(row["Website"])

    # # Use LLM to get deal summary and business summary
    cols = ["Deal_Intelligence_info","Business_Description"]
    col_summary = {}

    for rscol in cols:

        job = """
        SELECT
        ml_generate_text_result['predictions'][0]['content'] AS generated_text,
        ml_generate_text_result['predictions'][0]['safetyAttributes']
            AS safety_attributes,
        * EXCEPT (ml_generate_text_result)
        FROM
        ML.GENERATE_TEXT(
            MODEL `bqml_tutorial.llm_model`,
            (
            SELECT
                CONCAT('Summarize the following text in 50 words: ', {tgt_column}) AS prompt,
                *
            FROM
                `testing-bigquery-vertexai.web_UI.Rolling_08-09-23`
            WHERE Num = '{num}'
            ),
            STRUCT(
            0.2 AS temperature,
            200 AS max_output_tokens));""".format(num = row_num, tgt_column = rscol)

        result = bq.query(job)
        for r in result:
            # to get rid of ""
            col_summary[rscol] = r[0][1:len(r[0])-1]

    # do a split first on others to form a list -> allowing further process later
    other_info = others.split(";")

    return company_name, deal_intelligence, dominant_country, region, next_step, other_info, col_summary, link


# # # slide 1 - done
# slide index are 1-based
def gen_s1(presentation, content):
    """
    This function replaces all placeholder text in form of {information key} with respective
    values under the identical key in content dict.
    Placeholder text would be replaced by string of value associated with the information key
    in content dict.

    Inputs: 
        presentation: {PowerPoint.Presentations.Presentation}
            - slide 1/s1 - the slide we'll be working on in this function, 1st slide in
                template slide deck
        content: {dict}
            - "company_name" - {string}, name of company in interest
            - "date" - {string},asset pack creation date, Month Year

    Returns:
        presentation: {PowerPoint.Presentations.Presentation}
            - Placeholder text in Shape.TextFrame.TextRange replaced by values associated
                with information key, in form of string
    """
    s1 = presentation.Slides(1)
    for shape in range(1,s1.Shapes.Count+1):
        textframe = s1.Shapes(shape).TextFrame.TextRange
        # locate placeholder text in format of {information key}
        if textframe.Text.find("{") > -1 and textframe.Text.find("}") > -1:
            start_ind = textframe.Text.index("{")
            end_ind = textframe.Text.index("}")
            # get information key
            information_category = textframe.Text[start_ind+1:end_ind]

            # match information key with content dict
            if information_category in content.keys():
                pre_processed = textframe.Text
                # replace placeholder text with value associated with information key
                #   in content dict
                textframe.Text = pre_processed.replace("{"+information_category+"}",str(content[information_category]))
        else:
            continue

    return presentation


# # # slide 4
def gen_s4(presentation, content, bulleted_category_s4):
    """
    This function replaces all placeholder text in form of {information key} with respective
    values under the identical key in content dict.
    Placeholder text with infomation key in bulleted_category_s4 list would be replaced by bulleted
    list, and those with information key not found in bulleted_category_s4 list would be replaced by
    string associated with the information key.

    Inputs: 
        presentation: {PowerPoint.Presentations.Presentation}
            - slide 4 / s4 - the slide we'll be working on in this function, 4th slide in
                template slide deck
            - Shape - contain text in TextFrame.TextRange
            - Paragraphs - contains lines that ends with \n in TextFrame.TextRange
        content: {dict}
            - "company_name" - {string}, name of company in interest
            - "date" - {string}, asset pack creation date, Day Month Year
            - "sit_overview" - {list}, containing sentences filtered from deal_intelligence
                text
            - "deal_stage" - {list}, containing sentences filtered from deal_intelligence text
            - "deal_rationale" - {list}, containing sentences filtered from deal_intelligence
                text
            - "biz_desc" - {list}, containing sentences of what the company does
            - "next_step" - {string},one-liner of KPMG's plan on specified deal intelligence
        bulleted_category_s4: {list}, contains information key {string} that are to be in bulleted list
            - sit_overview
            - deal_stage
            - deal_rationale
            - biz_desc
            - next_step

    Returns:
        presentation: {PowerPoint.Presentations.Presentation}
            - Placeholder text in Shape.TextFrame.TextRange replaced by values associated
                with information key, in form of string or bulleted list
    """
    s4 = presentation.Slides(4)
    for shape in range(1,s4.Shapes.Count+1):

        textframe = s4.Shapes(shape).TextFrame.TextRange
        
        for parag in range(len(textframe.Paragraphs())):

            if textframe.Paragraphs(parag).Text.find("{") > -1 and textframe.Paragraphs(parag).Text.find("}") > -1:
                # get information_category
                start_ind = textframe.Paragraphs(parag).Text.index("{")
                end_ind = textframe.Paragraphs(parag).Text.index("}")

                information_category = textframe.Paragraphs(parag).Text[start_ind+1:end_ind]

                # set up bulleted list in respective paragraph
                # insert in empty paragraphs first, then input text in
                # change bullet character to 2024 (8212 in decimal)
                if information_category in content.keys() and information_category in bulleted_category_s4:
                    # insert in empty paragraph
                    for paragraph in range(1,len(content[information_category])):
                        textframe.Paragraphs(paragraph).Text += "\r\n"
                    # insert in text
                    for paragr in range(len(content[information_category])):
                        textframe.Paragraphs(paragr+1).ParagraphFormat.Bullet.Type = 1
                        # changed using unicode code point - decimal
                        textframe.Paragraphs(paragr+1).ParagraphFormat.Bullet.Character = 8212
                        textframe.Paragraphs(paragr+1).Font.Color.RGB = 51 * 256 + 141 *256**2
                        textframe.Paragraphs(paragr+1).ParagraphFormat.SpaceAfter = 0
                        textframe.Paragraphs(paragr+1).ParagraphFormat.SpaceBefore = 0
                        textframe.Paragraphs(paragr+1).Text = content[information_category][paragr]

                elif information_category in content.keys() and information_category not in bulleted_category_s4:
                    pre_processed = textframe.Paragraphs(parag).Text

                    # input as per normal
                    textframe.Paragraphs(parag).Text = pre_processed.replace("{"+information_category+"}",str(content[information_category]))
                if information_category == "deal_overview":
                    for tfp in range(len(textframe)):
                        textframe.Font.Italic = True

    return presentation


# # # slide 6
def gen_s6(presentation, content, bulleted_category_s6):
    """
    This function replaces all placeholder text in form of {information key} with respective
    values under the identical key in content dict.
    Placeholder text with infomation key in bulleted_category_s6 list would be replaced by bulleted
    list, and those with information key not found in bulleted_category_s6 list would be replaced by
    string associated with the information key.

    Inputs: 
        presentation: {PowerPoint.Presentations.Presentation}
            - slide 6 / s6 - the slide we'll be working on in this function, 4th slide in
                template slide deck
            - Shape - contain text in TextFrame.TextRange
            - Paragraphs - contains lines that ends with \n in TextFrame.TextRange
        content: {dict}
            - "company_name" - {string}, name of company in interest
            - "found_year" - {int}, year the company in interest is founded
            - "dominant_country" - {string}, country the company's headquarter is based in
            - "employee_number" - {int}, number of employees in the company
            - "biz_desc" - {list}, containing sentences of what the company does
            - "date" - {string}, asset pack creation date, Day Month Year
        bulleted_category_s6: {list}, contains information key {string} that are to be in bulleted list
            - biz_desc

    Returns:
        presentation: {PowerPoint.Presentations.Presentation}
            - Placeholder text in Shape.TextFrame.TextRange replaced by values associated
                with information key, in form of string or bulleted list
    """
    s6 = presentation.Slides(6)
    
    for shape in range(1,s6.Shapes.Count+1):

        # check that the {} exist in the shape
        if s6.Shapes(shape).hasTextFrame:

            textframe = s6.Shapes(shape).TextFrame.TextRange

            for parag in range(len(textframe.Paragraphs())):

                if textframe.Paragraphs(parag).Text.find("{") > -1 and textframe.Paragraphs(parag).Text.find("}") > -1:

                    start_ind = textframe.Paragraphs(parag).Text.index("{")
                    end_ind = textframe.Paragraphs(parag).Text.index("}")
                    # get information key
                    information_category = textframe.Paragraphs(parag).Text[start_ind+1:end_ind]

                    if information_category in content.keys() and information_category in bulleted_category_s6:

                        # add empty paragraph
                        for paragraph in range(1,len(content[information_category])):
                            textframe.Paragraphs(paragraph).Text += "\r\n"
                        # input info
                        for paragr in range(len(content[information_category])):
                            textframe.Paragraphs(paragr+1).ParagraphFormat.Bullet.Type = 1
                            textframe.Paragraphs(paragr+1).Font.Color.RGB = 0
                            textframe.Paragraphs(paragr+1).ParagraphFormat.SpaceAfter = 0
                            textframe.Paragraphs(paragr+1).ParagraphFormat.SpaceBefore = 0
                            textframe.Paragraphs(paragr+1).Text = str(content[information_category][paragr])
                    
                    elif information_category in content.keys() and information_category not in bulleted_category_s6:

                        pre_processed = textframe.Paragraphs(parag).Text
                        # input info
                        textframe.Paragraphs(parag).Text = pre_processed.replace("{"+information_category+"}",str(content[information_category]))
                else:
                    continue

    return presentation


# # # slide 7 - half-done (TODO check order of shape, hyperlink)
def gen_s7(presentation,content,country_flag,country,color_hier):
    """
    This function replaces all placeholder text in form of {information key} with respective
    values under the identical key in content dict.
    Placeholder text with infomation key would be replaced by string associated with the 
    information key.
    Shapes in map chart would be filled with colors in color_hier, in stated order, based on
    countries in region list.

    
    Inputs: 
        presentation: {PowerPoint.Presentations.Presentation}
            - slide 7 / s7, slide 13 / s13 - the slide we'll be working on in this function, 4th slide in
                template slide deck
            - Shape - contain text in TextFrame.TextRange
            - Characters - contain subset of TextFrame.TextRange
        content: {dict}
            - "company_name" - {string}, name of company in interest
            - "num_countries" - {int}, number of countries the company has set up offices
                in
            - "dominant_country" - {string}, country the company's headquarter is based in
            - "link" - {string}, link to company's official website
            - "date" - {string}, asset pack creation date, Day Month Year
            - "region" - {list}, containing countries the company has set up offices in
        country_flag: {dict} 
            - in format of country_name{string}:image_id{int}
            - image_id is unique in slide 13/s13
        country: {dict}
            - in format of country_name{string}:shape_id{int}
            - shape_id is unique in slide 7/s7

    Returns:
        presentation: {PowerPoint.Presentations.Presentation}
            - Placeholder text in Shape.TextFrame.TextRange replaced by values associated
                with information key, in form of string or bulleted list
            - Shape(s) of country/ies filled with color in color_hier in stated order
    """
    s13 = presentation.Slides(13)
    s7 = presentation.Slides(7)

    # edit slide
    for shape in range(1,s7.Shapes.Count+1):
        if s7.Shapes(shape).hasTextFrame:
            textframe = s7.Shapes(shape).TextFrame.TextRange

            if textframe.Text.find("{") > -1 and textframe.Text.find("}") > -1:

                # find all instance of placeholder text
                start_ind = [x.start() for x in re.finditer("{",textframe.Text)]
                end_ind = [y.start() for y in re.finditer("}",textframe.Text)]

                while len(start_ind) >= 1:
                    pre_processed = textframe.Text
                    ind_start = start_ind[0]
                    ind_end = end_ind[0]

                    information_category = pre_processed[ind_start+1:ind_end]

                    # change text
                    if information_category in content.keys():
                        # if information_category is link, add hyperlink
                        if information_category != "link":
                            textframe.Text = pre_processed.replace("{"+information_category+"}",str(content[information_category]))
                        else:
                            # # textframe.Character to get subset of text range
                            ind_start_link = textframe.Text.index(r"{link}")
                            hyperlink_range = textframe.Characters(Start=ind_start_link+1, Length=6)
                            hyperlink_range.ActionSettings(1).Hyperlink.Address = r"https://" + content["link"]

                    start_ind = [x.start() for x in re.finditer("{",textframe.Text)]
                    end_ind = [y.start() for y in re.finditer("}",textframe.Text)]
                    if ind_start in start_ind:
                        start_ind.remove(ind_start)
                        end_ind.remove(ind_end)

        # elif s7.Shapes(shape).hasTextFrame == False:
    # edit map chart shapes to 
    reg = content["region"]
    # target_country = ""
    for city in reg:
        for co in country.keys():
            # if co in city:
            if city == co:
                target_country = co
                # change fill of shape
                if reg.index(city) == 0:
                    color_ind = 6
                else:
                    color_ind = reg.index(city) % 7 - 1
                # color_ind = reg.index(city) % 7 
                chosen_color = color_hier[color_ind][0] + color_hier[color_ind][1]*256 + color_hier[color_ind][2]*256**2
                s7.Shapes(country[target_country]).Fill.ForeColor.RGB = chosen_color
                # add in country flag
                # get position of the map shape for target country
                left = s7.Shapes(country[target_country]).Left
                top = s7.Shapes(country[target_country]).Top
                if target_country in country_flag.keys():
                    # copy and paste flag image from slide 13 to slide 7
                    s13.Shapes(country_flag[target_country]).Copy()
                    s7.Shapes.Paste()
                    # the newly added image will always be last shape,
                    # adjust the location of image
                    s7.Shapes(len(s7.Shapes)).Left = left
                    s7.Shapes(len(s7.Shapes)).Top = top
                    # add name below flag, in 1x1 table
                    table_top = top + s7.Shapes(len(s7.Shapes)).Height
                    table_width = 40
                    table_height = 20
                    tb = s7.Shapes.AddTable(1,1,left,table_top,table_width,table_height)
                    tb.Table.Cell(1,1).Shape.TextFrame.TextRange.Font.Size = 9
                    tb.Table.Cell(1,1).Shape.Fill.ForeColor.RGB = 0 + 51*256 + 141*256**2
                    tb.Table.Cell(1,1).Shape.TextFrame.TextRange.Text = target_country
    return presentation

# # # slide 8 - done
def gen_s8(presentation, content, bulleted_category_s8, brand_counter=0):
    """
    This function replaces all placeholder text in form of {information key} with respective
    values under the identical key in content dict. 
    Placeholder text with infomation key in bulleted_category_s8 list would be replaced by bulleted
    list, and those with information key not found in bulleted_category_s8 list would be replaced by
    string associated with the information key.
    
    Inputs: 
        presentation: {PowerPoint.Presentations.Presentation}
            - slide 8/s8, the slide we'll be working on in this function, 8th slide in
                template slide deck
            - Shape - contain text in TextFrame.TextRange
        content: {dict}
            - "company_name" - {string}
            - "biz_summary" - {string}, summary of what the company does
            - "brand_name" - {list}, 
            - "brand_desc" - {dict}
        bulleted_category_s8: {list}, contains information key {string} that are to be in bulleted list
            - brand_desc
        brand_counter: {int}
            - set to 0 by default, counter used to keep track of which brand we are at

    Returns:
        presentation: {PowerPoint.Presentations.Presentation}
            - Placeholder text in Shape.TextFrame.TextRange replaced by values associated
                with information key, in form of string or bullted list
    """
    s8 = presentation.Slides(8)

    for shape in range(1,s8.Shapes.Count+1):
        if s8.Shapes(shape).hasTextFrame:
            textframe = s8.Shapes(shape).TextFrame.TextRange
            
            # locate placeholder text
            if textframe.Text.find("{") > -1 and textframe.Text.find("}") > -1:
                start_ind = textframe.Text.index("{")
                end_ind = textframe.Text.index("}")

                # get information key
                information_category = textframe.Text[start_ind+1:end_ind]

                if information_category in content.keys():
                    pre_processed = textframe.Text
                    if information_category == "brand_name":
                        textframe.Text = content[information_category][brand_counter]
                        brand_counter += 1
                    elif information_category in bulleted_category_s8:
                        # bulleted brand description
                        # insert empty paragraph
                        for paragraph in range(1,len(content[information_category][content["brand_name"][brand_counter]])):
                            textframe.Paragraphs(paragraph).Text += "\r\n"
                        # inputing data from content dict
                        for parag in range(len(content[information_category][content["brand_name"][brand_counter]])):
                            pre_processed = textframe.Paragraphs(parag+1).Text
                            textframe.Paragraphs(parag+1).ParagraphFormat.Bullet.Type = 1
                            # textframe.Paragraphs(parag+1).ParagraphFormat.Bullet.Character = 8226
                            textframe.Paragraphs(parag+1).Font.Color.RGB = 0
                            textframe.Paragraphs(parag+1).ParagraphFormat.SpaceAfter = 0
                            textframe.Paragraphs(parag+1).ParagraphFormat.SpaceBefore = 0
                            textframe.Paragraphs(parag+1).Text = content[information_category][content["brand_name"][brand_counter]][parag]
                    else:
                        # inputing data from content dict
                        textframe.Text = pre_processed.replace("{"+information_category+"}",str(content[information_category]))
                else:
                    continue
    
    return presentation, brand_counter


# # # slide 9
def gen_s9(presentation, content, bulleted_category_s9, brand_counter):
    """
    This function replaces all placeholder text in form of {information key} with respective
    values under the identical key in content dict. 
    Placeholder text with infomation key in bulleted_category_s9 list would be replaced by bulleted
    list, and those with information key not found in bulleted_category_s9 list would be replaced by
    string associated with the information key.
    
    Inputs: 
        presentation: {PowerPoint.Presentations.Presentation}
            - slide 9 / s9 - the slide we'll be working on in this function, 9th slide in
                template slide deck
            - Shape - contain text in TextFrame.TextRange
            - Paragraphs - contains lines that ends with \n in TextFrame.TextRange
        content: {dict}
            - "company_name" - {string}, name of company in interest
            - "date" - {string}, asset pack creation date, Day Month Year
            - "brand_desc" - {dict}, contain description of each brand in format of 
                brand_name:brand_description (list containing sentences of brand description)
            - "biz_summary" - {string}, summary of what the company does
            - "social_work_desc" - {list}, contains sentences that describes company's social
                work effort(s)
            - "sch_network_desc" - {list}, contains sentences that describes company's school
                network initiative(s)
            - "corp_solution_desc" - {list}, contains sentences that describes company's
                corporate solution(s)
        bulleted_category_s9: {list}, contains information key {string} that are to be in bulleted list
            - brand_desc
            - social_work_desc
            - sch_network_desc
            - corp_solution_desc
        brand_counter: {int}
            - set to 0 by default and adapted from gen_s8, counter used to keep track of 
                which brand we are at

    Returns:
        presentation: {PowerPoint.Presentations.Presentation}
            - Placeholder text in Shape.TextFrame.TextRange replaced by values associated 
                with the information key, in form of string or bulleted lists
    """
    s9 = presentation.Slides(9)

    for shape in range(1,s9.Shapes.Count+1):
        if s9.Shapes(shape).hasTextFrame:

            textframe = s9.Shapes(shape).TextFrame.TextRange

            # locate placeholder text in format of {information key}
            if textframe.Text.find("{") > -1 and textframe.Text.find("}") > -1:
                start_ind = textframe.Text.index("{")
                end_ind = textframe.Text.index("}")
                # get information key
                information_category = textframe.Text[start_ind+1:end_ind]

                if information_category in content.keys() and information_category in bulleted_category_s9:
                    # input empty paragraphs
                    # input data
                    if information_category == "brand_desc":
                        # insert in empty paragraphs for inputting lines later
                        for paragraph in range(len(content["brand_desc"][content["brand_name"][brand_counter]])-1):
                            textframe.Paragraphs(paragraph).Text += "\r\n" 

                        for parag in range(len(content["brand_desc"][content["brand_name"][brand_counter]])):
                            # state that it is unordered list, using the bullet character with unicode-hex 2014
                            #   (8226 in decimal)
                            textframe.Paragraphs(parag+1).ParagraphFormat.Bullet.Type = 1
                            textframe.Paragraphs(parag+1).ParagraphFormat.Bullet.Character = 8226
                            textframe.Paragraphs(parag+1).Text = content["brand_desc"][content["brand_name"][brand_counter]][parag]
                
                    else:
                        # insert empty paragraphs
                        for paragraph in range(len(content[information_category])-1):
                            textframe.Paragraphs(paragraph).Text += "\r\n" 

                        for parag in range(len(content[information_category])):
                            # format each line, stating that it is unordered list
                            textframe.Paragraphs(parag+1).ParagraphFormat.Bullet.Type = 1
                            # set font color to black
                            textframe.Paragraphs(parag+1).Font.Color.RGB = 0
                            # remove space before and after line
                            textframe.Paragraphs(parag+1).ParagraphFormat.SpaceAfter = 0
                            textframe.Paragraphs(parag+1).ParagraphFormat.SpaceBefore = 0
                            # input into paragraphs line by line
                            textframe.Paragraphs(parag+1).Text = content[information_category][parag]

                # for those placeholder text that are not to be replaced by bulleted list,
                #   replace with string
                elif information_category in content.keys() and information_category not in bulleted_category_s9:
                    pre_processed = textframe.Text
                    if information_category == "brand_name":
                        textframe.Text = content["brand_name"][brand_counter]
                    else:
                        textframe.Text = pre_processed.replace("{"+information_category+"}",str(content[information_category]))

    return presentation


# # # slide 12
def gen_s12(presentation,content,bulleted_category_s12):
    """
    This function replaces all placeholder text in form of {information key} with respective
    values under the identical key in content dict.
    Placeholder text with infomation key in bulleted_category_s12 list would be replaced by bulleted
    list.

    Inputs: 
        presentation: {PowerPoint.Presentations.Presentation}
            - slide 12/s12 - the slide we'll be working on for this function, 12th slide in
                the template slide deck
            - Shapes - contains text in TextFrame.TextRange
            - Paragraphs - contains lines that ends with \n in TextFrame.TextRange
        content: {dict}, information_category:data
            - "deal_Intelligence" - {list}, containing sentences from splitting of 
                row['Deal_Intelligence_info']
        bulleted_category_s12: {list}, contains information key {string} that are to be in bulleted list
            - deal_intelligence

    Returns:
        presentation: {PowerPoint.Presentations.Presentation}
            - Placeholder text in Shapes.TextFrame.TextRange.Text replaced by bulleted 
                list
    """
    s12 = presentation.Slides(12)
    bulleted_category_s12 = ["deal_intelligence"]
    for shape in range(1,s12.Shapes.Count+1):

        if s12.Shapes(shape).hasTextFrame:
            # access the TextRange to access properties such as Text and Paragraphs
            textframe = s12.Shapes(shape).TextFrame.TextRange

            # locate the placeholder text indicated by { and }
            if textframe.Text.find("{") > -1 and textframe.Text.find("}") > -1:
                start_ind = textframe.Text.index("{")
                end_ind = textframe.Text.index("}")
                # get information key
                information_category = textframe.Text[start_ind+1:end_ind]

                # condition for bulleted list
                if information_category in content.keys() and information_category in bulleted_category_s12:

                    # insert empty paragraph
                    for paragraph in range(1,len(content[information_category])):
                        textframe.Paragraphs(paragraph).Text += "\r\n"
                    # input into the empty paragraph line by line
                    for parag in range(len(content[information_category])):
                        textframe.Paragraphs(parag+1).ParagraphFormat.Bullet.Type = 1
                        textframe.Paragraphs(parag+1).ParagraphFormat.Bullet.Character = 8226
                        textframe.Paragraphs(parag+1).Text = content[information_category][parag]
    return presentation

def main(row_num):
    """
    This function will execute supplementary functions stated above to generate an asset pack for
    the chosen company.

    Inputs: 
        row_num: {int}
            - to be passed in from webpage

    Returns:
        presentation: {PowerPoint.Presentations.Presentation}
            - Placeholder text in Shapes.TextFrame.TextRange.Text replaced by string or bulleted 
                list
            - Shapes in map chart in slide 7 filled with color according to the region list and
                color_hier
    """
    # branding color scheme
    color_hier = [(0,51,141), (30,73,226), (172,234,255), (0,184,245), (12,35,60), (114,19,234), (253,52,156)]
    # country map shape dict, matching country names to the shape representing them
    # shapes 2, 18-24 can't be found
    country = {
        "South America":17, # correct 
        "America":13, # correct
        "Canada":14, # correct
        "Greenland":50, # correct
        "Haiti":18, # correct
        "Pakistan":34, # correct
        "Bhutan":35, # correct
        "Nepal":36, # correct
        "Bangladesh": 37, # correct
        "Sri Lanka":38, # correct
        "Japan":39, # correct
        "Russia":41, # correct
        "Middle East":42, # correct
        "Azerbaijan":48, # correct
        "India":92,
        "Italy":54, # correct
        "France":62, # correct
        "United Kingdom":63, # correct
        "Bosnia and Herzegovina":68, # correct
        "Norway":86,
        "Africa":90,
        "Philippines":94,
        "China":98,
        "Taiwan":99,
        "Nicaragua":5, # correct
        "Guatemala":7, # correct
        "Belize":8, # correct 
        "Costa Rica":10, # correct 
        "Honduras":11, # correct 
        "Mexico":12, # correct  
        "Jamaica":15, # correct
        "Panama":6, # correct 
        "El Salvador":9, # correct 
        "Dominican Republic":16, # correct 
        "Cuba":19, # correct
        "Bahamas":20, # correct
        "Lesser Antilles":22, # correct
        "Malaysia":97,
        "Indonesia":96,
        "Papua New Guinea":28, # correct
        "New Zealand":24, # correct
        "Australia":23, # correct 
        "Brunei":95,
        "Montenegro":89,
        "Serbia":88,
        "Sweden":87,
        "Denmark":85,
        "Finland":84,
        "Estonia":83,
        "Romania":82,
        "North Macedonia":81,
        "Slovenia":80,
        "Slovakia":79,
        "Lithuania":78,
        "Albania":77,
        "Croatia":76,
        "Bulgaria":75,
        "Czechia":74,
        "Latvia":73,
        "Poland":72,
        "Hungary":70, # correct 
        "Greece":69, # correct
        "Belarus":67, # correct 
        "Moldova":66, # correct
        "Ukraine":65, # correct
        "Ireland":64, # correct
        "Germany":61, # correct
        "Portugal":60, # correct
        "Netherlands":59, # correct
        "Spain":58, # correct
        "Belgium":57, # correct
        "Switzerland":56, # correct
        "Luxembourg":55, # correct
        "Liechtenstein":53, # correct
        "Andorra":52, # correct
        "Austria":51, # correct
        "Iceland":49, # correct
        "Turkmenistan":47, # correct
        "Uzbekistan":46, # correct
        "Georgia":45, # correct
        "Kazakhstan":44, # correct
        "Armenia":43, # correct
        "Mongolia":40, # correct
        "Laos":33, # correct
        "Cambodia":32, # correct
        "Myanmar":31, # correct
        "Thailand":30, # correct
        "Vietnam":29, # correct
        "Afghanistan":27, # correct
        "Kyrgyzstan":25, # correct
        "Tajikistan":26, # correct
        "Puerto Rico":21, # correct
    }

    # country flage image dict, matching image to the respective country
    country_flag = {
        "America":52,
        "Canada":14,
        "Greenland":24,
        "Haiti":26,
        "Pakistan":54,
        "Japan":53,
        "Russia":45,
        "Middle East":49,
        "Azerbaijan":7,
        "India":29,
        "Italy":31,
        "France":22,
        "United Kingdom":51,
        "Bosnia and Herzegovina":11,
        "Norway":40,
        "Africa":4,
        "Philippines":42,
        "China":15,
        "Taiwan":50,
        "Nicaragua":38,
        "Guatemala":25,
        "Belize":10,
        "Costa Rica":16,
        "Honduras":27,
        "Mexico":35,
        "Bahamas":8,
        "Lesser Antilles":29,
        "Malaysia":34,
        "Indonesia":30,
        "Papua New Guinea":41,
        "New Zealand":37,
        "Australia":6, 
        "Brunei":12,
        "Montenegro":36,
        "Serbia":46,
        "Sweden":49,
        "Denmark":19,
        "Finland":21,
        "Estonia":20,
        "Romania":44,
        "North Macedonia":39,
        "Slovenia":48,
        "Slovakia":47,
        "Lithuania":33,
        "Albania":5,
        "Croatia":17,
        "Bulgaria":13,
        "Czechia":18,
        "Latvia":32, 
        "Poland":43,
        "Hungary":28, 
        "Greece":23, 
        "Afghanistan":3,
        "Belarus":9, 
        "Moldova":77,
        "Ukraine":76,
        "Ireland":75,
        "Germany":74,
        "Portugal":73,
        "Netherlands":72,
        "Spain":71,
        "Belgium":70,
        "Switzerland":69,
        "Luxembourg":68,
        "Liechtenstein":67,
        "Andorra":66,
        "Austria":65,
        "Iceland":64,
        "Turkmenistan":63,
        "Uzbekistan":62,
        "Georgia":61,
        "Kazakhstan":60,
        "Armenia":59,
        "Mongolia":58,
        "Laos":57,
        "Cambodia":56,
        "Myanmar":55,
        "Thailand":80,
        "Vietnam":79,
        "Kyrgyzstan":78,
        "Puerto Rico":81,
        "Jamaica":86,
        "Panama":85,
        "El Salvador":84,
        "Dominican Republic":83,
        "Cuba":82,
        "Tajikistan":90,
        "Bhutan":89,
        "Nepal":88,
        "Bangladesh":87,
        "Sri Lanka":91,
    }

    # retrieve data from big query
    company_name, deal_intelligence, dominant_country, region, next_step, other_info, col_summary, link = retrieve_bigquery(row_num)

    # # content dict
    # save all information to be inputted here for easy access
    content = {
        "company_name":company_name,
        "month_year": datetime.today().strftime("%B %Y"),
        "date": datetime.today().strftime(r"%d %B %Y"),
        "biz_desc": [],
        "biz_summary": col_summary["Business_Description"],
        "deal_overview":col_summary["Deal_Intelligence_info"],
        "sit_overview":[],
        "deal_stage":[],
        "deal_rationale":[],
        "next_step":next_step.split("\n"),
        "deal_intelligence":deal_intelligence.split("\n"),
        "found_year":1900,
        "dominant_country": dominant_country,
        "region":region.split(";"),
        "num_countries":len(region.split(";")),
        "employee_number": 0,
        "link":link,
        "brand_name":["Sample brand name 1", "Sample brand name 2", "Sample brand name 3"],
        "brand_desc":{
            "Sample brand name 1": ["Sample brand description line","Sample brand description line"],
            "Sample brand name 2": ["Sample brand description line","Sample brand description line"],
            "Sample brand name 3": ["Sample brand description line","Sample brand description line"]
        },
        "social_work_desc":["Sample social work line", "Sample social work line"],
        "sch_network_desc":["Sample school network line","Sample school network line"],
        "corp_solution_desc":["Sample corporate solution line","Sample corporate solution line"]
    }

    # # # text processing
    to_be_processed = deal_intelligence.split("\n")

    # remove empty string, useless lines
    # get index of elements to be removed
    remove = []
    kw_remove = ["original source", "updated on", "checked on", "link to"]
    kw_startwith = ["by", "press release","about"]
    for tp in range(len(to_be_processed)):
        # remove empty lines
        if to_be_processed[tp] == "":
            remove.append(tp)
        # remove useless lines
        else:
            # obtain and store index of elements to be removed
            for kwr in kw_remove:
                if kwr in to_be_processed[tp].lower() and tp not in remove:
                    remove.append(tp)
            for kws in kw_startwith:
                if to_be_processed[tp].lower().startswith(kws) and tp not in remove:
                    remove.append(tp)

    # remove the element accordingly, note change in index
    for r in range(len(remove)):
        ind_remove = remove[r]
        to_be_processed.remove(to_be_processed[ind_remove-r])

    # keywords for sorting sentences
    kw = {
        "sit_overview":["launch","strategic review", "strategic","review", "shortlisted", "looking to", "sell", "business expansion", "expand", "mandate", "acquisition","acquired"],
        "deal_stage":["initial public offering", "expected to", "in a process", "plans to", "plan to", "early stage", "decision"],
        "deal_rationale":["funding used for", "funding would be used for", "investment", "funding", "focus on", "approach"],
        "biz_desc":["established as", "providing", "in house", "production", "largest", "market cap", "leading"],
    }

    # processing and sort Deal_intelligence sentences
    for l in range(len(to_be_processed)):
        for k in kw.values():
            # get key aka the information category with values in k
            information_category = list(kw.keys())[list(kw.values()).index(k)]
            # # for each word in k (containing all relevant words/phrases)
            for string in k:
                if string in to_be_processed[l].lower():
                    content[information_category].append(to_be_processed[l])

    # further process other info and store in content dict
    for information in other_info:
        info = information.split(":")
        if info[0] == "numOfEmployees":
            content["employee_number"] = info[1]
        elif info[0] == "yearFounded":
            content["found_year"] = info[1]

    
    # open an instance of PowerPoint
    ppt_app = comtypes.client.CreateObject("PowerPoint.Application")
    # open template asset pack slides
    template_path = "../../../results/input/templates/AssetPackSample_v0.51.pptx"
    template_abs_path = os.path.abspath(template_path)
    presentation_duplicate = ppt_app.Presentations.Open(template_abs_path)

    # # save slide, open duplicated slide
    save_path = """../../../results/output/asset_packs/{company_name}-AssetPack.pptx""".format(company_name=content["company_name"])
    save_abs_path = os.path.abspath(save_path)
    presentation_duplicate.SaveAs(save_abs_path)
    presentation = presentation_duplicate

    # editing content in slides
    presentation = gen_s1(presentation,content)
    bulleted_category_s4 = ["sit_overview","deal_stage","deal_rationale","biz_desc","next_step"]
    presentation = gen_s4(presentation,content,bulleted_category_s4)
    bulleted_category_s6 = ["biz_desc"]
    presentation = gen_s6(presentation,content,bulleted_category_s6)
    presentation = gen_s7(presentation,content,country_flag,country,color_hier)
    bulleted_category_s8 = ["brand_desc"]
    presentation, brand_counter = gen_s8(presentation, content, bulleted_category_s8)
    bulleted_category_s9 = ["brand_desc","social_work_desc","sch_network_desc","corp_solution_desc"]
    presentation = gen_s9(presentation, content, bulleted_category_s9, brand_counter)
    bulleted_category_s12 = ["deal_intelligence"]
    presentation = gen_s12(presentation,content,bulleted_category_s12)
    
    return presentation

main(1047)