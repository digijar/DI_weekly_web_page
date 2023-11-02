# # install comtypes, google-cloud-bigquery
# pip install comtypes --trusted-host=files.pythonhosted.org

import comtypes.client
import comtypes.gen.PowerPoint as ppt
# to get current date
from datetime import datetime
# text processing
import re
# service related libraries
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
import os
import uvicorn
import requests

app = FastAPI()

retrieve_data_URL = "http://127.0.0.1:5011"

# Enable CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

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
            - Paragraphs - contains lines that ends with \r\n in TextFrame.TextRange
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
            - Paragraphs - contains lines that ends with \r\n in TextFrame.TextRange
        content: {dict}
            - "company_name" - {string}, name of company in interest
            - "found_year" - {int}, year the company in interest is founded
            - "dominant_country" - {string}, country the company's headquarter is based in
            - "employee_num" - {int}, number of employees in the company
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
            ruler = s6.Shapes(shape).TextFrame.Ruler

            for parag in range(len(textframe.Paragraphs())):

                if textframe.Paragraphs(parag).Text.find("{") > -1 and textframe.Paragraphs(parag).Text.find("}") > -1:

                    start_ind = textframe.Paragraphs(parag).Text.index("{")
                    end_ind = textframe.Paragraphs(parag).Text.index("}")
                    # get information key
                    information_category = textframe.Paragraphs(parag).Text[start_ind+1:end_ind]

                    if information_category in content.keys() and information_category in bulleted_category_s6:

                        # set indentation
                        ruler.Levels(1).FirstMargin = 0
                        ruler.Levels(1).LeftMargin = 10
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


# # # slide 7
def gen_s7(presentation,content,country_flag,countries,color_hier,continent):
    """
    This function replaces all placeholder text in form of {information key} with respective
    values under the identical key in content dict.
    Placeholder text with infomation key would be replaced by string associated with the 
    information key.
    Shapes in map chart would be filled with colors in color_hier, in stated order, based on
    countries in region list.

    
    Inputs: 
        presentation: {PowerPoint.Presentations.Presentation}
            - slide 7 / s7, slide 13 / s12 - the slide we'll be working on in this function, 4th slide in
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
            - image_id is unique in slide 13/s12
        countries: {dict}
            - in format of country_name{string}:shape_id{int}
            - shape_id is unique in slide 7/s7
        color_hier: {list} - contain KPMG brand colors in order
        continent: {dict}
            - continent name {str}: countries names within geographical boundary of stated continent {list}

    Returns:
        presentation: {PowerPoint.Presentations.Presentation}
            - Placeholder text in Shape.TextFrame.TextRange replaced by values associated
                with information key, in form of string or bulleted list
            - Shape(s) of country/ies filled with color in color_hier in stated order
    """
    s12 = presentation.Slides(12)
    s7 = presentation.Slides(7)

    # edit slide text
    for shape in range(1,s7.Shapes.Count+1):
        if s7.Shapes(shape).hasTextFrame:
            textframe = s7.Shapes(shape).TextFrame.TextRange
            if textframe.Text.find("{") > -1 and textframe.Text.find("}") > -1:
                pre_process = textframe.Text
                matches = [x.group() for x in re.finditer(r"\{([^{}]+)\}",pre_process)]
                print(matches)
                for information_category in matches:
                    pre_processed = textframe.Text
                    if information_category == r"{link}":
                        if content["link"] != None:
                            textframe.Text = textframe.Text.replace(information_category, content["link"])
                            # ind_start_link = textframe.Text.index(information_category)
                            # hyperlink_range = textframe.Characters(Start=ind_start_link+1, Length=6)
                            # hyperlink_range.ActionSettings.Item(1).Hyperlink.Address = "http://" + content["link"]
                        else:
                            textframe.Text = textframe.Text.replace(information_category, "NA")

                    elif information_category == r"{region}":
                        textframe.Text = textframe.Text.replace(information_category,", ".join(content[information_category[1:len(information_category)-1]]))
                    else:
                        textframe.Text = textframe.Text.replace(information_category,str(content[information_category[1:len(information_category)-1]]))
  
    print("text replaced")

    # edit map chart shapes
    # create list of countries without duplicate
    reg = {
    }
    country_reg = []
    for cty in content["region"]:
        if cty in countries.keys():
            country_reg.append(cty)
        elif cty in continent.keys():
            reg[cty] = continent[cty]
    reg["countries"] = list(set(tuple(country_reg)))  
    print("sorting done - reg")

    for r in reg.keys():
        if r == "countries":
            # loop through reg list to fill shapes
            for country in reg["countries"]:
                if reg["countries"].index(country) == 0:
                    color_ind = 6
                else:
                    color_ind = reg["countries"].index(country) % 7 -1
                chosen_color = color_hier[color_ind][0]+color_hier[color_ind][1]*256 + color_hier[color_ind][2]*256**2
                s7.Shapes(countries[country]).Fill.ForeColor.RGB = chosen_color
                # add in country flag
                # get position of the map shape for target country
                left = s7.Shapes(countries[country]).Left + s7.Shapes(countries[country]).Width/2
                top = s7.Shapes(countries[country]).Top + s7.Shapes(countries[country]).Height/2
                if country in country_flag.keys():
                    # copy and paste flag image from slide 13 to slide 7
                    s12.Shapes(country_flag[country]).Copy()
                    s7.Shapes.Paste()
                    # the newly added image will always be last shape,
                    # adjust the location of image
                    s7.Shapes(len(s7.Shapes)).Left = left
                    s7.Shapes(len(s7.Shapes)).Top = top
                    # add name below flag, in 1x1 table
                    table_top = top + s7.Shapes(s7.Shapes.Count).Height
                    table_width = 40
                    table_height = 20
                    tb = s7.Shapes.AddTable(1,1,left,table_top,table_width,table_height)
                    tb.Table.Cell(1,1).Shape.TextFrame.TextRange.Font.Size = 9
                    tb.Table.Cell(1,1).Shape.Fill.ForeColor.RGB = 0 + 51*256 + 141*256**2
                    tb.Table.Cell(1,1).Shape.TextFrame.TextRange.Text = country
        else:
            color_ind = (len(reg["countries"]) +list(reg.keys()).index(r)) % 7
            chosen_color = color_hier[color_ind][0]+color_hier[color_ind][1]*256 + color_hier[color_ind][2]*256**2
            for ctys in reg[r]:
                s7.Shapes(countries[ctys]).Fill.ForeColor.RGB = chosen_color
            left = s7.Shapes(countries[reg[r][0]]).Left
            top = s7.Shapes(countries[reg[r][0]]).Top
            if r in country_flag.keys():
                s12.Shapes(country_flag[r]).Copy()
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
                tb.Table.Cell(1,1).Shape.TextFrame.TextRange.Text = r
    print("shape filling done")
         
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
            ruler = s8.Shapes(shape).TextFrame.Ruler
            
            # locate placeholder text
            if textframe.Text.find("{") > -1 and textframe.Text.find("}") > -1:
                start_ind = textframe.Text.index("{")
                end_ind = textframe.Text.index("}")

                # get information key
                information_category = textframe.Text[start_ind+1:end_ind]

                if information_category in content.keys():
                    # set indentation
                    ruler.Levels(1).FirstMargin = 0
                    ruler.Levels(1).LeftMargin = 10
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
            - Paragraphs - contains lines that ends with \r\n in TextFrame.TextRange
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
            ruler = s9.Shapes(shape).TextFrame.Ruler

            # locate placeholder text in format of {information key}
            if textframe.Text.find("{") > -1 and textframe.Text.find("}") > -1:
                start_ind = textframe.Text.index("{")
                end_ind = textframe.Text.index("}")
                # get information key
                information_category = textframe.Text[start_ind+1:end_ind]

                if information_category in content.keys() and information_category in bulleted_category_s9:
                    # set indentation
                    ruler.Levels(1).FirstMargin = 0
                    ruler.Levels(1).LeftMargin = 10
                    # input empty paragraphs
                    # input data
                    if information_category == "brand_desc":
                        # insert in empty paragraphs for inputting lines later
                        for paragraph in range(len(content["brand_desc"][content["brand_name"][brand_counter]])-1):
                            textframe.Paragraphs(paragraph).Text += "\r\n" 

                        for parag in range(len(content["brand_desc"][content["brand_name"][brand_counter]])):
                            # state that it is unordered list, using the bullet character with unicode-hex 2014
                            # (8226 in decimal)
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
def gen_s11(presentation,content,bulleted_category_s11):
    """
    This function replaces all placeholder text in form of {information key} with respective
    values under the identical key in content dict.
    Placeholder text with infomation key in bulleted_category_s12 list would be replaced by bulleted
    list.

    Inputs: 
        presentation: {PowerPoint.Presentations.Presentation}
            - slide 11/s11 - the slide we'll be working on for this function, 12th slide in
                the template slide deck
            - Shapes - contains text in TextFrame.TextRange
            - Paragraphs - contains lines that ends with \r\n in TextFrame.TextRange
        content: {dict}, information_category:data
            - "deal_Intelligence" - {list}, containing sentences from splitting of 
                row['Deal_Intelligence_info']
        bulleted_category_s11: {list}, contains information key {string} that are to be in bulleted list
            - deal_intelligence

    Returns:
        presentation: {PowerPoint.Presentations.Presentation}
            - Placeholder text in Shapes.TextFrame.TextRange.Text replaced by bulleted 
                list
    """
    s11 = presentation.Slides(11)
    for shape in range(1,s11.Shapes.Count+1):
        if s11.Shapes(shape).hasTextFrame:
            textframe = s11.Shapes(shape).TextFrame.TextRange

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
                    if information_category in content.keys() and information_category not in bulleted_category_s11:
                        textframe.Text = pre_processed.replace("{"+information_category+"}",str(content[information_category]))
                    elif information_category in content.keys() and information_category in bulleted_category_s11:
                        # insert empty paragraph
                        for paragraph in range(1,len(content[information_category])):
                            textframe.Paragraphs(paragraph).Text += "\r\n"
                        # input into the empty paragraph line by line
                        for parag in range(len(content[information_category])):
                            textframe.Paragraphs(parag+1).ParagraphFormat.Bullet.Type = 1
                            textframe.Paragraphs(parag+1).ParagraphFormat.Bullet.Character = 8226
                            textframe.Paragraphs(parag+1).Text = content[information_category][parag]

                        
                    start_ind = [x.start() for x in re.finditer("{",textframe.Text)]
                    end_ind = [y.start() for y in re.finditer("}",textframe.Text)]
                    if ind_start in start_ind:
                        start_ind.remove(ind_start)
                        end_ind.remove(ind_end)
    return presentation

@app.get("/apgen/{row_num}")
async def main(row_num:int):
# def main(row_num):
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
    # shapes 2 can't be found
    # shapes 53,54,56 (islands of russia) and (islands of norway) not filled
    countries = {
        "US":13, 
        "Canada":14, 
        "Greenland":79,
        "Haiti":30, 
        "Pakistan":46, 
        "Bhutan":47, 
        "Nepal":48, 
        "Bangladesh": 49, 
        "Sri Lanka":50, 
        "Japan":51, 
        "Russia":55,
        "Azerbaijan":77, 
        "India":169,  
        "Italy":82,
        "France":92,
        "United Kingdom":93,
        "Bosnia and Herzegovina":98,
        "Norway":116,
        "Philippines":171, 
        "China":175, 
        "Taiwan":176, 
        "Nicaragua":5, 
        "Guatemala":7, 
        "Belize":8, 
        "Costa Rica":10, 
        "Honduras":11, 
        "Mexico":12,  
        "Jamaica":15, 
        "Panama":6,  
        "El Salvador":9,  
        "Dominican Republic":16, 
        "Cuba":31, 
        "Bahamas":32, 
        "Lesser Antilles":34,  
        "Malaysia":174, 
        "Indonesia":173,  
        "Papua New Guinea":40, 
        "New Zealand":36, 
        "Australia":35, 
        "Brunei":169,  
        "Montenegro":119,
        "Republic of Serbia":118, 
        "Sweden":117,
        "Denmark":114, #116
        "Finland":113, 
        "Estonia":112, 
        "Romania":111, 
        "North Macedonia":110, 
        "Slovenia":109, 
        "Slovakia":108, 
        "Lithuania":107, 
        "Albania":106, 
        "Croatia":105, 
        "Bulgaria":104, 
        "Czech Republic":103,
        "Latvia":102, 
        "Poland":101,
        "Hungary":100,  
        "Greece":99, 
        "Belarus":97,    
        "Moldova":96,   
        "Ukraine":95, 
        "Ireland":94,  
        "Germany":91, 
        "Portugal":90,  
        "Netherlands":89, 
        "Spain":87,
        "Belgium":86, 
        "Switzerland":85,
        "Luxembourg":84,
        "Liechtenstein":83, 
        "Andorra":81,   
        "Austria":80,  
        "Iceland":78, 
        "Turkmenistan":76,
        "Uzbekistan":75, 
        "Georgia":74, 
        "Kazakhstan":73,  
        "Armenia":72, 
        "Mongolia":52, 
        "Laos":45, 
        "Cambodia":44, 
        "Myanmar":43, 
        "Thailand":42, 
        "Vietnam":41,  
        "Afghanistan":39, 
        "Kyrgyzstan":37, 
        "Tajikistan":38, 
        "Puerto Rico":33, 
        "Cyprus":71, 
        "Syria":70, 
        "Jordan":69, 
        "Turkey":68,
        "United Arab Emirates":67, 
        "Qatar":66, 
        "Iraq":65,
        "Iran":64,
        "Bahrain":63,
        "Oman":62, 
        "Yemen":61,  
        "Saudi Arabia":60,  
        "Kuwait":59, 
        "Israel":58, 
        "Lebanon":57, 
        "Singapore":179,
        "Ecuador":17,
        "Paraguay":18,
        "Chile":19,
        "Brazil":20,
        "Colombia":21,
        "Bolivia":22,
        "Uruguay":23,
        "Guyana":24,
        "French Guiana":25,
        "Peru":26,
        "Argentina":27,
        "Suriname":28,
        "Trinidad and Tobago":29,
        "Venezuela":88, 
        "Libya":120,
        "Western Sahara":121,
        "Ghana":122,
        "Burkina Faso":123,
        "Cameroon":124,
        "Benin":125,
        "Chad":126,
        "Eritrea":127,
        "Mali":128,
        "Tunisia":129,
        "Egypt":130,
        "Algeria":131,
        "Mauritania":132,
        "Morocco":133,
        "Somalia":134,
        "Niger":135,
        "Ethiopia":136,
        "Eswatini":137,
        "South Africa":138,
        "Malawi":139,
        "Madagascar":140,
        "Angola":141,
        "Lesotho":142,
        "Zambia":143,
        "Zimbabwe":144,
        "Namibia":145,
        "Mozambique":146,
        "Botswana":147,
        "Guinea":148,
        "Central African Republic":149,
        "Gabon":150,
        "Gambia":151,
        "Guinea Bissau":152,
        "Liberia":153,
        "Nigeria":154,
        "Republic of Congo":155,
        "Equatorial Guinea":156,
        "Burundi":157,
        "Uganda":158,
        "Rwanda":159,
        "Kenya":160,
        "Senegal":161,
        "Tanzania":162,
        "Sierra Leone":163,
        "Cote d Ivoire":164,
        "Togo":165,
        "Democratic Republic of Congo":166,
        "Djibouti":167,
        "Sudan":168,
    }

    # country flage image dict, matching image to the respective country
    country_flag = {
        "US":52, 
        "Canada":14, 
        "Greenland":24, 
        "Haiti":26, 
        "Pakistan":54, 
        "Japan":53, 
        "Russia":45, 
        "Azerbaijan":7, 
        "India":29, 
        "Italy":31, 
        "France":22, 
        "United Kingdom":51, 
        "Bosnia and Herzegovina":11, 
        "Norway":40, 
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
        "Malaysia":34, 
        "Indonesia":30, 
        "Papua New Guinea":41, 
        "New Zealand":37,
        "Australia":6, 
        "Brunei":12, 
        "Montenegro":36, 
        "Republic of Serbia":46, 
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
        "Czech Republic":18, 
        "Latvia":32,  
        "Poland":43, 
        "Hungary":28, 
        "Greece":23, 
        "Afghanistan":4, 
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
        "Sri Lanka":3, 
        "Cyprus":91,
        "Syria":104,
        "Jordan":103,
        "Turkey":102, 
        "United Arab Emirates":101,
        "Qatar":100,
        "Iraq":99,
        "Iran":105,
        "Bahrain":98,
        "Oman":97,
        "Yemen":96, 
        "Saudi Arabia":95, 
        "Kuwait":94, 
        "Israel":93, 
        "Lebanon":92, 
        "Singapore":106, 
        "Europe":107, 
        "SOUTHEAST Asia":108,
        # correct 109-170
        "Ecuador":170, # c
        "Paraguay":169,
        "Chile":168,
        "Brazil":167,
        "Colombia":166,
        "Bolivia":165,
        "Uruguay":164,
        "Guyana":163,
        "French Guiana":162,
        "Peru":161,
        "Argentina":160,
        "Suriname":159,
        "Venezuela":158, 
        "Libya":157,
        "Western Sahara":156,
        "Ghana":155,
        "Burkina Faso":154,
        "Cameroon":153,
        "Benin":152,
        "Chad":151,
        "Eritrea":150,
        "Mali":149,
        "Tunisia":148,
        "Egypt":147,
        "Algeria":146,
        "Mauritania":145,
        "Morocco":144,
        "Somalia":143,
        "Niger":142,
        "Ethiopia":141,
        "Eswatini":140,
        "South Africa":139,
        "Malawi":138,
        "Madagascar":137,
        "Angola":136,
        "Lesotho":135,
        "Zambia":134,
        "Zimbabwe":133,
        "Namibia":132,
        "Mozambique":131,
        "Botswana":130,
        "Guinea":129,
        "Central African Republic":128,
        "Gabon":127,
        "Gambia":126,
        "Guinea Bissau":125,
        "Liberia":124,
        "Nigeria":123,
        "Republic of Congo":122,
        "Equatorial Guinea":121,
        "Burundi":120,
        "Uganda":119,
        "Rwanda":118,
        "Kenya":117,
        "Senegal":116,
        "Tanzania":115,
        "Sierra Leone":114,
        "Cote d Ivoire":113,
        "Togo":112,
        "Democratic Republic of Congo":111,
        "Djibouti":110,
        "Sudan":109,
    }

    try:
        # retrieve data from big query
        response = requests.get(f"{retrieve_data_URL}/bq/{row_num}")
        info = response.json()
        
        if "\r\n" in info["company_name"]:
            after = info["company_name"].replace("\r\n"," ")
            info["company_name"] = after

        # # content dict
        # save all information to be inputted here for easy access
        content = {
            "company_name":"Company name not found",
            "month_year": datetime.today().strftime("%B %Y"),
            "date": datetime.today().strftime("%d %B %Y"),
            "biz_desc": "NA",
            "biz_summary": "NA",
            "deal_overview":"NA",
            "sit_overview":[],
            "deal_stage":[],
            "deal_rationale":[],
            "next_step":"NA",
            "deal_intelligence":"NA",
            "found_year":"NA",
            "dominant_country": "NA",
            "region":"NA",
            "num_countries":"NA",
            "employee_num": "NA",
            "link":None,
            "brand_name":["Sample brand name 1", "Sample brand name 2", "Sample brand name 3"],
            "brand_desc":{
                "Sample brand name 1": ["Sample brand description line","Sample brand description line"],
                "Sample brand name 2": ["Sample brand description line","Sample brand description line"],
                "Sample brand name 3": ["Sample brand description line","Sample brand description line"]
            },
            "social_work_desc":["Sample social work line", "Sample social work line"],
            "sch_network_desc":["Sample school network line","Sample school network line"],
            "corp_solution_desc":["Sample corporate solution line","Sample corporate solution line"],
            "capital_id":"NA",
            "orbis_id":"NA"
        }

        # saving info data into content dict
        country_match = {"US":["United States of America", "America", "US", "USA"],
                    "China":["China","Greater China"],
                    "SOUTHEAST Asia":["ASEAN"]}
        continent = {
            "Europe":["Albania","Andorra","Austria","Belarus","Belgium","Bosnia and Herzegovina","Bulgaria",
                      "Croatia","Czech Republic","Denmark","Estonia","Finland","France","Germany","Greece","Hungary","Iceland",
                      "Ireland","Italy","Latvia","Liechtenstein","Lithuania","Luxembourg","Moldova","Montenegro",
                      "Netherlands","North Macedonia","Norway","Poland","Portugal","Romania","Russia","Republic of Serbia",
                      "Slovakia","Slovenia","Spain","Sweden","Switzerland","Greenland","Ukraine","United Kingdom","Armenia","Azerbaijan","Cyprus","Georgia","Turkey"],
            "SOUTH Asia":["Bangladesh","Bhutan","India","Nepal","Pakistan","Sri Lanka","Afghanistan"],
            "SOUTHEAST Asia":["Brunei","Cambodia","Philippines","Indonesia","Laos","Malaysia","Myanmar","Singapore","Thailand","Vietnam"],
            "EAST Asia":["China","Japan","Mongolia","Taiwan"],
            "CENTRAL Asia":["Kazakhstan","Kyrgyzstan","Tajikistan","Turkmenistan","Uzbekistan"],
            "Middle East":["Saudi Arabia","Bahrain","United Arab Emirates","Yemen",
                           "Iraq","Iran","Israel","Jordan","Kuwait","Lebanon","Oman","Qatar","Syria"]
        }
        special_split = ["col_summary","region","other_info"]
        no_split = ["company_name","link","dominant_country"]
        norm_split = ["biz_desc","deal_intelligence","next_step"]
        for category in info.keys():
            if info[category] != None:
                if category in special_split:
                    if category == "col_summary":
                        content["deal_overview"] = info[category]["Deal Intelligence info"][2:]
                        content["biz_summary"] = info[category]["Business Description"][2:]
                    elif category == "region":
                        # get non-repeated list of country names
                        temp_cty = info["region"].split(";")
                        temp_cty += info["dominant_country"].split(", ")
                        for cty in temp_cty:
                            for country_name in country_match.values():
                                if cty in country_name:
                                    temp_cty[temp_cty.index(cty)] = list(country_match.keys())[list(country_match.values()).index(country_name)]
                        while "" in temp_cty or " " in temp_cty:
                            if "" in temp_cty:
                                temp_cty.remove("")
                            elif " " in temp_cty:
                                temp_cty.remove(" ")
                        content[category] = list(set(tuple(temp_cty)))
                        content["num_countries"] = len(content[category])
                    elif category == "other_info":
                        for inf in info["other_info"]:
                            inform = inf.split(":")
                            if inform[0] == "yearFounded":
                                content["found_year"] = inform[1]
                            elif inform[0] == "numOfEmployees":
                                content["employee_num"] = inform[1]
                    else:
                        content[category] = [info[category]]
                elif category in no_split:
                    if len(info[category]) != 0 and info[category] != " ":
                        content[category] = info[category]
                elif category in norm_split:
                    content[category] = re.split("[.!?]\s{1,}", info[category])
                    
        # # # text processing
        to_be_processed = []
        for sentence in content["deal_intelligence"]:
            to_be_processed += sentence.split("\r\n")

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
            "sit_overview":["launch","strategic review", "strategic","review", "shortlisted", "looking to", "sell", "business expansion", "expand", "mandate", "acquisition","acquired","invested","sales"],
            "deal_stage":["initial public offering", "expected to", "in a process", "plans to", "plan to", "early stage", "decision"],
            "deal_rationale":["funding used for", "funding would be used for", "investment", "funding", "focus on", "approach","potential","ebitda","gbp"]
        }

        # processing and sort Deal_intelligence sentences
        for line in range(len(to_be_processed)):
            for k in kw.values():
                # get key aka the information category with values in k
                information_category = list(kw.keys())[list(kw.values()).index(k)]
                # # for each word in k (containing all relevant words/phrases)
                for string in k:
                    if string in to_be_processed[line].lower():
                        content[information_category].append(to_be_processed[line])
        # remove duplicates in list
        for information_category in kw:
            content[information_category] = list(set(tuple(content[information_category])))

        # further process other info and store in content dict
        other_info = info["other_info"]
        for information in other_info:
            info = information.split(":")
            if info[0] == "numOfEmployees":
                content["employee_num"] = info[1]
            elif info[0] == "yearFounded":
                content["found_year"] = info[1]
            elif info[0] == "CapIQ_CompanyID":
                content["capital_id"] = info[1]
            elif info[0] == "Orbis_BvdID":
                content["orbis_id"] = info[1]

        # open an instance of PowerPoint
        ppt_app = comtypes.client.CreateObject("PowerPoint.Application")
        # open template asset pack slides
        template_path = "../../../results/input/templates/AssetPackSample_v0.61.pptx"
        template_abs_path = os.path.abspath(template_path)
        presentation_duplicate = ppt_app.Presentations.Open(template_abs_path)


        # # save slide, open duplicated slide
        save_path = """../../../results/output/asset_packs/{company_name}-AssetPack_fullTesting.pptx""".format(company_name=content["company_name"])
        save_abs_path = os.path.abspath(save_path)
        presentation_duplicate.SaveAs(save_abs_path)
        presentation = presentation_duplicate

        # editing content in slides
        presentation = gen_s1(presentation,content)
        print("slide 1 done")
        bulleted_category_s4 = ["sit_overview","deal_stage","deal_rationale","biz_desc","next_step"]
        presentation = gen_s4(presentation, content, bulleted_category_s4)
        print("slide 4 done")
        bulleted_category_s6 = ["biz_desc"]
        presentation = gen_s6(presentation,content,bulleted_category_s6)
        print("slide 6 done")
        presentation = gen_s7(presentation,content,country_flag,countries,color_hier,continent)
        print("slide 7 done")
        bulleted_category_s8 = ["brand_desc"]
        presentation, brand_counter = gen_s8(presentation, content, bulleted_category_s8)
        print("slide 8 done")
        bulleted_category_s9 = ["brand_desc","social_work_desc","sch_network_desc","corp_solution_desc"]
        presentation = gen_s9(presentation, content, bulleted_category_s9, brand_counter)
        print("slide 9 done")
        bulleted_category_s11 = ["deal_intelligence"]
        presentation = gen_s11(presentation,content,bulleted_category_s11)
        print("slide 12 done")

        success = True
    except Exception as e:
        print(e)
        success = False
    return success


if __name__ == '__main__':
    import uvicorn
    uvicorn.run("asset_pack_generator_v12:app", host='127.0.0.1', port=5010, reload=True)