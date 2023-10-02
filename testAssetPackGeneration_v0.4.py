# # install comtypes, google-cloud-bigquery
# pip install comtypes --trusted-host=files.pythonhosted.org

import comtypes.client
# to get current date
from datetime import datetime
# connection to google cloud bigquery
import os
from google.cloud import bigquery
import re

# # # big query
# # Get Num to get row
# TODO: get Num from webpage when they send request to create asset pack
row_num = 1094

# # TODO: Connection to Big Query
json_path = "src/mod/asset_pack_gen/testing_bigquery_vertexai_service_account.json"
json_abs_path = os.path.abspath(json_path)
os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = json_abs_path

# authenticate access
bq = bigquery.Client()
# sql_query - extract the whole row out
sql_query = """SELECT * FROM `testing-bigquery-vertexai.web_UI.Rolling_08-09-23` WHERE Num = '{num}'""".format(num = row_num) 
# run query
query_job = bq.query(sql_query)
# get result
results = query_job.result()

for row in results:
    cn = row["_Target"]
    di = row["Deal_Intelligence_info"]
    dc = row["Target_country"]
    # sbd = row["Short BD"]
    bd = row["Business_Description"]
    region = row["Target_Region"]
    ns = row["KPMG_View___Redacted"]
    others = row["Asset_pack_"]
    link = str(row["Website"])

# # deal summary and business summary
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
        do = r[0][1:len(r[0])-1]


# # # text processing
to_be_processed = di.split("\n")

# remove empty string, useless lines
# get index of elements to be removed
remove = []
kw_remove = ["original source", "updated on", "checked on", "link to"]
kw_startwith = ["by", "press release"]
for tp in range(len(to_be_processed)):
    # remove empty lines
    if to_be_processed[tp] == "":
        remove.append(tp)
    # remove useless lines
    else:
        for kwr in kw_remove:
            if kwr in to_be_processed[tp].lower() and tp not in remove:
                remove.append(tp)
        for kws in kw_startwith:
            if to_be_processed[tp].lower().startswith(kws) and tp not in remove:
                remove.append(tp)

for r in range(len(remove)):
    ind_remove = remove[r]
    to_be_processed.remove(to_be_processed[ind_remove-r])

# sort sentences based on keywords
kw = {
    "sit_overview":["launch","strategic review", "strategic","review", "shortlisted", "looking to", "sell", "business expansion", "expand", "mandate", "acquisition","acquired"],
    "deal_stage":["initial public offering", "expected to", "in a process", "plans to", "plan to", "early stage", "decision"],
    "deal_rationale":["funding used for", "funding would be used for", "investment", "funding", "focus on", "approach"],
    "biz_desc":["established as", "providing", "in house", "production", "largest", "market cap", "leading"],
    "found_year": ["established in", "found in", "founded in"]
}

# # content dict
# save all information here for easy access
# scrapping not possible for the remaining information, need
content = {
    "company_name":cn,
    "month_year": datetime.today().strftime("%B %Y"),
    "date": datetime.today().strftime(r"%d %B %Y"),
    "biz_desc": bd.split("\n"),
    "biz_summary": bd,
    "deal_overview":do.split("\n"),
    "sit_overview":[],
    "deal_stage":[],
    "deal_rationale":[],
    "next_step":ns.split("\n"),
    "deal_intelligence":di.split("\n"),
    "found_year":1900,
    "dominant_country": dc,
    "region":region.split(";"),
    "num_countries":len(region.split(";")),
    "emp_num": 0,
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
for l in range(len(to_be_processed)):
    for k in kw.values():
        info_cat = list(kw.keys())[list(kw.values()).index(k)]
        for string in k:
            if string in to_be_processed[l].lower():
                if info_cat == "found_year":
                    start_ind = to_be_processed[l].lower().index(string)
                    fy = int(to_be_processed[l][start_ind+len(string):start_ind+len(string)+5])
                    content[info_cat] = fy
                else:
                    content[info_cat].append(to_be_processed[l])

# check what is in others
other_info = others.split(";")
for oi in other_info:
    info = oi.split(":")
    if info[0] == "numOfEmployees":
        content["emp_num"] = info[1]
    elif info[0] == "yearFounded":
        content["found_year"] = info[1]


# open an instance of PowerPoint
ppt_app = comtypes.client.CreateObject("PowerPoint.Application")
# open template asset pack slides
template_path = "results/input/templates/AssetPackSample_v0.4.pptx"
template_abs_path = os.path.abspath(template_path)
prs_dup = ppt_app.Presentations.Open(template_abs_path)

# # save slide, open duplicated slide
save_path = """results/output/asset_packs/{company_name}-AssetPack.pptx""".format(company_name=content["company_name"])
save_abs_path = os.path.abspath(save_path)
prs_dup.SaveAs(save_abs_path)
prs = prs_dup

pp_constants = comtypes.client.Constants(prs)

# for i in pp_constants.enums.PpActionTypes.ppActionHyperlink:
#     print(i)

# # # # slide 1 - done
# # slide index are 1-based
# s1 = prs.Slides(1)
# for s in range(1,s1.Shapes.Count+1):
#     tf = s1.Shapes(s).TextFrame.TextRange
#     if tf.Text.find("{") > -1 and tf.Text.find("}") > -1:
#         start_ind = tf.Text.index("{")
#         end_ind = tf.Text.index("}")

#         info_cat = tf.Text[start_ind+1:end_ind]

#         if info_cat in content.keys():
#             pre_process = tf.Text
#             tf.Text = pre_process.replace("{"+info_cat+"}",str(content[info_cat]))
#     else:
#         continue


# # # # slide 4
# s4 = prs.Slides(4)
# # if info_cat in this list, will need to take extra care -> bulleted list
# bul_cat_s4 = ["sit_overview","deal_stage","deal_rationale","biz_desc","next_step"]

# for p in range(1,s4.Shapes.Count+1):

#     tf = s4.Shapes(p).TextFrame.TextRange
    
#     for pa in range(len(tf.Paragraphs())):

#         if tf.Paragraphs(pa).Text.find("{") > -1 and tf.Paragraphs(pa).Text.find("}") > -1:
#             # get info_cat
#             start_ind = tf.Paragraphs(pa).Text.index("{")
#             end_ind = tf.Paragraphs(pa).Text.index("}")

#             info_cat = tf.Paragraphs(pa).Text[start_ind+1:end_ind]

#             # set up bulleted list in respective paragraph
#             # insert in empty paragraphs first, then input text in
#             # change bullet character to 2024 (8212 in decimal)
#             if info_cat in content.keys() and info_cat in bul_cat_s4:
#                 # insert in empty paragraph
#                 for pa in range(len(content[info_cat])-1):
#                     tf.Paragraphs(pa).Text += "\r\n"
#                 # insert in text
#                 for par in range(len(content[info_cat])):
#                     tf.Paragraphs(par+1).ParagraphFormat.Bullet.Type = 1
#                     # changed using unicode code point - decimal
#                     tf.Paragraphs(par+1).ParagraphFormat.Bullet.Character = 8212
#                     tf.Paragraphs(par+1).Font.Color.RGB = 0
#                     tf.Paragraphs(par+1).ParagraphFormat.SpaceAfter = 0
#                     tf.Paragraphs(par+1).ParagraphFormat.SpaceBefore = 0
#                     tf.Paragraphs(par+1).Text = content[info_cat][par]

#             elif info_cat in content.keys() and info_cat not in bul_cat_s4:
#                 pre_process = tf.Paragraphs(pa).Text

#                 # input as per normal
#                 tf.Paragraphs(pa).Text = pre_process.replace("{"+info_cat+"}",str(content[info_cat]))
#             if info_cat == "deal_overview":
#                 for tfp in range(len(tf)):
#                     tf.Font.Italic = True


# # # # slide 6
# s6 = prs.Slides(6)
# bul_cat_s6 = ["biz_desc"]
# for a in range(1,s6.Shapes.Count+1):

#     # check that the {} exist in the shape
#     if s6.Shapes(a).hasTextFrame:

#         tf = s6.Shapes(a).TextFrame.TextRange

#         for p in range(len(tf.Paragraphs())):

#             if tf.Paragraphs(p).Text.find("{") > -1 and tf.Paragraphs(p).Text.find("}") > -1:

#                 start_ind = tf.Paragraphs(p).Text.index("{")
#                 end_ind = tf.Paragraphs(p).Text.index("}")
#                 # get info key
#                 info_cat = tf.Paragraphs(p).Text[start_ind+1:end_ind]

#                 if info_cat in content.keys() and info_cat in bul_cat_s6:

#                     # add empty paragraph
#                     for c in range(len(content[info_cat])-1):
#                         tf.Paragraphs(p).Text += "\r\n"
#                     # input info
#                     for par in range(len(content[info_cat])):
#                         tf.Paragraphs(par+1).ParagraphFormat.Bullet.Type = 1
#                         tf.Paragraphs(par+1).Font.Color.RGB = 0
#                         tf.Paragraphs(par+1).ParagraphFormat.SpaceAfter = 0
#                         tf.Paragraphs(par+1).ParagraphFormat.SpaceBefore = 0
#                         tf.Paragraphs(par+1).Text = str(content[info_cat][par])
                
#                 elif info_cat in content.keys() and info_cat not in bul_cat_s6:

#                     pre_process = tf.Paragraphs(p).Text
#                     # input info
#                     tf.Paragraphs(p).Text = pre_process.replace("{"+info_cat+"}",str(content[info_cat]))
#             else:
#                 continue


# # # slide 7 - half-done (TODO check order of shape, hyperlink)
# # # country flag
s13 = prs.Slides(13)
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
    "Belarus":9, # latest flag collection
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
}

# # map pieces
s7 = prs.Slides(7)
groups = []
colour_hier = [(0,51,141), (30,73,226), (172,234,255), (0,184,245), (12,35,60), (114,19,234), (253,52,156)]
country = {
    "South America":5,
    "America":16,
    "Canada":17,
    "Greenland":57,
    "Haiti":22,
    "Pakistan":50,
    "Japan":51,
    "Russia":53,
    "Middle East":49,
    "Azerbaijan":55,
    "India":92,
    "Italy":60,
    "France":68,
    "United Kingdom":69,
    "Bosnia and Herzegovina":75,
    "Norway":86,
    "Africa":90,
    "Philippines":94,
    "China":98,
    "Taiwan":99,
    "Nicaragua":8,
    "Guatemala":10,
    "Belize":11,
    "Costa Rica":13,
    "Honduras":14,
    "Mexico":15,
    "Bahamas":25,
    "Lesser Antilles":29,
    "Malaysia":97,
    "Indonesia":96,
    "Papua New Guinea":38,
    "New Zealand":31,
    # all corrected below
    "Australia":28, 
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
    "Hungary":71, 
    "Greece":70, 
    "Belarus":68, 
    "Moldova":67,
    "Ukraine":66,
    "Ireland":65,
    "Germany":62,
    "Portugal":61,
    "Netherlands":60,
    "Spain":59,
    "Belgium":58,
    "Switzerland":57,
    "Luxembourg":56,
    "Liechtenstein":54,
    "Andorra":53,
    "Austria":52,
    "Iceland":50,
    "Turkmenistan":48,
    "Uzbekistan":47,
    "Georgia":46,
    "Kazakhstan":45,
    "Armenia":44,
    "Mongolia":41,
    "Laos":38,
    "Cambodia":37,
    "Myanmar":36,
    "Thailand":35,
    "Vietnam":34,
    "Afghanistan":32,
    "Kyrgyzstan":30,
    "Puerto Rico":26,
}


# s7.Shapes(100).TextFrame.TextRange.Characters(s7.Shapes(100).TextFrame.TextRange.Text.find(r"website {link}")+1,len(r"website {link}")).ActionSettings(1).Action = 7
# print(s7.Shapes(100).TextFrame.TextRange.Characters(s7.Shapes(100).TextFrame.TextRange.Text.find(r"website {link}")+1,len(r"website {link}")).ActionSettings(1).Action == 0)

# edit slide
for mp in range(1,s7.Shapes.Count+1):
    # print(mp, s7.Shapes(mp).ActionSettings(1))
    if s7.Shapes(mp).hasTextFrame:
        tf = s7.Shapes(mp).TextFrame.TextRange


        if tf.Text.find("{") > -1 and tf.Text.find("}") > -1:

            # find all instance of { and } in text
            start_ind = [x.start() for x in re.finditer("{",tf.Text)]
            end_ind = [y.start() for y in re.finditer("}",tf.Text)]

            while len(start_ind) >= 1:
                pre_processed = tf.Text
                ind_start = start_ind[0]
                ind_end = end_ind[0]

                info_cat = pre_processed[ind_start+1:ind_end]

                # change text
                if info_cat in content.keys():
                    # tf.Text = pre_processed.replace("{"+info_cat+"}",str(content[info_cat]))
                    # if info_cat is link, add hyperlink
                    if info_cat == "link":
                        # pass
                        # select the word
                        # tf.Text = pre_processed.replace(r"{link}","link")
                        ind_start_link = tf.Text.index(r"{link}")
                        hyperlink_range = tf.Characters(Start=ind_start_link+1, Length=6)
                        hyperlink_range.ActionSettings(1).Hyperlink.Address = r"https://" + content["link"]
                        hyperlink_range.Text = "link"
                    else:
                        tf.Text = tf.Text.replace("{"+info_cat+"}",str(content[info_cat]))

                start_ind = [x.start() for x in re.finditer("{",tf.Text)]
                end_ind = [y.start() for y in re.finditer("}",tf.Text)]
                if ind_start in start_ind:
                    start_ind.remove(ind_start)
                    end_ind.remove(ind_end)


    # elif s7.Shapes(mp).hasTextFrame == False:
        # reg = content["region"]
    #     # tgt_country = ""
    #     for cty in reg:
    #         for co in country.keys():
    #             if co in cty:
    #                 tgt_country = co
    #                 # change fill of shape
    #                 color_ind = reg.index(cty) % 7
    #                 chosen_color = colour_hier[color_ind][0] + colour_hier[color_ind][1]*256 + colour_hier[color_ind][2]*256**2
    #                 s7.Shapes(country[tgt_country]).Fill.ForeColor.RGB = chosen_color
    #                 # add in country flag
    #                 left = s7.Shapes(country[tgt_country]).Left
    #                 top = s7.Shapes(country[tgt_country]).Top
    #                 if tgt_country in country_flag.keys():
    #                     s13.Shapes(country_flag[tgt_country]).Copy()
    #                     s7.Shapes.Paste()
    #                     # get pic index
    #                     s7.Shapes(len(s7.Shapes)).Left = left
    #                     s7.Shapes(len(s7.Shapes)).Top = top
    #                     # add name below flag
    #                     table_top = top + s7.Shapes(len(s7.Shapes)).Height
    #                     table_width = 40
    #                     table_height = 20
    #                     tb = s7.Shapes.AddTable(1,1,left,table_top,table_width,table_height)
    #                     tb.Table.Cell(1,1).Shape.TextFrame.TextRange.Font.Size = 9
    #                     tb.Table.Cell(1,1).Shape.Fill.ForeColor.RGB = 0 + 51*256 + 141*256**2
    #                     tb.Table.Cell(1,1).Shape.TextFrame.TextRange.Text = tgt_country


# # # # slide 8 - done
# s8 = prs.Slides(8)
# brand_counter = 0
# bul_cat_s8 = ["brand_desc"]
# for sp in range(1,s8.Shapes.Count+1):
#     if s8.Shapes(sp).hasTextFrame:
#         tf = s8.Shapes(sp).TextFrame.TextRange
        
#         if tf.Text.find("{") > -1 and tf.Text.find("}") > -1:
#             start_ind = tf.Text.index("{")
#             end_ind = tf.Text.index("}")

#             info_cat = tf.Text[start_ind+1:end_ind]

#             if info_cat in content.keys():
#                 pre_processed = tf.Text
#                 if info_cat == "brand_name":
#                     tf.Text = content[info_cat][brand_counter]
#                     brand_counter += 1
#                 elif info_cat in bul_cat_s8:
#                     # bulleted brand description
#                     # insert empty paragraph
#                     for lbd in range(len(content[info_cat][content["brand_name"][brand_counter]])-1):
#                         tf.Text += "\r\n"
#                     for par in range(len(content[info_cat][content["brand_name"][brand_counter]])):
#                         pre_processed = tf.Paragraphs(par+1).Text
#                         tf.Paragraphs(par+1).ParagraphFormat.Bullet.Type = 1
#                         # tf.Paragraphs(par+1).ParagraphFormat.Bullet.Character = 8226
#                         tf.Paragraphs(par+1).Font.Color.RGB = 0
#                         tf.Paragraphs(par+1).ParagraphFormat.SpaceAfter = 0
#                         tf.Paragraphs(par+1).ParagraphFormat.SpaceBefore = 0
#                         tf.Paragraphs(par+1).Text = content[info_cat][content["brand_name"][brand_counter]][par]
#                 else:
#                     tf.Text = pre_processed.replace("{"+info_cat+"}",str(content[info_cat]))
#             else:
#                 continue


# # # # slide 9
# s9 = prs.Slides(9)
# bul_cat_s9 = ["brand_desc","social_work_desc","sch_network_desc","corp_solution_desc"]
# for sm in range(1,s9.Shapes.Count+1):
#     if s9.Shapes(sm).hasTextFrame:

#         tf = s9.Shapes(sm).TextFrame.TextRange

#         if tf.Text.find("{") > -1 and tf.Text.find("}") > -1:
#             start_ind = tf.Text.index("{")
#             end_ind = tf.Text.index("}")

#             info_cat = tf.Text[start_ind+1:end_ind]

#             if info_cat in content.keys() and info_cat in bul_cat_s9:
#                 # input empty paragraphs
#                 # input data
#                 if info_cat == "brand_desc":
#                     for ci in range(len(content["brand_desc"][content["brand_name"][brand_counter]])-1):
#                         tf.Text += "\r\n" 

#                     for cic in range(len(content["brand_desc"][content["brand_name"][brand_counter]])):
#                         tf.Paragraphs(cic+1).ParagraphFormat.Bullet.Type = 1
#                         tf.Paragraphs(cic+1).ParagraphFormat.Bullet.Character = 8226
#                         tf.Paragraphs(cic+1).Text = content["brand_desc"][content["brand_name"][brand_counter]][cic]
            
#                 else:
#                     for ci in range(len(content[info_cat])-1):
#                         tf.Text += "\r\n" 
#                     for cic in range(len(content[info_cat])):
#                         tf.Paragraphs(cic+1).ParagraphFormat.Bullet.Type = 1
#                         tf.Paragraphs(cic+1).Font.Color.RGB = 0
#                         tf.Paragraphs(cic+1).ParagraphFormat.SpaceAfter = 0
#                         tf.Paragraphs(cic+1).ParagraphFormat.SpaceBefore = 0
#                         tf.Paragraphs(cic+1).Text = content[info_cat][cic]

#             elif info_cat in content.keys() and info_cat not in bul_cat_s9:
#                 pre_processed = tf.Text
#                 if info_cat == "brand_name":
#                     tf.Text = content["brand_name"][brand_counter]
#                 else:
#                     tf.Text = pre_processed.replace("{"+info_cat+"}",str(content[info_cat]))



# # # # slide 12
# s12 = prs.Slides(12)
# bul_cat_s12 = ["deal_intelligence"]
# for x in range(1,s12.Shapes.Count+1):

#     if s12.Shapes(x).hasTextFrame:

#         tf = s12.Shapes(x).TextFrame.TextRange

#         if tf.Text.find("{") > -1 and tf.Text.find("}") > -1:
#             start_ind = tf.Text.index("{")
#             end_ind = tf.Text.index("}")
#             # get info key
#             info_cat = tf.Text[start_ind+1:end_ind]

#             if info_cat in content.keys() and info_cat in bul_cat_s12:

#                 # insert empty paragraph
#                 for c in range(len(content[info_cat])):
#                     tf.Text += "\r\n"
#                 # input into the empty paragraph
#                 for par in range(len(content[info_cat])):
#                     tf.Paragraphs(par+1).ParagraphFormat.Bullet.Type = 1
#                     tf.Paragraphs(par+1).ParagraphFormat.Bullet.Character = 8226
#                     tf.Paragraphs(par+1).Text = content[info_cat][par]

