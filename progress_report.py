"""
Takes the desired columns from formsite csv file,
and puts pre-determined info into docx file where the students are separated by teams
with bullet lists of the PM, members, added up scores, and comments.

INSTALL:
pip install pandas
pip install python-docx
pip install numpy
"""
import pandas as pd
import numpy as np
from docx import Document
from docx.shared import Pt

document = Document()
reports = pd.read_csv("progress_report_form_2.csv")
report_dict = {}
num = int(input("Enter progress report num: "))
reports = reports.loc[reports['Progress Report #'] == num]
reports.to_csv("progress_report_#{}.csv".format(str(num)),index=False)
print(reports["PM Name"])

##{
##    Row # : num,
##    Row Data: {
##        Project Type: Project Type,
##        Project Title : Project Title,
##        PM Name : Name,
##        Team Members : [Team Member Name: Team member name,
##             data: {
##                Overall: Score,
##                Professionalism: Score,
##                Technical: Score,
##                Communication: Score
##                Promptness: Score,
##                Ability to Get Along: Score,
##                Ability to Learn: Score
##            },
##            ...
##            ]
##       Comments : comment
##    }
##}

reports_cols = list(reports.columns)

project_type = [reports_cols[3]]
pm_name = [reports_cols[5]]
tm_1_evals = reports_cols[9:19]
tm_2_evals = reports_cols[19:28]
tm_3_evals = reports_cols[28:37]
tm_4_evals = reports_cols[37:46]
tm_5_evals = reports_cols[46:55]
comments_col = ["Any obstacles or problems facing your project?"]
team_member_evals = tm_1_evals+tm_2_evals+tm_3_evals+tm_4_evals+tm_5_evals
working_cols = project_type + pm_name + team_member_evals+comments_col #takes desired columns


num_rows = reports.index.tolist()

tm_reports = reports.loc[:, working_cols]
print(tm_reports)
drop_suffix = [" \(Item #31\)"," \(Item #34\)"," \(Item #38\)"," \(Item #41\)"," \(Item #44\)"]
for suf in drop_suffix:
    tm_reports.columns = tm_reports.columns.str.replace(suf, "")


for i in num_rows: #makes hashmaps of each row, then populates the report data hashmap
    row = tm_reports.loc[i]
    pm = row["PM Name"]
    row_dict = None
    project_type = row["What type of project do you have?"]
    project_title = row["Project Title"]
    tm_list = []
    for j in range(1,6):
        tm_dict = {}
        overall = None
        pro = None
        tech = None
        comm = None
        prompt = None
        g_along = None
        learn = None
        comment = None
        if j == 1:
            name = row["Team member name"]
            comment = row["Comments / Suggestions"]
        else:
            name = row["Team member name.{}".format(str(j-1))]
            comment = row["Comments / Suggestions.{}".format(str(j-1))]
        
        if j != 2:
            overall = row["Team Member#{} Overall".format(str(j))]
            pro = row["Team Member#{} Professionalism".format(str(j))]
            tech = row["Team Member#{} TechnicalSkills".format(str(j))]
            comm = row["Team Member#{} CommunicationSkills".format(str(j))]
            prompt = row["Team Member#{} Promptness".format(str(j))]
            g_along = row["Team Member#{} Ability toGet Along with Others".format(str(j))]
            learn = row["Team Member#{} Ability to Learn".format(str(j))]
        else:
            overall = row["Team Member # {} Overall".format(str(j))]
            pro = row["Team Member # {} Professionalism".format(str(j))]
            tech = row["Team Member # {} TechnicalSkills".format(str(j))]
            comm = row["Team Member # {} CommunicationSkills".format(str(j))]
            prompt = row["Team Member # {} Promptness".format(str(j))]
            g_along = row["Team Member # {} Ability toGet Along with Others".format(str(j))]
            learn = row["Team Member # {} Ability to Learn".format(str(j))]

        comment = row["Any obstacles or problems facing your project?"]
        tm_dict = {"Team Member Name": name,
                   "data" : {"Overall": overall,
                           "Professionalism": pro,
                            "Technical": tech,
                            "Communication": comm,
                            "Promptness": prompt,
                            "Ability to Get Along": g_along,
                            "Ability to Learn": learn,
                           }
                   }
        
        tm_list.append(tm_dict)
    row_dict = {
        "Row #" : i,
        "Row Data": {
            "Project Type" : project_type,
            "Project Title" : project_title,
            "PM Name" : pm,
            "Team Members" : tm_list,
            "Comments": comment
        }
    }
    report_dict["Row #{}".format(i)] = row_dict
#print(report_dict)

#Make docx version
for row in report_dict:
    row_num = report_dict[row]
    r_data = row_num["Row Data"]
    #print(r_data)
    p_title = "{} ({})".format(r_data["Project Title"], r_data["Project Type"]) #Makes project title bold
    p_pm_name = "PM: {}".format(r_data["PM Name"])
    p_tm_title = "Team Members:"

    p = document.add_paragraph(style='List Bullet')
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    p_bold = p.add_run(p_title)
    p_bold.bold = True

    p_pm_name = document.add_paragraph(p_pm_name, style='List Bullet 2')
    p_pm_name.paragraph_format.space_before = Pt(3)
    p_pm_name.paragraph_format.space_after = Pt(3)

    p_tm_title = document.add_paragraph(p_tm_title, style='List Bullet 2')
    p_tm_title.paragraph_format.space_before = Pt(3)
    p_tm_title.paragraph_format.space_after = Pt(3)

    for n in r_data["Team Members"]: 
        tm_name = "{}".format(n["Team Member Name"])
        score = 0
        #print(type(n["data"]["Overall"]))
        try:
            score += int(n["data"]["Overall"])
            score += int(n["data"]["Professionalism"])
            score += int(n["data"]["Technical"])
            score += int(n["data"]["Communication"])
            score += int(n["data"]["Promptness"])
            score += int(n["data"]["Ability to Get Along"])
            score += int(n["data"]["Ability to Learn"])
        except ValueError:
            score = np.nan
        tm_name = tm_name.rsplit('\t', 1)[0]
        #print(tm_name + ' ' + str(score))
        tm_name += " ({}/70)".format(str(score))
        tm_name = document.add_paragraph(tm_name, style='List Bullet 3')
        tm_name.paragraph_format.space_before = Pt(3)
        tm_name.paragraph_format.space_after = Pt(3)

    comment_title = document.add_paragraph("Comments/Issues", style='List Bullet 2')
    comment_title.paragraph_format.space_before = Pt(3)
    comment_title.paragraph_format.space_after = Pt(3)

    comment_text = r_data["Comments"]
    if r_data["Comments"] is np.nan:
        comment_text = "N/A"
    else:
        comment_text = r_data["Comments"]
    comments = document.add_paragraph(comment_text, style='List Bullet 3')
    comments.paragraph_format.space_before = Pt(3)
    comments.paragraph_format.space_after = Pt(3)
    space = document.add_paragraph("\n")
        

document.save("progress_report_#{}.docx".format(str(num)))






