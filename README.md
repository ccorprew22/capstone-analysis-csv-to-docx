# Capstone Progress Report Analysis CSV to Docx
This program takes the PM progress report in a CSV format and turns it into a docx file in a bullet point format with the team member scores and comments.

+ To get the CSV file, go on to the Formsite and export a CSV file.
+ To install dependencies, enter in the terminal `pip install -r requirements`.
+ The format of the column names may change, so you will need to adjust the code starting at `line 73` (for loop to change column titles)
+ Some names will need to be manually added after making the doc. Most of the time the PM will put the name in the team member comment section.
+ Also some scores will also need to be manually add due to "?" being found
occasionally where numbers are supposed to be.
+ If you make any changes, feel free to clone this repo and push to this one.
