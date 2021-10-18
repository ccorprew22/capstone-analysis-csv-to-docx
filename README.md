# Capstone Progress Report Analysis CSV to Docx
[![Build Status](https://app.travis-ci.com/ccorprew22/capstone-analysis-csv-to-docx.svg?branch=main)](https://app.travis-ci.com/ccorprew22/capstone-analysis-csv-to-docx)

This program takes the PM progress report in a CSV format and turns it into a docx file in a bullet point format with the team member scores and comments, sorted by executive team mentor. Also makes a CSV file with the selected progress report number. Look at sample to see an example. Scores/Names aren't real.

This program currently works if it matches column headers and the Formsite template provided. However, the professor may make changes.

## Note!!!
You will still have to go through the progress report file, this program is only to reduce the amount of typing, formatting, and busy work.

+ To get the CSV file, go on to the Formsite and export a CSV file.
+ To install dependencies, enter in the terminal `pip install -r requirements`.
+ The format of the column names may change, so you will need to adjust the code starting at `line 73` (for loop to change column titles)
+ Some names will need to be manually added after making the doc. Most of the time the PM will put the name in the team member comment section. Also you have to remove any extra empty spaces because I checked for five team members since there was a large amount of empty columns/unused columns.
+ Also some scores will also need to be manually add due to "?" being found
occasionally where numbers are supposed to be.
+ You have to manually go through and
+ If you make any changes, feel free to clone this repo and push to this one.
