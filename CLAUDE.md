# Project Overview
This project is a script to build bulletins for St. Andrew’s Episcopal Church, taking data from the planning spreadsheets, stored data (mostly in YAML), and various other document templates and images. 

## Project Pieces Completed:
9 am Sunday bulletin - generates successfully! 
11 am Sunday bulletin - generates successfully!
8 am Sunday bulletin - generates successfully!
Reading sheets - generate successfully!
Holy Week bulletins - in progress.

## Scripture Poetry Formatting (Resolved)
* Scripture readings with poetry (e.g., Matthew 21:5/9, Philippians 2:6-11, Ephesians 5:14) are now properly detected and rendered with correct indentation.
* The solution uses Oremus’s own HTML `<br>` tag CSS classes to detect poetry structure and indent levels — no BibleGateway needed.
* Poetry is rendered with three indent levels via dedicated styles: "Reading (Poetry)", "Reading (Poetry Indent 1)", "Reading (Poetry Indent 2)".

## Project Pieces Outstanding (detailed notes for next piece to work on are below):
* Maundy Thursday (bugs)
* Good Friday (bugs)
* Hidden Springs Bulletins (not started)

# Remember: 
All the the propers (which is the Episcopal word for the readings and prayers that are assigned for that day) and all the formatting styles are consistent across the 8 am, 9 am, and 11 am bulletins, as well as special service bulletins.

Architect the code so that as much of the code base can be reused as possible. We will need to be able to specify when running generate.py which bulletin to generate, or generate all of them.