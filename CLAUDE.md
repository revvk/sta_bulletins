# Project Overview
This project is a script to build bulletins for St. Andrew’s Episcopal Church, taking data from the planning spreadsheets, stored data (mostly in YAML), and various other document templates and images. 

## Project Pieces Completed:
9 am Sunday bulletin - generates successfully! 
11 am Sunday bulletin - generates successfully!
8 am Sunday bulletin - generates successfully!
Reading sheets - generate successfully!

## Outstanding Bugs (Fix First Before Project Work)
The 11 am bulletin for 2026-03-15 has the following errors:
* A blank spacer line between the rubric “Be seated” and the heading for the scriptures, “The Scriptures: Ephesians 5:8-14”. The rubric should be on the line immediately above “Be seated” (this is not a problem at 9 am because there is a Children’s Sermon header first)
* The preacher is Tim Jenkins; in the name mapping for the preacher, we should include Tim —> Tim Jenkins. (Also in 9 am bulletin)
* The prayers of the people are Form V; but for Form V, after the reader ways “we pray to you, O Lord.” The “Lord, have mercy.” should be bold (the people say it together). (Also in 9 am bulletin)
* The closing hymn at 11 am is “O for a thousand tongues to sing” #493 — but the script generated it with the 9 am praise band version of of the song. I think at 11 am, if the song exists in the hymnal list, then we sing the hymnal version — even if there is a praise band version. At 9 am, if the song exists in the praise band lyrics we use those — even if there is a hymnal version.

## Project Pieces Outstanding (detailed notes for next piece to work on are below):
* Special services - the proper liturgies for major feasts — Ash Wednesday, Maundy Thursday, Good Friday, Easter Sunrise, etc each have unique prayers and characteristics. Most of them — Palm Sunday is the exception — do not vary from year to year.
* Hidden Springs bulletin - this bulletin is for our weekly service at a senior living center, and follows a slightly different liturgical pattern, has a different music selection process, and is printed in large print
* Special Covers - seasonally, at special services, and for Hidden Springs & Chateau bulletins; we need to be able to specify special bulletin covers
* Song Import and Insertion - need to build a script that can insert new music, simply formatted in markdown. 

# Remember: 
All the the propers (which is the Episcopal word for the readings and prayers that are assigned for that day) and all the formatting styles are consistent across the 8 am, 9 am, and 11 am bulletins, as well as special service bulletins.

Architect the code so that as much of the code base can be reused as possible. We will need to be able to specify when running generate.py which bulletin to generate, or generate all of them.

# Proper Liturgies for Special Days & Special Covers
Let’s work on the Proper Liturgies for Special Days, since Holy Week is coming up soon.
* I have put into the source_documents folder Word versions of the pages documents of the services for Palm Sunday, Maundy Thursday, and Good Friday from 2025. 
* You also already have the BCP text for the liturgies, which are on pp. 264-282 (let’s not worry about Holy Saturday or Easter Vigil for now).
* The music for Palm Sunday will be based on the same planning documents as normal, but for the non-Sunday services it will only be in the Service Music tab of the Google Sheet. Currently all the musical selections are empty — so for now we won’t have hymns or lyrics as we generate the bulletins
* Styles generally remain consistent with other bulletins, but the long readings from the gospels have specific styling. I can provide special files for the gospel readings by parts, so that you don’t have to parse those out from the bulletins. Let me know when you’re ready for those, and you can parse them into a YAML file for reuse going forward.
* These special liturgies will force us to build in support for special covers. So far there are two special covers made, for Palm Sunday and Good Friday, in the templates folder.