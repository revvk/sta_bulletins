# Project Overview
This project is a script to build bulletins for St. Andrew’s Episcopal Church, taking data from the planning spreadsheets, stored data (mostly in YAML), and various other document templates and images. 

## Project Pieces Completed:
9 am Sunday bulletin - generates successfully! 
11 am Sunday bulletin - generates successfully!
8 am Sunday bulletin - generates successfully!
Reading sheets - generate successfully!
Holy Week bulletins - in progress.

## Outstanding Bugs (Fix First Before Project Work)
### The bulletins for 2026-03-22 had the following errors:
* The first verse of Psalm 130 has two lines in the first half of the verse. “Out of the depths I have called to you, O LORD;” then “Lord, hear my voice.” Looking through the YAML file, I believe this error is systemic. The first half of the verse is never properly broken into two lines when it is two lines in the Book of Common Prayer. So rather than fixing Psalm 130, we need to update the YAML for all verses that have more than one line in the first half of the verse.
* The list of clergy in the Prayers of the People: “Andrew, Logan, Paulette, Mike, Gene” is missing Katie. This happened because the source data was somewhat outdated, but all POP forms that list the clergy should include Katie between Paulette and Mike.”
* The last petition in the Prayers of the People begins, “We pray for all who have died ,”. That space before the comma is correct, but the cross that is there is not printed.
### The Palm Sunday (03-29-2026) bulletins have the following errors:
* Because the Palm Sunday service at 9 and 11 am begins outside, we do not need the initial rubric or the Prelude heading. (8 am doesn’t have those to begin with)
* There should be a space between the heading for “The Liturgy of the Palms” and the rubric. The rubric itself should read, “After getting their palms, the congregation with gather at the designated spot forming a processional route back towards the church.
* After the Celebrant says “Blessed is he who comes in the name of the Lord” in the Liturgy of the Palms, there should be a cross symbol.
* As an exception to the general rule, the lyrics to “All glory, laud and honor” should be printed even in the 11 am bulletin, because the congregation will be outside without hymnals.
* The quotes from the prophets in the Matthew reading should be rendered as poetry.
* The poetry section in the Philippians readings is still not rendering as poetry.
* Make the font size for the part names (Narrator, Pilate, etc) be 9 pt instead of 11 pt.
* In place of the Sermon, we will have a Musical Reflection. The rubric below the sermon should read, “In response to our Lord’s passion and death, we will observe a time of quiet prayer and reflection.”
* In place of the Prayers of the People, on Palm Sunday we say the Great Litany — BCP pp. 148-153 (ending with “Lord, have mercy upon us.”) It the petition about the fellowship of saints, it should read “That it may please thee to grant that, in the fellowship of Saint Andrew and all the saints…”
## Project Pieces Outstanding (detailed notes for next piece to work on are below):
* None at present. Fix bugs.

# Remember: 
All the the propers (which is the Episcopal word for the readings and prayers that are assigned for that day) and all the formatting styles are consistent across the 8 am, 9 am, and 11 am bulletins, as well as special service bulletins.

Architect the code so that as much of the code base can be reused as possible. We will need to be able to specify when running generate.py which bulletin to generate, or generate all of them.