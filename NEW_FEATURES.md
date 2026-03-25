# Hidden Springs Bulletin Creation

Hidden Springs is a senior living center. St. Andrew’s holds a weekly service on Wednesday mornings. Most of the congregants are not Episcopalian, and so we have monthly communion on the first Sunday of the month rather than weekly communion. The changes to the bulletin generation are derive from these basic facts:

* Because the congregation is elderly, the bulletin is printed on bigger, US Letter, sized paper with bigger fonts.
* Preface all new styles with LP_ for large print, keeping the rest of the style name the same.
* There are custom front and back covers in the /templates folder. They are senior_living_front_cover.docx and senior_living_back_cover.docx because the general format is also used for occasional services at other senior living centers.
* Generally speaking, the service uses the propers for the Sunday prior, and the name of the day for the front cover is “Wednesday after the [Sunday Name]. 
** The exception is when the Wednesday is a feast day, in which case we use the Feast Day propers. In that case, the name of the day is just the name of the feast. March 25, the Feast of the Annunciation, is one such case.
* The propers for the service can be found in the “Hidden Springs Planner” sheet in the Rota & Liturgical Schedule Google Sheet.
* For the most part, except on first Sundays when we have Holy Communion, the service follows a pattern we call Liturgy of the Word (LOW). This is not actually in the BCP, but is more or less the proanophora from Holy Eucharist Rite II. The service type (LOW or HEII) is specified in column A of the “Hidden Springs Planner” sheet. 
* I’ve included five bulletin examples with the LOW in /source_documents/Hidden_Springs/ — one each from the seasons of Advent (LOW), Epiphany (LOW), Lent (HEII), Easter (LOW), and after Pentecost (LOW). You should be able to derive the structure and seasonal variations from those, but please ask any questions you need. HEII follows the basic structure of the existing bulletins.
* Music at this service is played via AAC files that have been generated from MIDI files
** All the music can be found in the folder /hidden_springs_music
** We are going to need to build a sub-plan in order to properly generate the bulletins with this music.
*** The music is in three sections (each a directory): Hymnal 1982 hymns; Other music (praise songs and hymns from other traditions); Prelude and Postludes
*** We do not have physical hymnals at Hidden Springs, so all the lyrics need to be printed.
*** Some of the music files have a limited number of verses (fewer than in the hymnal), and some songs are available in multiple instruments or with different numbers of verses, so there are multiple versions. 
*** Most, but not quite all, of the lyrics to the Hidden Springs music can be found in this Google Sheet https://docs.google.com/spreadsheets/d/1vX9llfMg0bAWZiSM10RaAsgExyKIN5YflKaaaEPoOOI/edit?usp=sharing — unfortunately, the names of the hymns there aren’t always perfect matches for the AAC files, so you’ll have to be careful with the matching
*** Feel free to rename and better organize the MIDI and AAC files.
*** Tell me if you think I’m wrong, but I think it would be best to create a new hidden_springs_songs.yaml file in the /data/hymns directory rather than adding these to the songs.yaml. However, when making that new yaml file, if we are missing the lyrics to a song in the Google Sheet and we can find the lyrics to a piece already in our songs.yaml, then let’s copy them over.
* Hidden Springs has special versions of all the prayers of the people that reference their specific ministry context. These will be in the pop.yaml file, though if one is not found the fallback should be to go with the generic form of that prayer.
* At the end of the bulletin on the back cover, we include a schedule of upcoming services (see example bulletins). This should be able to generate automatically from the data in the Rota’s ‘Hidden Springs Planner’ sheet. We include data for the upcoming three Wednesdays. The clergy person is the “Preacher and Officiant” on LOW Sundays, and “Preacher and Celebrant” on Holy Communion Sundays. There are three temple fields on the back cover template to insert the service data: {{Upcoming_Service_1}}, {{Upcoming_Service_2}}, {{Upcoming_Service_3}}
* There is no offering collected at the service except on first Sunday of the month communion Sundays. And in any case, no rubrics are needed about giving or the QR code.

The immediate goal is to get the Annunciation bulletin for March 25 generating correctly.
