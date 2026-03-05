The bulletins for both the 9 am and 11 am services are now mostly being generated correctly! 

There are some small issues outstanding:
# 11 am: 
* The lyrics didn't print with the Kyrie. Even at 11 am, we should print the lyrics to this piece of service music. (Service music generally should have the lyrics printed, even if we have the hymnal reference).
* In the same spot, the title “Lord, have mercy upon us (Willan) was missing the parentheses around Willan that are in the YAML file? Where did they go?
* The opposite problem also happened. The lyrics to ‘Come thou fount of every blessing’ printed? That's song is in the hymnal, and so it should just have the First line / title and the hymnal reference. 
# Both services: 
* The Oremus text uses British spellings — i.e. Saviour instead of Savior. We don’t need a perfect solution, but can we find a list of most common British vs American spellings and run the texts through a script to Americanize them?
* Different Versions of the Prayers of the People
** We need to a way to have different versions of the Prayers of the People. For example, there might be several versions of Forms III, IV, V: One that we use regularly, one that we use while the bishop has asked us to pray for immigrants, and one we use at Hidden Springs. I’m happy to add versions to the YAML file for the prayers, but we’ll need some way to handle that (and to prompt for it in the bulletin creation project).

When these todos are taken care of, we’ll move on to the 8 am bulletin generation. 

# Remember: 
All the formatting styles are consistent across the bulletins, the front and back covers are the same, and the propers (which is the Episcopal word for the readings and prayers that are assigned for that day) are the same. 

Architect the code so that as much of the code base can be reused as possible. We will need to be able to specify when running generate.py which bulletin to generate, or generate all of them.
