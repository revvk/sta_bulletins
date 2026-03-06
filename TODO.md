# Remember: 
All the formatting styles are consistent across the bulletins, the front and back covers are the same, and the propers (which is the Episcopal word for the readings and prayers that are assigned for that day) are the same. 

Architect the code so that as much of the code base can be reused as possible. We will need to be able to specify when running generate.py which bulletin to generate, or generate all of them.

# 8 am Bulletin
The next type of bulletin to add is the last of the Sunday morning bulletins. 
* The 8 am service has no music. These means that all the music slots from the other services can be removed. It also means that where there is service music (in Lent that is the Kyrie/Psalm 51, the Sanctus, and the Fraction Anthem) it needs to be rendered as spoken text for the people to recite. 
** The Kyrie looks like this: 
		Celebrant	Lord, have mercy.
		**People	Christ, have mercy.**
		Celebrant	Lord, have mercy.
** At the Sanctus, the whole thing is bold.
** At the Fraction it looks like this:
		Lamb of God, you take away the sins of the world: **have mercy on us.**
		Lamb of God, you take away the sins of the world: **have mercy on us.**
		Lamb of God, you take away the sins of the world: **grant us peace.**
* The beginning rubric at the 8 am is shorter than the rubric at the other services. It should read: “Prayers before worship can be found in the Book of Common Prayer, p. 833-35.”
* The Offertory rubric at the 8 am is the shorter one: Please place your offerings and Connection Cards in the offering plates as they are passed. You can also easily give online at give.standrewsmckinney.org. (This should be the one used at 11 am too, but I’m not sure that’s happening.)