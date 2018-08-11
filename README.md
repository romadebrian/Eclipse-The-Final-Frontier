Download Runtime = https://www.dropbox.com/s/er8k0223tq2x2ku/Runtime.rar?dl=0

|------------------------|
|       ~~KEY~~          |
|                        |
|  [X] Complete          |
|  [!] ToDo              |
|  [?] In Progress       |
|                        |
|------------------------|
									In the source, Search QUICKCHANGE for 
									changes made that weren't necessarily 
									meant to be permanent but can stay.
									Search LEFTOFF to return to where I
									left off programming.



[X]Fixed lstCommands in Editor_Events still visible during editing other items

'Already added

[?]Find and fix any and all bugs/errors that I can

[!]Make a converter for items/npcs/quests/players

[!]Marriage System

[!]Make different chatbox sizes

[!]Resource Stages for different activities/tools/items required/given

[?]Add FullScreen option.



[X]New dinky GUI Bars to fit the Eclipse Theme. Equipped with main gui buttons/minimap toggles.
   [] You can enable/disable the new gui bars from the Server in the controls tab.

[X]Option Pane -- Minimap on/off

[X]Option Pane -- Buttons on/off

[X]Quick fix for NPC Criticals.

[X]Quick fix for NPC Drop Item Num scrollbar.

[X]Projectiles have been added. Causes some slower cps, but it's fully functional.

[!]Buff system. (Attributes/Skills can be temporarily advanced.)

[!]Add reflect combat capability. Damage done has a small chance of being turned over to the dealer instead.
   [] Possibly make it require an item.

[?]New Item Type -- Books


~ADDED~

[!]Fix bug in stackable items

[X]NPC random respawn times

[X]Removed having to target npc to cast AOE spell

[X]Added Drop Items On Death option to map editor properties.

[X]Changed item damage scrl max value to integer (32,767)

[X]Make quests capable of giving skill exp.

[X]New combat system for elements and damage knockoffs
	[X]NPC's
	[X]Items(armour and weapons need to utilize this feature)


[]Make items that aren't stackable, stay in hotbar if player has more left
	[]I think I'm leaving this the way it is.

[X]Fixed blockVar bug in CheckDirection.
	Subscript out of Range because the Map.Tile array range wasn't reset to fit the parameters properly.

[X]Fixed Spell Editor bug when damage textbox was empty

[X]Remove ability to open more than one of an editor

[X]Add Shop Editor and Combo Editor to Editor Run-Through feature...

[X]Made HotBar keep Stackable items in the bar when used.

[X]Combination System, whiten items, click to combine, client-side editor
	~Will Need Updated if Added~
	[-]Custom messages upon completion/no combo available
	[-]Allow more than two items to be selected creating larger combos

[X]Add 'Quest Progress' Conditional Branch option in the Event Editor

[X]Fixed player target when the target logs out.

[X]Fixed editing event conditional statement for player level

[X]Made 'E' toggle through editors with nice and simple GUI (ADMIN_MODERATOR and up.)

[X]Server Panel Option for dropping items on death.

[X]Created a Friends System with working private messages and GUI ('B' opens the Buddy List)
	-[X] If you have a player targeted, the letter 'B' sends a buddy request.
	-[X] Buddy list updates upon arrival of data.
	-[X] Displays Online/Offline status beside name.
	-[X] GUI for Messaging/Editing friend status(Delete)
	-[X] Double Clicking Name opens panel with Friend Details such as lvls and other stats
	-[X] Only friends can PM each other./deactivate the system.
	-[X] Server panel option to activate/deactivate
	-[X] Limit to 5 requests per every 30 minutes
		(Avoid stalking so to speak.)
		(You get 1 request point back every 5 minutes.)

[X] (+)(-)20 New Packets

[X]Fixed PM's.

[X]Fixed GUI occasionally still visible after logout and log back in.

[X]The button that opens the inventory has been switched to the letters 'V' and 'N' for whichever you prefer

[X]Added custom Success/Empty message colors in the resource editor

[X]Skill System is now easily customized Server-Side in modSkills with simple directions and easily noticable notations.

[X]Fixed NPC Check movement error, Subscript out of range. MapNpc(MapNpcNum).num was coming in as a 0

[X]Fixed Index check error, Subscript out of range. Functions weren't making sure index was > 0

[X]Darkened up the background colors of the editors a bit. 
	-[]I like it better this way since it's not so bright. And when you're making games all day, you're looking at these forms quite often.
	-[]Hard to tell a difference unless you know what you're looking for, but trust me, your eyes will thank me.

[X]Added skill levels/exp/config file for crafting, mining, woodcutting, and fishing.
	-[X]Events can give exp to any of these skills.
	-[X]Resources can give exp to any of these skills.
	-[X]Events can require a certain skill level.
	-[X]Resources can require a certain skill level.
	-[X]Items can require a certain skill level.

[X]Don't allow polearms to attack through blocked tiles/resources/events/npcs that aren't attackable

[X]Fixed Quest Dialog visibility leaving when the space key is pressed, but popping up above all other dialogs.
	-[X]Can no longer pick up an item while the Quest Dialog is open.

[X]Fixed little bug with event chat bubble text -Event- option not correctly setting to true when needed.

[X]In case of accidental Rand function misuse, this function now corrects the high and low values to MAKE SURE
   the high is the high and the low is the low. If they're mixed up, it simply switches them and continues.

[X]Pressing enter after entering password on the login screen will now log you in.

[X]GUI hidden when in map editor (health bars and such along with chatbox and minimap)

[X]NPC Avoid tiles now block events too.

[X]Fixed item (amount needed) event condition...

[X]Added HasItems(index, itemnum, itemamount) boolean function.

[X]Added function to allow resource to give more than one of an item...

[X]Added feature to allow spacing resource reward amount throughout the attacking/damaging process.

[X]Added feature for resource health to be random between given static numbers.

[X]Changed frmEditor_Spell.scrlCool.Max AKA spell Casting time capabilities from 60 seconds to 300 which is 5 minutes

[X]Changed frmEditor_Spell.scrlCast.Max AKA spell cooldown time capabilities from 60 seconds to 300 which is 5 minutes

[X]Changed frmEditor_Spell.scrlVital to a texbox to allow numbers higher than 32767.

[X]Changed frmEditor_Spell.lblDir caption to "Dir: Up" because that's the default Index on startup.

[X]Changed frmEditor_Spell.scrlDir.Max to 3 because that's all it should be. Otherwise an error occurs.

[X]Changed NPC death exp given calculation. Random within 5, 10, or 20% of set value, options included within npc editor. (Ex. 150 = 142-158) The higher the original number, the greater the difference.

[X]Added Random NPC Health feature (Took a good bit of work.. Nothing like the resource was.)

[X]Added damage cap, will not show damage higher then the amount of hp the npc has(for players who like to know how much health the kill had)
	-[!] Will probably make this a player option Client-Side

[X]Removed ability to delete or edit currency type item on item 1 (What's the point? It's always currency lol) For those who must know why
	it's also because another feature I added requires a currency be on itemslot 1. NOT a big deal though. You can still edit everything
	else about the item just not the type of item and you can't use the delete button on it. ;)

[X]Added damage boost for combat level of the currently equiped weapon

[X]Added Walkthrough Toggle command "/walkthrough" and button on admin panel so you can turn on and off the ability to "walk through" shit (for ADMIN_MONITOR and up)

[X]Changed heal button on admin panel to heal self when no name is entered

[X]Add kill counter/death counter (Have not made any ways to display them yet)

[X]Fixed Spell Class not having 'None' option.

[X]Fixed Healing Spell with Intervals not updating HP.

[X]Fixed Quest copymemory error. Server/Client QuestRec's were different. Causing an improper placement of variables.

[X]Fixed Spell editor array data being all screwed up...(Wiped all spells and started over, works fine now, something wasn't saved right)

[X]Added 'OPTION' for random currency drop, include random percent...


[X]Brought back ability to attack npc without a weapon

[X]Removed the ability to Walkthrough almost anything just because the map is a safezone, now you'll only walk through players

[X]Added dynamic colors for player's health decreasing during attack(70%hp+ Green, 35%hp+ Yellow, 34.99%hp- BrightRed)
	-[X]Shows Players Health before the hit and the damage done by the npc(Ex. 66 -22)
	-[X]Not sure if it would be better to show the what the player's hp was before the hit; then damage, or show what it is after
		and probably not show the damage.

[X]Add follow feature. Uses the target system. Must be beside the player you want to follow
	-[X]Automatically corrects direction, doesn't just MOCH the other player (that would be silly :p)
	-[X]Stops following if more than 2 spaces away.
	-[!]Unfortunately it is a little laggy right now, the player you follow cannot run or walk full speed, you'll lose 
		the following connection. I'm thinking about redoing it though. I know I didn't do it the most efficient
		way, I just wanted a WORKING way. And here it is. I'll clean it up sometime, or you can. I'm leaning towards
		using the code from CanEventMoveTowardsPlayer and just making it better. Dunno yet.

[X]Added TopKill Event on the Server
	-[X]1st through 5th placements are available
	-[X]Custom exp rewards
	-[X]Dinky/nifty server gui
	-[X]3rd, 4th, and 5th placements are optional
	-[X]Custom start message
	-[X]Custom end message
	-[X]Custom Action Message after every kill (optional)
	-[X]Custom Player Message after every kill (optional)
	-[X]Custom messages have built in data inserts:
		(a) #1stname#     will automatically be replaced with the name of the player who was in first place.
		(b) #2ndname#     ^             ^            ^             ^            ^           second place.
		(c) #3rdname#     ^             ^            ^             ^            ^           third place.
		(d) #4thname#     ^             ^            ^             ^            ^           fourth place.
		(e) #5thname#     ^             ^            ^             ^            ^           fifth place.
		
		(f) #1stkills#     will automatically be replaced with the amount of kills first place had.
		(g) #2ndkills#     ^             ^            ^             ^           ^  second place ^
		(h) #3rddkills#    ^             ^            ^             ^           ^  third place  ^
		(i) #4thkills#     ^             ^            ^             ^           ^  fourth place ^
		(j) #5thkills#     ^             ^            ^             ^           ^  fifth place  ^
		
		(k) #1stexp#      will automatically be replaced with the amount of exp awarded to first place.
		(l) #2ndexp#      ^             ^            ^             ^            ^          second place.
		(l) #3rdexp#      ^             ^            ^             ^            ^          third place.
		(l) #4thexp#      ^             ^            ^             ^            ^          fourth place.
		(l) #5thexp#      ^             ^            ^             ^            ^          fifth place.
		
		(p) #totalkills#  will be replaced with all of the kills added up.
		(q) #getkills#    will be replaced with the amount of kills needed to end the game.
		(r) #placement#   will be replaced with the player's current placement (only for action/player message settings)
		(s) #playerkills# will be replaced with the player's current kills     (only for action/player message settings)
	
	-[X]Custom Msg Colors. All ingame colors available
	-[X]Built with default Messages (Just leave the message box's blank)
	-[!}At some point this might have the capability to give away more than exp, idk yet. I like experiece :p
	-[!]I like the idea of an option to play an animation on the character after each kill, during the attacking process, 
		and/or when the game ends. Thinking about it. Won't be hard I think, just not sure if I want to release it to the public. :P

----------------------------------------------
-----------Resource Random Feature------------
----------------------------------------------
Pro's:
	-Allows for random health. You set the high and the low. This way the amount of 'hits' it will take to chop down is unknown.
	-Like runescape skilling, resources can give more than one of an item.
	-Allows for random item loot. You set the high and the low. You'll get a random amount of the item you have chosen.
	-Finally, my favorite, loot can be distributed while attacking, not just when the resource is destroyed... That's just
		saying that you will get the specified amount of items given to you one by one throughout the health of the resource.
		So with random health, random amount of items given, and the distribution option, resources are now more realistic 
		in effect and create a more enjoyable aspect in the game.

Con's:
	-If the distribution option is activated, the damage of the tool needs to coincide relatively evenly with the health of the
		resource, meaning if your resource's health is 10 and you want to give 5 items throughout the health, you'll want
		your weapon to do NO more than 1 or 2 damage. The reason for this is because of the algorithm. The code
		divides the amount of health by the amount of items. So if the health is 10, it will be divided by 5. As you should
		already know, the result is 2. So for every 2 health points taken away, you will be given an item. If your weapon does
		10 damage in one hit, you'll only be awarded the final item and not the other 4. So with a little bit of math, you'll
		be able to figure out how to set up your tools and resources appropriately. Sorry about this but honestly, if you
		don't like it, change it. It's fully functional as is so I don't 





<p>&nbsp;</p>
<p>&nbsp;</p>
<div>
  <center>
  <p align="center"><img src=http://i1348.photobucket.com/albums/p737/roma_coll04/Kota%20Roma_zpsgvxftepu.png?t=1532940534 width=1000 height=700 /></p>
  <p align="center"><strong> Use Visual Basic 6 </strong></p>
</div>
