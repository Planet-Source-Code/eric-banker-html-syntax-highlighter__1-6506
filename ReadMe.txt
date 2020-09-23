Welcome to the EzColorCode Html Syntax Highlight Demonstration.
------------------------------------------------------------------------------------------

Last Updated: 3-10-00

Latest Updates:
	1. Added Undo/Redo to the program. A little sloppy still. Has some problems 
		with undoing inserted tags (hit undo 3 times)
	2. Fixed up some color coding problems with Pasting
	3. Fixed problem with large files. now it just takes a while to load
	4. added activegui to emulate office toolbar
	5. made undo/redo enable and disable when it's correct
	6. added line and column numbers
	7. Fixed undo/redo when inserting tags

Known Issues:
	1. Problems when the user deletes a ">" character. it can't remember coloring
	2. Some problems with pasting text
	3. It looses color coding when "Redoing" an undone action
	5. Coloring a file initially can be slow on large files

##########################################################################################

Below you will find the information you need to run this application.

******************************************************************************************
##########################################################################################

This applications uses 2 activeX controls. These controls are included in the distribution
and both of these controls are freeware but are not covered under the GPL License included
with this program.
 
You can always find the latest versions of the controls at:

AS-IFCE1.ocx is located at: http://www.users.globalnet.co.uk/~ariad/
I am unable to find a web address for ActiveGUI. It seems that it is no longer supported.

##########################################################################################
******************************************************************************************

System Requirements to run and compile the control:

Pentium 100 w/ 16 meg of ram (estimate)
Windows 95/98/Nt/2000
at least Visual Basic 5.0 (included in Visual Studio 97) to open and compile the project

You will need to register the AS-IFCE1.ocx and ActiveGUI.ocx controls. 
Do this by copying them to the

C:\windows\system
C:\winnt\system32

directory and then open up a command prompt. Once in the command prompt type:

If windows 98 or 95 type this:
	cd c:\windows\system <enter>
	regsvr32 AS-IFCE1.ocx <enter>
	regsvr32 ActiveGUI.ocx <enter>

If Windows NT or 2000 type this:
	cd c:\winnt\system32 <enter>
	regsvr32 AS-IFCE1.ocx <enter>
	regsvr32 ActiveGUI.ocx <enter>

If everything goes well it will come back and say everything worked. You should now 
be able to open the project and play with the program


##########################################################################################

About the Program

EzColorCode was written because of a lack of good color coding Ability for visual
basic. The program uses several properties that you should be aware of

-----------------------------------------------------------------------------------------
EzColorCode Methods:

.HtmlHighlight 
	This method color codes the current textbox with the colors selected above

.KeyPressEvent(KeyAscii) 
	This method color codes as the user is typing

.InsertTag (Tag) 
	This method is for color coding tags that are to be inserted through buttons and menus

.InsertAspTag (Tag)
	This method does the same thing as above but with ASP specific tags

.IsOutsideTag ()
	This method figures out if the cursor is in a tag or not and sets the color correctly.

.PlaceCursor (Tag, CursorPostition)
	This basically puts the cursor in the middle of the tag.
-----------------------------------------------------------------------------------------

There are other methods and properties but those are the most important. All of the Methods are 
demonstrated in the demonstration program. Setting colors is not demonstrated but is easy to do. 
To do this for any of the colors do the following:

In the form_load() sub after you define the RTB for the control change the 
colors for each tag like so:

m_ASPCol = 128

Which is a nice dark red color taken straight from the color common dialog. This should make it 
Easy for you to add selectable colors for users of your program.

-----------------------------------------------------------------------------------------

This code is GPL'd so please read the documentation file that comes with this control and app.

That's about it. Any questions or comments please contact me, Eric Banker, at:
ebanker@gmu.edu

Flames will be ignored ... sometimes :)

enjoy

This code is copyright me Eric Banker in the year 2000
Eric Banker