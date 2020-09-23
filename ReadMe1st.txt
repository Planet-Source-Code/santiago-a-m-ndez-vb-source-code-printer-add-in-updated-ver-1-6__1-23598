' Author:       Santiago A. Méndez  (Guatemala, C.A.)
' email:        Santiago@InternetDeTelgua.com.gt
'27-Abril-01

The purpose of this VB project Add-In is to bring the user the option to print source code of a module or selected lines of code in the actual window as visual basic does.  The difference is that this project prints in bold: name of functions/sub procedures, end of functions/sub procedures;  in bold + italic: comments; prints a line after and end sub/function, select printer properties for printing.

I've not seen a VB project Add-In that prints source code in PSC before.

I took this idea from:
	A Project’s Source Code Printer 
	Submitted on: 3/27/2001 10:15:12 AM
	By: Morgan Haueisen
	Level: Intermediate


I have seen in PSC more than one VB project that prints code of an entire VB project or selected modules formatting the output.  They work well but one inconvenience is that they only prints the source code of files thas has been saved to disk. You cannot print the module you are working on if you don't save first, or print only a selection of text.

What this vb add-in project provides is the option to print your current module or the selected text of it, in a formatted way to facilitate reading.  Optionaly you can print an index page of the sub/functions.  Also you can select a printer, define some printer properties as: orientation, print quality, tray, etc., using the print setup through API.


'HOW TO INSTALL ADD-IN

1. Open the visual basic project
2. Compile it as SCPAddIn.dll					This register your add-in
3. Run the procedure AddToINI() located in Module1		This adds the add-in to "Add-in manager" in VB
        a. go to inmediate window (Ctrl + G)
        b. type "addtoini" (without quotes) and press enter 
 
        The steps before must be run only once, this makes appear in the "Add-in manager" the item "PrintSourceCodeAddIn.PrintSource"
        which is were you enable it in the next step.  Because this new item is saved when you leave VB, its not necessary to run them 
        again when your enter to VB next time.   

4. Open Add-In Manager and load the PrintSourceCodeAddIn.PrintSource Add-In
	You can tell Add-In manager to load this addin manually or automatically when you enter VB  throug it´s "Load Behavior" frame.

5. In Tools Menu of VB appears the menu item "Print Current Module", click-it, an enjoy.

        If you enter to visual basic and the option "Print Current Module" doesn´t appear, you must open Add-In manager dialog and select
        PrintSourceCodeAddIn.PrintSource in the list of available Add-Ins, in the frame "Load Behavior" you can tell add-in manager to load 
        the selected add-in manually or automatically (when you enter in VB) and press OK.
        This will show the menu option under tools menu.
 
        Now you can open any VB project and print.

Any questions/suggestions tellme.


*------------------------------------------------------ SECOND UPDATE ---------------------------------------------------------------*
'25-Mayo-01

New:
	1. Select print quality
	2. Define page margins

 
IF YOU HAVE ANY QUESTIONS, JUST ASK,
 
    Santiago Méndez
