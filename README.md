# vbnotescript
A very light note taking and searching app written in VBScript for Windows.
Simple yet powerfull script with just around 10KB in size.

## Concepts
* vbnotescript app stores entire notes in a plain text main file (vbnotescript.txt)
* On startup, it loads whole of the file content in memory.
* On shutdown, it makes a backup of initial main file and dumps the in memory updated content to it
* Notes can be added and edited with Notepad++ or Notepad, default is notepad++
* Text searching returns the matched text and a unique number as a refernce to the note
* Note editing with single save, as on committing save, temporary file is marked for updation in memory and deletion.  

## Prerequisites
* Notepad++ or even Notepad
* Notepad++ with File Status auto detection enabled
   * Settings -> Preferences -> MISC. tab -> File Status Auto Detection -> Enabled

* Note: Default is notepad++, to use notepad, replace it in the vbs file 



## Getting Started
1. Create a folder in your system **vbnotescript**
1. Copy the vbnotescript.vbs file. [vbnotescript.vbs](https://github.com/sourceplant/vbnotescript/blob/master/vbnotescript.vbs)
1. Create a desktop shortcut 
   1. Right click anywhere inside the folder - New - Shortcut
   1. Browse to vbnotescript.vbs script location
   1. Enter the name of shortcut and press okay **vbnotescript**
1. Assign keyboard shortcut key to the previously created Shortcut Icon
   1. Right click on Shortcut file and select properties
   1. Click on shortcut key text box and select your key combination **Ctrl+Alt+F**
   1. Press Okay
1. Now try to execute the vbnotescript
   1. Press your keyboard shortcut combination **Ctrl+Alt+F**
   1. A input text box is displayed
   1. Ready for the use
  
![Image description](https://github.com/sourceplant/vbnotescript/blob/master/INSTALLATION.jpg)

## Working with vbnotescript

* Press your keyboard shortcut combination **Ctrl+Alt+F** to start the app
* To create a new Note, type ":new" in the inputBox.
   * A blank temp file opens up in Notepad++
   * Save it once with Ctrl +s , only when you are done with the file, else leave it open
   *Note* - Saving the file, flags it for updation and deletion the temp file.
   *Note* - There is no connecpt of Notes name, however you can take first line of each note as note name by marking with eg. ##### Note1
* To search for some text, type your search in input box and press ok
   * A notepad file appears with all matched items and a refernce number to the notes having those items
   * To view or edit the note having the matched items, enter the reference number in new inputbox marked with Enter Selection.
   * Same way multiple notes linked to the search can be opened for viewing or editing.
   * Once you are done with that particular search, press cancel button and repeat for next search.
* To edit a Note
   * Repeat the procedure to search for the note texts or notes name.
   * Save the note only once you are done with it.
* Saving the clipboard content as a new note
   * To save the clipboard, press okay while input box is having text "Saving the Clipboard"
