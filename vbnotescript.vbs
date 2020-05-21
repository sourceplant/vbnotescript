' Revision A1 - 12/10/2019 23:44 UTC - Updated on email
' Simple note taking
' Prompts to either save or search
' By default, it captures the clipboard if "Saving the Clipbord" is in input box.
'
Public WrdArray
Public input
Public size
' Stores KB
Public detailArr
' Timestamp of last written file
Public last
Public oldClipboardText
Public newfile
delimiter = "--++======#======++--"
Set refdict = CreateObject("Scripting.Dictionary")
'Storing Address of Note start and end line
Set dict = CreateObject("Scripting.Dictionary") 
'Soring file name linked to a note
Set filedict = CreateObject("Scripting.Dictionary") 
'Set resultdict = CreateObject("Scripting.Dictionary")  
set objHTML = CreateObject("htmlfile")

set objShell = WScript.CreateObject("Wscript.Shell")

set objFSO = CreateObject("Scripting.FileSystemObject")   
'set objFile = objFSO.OpenTextFile("commands.txt", 1)
'commandArr = Split(objFile.ReadAll, vbLf)
'objFile.Close

' Checking if File Exists
If objFSO.FileExists("vbnotescript.txt") = False Then
    set objFile1 = objFSO.OpenTextFile("vbnotescript.txt", 2, True)
	objFile1.Write delimiter & vbLf & delimiter & vbLf
	objFile1.Close
End If

' Reading the knowledge base file
set objFile1 = objFSO.OpenTextFile("vbnotescript.txt", 1)
detailArr = Split(objFile1.ReadAll, vbLf)
objFile1.Close

'Storing size of input File, will be used later to check if write is needed.
initialsize = UBound(detailArr)

'
commandArr = Filter(detailArr,"#")

 ' regEx is used to avoid problem to UNICODE text in clipboard, it removes the unicode texts
	Set regEx = CreateObject("vbscript.regexp")
	regEx.Global = True
	regEx.MultiLine = True
	regEx.IgnoreCase = True
	regEx.Pattern = "[^\u0000-\u007F]"

' Taking input from InputBox and doing various thing and for searchig calls search function at end
function ReadInput()
	' Read the input from box 
	input = InputBox("Enter Text: ", "(-)", "Saving the Clipboard") 
	' If Cancel button is pressed, InputBox gives "" ; hence making function to return false so that script proceeds and indirectly exit the script
	If input = "" Then
	        ' This is how function return to caller
			readInput = False
	' This section handles if you want to add some content so creates a file for you, which later will be merged to knowledge base array
	ElseIf input = ":new" Then
	        newfile = newfile + 1
			                nkey="0"&newfile
							dict.Add nkey, "new"
							tfile = nkey &"_"& dict(nkey) & ".sh"
							filedict.Add nkey &"_"& dict(nkey), tfile
							'MsgBox tfile
							set objFile2 = objFSO.OpenTextFile(tfile, 8, True)
							objShell.Run "notepad++" & " "& tfile
							objFile2.Close
							last=now
							readInput = True
	'This section handles capturing clipboard content and adding it right away to knowledge base array 
	ElseIf input = "Saving the Clipboard" Then
		'This section handles the one time copy of clipboard
		' If clipboard has nothing or images , ignore it
		ClipboardText = objHTML.ParentWindow.ClipboardData.GetData("text")
		If IsNull (ClipboardText) Then
			readInput = True
			MsgBox "Empty Clipboard"
			Exit Function
		End If
		'if not equal to null, use input and split into array , else reference clipboard data
		ClipboardText = Replace(ClipboardText,vbCrLf,vbLf)
	    objShell.Popup clipboardText, 1, "Saved"
		tempArray = Split(ClipboardText, vbLf)
		Redim preserve detailArr(Ubound(detailArr) + 1)
		detailArr(Ubound(detailArr)) = delimiter
		For i = 0 To UBound(tempArray)
			Redim preserve detailArr(Ubound(detailArr) + 1)
			detailArr(Ubound(detailArr)) = regEx.Replace(tempArray(i), " ") 
		Next 
		Redim preserve detailArr(Ubound(detailArr) + 1)
		detailArr(Ubound(detailArr)) = delimiter
		For i = 0 To UBound(tempArray)
			Redim preserve commandArr(Ubound(commandArr) + 1)
			If (InStr(1,tempArray(i), "#", vbTextCompare) <> 0) Then
				commandArr(Ubound(commandArr)) = tempArray(i)
			End If
		Next
	readInput = True
	else
	' Here we are handling the searching, so first deleting leading and trailing space from input and going ahead with searching
		WrdArray = Split(Trim(input))
		' This update function is called here to update the KB array with new information, when new files are created or opened and edited the existing
		update()
		search()
		readInput = True
	End If
End Function


' This function lists the matched pattern and allows to view whole files marked with custom SOF and EOF
Sub search()
    Set resultdict = CreateObject("Scripting.Dictionary") 
    ' Any matched result
	 found=false
	' Counter holding successfull hits in a line
	count=0
	searched = "------------ Result -------------"
	tempholder = "-----------Result --------------------------------------------------------------------"
	selected = delimiter
	sameFile = "false"
	' This sub-section finds and display matched patterns and while doing so it records the index of SOF and mattched pattern, which is later used 
	' to display the file having the matched pattern quickly.
	' Array size
	size=UBound(WrdArray)
	set objFile2 = objFSO.OpenTextFile("Result.sh", 2, True)
	' Traversing through the array to find the matched patterns and putting the matched result into a file to display, 
	' Putting to file and viewing with notepad  because VBA does not have a good text displayer.
	'It also stores the index of matched pattern along with SOF in refdict
	For i = LBound(detailArr) To UBound(detailArr)
			For Each y In WrdArray
				If (InStr(1,detailArr(i), y, vbTextCompare) <> 0) Then
				count = count+1
				'Msgbox count+size
				End If
			Next
			If (StrComp(detailArr(i), delimiter) = 0) Then
			closest = i
			sameFile = "false"
			End If
			'Checks if all delimted search patterns are having a match in a line
			'If ((count-1 = size) or (count-1 >= size-Cint(size/2)))  Then
			If (count >= Cint(size/2)+1)  Then
			'Msgbox Cint(size+1/count) & size+1 mod count
				' Saving the index of pattern and its SOF
				refdict(i) = closest
				If Not resultdict.Exists(size+1 - count) Then
					resultdict.Add size+1 - count, i & "-->|  "& detailArr(i) & vbCRLf
				else
					resultdict(size+1 - count) = resultdict(size+1 - count) & i & "-->|  "& detailArr(i) & vbCRLf
				End If
			'	objFile2.Write size+1 - count & "-->| " & i & "-->|  "& detailArr(i) & vbCRLf
				count=0
				found=true
				sameFile = "true"
			End If
			' Reset the match counter
			count=0
	Next
	'For Each key In resultdict.Keys
	'MsgBox key
	For key = 0 To size 
	objFile2.Write resultdict(key)
    Next
	objFile2.Close	
	' If no matches found, no need to show further options, exit the subroutine
	' Display the result
	If found = true Then
	 objShell.Run "notepad++ Result.sh"
		' vbCancel aborts the script
		' Let user view details by opening complete file with matched pattern, this basically uses the index of matched pattern from dictionary "refdict"
		' and creates a new dictionary "dict" which will have SOF and EOF for each viewed files.
		do
			input=InputBox("Enter Selection: ", "Find - ")
			if input = "" Then
				'MsgBox "Not an Integer, try again"
				check = False
				Exit Sub
			End If
			If Not IsNumeric(input) Then 
				MsgBox "Not an Integer, try again"
				check = True	
				'Exit Sub
			ElseIf refdict.Exists(CLng(input)) Then
				'If CLng(input) <= UBound(detailArr) Then
					selected=CLng(input)
							set objFile2 = objFSO.OpenTextFile(".result.sh", 2, True)
				'objFile2.Write "++++" 
				For i = refdict(selected)+1 To UBound(detailArr)
				'MsgBox i
					objFile2.Write detailArr(i) & vbLf
					If (StrComp(detailArr(i+1), delimiter) = 0) Then
						tfile = refdict(selected) &"_"&i+1&".sh"
						If Not dict.Exists(refdict(selected)) Then 
							objFile2.Close
							' Storing start and end line no. of the note in dict
							' refdict(selected) is address of start of Note and i+1 is the address of End of Note
							dict.Add refdict(selected), i+1
							' Storing note and its linked filename eg. "startline_endline" = startline_endline.sh
							filedict.Add refdict(selected)&"_"&i+1, tfile
							if Not objFso.FileExists(tfile) Then
							   objFso.MoveFile ".result.sh",tfile
							End If
							objShell.Run "notepad++" & " "&tfile
						else
							objFile2.Close
							'objShell.Run "notepad++" & " "&tfile
							objShell.Run "notepad++" & " "&filedict(refdict(selected)&"_"& dict(refdict(selected)))
					End If
					Exit For
					End If
				Next
			objFile2.Close
                        ' Time when last file updated, before this time all are processed already
                        last=now
		    check = True
			else
				MsgBox "Out of bound, no such options, try again"
				check = True
				'Exit Sub
			End If
		Loop until check = False
	else
	 found=false
	 MsgBox "No Matching result found"
	End If
End Sub

' This function allows to modify the detailArray knowledge base in realtime by marking old lines to Garbage with adding updated to array at end
' In simple term, it looks for the saved file and update in memory.
Sub update()
    If dict.count <> 0 Then
		'MsgBox dict.count
		For Each key In dict.Keys
		'set f = objFSO.GetFile(key&"_"&dict(key)&".sh")
		set f = objFSO.GetFile(filedict(key&"_"&dict(key)))
		'MsgBox key&"_"&dict(key)&".sh is last updated in " & DateDiff("s",last,f.DateLastModified) & " seconds "
                ' Any file updated since last file written by script which is tracked by last variable
				'MsgBox key
           if (DateDiff("s",last,f.DateLastModified) > 0) Then
			       if dict(key) <> "new" Then
			         For i = key To dict(key)
				        detailArr(i) = "+++++GARBAGE+++++"
			         Next
			       End If
			set objFile1 = objFSO.OpenTextFile(filedict(key&"_"&dict(key)), 1)
			'MsgBox key & "---" & dict(key)
			'Storing the new start line of note in memory to accommodate further saves
			newAddrStart=Ubound(detailArr) + 1
			If Not objFile1.AtEndOfStream Then
			updateArr = Split(objFile1.ReadAll, vbLf)
			Redim preserve detailArr(Ubound(detailArr) + 1)
			detailArr(Ubound(detailArr)) = delimiter
			For i = 0 To UBound(updateArr)
			 ' If (StrComp(updateArr(i), "++++") <> 0) Then
					Redim preserve detailArr(Ubound(detailArr) + 1)
					detailArr(Ubound(detailArr)) = updateArr(i)
				'End If
			Next
			Redim preserve detailArr(Ubound(detailArr) + 1)
			detailArr(Ubound(detailArr)) = delimiter
			End If
			'Storing the new end line of note in memory to accommodate further saves
			newAddrEnd=Ubound(detailArr)
			'Making and start to end line mapping in dictionary
			If Not dict.Exists(newAddrStart) Then 
			dict.Add newAddrStart, newAddrEnd
			End If
			'Making new note address mapping to old filename, so that user can still keep editing that
			If Not filedict.Exists(newAddrStart&"_"&newAddrEnd) Then 
			filedict.Add newAddrStart&"_"&newAddrEnd, filedict(key&"_"&dict(key))
			' Deleting old note address to file name mapping
			filedict.Remove(key&"_"&dict(key))
			dict.Remove(key)
			End If
			objFile1.Close
			'objFSO.DeleteFile(key&"_"&dict(key)&".sh")
          End If
		Next
	
	End If
End Sub

' Main
newfile=1
last=now
do
' Calling update to check and update any edited files
	update()
' Get Input from user and display data
	check = readInput()
Loop until check = False

' Calling update to check and update any edited files last time before exiting
update()
' Cleaning up the files which were created but not updated.
For Each key In filedict.Keys
    'MsgBox key
	objFSO.DeleteFile(filedict(key))
	filedict.Remove(key)
Next

' Checking if any change in array since startup.
if initialsize<>UBound(detailArr) Then
' Making backup before opening
Set f = objFSO.GetFile("vbnotescript.txt")
 tp = "vbnotescript.txt" & f.DateLastModified
old=replace(replace(replace(tp,"/","_"),":","_")," ","_")
objFso.CopyFile "vbnotescript.txt", old 

' Writing up the loaded and populated detailArr content to file, ignoring garbage marked lines which were basically edited files.
set objFile = objFSO.OpenTextFile("vbnotescript.txt", 2, true)
For i = 0 To UBound(detailArr)
	If detailArr(i) <> "+++++GARBAGE+++++" Then
		objFile.Write detailArr(i)
		if i <> UBound(detailArr) Then
		objFile.Write vbLf
	End If 
	End If
Next
objFile.close
End If


' Object garbase collection

set objFSO = Nothing
set regEx = Nothing
set objHTML = Nothing
