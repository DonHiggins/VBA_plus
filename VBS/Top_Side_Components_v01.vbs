Option Explicit
'- - - - - - - - - - - - - - -
' Top_Side_Components (from Wav_File_Splitter_V3.vbs, and Pin_Assignments.vbs)
'   -- User Selects a file to open -- a ViewDraw schematic file
'      We search out bottom-side components: R & C, 1206/805/603, horiz/vert,
'      And we substitute the  corresponding top-side component.
'      As of 2015, we no longer do wave-solder on bottom side of boards at AMC, consequently we no longer 
'      need bottom-side components with extra surrounding space to accommodate wave-solder. 
'- - - - - - - - - - - - - - -
'
Dim schematic_file_path        'For ex: "Z:\Source\VBS\Top_Side_Components"
Dim input_schematic_file_name  'For ex: "xDEMO.1"
Dim fso                  'the file system object
Dim In_f
Dim component_name_array(500) 
Dim component_name_count
Dim g_change_refdes_bool
Dim g_output_in_same_folder_bool

'Initialize global parameters
'schematic_file_path = "Z:\Source\VBS\Top_Side_Components" 'default input file folder
schematic_file_path = "G:\ECAD\Projects\Aut_TS3\xDEMO\sch"
input_schematic_file_name = "xDEMO.1" 'default ViewDraw schematic file Name
component_name_count = 0 'start with empty array

call Top_Side_Cmpnt_Start

WScript.Quit 'Normal end to script
'- - - - - - - - - - - - - - -
'
Function fix_prompt_string(prompt_string) 
'Insert a " " after each "\" so that long path names are wrapped,
' instead of being truncated
  Dim Str_1
  Dim Str_2
  Dim I

  Str_1 = prompt_string
  I = 1
  Do While (I > 0)
    I = InStr(I,str_1,"\")
    if (I > 0) Then
      str_2 = Mid(str_1,1,I) & " " & Mid(str_1,I+1)
      str_1 = str_2
      I = I + 1
    end if
  Loop

  fix_prompt_string = str_1
End Function
'- - - - - - - - - - - - - - -
'
Function Get_Signal_and_Pin_Number(Selected_Component)
'
'Ask user to enter signal name and pin # 
'Quit script if user clicks CANCEL
'
  Dim Prompt_String
  Dim ok_or_cancel
  Dim I
  Dim Signal, Pin_Num

  '- - - - - Get Signal - - - - -
  prompt_string = "Enter the name of the signal " & Chr(13) & Chr(10)_
                & "(net name) that you want to connect to a component pin."  & Chr(13) & Chr(10)_
                & " " & Chr(13) & Chr(10)_
                & " " 

  prompt_string = fix_prompt_string(prompt_string) 'to work better w/ InputBox
  Signal = "enter signal here"
  Signal = InputBox(prompt_string, "Top_Side_Cmpnt",_
                 Signal)
  'If operator clicks CANCEL, InputBox returns zero length string ("")
  if (Signal = "") Then WScript.Quit 'operator CANCELED script

 '- - - - - Get Pin # - - - - -
  prompt_string = "Enter the pin # " & Chr(13) & Chr(10)_
                & "that you want to attach the signal to."  & Chr(13) & Chr(10)_
                & " " & Chr(13) & Chr(10)_
                & " " 

  prompt_string = fix_prompt_string(prompt_string) 'to work better w/ InputBox
  Pin_Num = "enter pin # here"
  Pin_Num = InputBox(prompt_string, "Top_Side_Cmpnt",_
                 Pin_Num)
  'If operator clicks CANCEL, InputBox returns zero length string ("")
  if (Pin_Num = "") Then WScript.Quit 'operator CANCELED script

'- - - - - Get Pin # - - - - -
  prompt_string = "Next step is to create the edited file. " & Chr(13) & Chr(10)_
                &  Chr(13) & Chr(10)_
                & "Verify that you want us to add attribute " & Chr(13) & Chr(10)_
				& Chr(13) & Chr(10)_
				& "   " & "SIGNAL=" & UCase(Signal) & ";" & Pin_Num & Chr(13) & Chr(10)_
				& Chr(13) & Chr(10)_
				& "to all " & Selected_Component & " instances." & Chr(13) & Chr(10)_
				& Chr(13) & Chr(10)_
                & "Ok or Cancel ?" 

  prompt_string = fix_prompt_string(prompt_string) 'to work better w/ InputBox
  ok_or_cancel = "Ok"
  ok_or_cancel = InputBox(prompt_string, "Top_Side_Cmpnt",_
                 ok_or_cancel)
  'If operator clicks CANCEL, InputBox returns zero length string ("")
  if (ok_or_cancel = "") Then WScript.Quit 'operator CANCELED script
  if (Not(LCase(ok_or_cancel) = "ok")) Then WScript.Quit 'operator CANCELED script
  
  Get_Signal_and_Pin_Number = UCase(Signal) & ";" & Pin_Num
End Function
'- - - - - - - - - - - - - - -
'
Function Get_Change_Refdes_YesNo()
'
'Ask user to choose (Yes / NO) whether or not to change REFDES entries
'Quit script if user clicks CANCEL
'Returns true if user wants to reset REFDES entries

  Dim Prompt_String
  Dim ok_or_cancel
  Dim I
  Dim Y_or_N, Pin_Num

  '- - - - - Get Signal - - - - -
  prompt_string = "Do you want to reset the RefDes " & Chr(13) & Chr(10)_
                & "for all bottom side R's & C's ?"  & Chr(13) & Chr(10)_
                & " " & Chr(13) & Chr(10)_
                & "Change them all to R?, RS?, RE?,"  & Chr(13) & Chr(10)_
                & "C?, CS?, or CE?"  & Chr(13) & Chr(10)_
                & " " & Chr(13) & Chr(10)_
                & "Yes or No ?"  & Chr(13) & Chr(10)_
                & " " 

  prompt_string = fix_prompt_string(prompt_string) 'to work better w/ InputBox
  Y_or_N = "No"
  Y_or_N = InputBox(prompt_string, "Top_Side_Cmpnt",_
                 Y_or_N)
  'If operator clicks CANCEL, InputBox returns zero length string ("")
  if (Y_or_N = "") Then WScript.Quit 'operator CANCELED script

  if (LCase(mid(Y_or_N,1,1)) = "y") Then 
		Get_Change_Refdes_YesNo = true
	Else
		Get_Change_Refdes_YesNo = false
	End If	

End Function
'- - - - - - - - - - - - - - -
'
Function Get_Output_In_Same_Folder_YesNo()
'
'Ask user to choose (Yes / NO) whether or not to 
'write output files to same folder as input files (after appending a suffix to the file name)
'The NO option implies we will create a separate folder for the output files.
'Quit script if user clicks CANCEL
'Returns true if user wants to output files to same folder

  Dim Prompt_String
  Dim ok_or_cancel
  Dim I
  Dim Y_or_N, Pin_Num

  '- - - - - Get Signal - - - - -
  prompt_string = "Do you want the output files written " & Chr(13) & Chr(10)_
                & "to the same folder as the input files? "  & Chr(13) & Chr(10)_
                & "(If so, we add a ""_TS"" suffix"  & Chr(13) & Chr(10)_
                & "to each output file name.)   : ""Yes"" "  & Chr(13) & Chr(10)_
				& " " & Chr(13) & Chr(10)_
                & "Or ""No"" to have output files written "  & Chr(13) & Chr(10)_
                & "to a separate folder."  & Chr(13) & Chr(10)_
                & " " & Chr(13) & Chr(10)_
                & "Yes or No ?"  & Chr(13) & Chr(10)_
                & " " 

  prompt_string = fix_prompt_string(prompt_string) 'to work better w/ InputBox
  Y_or_N = "No"
  Y_or_N = InputBox(prompt_string, "Top_Side_Cmpnt",_
                 Y_or_N)
  'If operator clicks CANCEL, InputBox returns zero length string ("")
  if (Y_or_N = "") Then WScript.Quit 'operator CANCELED script

  if (LCase(mid(Y_or_N,1,1)) = "y") Then 
		Get_Output_In_Same_Folder_YesNo = true
	Else
		Get_Output_In_Same_Folder_YesNo = false
	End If	

End Function

'- - - - - - - - - - - - - - -
'
Sub Get_Input_File_Name()
'
'Ask for top level project folder and schematic file
'Quit script if user clicks CANCEL
'

  Dim Prompt_String
  Dim ok_or_cancel
  Dim I

  '- - - - - Get Project Path - - - - -

  prompt_string = "Schematic File Path  : " & Chr(13) & Chr(10)_
                & "Input Schematic File : " & Chr(13) & Chr(10)_
                & " " & Chr(13) & Chr(10)_
                & "Enter full path to folder containing the Schematic file: " 

  prompt_string = fix_prompt_string(prompt_string) 'to work better w/ InputBox
  schematic_file_path = InputBox(prompt_string, "Top_Side_Cmpnt",_
                 schematic_file_path)
  'If operator clicks CANCEL, InputBox returns zero length string ("")
  if (schematic_file_path = "") Then WScript.Quit 'operator CANCELED script

  '- - - - - Get Schematic Name - - - - -

  prompt_string = "Schematic File Path  : " & Chr(13) & Chr(10) _
                & Chr(13) & Chr(10)_               
			    & schematic_file_path & Chr(13) & Chr(10)_
				& Chr(13) & Chr(10)_
                & "Input Schematic File : " & Chr(13) & Chr(10)_
                & " " & Chr(13) & Chr(10)_
                & "Enter Schematic file name: " 

  prompt_string = fix_prompt_string(prompt_string) 'to work better w/ InputBox
  input_schematic_file_name = InputBox(prompt_string, "Top_Side_Cmpnt",_
                    input_schematic_file_name)
  'If operator clicks CANCEL, InputBox returns zero length string ("")
  if (input_schematic_file_name = "") Then WScript.Quit 'operator CANCELED script

 End Sub
'- - - - - - - - - - - - - - -
Function Select_Component()
  Dim Prompt_String
  Dim Response
  Dim I

  prompt_string = "I just brought up a list of components in that schematic page." & Chr(13) & Chr(10)_
                & Chr(13) & Chr(10)_
                & "Pick a component to modify" & Chr(13) & Chr(10)_
                & Chr(13) & Chr(10)_
                & "Then click OK or CANCEL :" 

  prompt_string = fix_prompt_string(prompt_string) 'to work better w/ InputBox
  Response = InputBox(prompt_string, "Top_Side_Cmpnt",_
                    "paste component name here")
  'If operator clicks CANCEL, InputBox returns zero length string ("")
  if (Response = "") Then WScript.Quit 'operator CANCELED script
  Select_Component = Response
End Function 
'- - - - - - - - - - - - - - -
'
Function get_space_separated_token(Line,nth_token)
'
'Return the nth space-separated token from the Line of text
'
  Dim I, L
  Dim token_number

  L = Len(Line)
  token_number = 1
  Do While (L > 0)
    get_space_separated_token = "" 'default return value is 0-length string
    I = InStr(1,Line," ")

    if (I = 1) Then
      if (L = 1) Then Exit Function 'didn't find nth token
      Line = Mid(Line,2) 'get rid of leading blank
      L = L-1
    
    elseif (token_number = nth_token) then 'this is the one we're looking for
      if (I = 0) Then
        get_space_separated_token = Line
      else 
        get_space_separated_token = Mid(Line,1,I-1)
      end if
      Exit Function 'OK, found nth token

    else 'skip over this token
      if (I = 0) then Exit Function 'didn't find nth token
      Line = Mid(Line,I) 'this leaves us with a space as first character
                         '  but it's handled OK, and we avoid a 0-length string
      L = L-I
      token_number = token_number + 1
    end if
  Loop
  'if we exit the loop, it means we didn't find the nth token
  '  and we return a 0-length string
End Function 
'- - - - - - - - - - - - - - -
'
Function component_already_in_array(token)
' See if <token> is already contained in array
' component_name_array() up to element: component_name_count
' Return True / False
  Dim I
  
  component_already_in_array = False ' Default False until True
  If (component_name_count = 0) Then Exit Function
 
  For I = 1 To component_name_count
  If (component_name_array(I) = LCase(token)) Then
      component_already_in_array = True
      Exit Function
    End If
  Next
  Exit Function 'Got to end of array w/out finding a match
  
End Function
'- - - - - - - - - - - - - - -
sub Read_Schematic_File_Collect_Components(fname)
'
'Called w/ fname is the fully qualified path and
'  file name (with extension) of an existing schematic page file. 
'Read the file, a line at a time
'Look for "I <n> <NAME> <n> ..." records
'where <name> contains a ":" 
'then add it to our list of components.
'
  Dim f
  Dim Line
  Dim Input_String
  Dim Line_Num
  Dim prompt_string
  Dim token
  Dim I
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
  Const Show_Debug_Messages = False ' ***** FOR DEBUGGING *****

  Set f = fso.OpenTextFile(fname, ForReading)
  
  Do While (f.AtEndOfStream <> True)  
    'Read file one line at a time
    Line = f.ReadLine
    If (Show_Debug_Messages) Then
      prompt_string = Line & Chr(13) & Chr(10)_
         & "+" & Mid(Line,1,1)_
         & ":" & Mid(Line,2,1)_
         & ":" & Mid(Line,3,1)_
         & ":" & Mid(Line,4,1)_
         & ":" & Mid(Line,5,1)_
         & "+" & Mid(Line,6,1)_
         & ":" & Mid(Line,7,1)_
         & ":" & Mid(Line,8,1)_
         & ":" & Mid(Line,9,1)_
         & ":" & Mid(Line,10,1)_
         & "+"
      prompt_string = fix_prompt_string(prompt_string) 'to work better w/ InputBox
      Input_String = InputBox(prompt_string, "Top_Side_Cmpnt",_
                    "continue ...")
      'If operator clicks CANCEL, InputBox returns zero length string ("")
      if (Input_String = "") Then WScript.Quit 'Operator CANCELED script
    End If ' end "If (Show_Debug_Messages)"
    
    '- - - - - -
    'Look for a line starting with "I", containing a component type 
    '  ( embedded ":") as the 3rd token
    Do 'actually, we aren't looping, really, just a creative
       '  use of the "Exit Do" construct to jump to the end of a block
      If (Mid(Line,1,1) <> "I") Then Exit Do 
      token = get_space_separated_token(Line,3) 'third token from Line
      If (token = "") Then Exit Do
      I = InStr(1,token,":")
      If (Not(I > 0)) Then Exit Do 'do want it if it has a ":" in it
      If (component_already_in_array(token)) Then Exit Do ' already found this component
      If (UBound(component_name_array) = component_name_count) Then
        'sch_name_array is filled up
        call MsgBox ("Top_Side_Components.vbs: " & Chr(13) & Chr(10)_
             & "Read_Schematic_File_Collect_Components()" & Chr(13) & Chr(10)_
             & "***FATAL ERROR***:"& Chr(13) & Chr(10)_
             & "     overflowed hard-coded size of component_name_array: "_
             & UBound(component_name_array))
        WScript.Quit 'Fatal Error
      end if      
      component_name_count = component_name_count + 1
      component_name_array(component_name_count) = LCase(token) 
      If (Show_Debug_Messages) Then call MsgBox ("found token: " _
            & token)
    Loop Until (True)
    '- - - - - -

  Loop

End Sub 
'- - - - - - - - - - - - - - -
'
Function create_folder_unless_it_already_exists(folder_spec)
'
' Creates a new folder if it doesn't already exist
' folder_spec is a completely specified path to a folder
' Return True if new folder was created
'
  Dim new_fldr            'a folder object

  create_folder_unless_it_already_exists = False
  If (fso.FolderExists(folder_spec)) Then
    'MsgBox("Folder, " & folder_spec & ", already exists.")
    Exit Function
  End If 
  Set new_fldr = fso.CreateFolder(folder_spec)
  'call MsgBox ("Created folder: " & new_fldr.Name)
  'call MsgBox ("Created folder : " & Chr(13) & Chr(10) & folder_spec)
  create_folder_unless_it_already_exists = True
End Function 
'- - - - - - - - - - - - - - -
'
Function LittleEndian4Bytes(Num_In)
  LittleEndian4Bytes = Chr((Num_In And &H0FF&) / &H01&) _
                     & Chr((Num_In And &H0FF00&) / &H0100&) _
                     & Chr((Num_In And &H0FF0000&) / &H010000&) _
                     & Chr((Num_In And &H0FF000000&) / &H01000000&) 
End Function 
'- - - - - - - - - - - - - - -
'
Function LittleEndian2Bytes(Num_In)
  LittleEndian2Bytes = Chr((Num_In And &H0FF&) / &H01&) _
                     & Chr((Num_In And &H0FF00&) / &H0100&) 
End Function 
'- - - - - - - - - - - - - - -
'
Sub Write_Components_To_Text_File(file_name)
' We have a list of components in an array.
' Write the list into a text file so we can look at them
' file_name parameter is something like "Components.txt"
' We write file into directory from which VBS script is running.
  Dim I
  Dim objFile, objFSO
  Dim strScriptPath, objScriptFile, objShell, strScriptFolder
  
  Set objShell = CreateObject("Wscript.Shell")
  Set objFSO=CreateObject("Scripting.FileSystemObject")
  strScriptPath = Wscript.ScriptFullName
  Set objScriptFile = objFSO.GetFile(strScriptPath)
  strScriptFolder = objFSO.GetParentFolderName(objScriptFile) 
  
  Set objFile = objFSO.CreateTextFile(strScriptFolder & "\" & file_name,True)
  'MsgBox(strScriptFolder & "COMP.txt")
  For I = 1 to component_name_count
    objFile.Write component_name_array(I) & vbCrLf
  Next
  objFile.Close
  
  'Run Notepad to open and display the new text file
  objShell.Run "Notepad.exe " & strScriptFolder & "\" & file_name

End Sub
'- - - - - - - - - - - - - - -
'
Sub  Read_File_Write_Edited_File(schematic_file_path, input_schematic_file_name, Output_Subfolder)
  Dim f_read, Line, fso, f_write, Add_Line, Temp_Line, Saved_I_Line
  Dim Val, token, I, token_4
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
  Const R1206 = 1, R805 = 2, R603 = 3, C1206 = 4, C805 = 5, C603 = 6
  Dim Component_Type
  Dim line_has_been_written_to_output
  Dim Changed_Bottom_Side_Component
  Dim output_schematic_file_name

  Set fso = CreateObject("Scripting.FileSystemObject")
  Set f_read = fso.OpenTextFile(schematic_file_path & "\" & input_schematic_file_name, ForReading)
 
  If (g_output_in_same_folder_bool = true) Then
	'append suffix "_TS" to the file name
	output_schematic_file_name = append_suffix_TS(input_schematic_file_name)
  Else
	output_schematic_file_name = input_schematic_file_name
  End If
 
  Set f_write = fso.CreateTextFile(schematic_file_path & Output_Subfolder & "\" & output_schematic_file_name,True)

  Changed_Bottom_Side_Component = false ' until we decide otherwise
 
  Do While (f_read.AtEndOfStream <> True)  
    'Read file one line at a time
    Line = f_read.ReadLine
	line_has_been_written_to_output = false ' until we do write it

    'Look for a line starting with "I", containing a component type 
    '  ( embedded ":") in the 3rd token
    Do 'actually, we aren't looping, really, just a creative
       '  use of the "Exit Do" construct to jump to the end of a block
        If (Mid(Line,1,1) <> "I") Then Exit Do
	    Temp_Line = Line
		Saved_I_Line = Line 'keep a 2nd copy
		Changed_Bottom_Side_Component = false ' until we decide otherwise
												' And only set/reset this after reading an "I" component line
        I = InStr(1,Line,"SMD_Descretes")
	    Component_Type = 0 'until we find actual component type
	    If (I <= 0) Then Exit Do
		
	    ' - - - - - - - we have "I", "SMD_Descretes", and now we ditch bottom side "B" - - - - - -
		I = InStr(1,Line,"R1206B")
		If (I > 0) Then
			Component_Type = R1206
			Line = mid(Temp_Line,1,I+4) & mid(Temp_Line,I+6) ' ditch the "B"
		End If
		I = InStr(1,Line,"R805B")
		If (I > 0) Then
			Component_Type = R805
			Line = mid(Temp_Line,1,I+3) & mid(Temp_Line,I+5) ' ditch the "B"
		End If
		I = InStr(1,Line,"R603B")
		If (I > 0) Then
			Component_Type = R603
			Line = mid(Temp_Line,1,I+3) & mid(Temp_Line,I+5) ' ditch the "B"
		End If
		I = InStr(1,Line,"C1206B")
		If (I > 0) Then
			Component_Type = C1206
			Line = mid(Temp_Line,1,I+4) & mid(Temp_Line,I+6) ' ditch the "B"
		End If
		I = InStr(1,Line,"C805B")
		If (I > 0) Then
			Component_Type = C805
			Line = mid(Temp_Line,1,I+3) & mid(Temp_Line,I+5) ' ditch the "B"
		End If
		I = InStr(1,Line,"C603B")
		If (I > 0) Then
			Component_Type = C603
			Line = mid(Temp_Line,1,I+3) & mid(Temp_Line,I+5) ' ditch the "B"
		End If

		If (Component_Type > 0) Then
			'write out "I" component-type line we just changed
			f_write.Write Line & vbCrLf
			line_has_been_written_to_output = true
			Changed_Bottom_Side_Component = true 
			'Read in the next line
			Line = f_read.ReadLine
			line_has_been_written_to_output = false
			'Here we may discard or change a date-stamp record
			'Is this a date-stamp record?
			If Not(InStr(1,Line,"|R") = 1) Then Exit Do
			If (Component_Type = R805) Then
				'Here we just remove the line
				line_has_been_written_to_output = true
			End If
			If (Component_Type = R603) Then
					'Here we substitute a different date stamp record
				Line = "|R 19:55_5-15-03"
				f_write.Write Line & vbCrLf
				line_has_been_written_to_output = true
			End If
			If (Component_Type = C603) Then
				'Here we substitute a different date stamp record
				'But there are different date stamps for C603.1 and C603.2 (horiz and vert)
				token_4 = get_space_separated_token(Temp_Line,4) '4th token from Line differentiates horiz from vert
				'	Call MsgBox ("T_S_C.vbs: " & Chr(13) & Chr(10)_
				'		& "token_4: " & token_4 & Chr(13) & Chr(10)_
				'		& "Saved_I_Line : " & Saved_I_Line & Chr(13) & Chr(10))
				if (token_4 = 2) Then
					Line = "|R 19:53_5-15-03" 'C603.2
				else
					Line = "|R 19:52_5-15-03" 'C603.1
				end if
				f_write.Write Line & vbCrLf
				line_has_been_written_to_output = true
			End If
			
		End If
		
		Exit Do
	Loop Until (True)
    '- - - - - -
	
	
	If Not (line_has_been_written_to_output) Then
	
		If (Changed_Bottom_Side_Component = true) And (g_change_refdes_bool = true) Then
			'Here we are looking at a line following the "I" line specifying
			'one of our formerly bottom-side components.
			'If it contains "REFDES=", and our global flag is set indicating that the
			'user desires us to reset the refdes, then that's what we do
				'Call MsgBox ("T_S_C.vbs: " & Chr(13) & Chr(10)_
				'	& Chr(13) & Chr(10)_
				'	& "(Changed_Bottom_Side_Component = true)" & Chr(13) & Chr(10)_
				'	& "And (g_change_refdes_bool = true)"& Chr(13) & Chr(10)_
				'	& Chr(13) & Chr(10)_
				'	& "Line : "& Line & Chr(13) & Chr(10))
				'	WScript.Quit
			I = InStr(1,Line,"REFDES=")
			If (I > 1) Then
				Line = mid(Line,1,(I+6))
				'Call MsgBox ("T_S_C.vbs: " & Chr(13) & Chr(10)_
				'	& Chr(13) & Chr(10)_
				'	& "I = InStr(1,Line,""REFDES="") and I > 1"& Chr(13) & Chr(10)_
				'	& Chr(13) & Chr(10)_
				'	& "Line : "& Line & Chr(13) & Chr(10))
				'	WScript.Quit
				If (Component_Type = R1206) Then
					Line = Line & "R?"
				ElseIf (Component_Type = R805) Then
					Line = Line & "RS?"
				ElseIf (Component_Type = R603) Then
					Line = Line & "RE?"
				ElseIf (Component_Type = C1206) Then
					Line = Line & "C?"
				ElseIf (Component_Type = C805) Then
					Line = Line & "CS?"
				ElseIf (Component_Type = C603) Then
					Line = Line & "CE?"
				End If
			End If
		End If
	
		'write out this line
		f_write.Write Line & vbCrLf
		line_has_been_written_to_output = true
    End If

  Loop 'bottom of the "  Do While (f_read.AtEndOfStream <> True) "
   f_read.Close
   f_write.Close

End Sub
'- - - - - - - - - - - - - - -
'
Function Create_Unique_Sub_Folder(schematic_file_path, subfolder_root)
'Create a new subfolder named subfolder_root in schematic_file_path
'If that subfolder already exists, then append "_nn" to new subfolder name and try again.
  Dim path_and_folder
  Dim folder_name
  Dim fso
  Dim done
  Dim I
  Dim new_folder
  
  Set fso = CreateObject("Scripting.FileSystemObject")
  done = false
  I = 0
  folder_name = subfolder_root
  Do
    path_and_folder = schematic_file_path & folder_name
    If (fso.FolderExists(path_and_folder)) Then
	   'call MsgBox ("Top_Side_Components.vbs: " & Chr(13) & Chr(10)_
		'& "Folder Exists" & Chr(13) & Chr(10))
		I = I + 1
		folder_name = subfolder_root & "_" & Right("00" & I,2)
	Else
	   done = true
	End If
	Loop until (done)
	Set new_folder = fso.CreateFolder(path_and_folder)

	Create_Unique_Sub_Folder = folder_name
End Function 

'- - - - - - - - - - - - - - -
'
Function append_suffix_TS(file_name)
'Append the suffix "_TS" to the file name, so that the resultimg file can
'coexist with the original file in the same folder.
Dim DotPos, Suffix 

	DotPos = InstrRev(file_name, ".", -1, 1)
	If ((DotPos < 2) Or (DotPos = Len(file_name))) Then
		Call MsgBox ("T_S_C.vbs: " & Chr(13) & Chr(10)_
			& Chr(13) & Chr(10)_
			& "Problem adding ""_TS"" suffix"& Chr(13) & Chr(10)_
			& "to file name : " & file_name & Chr(13) & Chr(10)_
			& Chr(13) & Chr(10))
		WScript.Quit
	End If
	Suffix = Mid(file_name,(DotPos+1))
	append_suffix_TS = Mid(file_name,1,(DotPos-1)) _
		& "_TS." _
		& Suffix
End Function 
'- - - - - - - - - - - - - - -
'
Function Increment_Schematic_Page(schematic_file_name)
'Schematic file name ends in a dot "." followed by a numeric suffix,
'in place of what is normally a DOS file type.
'We return the file name after incrementing the numeric suffix,
'equivalent to advancing to the next page of the schematic.
Dim DotPos, Suffix 

	DotPos = InstrRev(schematic_file_name, ".", -1, 1)	
	Suffix = Mid(schematic_file_name,(DotPos+1))
	
	Increment_Schematic_Page = Mid(schematic_file_name,1,DotPos) _
		& (Suffix + 1)
End Function 
'- - - - - - - - - - - - - - -
'
Sub Top_Side_Cmpnt_Start()
'
'Locate the input file
'Read File and interpret its contents as a viewdraw schematic file
'Do it
'
  Dim prompt_string
  Dim I
  Dim debug_input
  Dim Opt_1_Split_2_Crop
  Dim Selected_Component
  Dim Signal_and_Pin_Number
  Dim Output_Subfolder
  Dim no_more_pages
  
  call Get_Input_File_Name ' ask user for info, leave it in globals
  Set fso = CreateObject("Scripting.FileSystemObject")
   
   If (Not fso.FileExists(schematic_file_path & "\" & input_schematic_file_name)) Then
        prompt_string = schematic_file_path & "\" & input_schematic_file_name
		prompt_string = fix_prompt_string(prompt_string)
        call MsgBox ("Top_Side_Components.vbs: " & Chr(13) & Chr(10)_
         & "Top_Side_Cmpnt_Start()" & Chr(13) & Chr(10)_
         & "file does not exist : " & Chr(13) & Chr(10)_
		 & Chr(13) & Chr(10)_
         & prompt_string & Chr(13) & Chr(10))
		WScript.Quit
   end if
 
	g_change_refdes_bool = Get_Change_Refdes_YesNo()

	'Call MsgBox ("T_S_C.vbs: " & Chr(13) & Chr(10)_
	'	& Chr(13) & Chr(10)_
	'	& "g_change_refdes_bool : "& g_change_refdes_bool & Chr(13) & Chr(10))
	
	g_output_in_same_folder_bool = Get_Output_In_Same_Folder_YesNo()

	'Call MsgBox ("T_S_C.vbs: " & Chr(13) & Chr(10)_
	'	& Chr(13) & Chr(10)_
	'	& "g_output_in_same_folder_bool : "& g_output_in_same_folder_bool & Chr(13) & Chr(10))
	
	If (g_output_in_same_folder_bool = false) Then
		Output_Subfolder = Create_Unique_Sub_Folder(schematic_file_path, "\TOP_SIDE")
	Else
		Output_Subfolder = "" ' same as schematic file path
	End If

	'Now we call Read_File_Write_Edited_File() for each page in the schematic
	no_more_pages = false
	Do
		If (Not fso.FileExists(schematic_file_path & "\" & input_schematic_file_name)) Then
			no_more_pages = true
		Else	
			'call MsgBox ("Top_Side_Components.vbs: " & Chr(13) & Chr(10)_
			'	& "ready to edit schematic file" & Chr(13) & Chr(10)_
			'	& input_schematic_file_name &Chr(13) & Chr(10))
			call Read_File_Write_Edited_File(schematic_file_path, input_schematic_file_name, Output_Subfolder)
			'increment the numeric suffix on the schematic file name.
			input_schematic_file_name = Increment_Schematic_Page(input_schematic_file_name)
		End If

	Loop until (no_more_pages)
   
   prompt_string = "Top_Side_Components.vbs completed successfully."_
                & Chr(13) & Chr(10)
   call MsgBox (prompt_string)

End Sub 
