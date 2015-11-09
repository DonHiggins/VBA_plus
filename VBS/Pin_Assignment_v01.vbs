Option Explicit
'- - - - - - - - - - - - - - -
' Pin_Assignment (from Wav_File_Splitter_V3.vbs)
'   -- User Selects a file to open -- a ViewDraw schematic file
'      We try to help add or delete SIGNAL=<net-name>;<pin#> attributes.
'      Originally designed to allow you to specify a SHELL net connection
'      without having to manually go in and do it on every single pin in the connector.
'- - - - - - - - - - - - - - -
'
Dim schematic_file_path        'For ex: "Y:\DON_H\VBScript\Wav_File_Splitter"
Dim input_schematic_file_name  'For ex: "Fishing_Epic_by_Kevin_Kling.wav"
Dim fso                  'the file system object
Dim In_f
Dim component_name_array(100) 
Dim component_name_count

'Initialize global parameters
schematic_file_path = "G:\ECAD\Projects\Aut_TS3\TB3IOM\TB3IOMB_Pin_Assignment\sch" 'default input file folder
input_schematic_file_name = "TB3_Connectors.1" 'default ViewDraw schematic file Name
component_name_count = 0 'start with empty array

call Pin_Assignment_Start

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
  Signal = InputBox(prompt_string, "Pin_Assignment",_
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
  Pin_Num = InputBox(prompt_string, "Pin_Assignment",_
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
  ok_or_cancel = InputBox(prompt_string, "Pin_Assignment",_
                 ok_or_cancel)
  'If operator clicks CANCEL, InputBox returns zero length string ("")
  if (ok_or_cancel = "") Then WScript.Quit 'operator CANCELED script
  if (Not(LCase(ok_or_cancel) = "ok")) Then WScript.Quit 'operator CANCELED script
  
  Get_Signal_and_Pin_Number = UCase(Signal) & ";" & Pin_Num
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
  schematic_file_path = InputBox(prompt_string, "Pin_Assignment",_
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
  input_schematic_file_name = InputBox(prompt_string, "Pin_Assignment",_
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
  Response = InputBox(prompt_string, "Pin_Assignment",_
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
      Input_String = InputBox(prompt_string, "Wav_File_Split",_
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
        call MsgBox ("Pin_Assignment.vbs: " & Chr(13) & Chr(10)_
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
Sub  Read_File_Write_Edited_File(schematic_file_path, input_schematic_file_name, Selected_Component, Signal)
  Dim f_read, Line, fso, f_write, Add_Line, Temp_Line
  Dim Val, token, I
  Const ForReading = 1, ForWriting = 2, ForAppending = 8

  Set fso = CreateObject("Scripting.FileSystemObject")
  Set f_read = fso.OpenTextFile(schematic_file_path & "\" & input_schematic_file_name, ForReading)
  
  Set f_write = fso.CreateTextFile(schematic_file_path & "\" & input_schematic_file_name & "_PA",True)

  
  Do While (f_read.AtEndOfStream <> True)  
    'Read file one line at a time
    Line = f_read.ReadLine

    'Look for a line starting with "I", containing a component type 
    '  ( embedded ":") as the 3rd token
    Do 'actually, we aren't looping, really, just a creative
       '  use of the "Exit Do" construct to jump to the end of a block
      If (Mid(Line,1,1) <> "I") Then Exit Do
	  Temp_Line = Line
      token = get_space_separated_token(Temp_Line,3) 'third token from Line
      If (token = "") Then Exit Do
	  If(LCase(token) = Selected_Component) Then
	    'write out this "I" line
		f_write.Write Line & vbCrLf
		'read in next line
		Line = f_read.ReadLine
		'if next line has a "REFDES=" string, then we use it to form the "A" attribute line for our Pin Assignment
        I = InStr(1,Line,"REFDES=")
        If (Not(I > 0)) Then 
		  'This means there are things we don't understand about schematic file structure, so we abort
          call MsgBox ("Pin_Assignment.vbs: " & Chr(13) & Chr(10)_
            & "Read_File_Write_Edited_File()" & Chr(13) & Chr(10)_
		    & Chr(13) & Chr(10)_
		    & "First record following an ""I"" record does not contain ""REFDES=""" & Chr(13) & Chr(10)_
		    & Chr(13) & Chr(10)_
		    & Line & Chr(13) & Chr(10)_
		    & Chr(13) & Chr(10)_
            & "We don't understand the format.  Exiting." & Chr(13) & Chr(10))
		   WScript.Quit
		End If
		'get 3rd token from "REFDES=" line, it is a # we need to form the new line.
		Temp_Line = Line
		token = get_space_separated_token(Temp_Line,3) 'third token from Line
		Val = token - 20
		Add_Line = "A 50 " & Val & " 10 0 3 0 SIGNAL=" & Signal
		f_write.Write Add_Line & vbCrLf
      End If
	Loop Until (True)
    '- - - - - -
    f_write.Write Line & vbCrLf
	
  Loop
   f_read.Close
   f_write.Close

End Sub
'- - - - - - - - - - - - - - -
'
Sub Pin_Assignment_Start()
'
'Locate the input file
'Read File and interpret its contents as a WAV file
'Ask for info on Splitting file into multiple outputs
'Do it
'
  Dim prompt_string
  Dim I
  Dim debug_input
  Dim Opt_1_Split_2_Crop
  Dim Selected_Component
  Dim Signal_and_Pin_Number

  call Get_Input_File_Name ' ask user for info, leave it in globals
  Set fso = CreateObject("Scripting.FileSystemObject")
   
   If (Not fso.FileExists(schematic_file_path & "\" & input_schematic_file_name)) Then
        prompt_string = schematic_file_path & "\" & input_schematic_file_name
		prompt_string = fix_prompt_string(prompt_string)
        call MsgBox ("Pin_Assignment.vbs: " & Chr(13) & Chr(10)_
         & "Pin_Assignment_Start()" & Chr(13) & Chr(10)_
         & "file does not exist : " & Chr(13) & Chr(10)_
		 & Chr(13) & Chr(10)_
         & prompt_string & Chr(13) & Chr(10))
		WScript.Quit
   end if
  call Read_Schematic_File_Collect_Components(schematic_file_path & "\" & input_schematic_file_name)
  call Write_Components_To_Text_File("Components.txt")
  Selected_Component = Select_Component() 'user identified component to modify
  'Make sure it IS one of the components in our array
  If Not(component_already_in_array(Selected_Component)) Then
       call MsgBox ("Pin_Assignment.vbs: " & Chr(13) & Chr(10)_
         & "Pin_Assignment_Start()" & Chr(13) & Chr(10)_
		 & Chr(13) & Chr(10)_
		 & "Selected_Component : " & Selected_Component & Chr(13) & Chr(10)_
		 & Chr(13) & Chr(10)_
         & "Couldn't find that component in my list.  Exiting." & Chr(13) & Chr(10))
		WScript.Quit
   end if
   Signal_and_Pin_Number = Get_Signal_and_Pin_Number(Selected_Component) 'function call
   call Read_File_Write_Edited_File(schematic_file_path, input_schematic_file_name, Selected_Component, Signal_and_Pin_Number)
 
  prompt_string = "Pin_Assignment.vbs completed successfully."_
                & Chr(13) & Chr(10)
  call MsgBox (prompt_string)

End Sub 
