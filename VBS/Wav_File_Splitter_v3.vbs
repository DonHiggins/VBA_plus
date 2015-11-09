Option Explicit
'- - - - - - - - - - - - - - -
' Wav_File_Splitter_V3.vbs
'   -- User Selects a file to open
'      We try to interpret its contents according to WAV format.
'      We split the file into several smaller files
'      new in v3 -- split the file at a point in time
'- - - - - - - - - - - - - - -
'
Dim wav_file_path        'For ex: "Y:\DON_H\VBScript\Wav_File_Splitter"
Dim input_wav_file_name  'For ex: "Fishing_Epic_by_Kevin_Kling.wav"
Dim fso                  'the file system object
Dim In_f

Dim Number_of_Tracks
Dim Sample_Size
Dim Num_Samples
Dim Samples_per_Track

' Info read from WAV File header
Dim Chunk_ID             'Info from Wav File Header
Dim Chunk_Size
Dim Chunk_Format
Dim Fmt_SubChunk_ID
Dim SubChunk1_Size
Dim Audio_Format
Dim Num_Channels
Dim Sample_Rate
Dim Byte_Rate
Dim Block_Align
Dim Bits_Per_Sample
Dim Data_SubChunk_ID
Dim SubChunk2_Size
Dim Time_To_Split_In_Sec

'Initialize global parameters
wav_file_path = "Y:\DON_H\VBScript\Wav_File_Splitter" 'default input_wav_file folder
input_wav_file_name = "Fishing_Epic_by_Kevin_Kling.wav" 'default ViewDraw schematic Name

call Wav_File_Split

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
Sub Get_Input_File_Name()
'
'Ask for top level project folder and schematic file
'Quit script if user clicks CANCEL
'

  Dim Prompt_String
  Dim ok_or_cancel
  Dim I

  '- - - - - Get Project Path - - - - -

  prompt_string = "Wav File Path  : " & Chr(13) & Chr(10)_
                & "Input WAV File : " & Chr(13) & Chr(10)_
                & " " & Chr(13) & Chr(10)_
                & "Enter full path to folder containing the WAV file: " 

  prompt_string = fix_prompt_string(prompt_string) 'to work better w/ InputBox
  wav_file_path = InputBox(prompt_string, "Wav_File_Split",_
                 wav_file_path)
  'If operator clicks CANCEL, InputBox returns zero length string ("")
  if (wav_file_path = "") Then WScript.Quit 'operator CANCELED script

  '- - - - - Get Schematic Name - - - - -

  prompt_string = "Wav File Path  : " & wav_file_path & Chr(13) & Chr(10)_
                & "Input WAV File : " & Chr(13) & Chr(10)_
                & " " & Chr(13) & Chr(10)_
                & "Enter WAV file name: " 

  prompt_string = fix_prompt_string(prompt_string) 'to work better w/ InputBox
  input_wav_file_name = InputBox(prompt_string, "Wav_File_Split",_
                    input_wav_file_name)
  'If operator clicks CANCEL, InputBox returns zero length string ("")
  if (input_wav_file_name = "") Then WScript.Quit 'operator CANCELED script

  '- - - - - Get OK or Cancel - - - - -

  prompt_string = "Wav File Path  : " & wav_file_path & Chr(13) & Chr(10)_
                & "Input WAV File : " & input_wav_file_name & Chr(13) & Chr(10)_
                & " " & Chr(13) & Chr(10)_
                & "Ready to read file and parse as WAV file. OK or CANCEL? :" 

  prompt_string = fix_prompt_string(prompt_string) 'to work better w/ InputBox
  ok_or_cancel = InputBox(prompt_string, "Wav_File_Split",_
                    "OK")
  'If operator clicks CANCEL, InputBox returns zero length string ("")
  if (ok_or_cancel = "") Then WScript.Quit 'operator CANCELED script
End Sub 
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

Sub Read_File_Find_Components(fname)
'
'Called w/ fname is the fully qualified path and
'  file name (eith extension) of an existing schematic page file. 
'Read the file, a line at a time
'Look for "I <n> <NAME> <n> ..." records
' if <NAME>.1 exists as a file, then add it to our list of components.
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
    'Look for a line starting with "I", containing a schematic 
    '  file name as the 3rd token
    Do 'actually, we aren't looping, really, just a creative
       '  use of the "Exit Do" construct to jump to the end of a block
      If (Mid(Line,1,1) <> "I") Then Exit Do 
      token = get_space_separated_token(Line,3) 'third token from Line
      If (token = "") Then Exit Do
      I = InStr(1,token,":")
      If (I > 0) Then Exit Do 'don't want it if it has a ":" in it
      If (Not fso.FileExists(project_path & "\sch\" & token & "." & 1)) Then 
        If (fso.FileExists(project_path & "\sym\" & token & "." & 1)) Then
          call found_sym_w_out_sch(LCase(token),Line)
        End If 
      Exit Do
      End If
      If (sch_already_in_array(token)) Then Exit Do ' already found this component
      If (UBound(sch_name_array) = sch_name_count) Then
        'sch_name_array is filled up
        call MsgBox ("Roll_ViewDraw_Version.vbs: " & Chr(13) & Chr(10)_
             & "Read_File_Find_Components()" & Chr(13) & Chr(10)_
             & "***FATAL ERROR***:"& Chr(13) & Chr(10)_
             & "     overflowed hard-coded size of sch_name_array: "_
             & UBound(sch_name_array))
        WScript.Quit 'Fatal Error
      end if      
      sch_name_count = sch_name_count + 1
      sch_name_array(sch_name_count) = LCase(token) 
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
Sub Ask_How_Many_Tracks( )
'
'Quit script if user clicks CANCEL
'

  Dim Prompt_String
  Dim ok_or_cancel
  Dim I

  Number_of_Tracks = 0
  While ((Number_of_Tracks < 2) Or (Number_of_Tracks > 30))
     prompt_string = "How many tracks shall I split it into ? " 

     Number_of_Tracks = InputBox(prompt_string, "Wav_File_Split",_
                 10)
     'If operator clicks CANCEL, InputBox returns zero length string ("")
     if (Number_of_Tracks = "") Then WScript.Quit 'operator CANCELED script
  Wend

End Sub 
'- - - - - - - - - - - - - - -
'
Sub Ask_What_Time_To_Split_At( )
'
'Quit script if user clicks CANCEL
'

  Dim Prompt_String
  Dim MM_SS_Input
  Dim I
  Dim MM
  Dim SS
  Dim C
  Dim parsing_minutes

  Time_To_Split_In_Sec = 0
  prompt_string = "Enter time to split as MM:SS" 

  MM_SS_Input = InputBox(prompt_string, "Wav_File_Split",_
                 "MM:SS")
  'If operator clicks CANCEL, InputBox returns zero length string ("")
  if (MM_SS_Input = "") Then WScript.Quit 'operator CANCELED script

  '--- Now parse the MM:SS input
  SS = 0
  MM = 0
  parsing_minutes = true
  Do 
   C = mid(MM_SS_Input,1,1)
   MM_SS_Input = mid(MM_SS_Input,2)
   If (IsNumeric(C)) Then
     If (parsing_minutes) Then
	   MM = (MM * 10) + C
	 Else
	   SS = (SS * 10) + C
	 End If
   Elseif (C = ":") Then
     parsing_minutes = false
   Else
     MsgBox("Exiting, Problem parsing MM:SS input")
     WScript.Quit 'problem parsing MM:SS input
   End If
  Loop until (len(MM_SS_Input) = 0)

  '--- sucessful parsing MM:SS input
  'MsgBox("Successful parsing MM:SS input : " & MM & ":" & SS)
  Time_To_Split_In_Sec = (MM * 60) + SS
End Sub 
'- - - - - - - - - - - - - - -
'
Sub Display_WAV_File_Header( )
  Dim Prompt_String
  Dim ok_or_cancel
  Dim I

  prompt_string = "Chunk_ID             : " & Chunk_ID & Chr(13) & Chr(10) _
                & "Chunk_Size          : " & Chunk_Size & Chr(13) & Chr(10)_
                & "     (Hex)                : " & "0x" & Right("00000000" & Hex(Chunk_Size),8) _
                                           & Chr(13) & Chr(10)_
                & "Chunk_Format       : " & Chunk_Format & Chr(13) & Chr(10) _
                & "fmt_subchunk_id   : " & Fmt_SubChunk_ID & Chr(13) & Chr(10)_
                & "SubChunk1_Size   : " & SubChunk1_Size & Chr(13) & Chr(10)_
                & "Audio_Format        : " & Audio_Format & Chr(13) & Chr(10)_
                & "Num_Channels      : " & Num_Channels & Chr(13) & Chr(10)_
                & "Sample_Rate         : " & Sample_Rate & Chr(13) & Chr(10)_
                & "Byte_Rate              : " & Byte_Rate & Chr(13) & Chr(10)_
                & "Block_Align            : " & Block_Align & Chr(13) & Chr(10)_
                & "Bits_Per_Sample    : " & Bits_Per_Sample & Chr(13) & Chr(10)_
                & "data_subchunk_id : " & Data_SubChunk_ID & Chr(13) & Chr(10)_
                & "SubChunk2_Size   : " & SubChunk2_Size & Chr(13) & Chr(10)_
                & " " & Chr(13) & Chr(10) _
                & "OK ?" 

  prompt_string = fix_prompt_string(prompt_string) 'to work better w/ InputBox
  ok_or_cancel = InputBox(prompt_string, "Wav_File_Split",_
                 "OK")
  'If operator clicks CANCEL, InputBox returns zero length string ("")
  if (ok_or_cancel = "") Then WScript.Quit 'operator CANCELED script

End Sub
'- - - - - - - - - - - - - - -
'
Sub Read_WAV_File_Parse_Info(fname)
'
' Open the file for reading
' Read the byte-fields in the file header
' Leave values in globals
'
  Dim Input_String
  Dim I
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
  Const Show_Debug_Messages = False ' ***** FOR DEBUGGING *****
  Dim Bytes_Read_In_Fmt_Chunk

  Set In_f = fso.OpenTextFile(fname, ForReading)
  
  Chunk_ID = In_f.Read(4)

  Input_String = In_f.Read(4)

  Chunk_Size = Asc(Mid(Input_String,1,1)) _
             + (&H100& * Asc(Mid(Input_String,2,1))) _
             + (&H10000& * Asc(Mid(Input_String,3,1))) _
             + (&H1000000& * Asc(Mid(Input_String,4,1)))

  Chunk_Format = In_f.Read(4)

'In simple WAV file, next subChunk is a Format "fmt" subChunk, but we
'may encounter other types of subChunks, and have to skip over them.

  Do
    Fmt_SubChunk_ID = In_f.Read(4)

    Input_String = In_f.Read(4)
    SubChunk1_Size = Asc(Mid(Input_String,1,1)) _
             + (&H100& * Asc(Mid(Input_String,2,1))) _
             + (&H10000& * Asc(Mid(Input_String,3,1))) _
             + (&H1000000& * Asc(Mid(Input_String,4,1)))

    If (LCase(Fmt_SubChunk_ID) <> "fmt ") Then

      'MsgBox("Skipping over SubChunk : " &  Fmt_SubChunk_ID & Chr(13) & Chr(10) _
	  '   & "Chunk Data Size : " & SubChunk1_Size)

      If (SubChunk1_Size > 1000000) Then
        MsgBox("Exiting, SubChunk size too big.")
        WScript.Quit 'Normal end to script ' exit, problemintermpeting WAVE Format
      End If
     
      Input_String = In_f.Read(SubChunk1_Size)
    End If

  Loop Until (LCase(Fmt_SubChunk_ID) = "fmt ")

  'MsgBox("Reading Fmt SubChunk : " &  Fmt_SubChunk_ID & Chr(13) & Chr(10) _
  '	   & "Chunk Data Size : " & SubChunk1_Size)

  Bytes_Read_In_Fmt_Chunk = 0
  Input_String = In_f.Read(4)
  Bytes_Read_In_Fmt_Chunk = Bytes_Read_In_Fmt_Chunk + 4
  Audio_Format = Asc(Mid(Input_String,1,1)) _
             + (&H100& * Asc(Mid(Input_String,2,1))) 
  Num_Channels = Asc(Mid(Input_String,3,1)) _
             + (&H100& * Asc(Mid(Input_String,4,1)))

  Input_String = In_f.Read(12)
  Bytes_Read_In_Fmt_Chunk = Bytes_Read_In_Fmt_Chunk + 12
  Sample_Rate = Asc(Mid(Input_String,1,1)) _
             + (&H100& * Asc(Mid(Input_String,2,1))) _
             + (&H10000& * Asc(Mid(Input_String,3,1))) _
             + (&H1000000& * Asc(Mid(Input_String,4,1)))
  Byte_Rate = Asc(Mid(Input_String,5,1)) _
             + (&H100& * Asc(Mid(Input_String,6,1))) _
             + (&H10000& * Asc(Mid(Input_String,7,1))) _
             + (&H1000000& * Asc(Mid(Input_String,8,1)))
  Block_Align = Asc(Mid(Input_String,9,1)) _
             + (&H100& * Asc(Mid(Input_String,10,1))) 
  Bits_Per_Sample = Asc(Mid(Input_String,11,1)) _
             + (&H100& * Asc(Mid(Input_String,12,1)))
             
'  MsgBox("Fmt_SubChunk : 0x" & Right("00" & Hex(Asc(Mid(Input_String,1,1))),2)  & Chr(13) & Chr(10) _
'	      &  "               0x" & Right("00" & Hex(Asc(Mid(Input_String,2,1))),2)  & Chr(13) & Chr(10) _
'	      &  "               0x" & Right("00" & Hex(Asc(Mid(Input_String,3,1))),2)  & Chr(13) & Chr(10) _
'	      &  "               0x" & Right("00" & Hex(Asc(Mid(Input_String,4,1))),2)  & Chr(13) & Chr(10) _
'	      &  "               0x" & Right("00" & Hex(Asc(Mid(Input_String,5,1))),2)  & Chr(13) & Chr(10) _
'	      &  "               0x" & Right("00" & Hex(Asc(Mid(Input_String,6,1))),2)  & Chr(13) & Chr(10) _
'	      &  "               0x" & Right("00" & Hex(Asc(Mid(Input_String,7,1))),2)  & Chr(13) & Chr(10) _
'	      &  "               0x" & Right("00" & Hex(Asc(Mid(Input_String,8,1))),2)  & Chr(13) & Chr(10) _
'	      &  "               0x" & Right("00" & Hex(Asc(Mid(Input_String,9,1))),2)  & Chr(13) & Chr(10) _
'	      &  "               0x" & Right("00" & Hex(Asc(Mid(Input_String,10,1))),2)  & Chr(13) & Chr(10) _
'	      &  "               0x" & Right("00" & Hex(Asc(Mid(Input_String,11,1))),2)  & Chr(13) & Chr(10) _
'	      &  "               0x" & Right("00" & Hex(Asc(Mid(Input_String,12,1))),2)  & Chr(13) & Chr(10) _
'	      & "Bits_Per_Sample : " & Bits_Per_Sample)

  'If Fmt chunk contains data beyond what we are using -- in other words
  'if SubChunk1_Size > # of bytes we have read so far, then read (and discard)
  'the additional bytes to get to the next chunk
  If (Bytes_Read_In_Fmt_Chunk < SubChunk1_Size) Then
     Input_String = In_f.Read(SubChunk1_Size - Bytes_Read_In_Fmt_Chunk)
  End If 

  Do
    Data_SubChunk_ID = In_f.Read(4)

    Input_String = In_f.Read(4)
    SubChunk2_Size = Asc(Mid(Input_String,1,1)) _
             + (&H100& * Asc(Mid(Input_String,2,1))) _
             + (&H10000& * Asc(Mid(Input_String,3,1))) _
             + (&H1000000& * Asc(Mid(Input_String,4,1)))

    If (LCase(Data_SubChunk_ID) <> "data") Then

      MsgBox("Skipping over SubChunk : " &  Data_SubChunk_ID & Chr(13) & Chr(10) _
	     & "Chunk Data Size : " & SubChunk2_Size)

      MsgBox("Data_SubChunk_ID : 0x" & Right("00" & Hex(Asc(Mid(Data_SubChunk_ID,1,1))),2)  & Chr(13) & Chr(10) _
	      &  "                   0x" & Right("00" & Hex(Asc(Mid(Data_SubChunk_ID,2,1))),2)  & Chr(13) & Chr(10) _
	      &  "                   0x" & Right("00" & Hex(Asc(Mid(Data_SubChunk_ID,3,1))),2)  & Chr(13) & Chr(10) _
	      &  "                   0x" & Right("00" & Hex(Asc(Mid(Data_SubChunk_ID,4,1))),2)  & Chr(13) & Chr(10) _
	      & "Chunk Data Size : " & SubChunk2_Size)

      If (SubChunk2_Size > 1000000) Then
        MsgBox("Exiting, SubChunk size too big.")
          WScript.Quit 'Normal end to script ' exit, problemintermpeting WAVE Format
        End If
     
       Input_String = In_f.Read(SubChunk2_Size)
    End If

  Loop Until (LCase(Data_SubChunk_ID) = "data")

End Sub 
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
Sub Write_Next_Track(Track_Num)
  Dim f
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
  Const Show_Debug_Messages = False ' ***** FOR DEBUGGING *****
  Dim In_fname 
  Dim Out_fname
  Dim New_Chunk_Size
  Dim New_Subchunk2_Size
  Dim Count
  Dim Block_Size
  Dim This_Block_Size
  Dim Count_at_last_Report 
  Dim Prompt_String
  Dim ok_or_cancel
  Dim I
  Dim Bytes_Written_In_Fmt_Chunk_So_Far

  In_fname = wav_file_path & "\" & input_wav_file_name

  'Useful Illustration: gets (file.extension), (extension), (file w/o extension)
  'MsgBox("GetFileName : " &  fso.GetFileName(fname) & Chr(13) & Chr(10) _
  '     & "GetExtensionName : " & fso.GetExtensionName(fname) & Chr(13) & Chr(10) _
  '     & "GetBaseName : " & fso.GetBaseName(fname) & Chr(13) & Chr(10))

  Out_fname =  wav_file_path _
            & "\" & fso.GetBaseName(In_fname) _
            & "_" & Right("00" & Track_Num ,2) _
            & "." & fso.GetExtensionName(In_fname)

  New_Subchunk2_Size = Samples_per_Track * Sample_Size
  If (Track_Num = Number_of_Tracks) Then
     'Add any left-over samples to last track
     'New_Subchunk2_Size = Samples_per_Track * Sample_Size + (Num_Samples Mod Number_of_Tracks)
	 'Above probably wrong anyway, mixes bytes and samples
	 'Below fulfills origimal intention, and works for split-at-time, as well as split-in-equal-tracks
     New_Subchunk2_Size = (Num_Samples - (Samples_per_Track * (Track_Num - 1))) * Sample_Size 
  End If

  New_Chunk_Size = New_Subchunk2_Size - 36

  Set f = fso.OpenTextFile(Out_fname, ForWriting, True)
  f.Write Chunk_ID 
  f.Write LittleEndian4Bytes(New_Chunk_Size)      'Chunk_Size
  f.Write Chunk_Format 
  f.Write Fmt_SubChunk_ID 
  f.Write LittleEndian4Bytes(SubChunk1_Size) 
  f.Write LittleEndian2Bytes(Audio_Format) 
  f.Write LittleEndian2Bytes(Num_Channels) 
  f.Write LittleEndian4Bytes(Sample_Rate) 
  f.Write LittleEndian4Bytes(Byte_Rate) 
  f.Write LittleEndian2Bytes(Block_Align) 
  f.Write LittleEndian2Bytes(Bits_Per_Sample)
  Bytes_Written_In_Fmt_Chunk_So_Far = 16
  
  'Guessing: Maybe we have to add extra bytes in here to match our Subchunk size !!!!
  Do
    if (Bytes_Written_In_Fmt_Chunk_So_Far < SubChunk1_Size) Then
      f.Write Chr(0) & Chr(0)
	  Bytes_Written_In_Fmt_Chunk_So_Far = Bytes_Written_In_Fmt_Chunk_So_Far + 2
    End If
  Loop Until (Bytes_Written_In_Fmt_Chunk_So_Far >= SubChunk1_Size)
   
  f.Write Data_SubChunk_ID 
  f.Write LittleEndian4Bytes(New_SubChunk2_Size)   'SubChunk2_Size

  'f.Write In_f.Read(New_SubChunk2_Size) 
  
  'MsgBox("Track_Num: " & Track_Num  & Chr(13) & Chr(10) _ 
  '     & "New_SubChunk2_Size: " & New_SubChunk2_Size)

  Count = 0
  Count_at_last_Report = 0
  Block_Size = 10000
  While (Count < New_SubChunk2_Size)
     If (Count + Block_Size <= New_SubChunk2_Size) Then
        This_Block_Size = Block_Size
     Else
        This_Block_Size = New_SubChunk2_Size - Count
     End If

     f.Write In_f.Read(This_Block_Size) 
     Count = Count + This_Block_Size

     '- - - - - - - - - -
     'If (Count > Count_at_last_Report + 50000) Then
     '   prompt_string = "Count So Far  : " & Count & Chr(13) & Chr(10) _
     '           & "Remaining     : " & (New_SubChunk2_Size - Count) & Chr(13) & Chr(10)_
     '           & " " & Chr(13) & Chr(10) _
     '           & "OK ?" 
     '
     '   ok_or_cancel = InputBox(prompt_string, "Wav_File_Split",_
     '            "OK")
     '   'If operator clicks CANCEL, InputBox returns zero length string ("")
     '   if (ok_or_cancel = "") Then WScript.Quit 'operator CANCELED script
     '   Count_at_last_Report = Count
     'End If
     '- - - - - - - - - -

  Wend


End Sub 
'- - - - - - - - - - - - - - -
'
Function Display_Option_To_Split_or_Crop( )
  Dim Prompt_String
  Dim Option_To_Split_or_Crop
  Dim I

  prompt_string = "Enter a number to select function: " & Chr(13) & Chr(10) _
                &  Chr(13) & Chr(10)_
                & " 1 Split the file to equal length segments" & Chr(13) & Chr(10)_
                &  Chr(13) & Chr(10)_
                & " 2 Split the file at a point in time" & Chr(13) & Chr(10)_
                &  Chr(13) & Chr(10)

  prompt_string = fix_prompt_string(prompt_string) 'to work better w/ InputBox
  Option_To_Split_or_Crop = InputBox(prompt_string, "Wav_File_Split",_
                 "1")
  'If operator clicks CANCEL, InputBox returns zero length string ("")
  if (Option_To_Split_or_Crop = "") Then WScript.Quit 'operator CANCELED script

  if (Option_To_Split_or_Crop <> 1) and (Option_To_Split_or_Crop <> 2)Then 
	MsgBox("Exiting. Legal responses: 1, 2, or Cancel")
    WScript.Quit 'operator CANCELED script
  End If

  Display_Option_To_Split_or_Crop = Option_To_Split_or_Crop

End Function
'- - - - - - - - - - - - - - -
'
Sub Wav_File_Split()
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

  call Get_Input_File_Name ' ask user for info, leave it in globals
  Set fso = CreateObject("Scripting.FileSystemObject")

  call Read_WAV_File_Parse_Info(wav_file_path & "\" & input_wav_file_name)
  call Display_WAV_File_Header

  'Calculate these global values from info we read from the "fmt" chunk
  Sample_Size = (Bits_Per_Sample/8)*Num_Channels
  Num_Samples = SubChunk2_Size/Sample_Size
  
  'Ask if we want to split the file into even pieces, or crop (split at one point in time)
  Opt_1_Split_2_Crop = Display_Option_To_Split_or_Crop()
  If (Opt_1_Split_2_Crop = 2) then
  
  '--- Here we are splitting into just 2 tracks at a point in time ---
    call Ask_What_Time_To_Split_At
	' leaves result in global Time_To_Split_In_Sec
	
	Number_of_Tracks = 2
    'Here we set Samples_per_Track equal to the number of samples in the first track
	'All remaining samples are rolled up into the 2nd track
    Samples_per_Track = sample_rate * Time_To_Split_In_Sec

    For I = 1 to Number_of_Tracks
      'If (I) > 2 Then Exit For '*** early exit
      call Write_Next_Track(I)
    Next

    prompt_string = "WAV_File_Splitter.vbs completed split-at-time successfully."_
                & Chr(13) & Chr(10)
    call MsgBox (prompt_string)
    WScript.Quit 'Normal end to script
  End If

  '--- Here we are splitting into equal length tracks ---
  
  call Ask_How_Many_Tracks

  'MsgBox("(Bits_Per_Sample/8)*Num_Channels : " &  (Bits_Per_Sample/8)*Num_Channels & Chr(13) & Chr(10) _
  '     & "SubChunk2_Size/Number_of_Tracks  : " &  (SubChunk2_Size/Number_of_Tracks) & Chr(13) & Chr(10) _
  '     & "(SubChunk2_Size/Number_of_Tracks)\((Bits_Per_Sample/8)*Num_Channels)  : " & ((SubChunk2_Size/Number_of_Tracks)\((Bits_Per_Sample/8)*Num_Channels)) & Chr(13) & Chr(10) _
  '     & "(SubChunk2_Size Mod Number_of_Tracks) : " &  (SubChunk2_Size Mod Number_of_Tracks)) & Chr(13) & Chr(10) _
  '	   & "Number_of_Tracks :" & Number_of_Tracks

  Samples_per_Track = (Num_Samples - (Num_Samples Mod Number_of_Tracks))/Number_of_Tracks
     'Note: this rounds down to an integral number of samples,
     '      any samples left at the end get added to the last track

  For I = 1 to Number_of_Tracks
  
     'If (I) > 2 Then Exit For '*** early exit
     call Write_Next_Track(I)

  Next

  prompt_string = "WAV_File_Splitter.vbs completed successfully."_
                & Chr(13) & Chr(10)


  call MsgBox (prompt_string)

End Sub 
