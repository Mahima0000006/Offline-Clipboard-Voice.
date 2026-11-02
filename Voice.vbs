Dim voice, fileStream, fso, timestamp, filePath, userText, clipboard, clipText

' Create clipboard object
Set clipboard = CreateObject("htmlfile")

' Prompt user for input text
userText = InputBox("Enter the text to convert to speech (leave blank to use clipboard):", "Text to Speech")

' If user leaves input empty → take from clipboard
If Trim(userText) = "" Then
    On Error Resume Next
    clipText = clipboard.ParentWindow.ClipboardData.GetData("text")
    On Error GoTo 0
    If Trim(clipText) = "" Then
        MsgBox "No input provided and clipboard is empty. Exiting.", vbExclamation, "Error"
        WScript.Quit
    Else
        userText = clipText
    End If
Else
    ' If user typed something → copy to clipboard
    clipboard.ParentWindow.ClipboardData.SetData "text", userText
End If

' === Text-to-Speech to WAV ===
Set voice = CreateObject("SAPI.SpVoice")
Set fileStream = CreateObject("SAPI.SpFileStream")
Set fso = CreateObject("Scripting.FileSystemObject")

' Create a safe timestamp string for filename
timestamp = Replace(Replace(Replace(Now, ":", "-"), "/", "-"), " ", "_")
filePath = fso.GetAbsolutePathName(".") & "\tts_" & timestamp & ".wav"

' Open stream for writing (3 = SSFMCreateForWrite)
fileStream.Open filePath, 3, False

' Assign the stream and speak the input text
Set voice.AudioOutputStream = fileStream
voice.Speak userText

' Close the stream
fileStream.Close

' Notify user
MsgBox "Speech exported to:" & vbCrLf & filePath & vbCrLf & vbCrLf & _
       "(Text also copied to clipboard)", vbInformation, "Done"
