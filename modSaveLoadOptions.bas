Attribute VB_Name = "modSaveLoadOptions"
Option Explicit

Public Function LoadText(FromFile As String) As String
On Error GoTo Handle
'Checking if the file currently exists
If FileExists(FromFile) = False Then MsgBox "File not found. Check if the file is Currently exists.", vbCritical, "Sorry": Exit Function
Dim sTemp As String
Dim filenum As Integer

    filenum = FreeFile
    Open FromFile For Input As #filenum   'Open the file to read
        sTemp = Input(LOF(1), filenum)    'Getting the text
    Close #filenum                        'Closing the file
    LoadText = sTemp
Exit Function
Handle:
MsgBox "Error " & Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
End Function


Public Function SaveText(Text As String, FileName As String) As Boolean
On Error GoTo Handle
Dim sTemp As String
Dim filenum As Integer
Dim rply As String

    filenum = FreeFile
    sTemp = Text

    If FileExists(FileName) = False Then    'Check whether the file created
        Open FileName For Output As #filenum  'Opening the file to SaveText
        Print #filenum, sTemp             'Printing  the text to the file
        Close #filenum                        'Closing
        SaveText = False     'Returns 'False'
    Else
       ' rply = MsgBox("File already exists. Do you want to overwrite?", vbYesNo, "File exists already")
       ' If rply = vbYes Then
           Open FileName For Output As #filenum  'Opening the file to SaveText
           Print #filenum, sTemp             'Printing  the text to the file
           Close #filenum                        'Closing
        'Else
        '   Open FileName For Append As #filenum  'Opening the file to SaveText
        '   Print #filenum, sTemp             'Printing  the text to the file
         '  Close #filenum                        'Closing
       'End If
       SaveText = True     'Returns 'True'
       MsgBox "Option change has been saved."
    End If
Exit Function
Handle:
    SaveText = False
    MsgBox "Error " & Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
End Function


Public Function FileExists(FileName As String) As Boolean
'This function checks the existance of a file
On Error GoTo Handle
    If FileLen(FileName) >= 0 Then: FileExists = True: Exit Function
Handle:
    FileExists = False
End Function

