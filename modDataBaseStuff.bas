Attribute VB_Name = "modDataBaseStuff"
Public Db As Database
Public Rs As Recordset
Public findstart As Integer
Public postnr As Integer
Public g As Long

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXPLORER = &H80000
'UDT that makes calling the commondialog easier
Public Type CMDialog
    Ownerform As Long
    Filter As String
    Filetitle As String
    FilterIndex As Long
    FileName As String
    DefaultExtension As String
    OverwritePrompt As Boolean
    AllowMultiSelect As Boolean
    Initdir As String
    Dialogtitle As String
    Flags As Long
End Type
Public cmndlg As CMDialog

Public Sub ShowOpen()
    Dim OFName As OPENFILENAME
    Dim temp As String
    
    With cmndlg
        .Filter = "All Files (.*)|*.*"
        OFName.lStructSize = Len(OFName)
        OFName.hWndOwner = .Ownerform
        OFName.hInstance = App.hInstance
        OFName.lpstrFilter = Replace(.Filter, "|", Chr(0))
        OFName.lpstrFile = Space$(254)
        OFName.nMaxFile = 255
        OFName.lpstrFileTitle = Space$(254)
        OFName.nMaxFileTitle = 255
        OFName.lpstrInitialDir = .Initdir
        OFName.lpstrTitle = .Dialogtitle
        OFName.nFilterIndex = .FilterIndex
        OFName.Flags = .Flags Or OFN_EXPLORER Or IIf(.AllowMultiSelect, OFN_ALLOWMULTISELECT, 0)
        If GetOpenFileName(OFName) Then
            .FilterIndex = OFName.nFilterIndex
            If .AllowMultiSelect Then
                temp = Replace(Trim$(OFName.lpstrFile), Chr(0), ";")
                If RIGHT(temp, 2) = ";;" Then temp = LEFT(temp, Len(temp) - 2)
                .FileName = temp
            Else
                .FileName = StripTerminator(Trim$(OFName.lpstrFile))
                .Filetitle = StripTerminator(Trim$(OFName.lpstrFileTitle))
            End If
        Else
            .FileName = ""
        End If
    End With

End Sub

Public Function StripTerminator(ByVal strString As String) As String
    'Removes chr(0)'s from the end of a string
    'API tends to do this
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = LEFT$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

Public Sub CreateNewDB(FileName As String)
Dim NewDB As Database
Dim NewTable As TableDef
Dim DBName As String
        
    DBName = App.Path & "\" & FileName
        
    Close

    If Dir(DBName) <> "" Then
        Kill DBName
    End If
    

    Set NewDB = CreateDatabase(DBName, dbLangGeneral)
                
       
    Set NewTable = NewDB.CreateTableDef("data")
    
    With NewTable
        .Fields.Append .CreateField("ZipName", dbMemo)
        .Fields.Append .CreateField("AuthorName", dbMemo)
        .Fields.Append .CreateField("PicName", dbMemo)
        .Fields.Append .CreateField("picture", dbMemo)
        .Fields.Append .CreateField("Comments", dbText)
        For t = 0 To 4
            .Fields(t).AllowZeroLength = True
        Next t
    End With
        
        NewDB.TableDefs.Append NewTable
    
    NewDB.Close
End Sub
Public Sub find(findstr As String, Start As Integer)
Dim found As Boolean, t As Integer
Set Db = OpenDatabase(App.Path & "\klf.mdb")
Set Rs = Db.OpenRecordset("data")
Rs.MoveFirst
For t = Start To Rs.RecordCount - 1
    With Rs
        For f = 0 To .Fields.Count - 1
            If .Fields(f) <> "" Then
            test = InStr(1, findstr, .Fields(f), vbTextCompare)
            'MsgBox test
            If InStr(1, findstr, Trim(.Fields(f))) > 0 Then GoTo found
            End If
        Next f
        Rs.MoveNext
    End With
    
Next t

Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
Exit Sub
found:
Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
findstart = t
visapost t + 1
postnr = t + 1
End Sub

Public Sub visapost(post As Integer)
Set Db = OpenDatabase(App.Path & "\klf.mdb")
Set Rs = Db.OpenRecordset("data")
Dim Data As String
    Rs.Move post - 1
    FrmMain.txtZipName = Trim(Rs.Fields(0))
    FrmMain.txtAuthorName = Trim(Rs.Fields(1))
    FrmMain.txtPicName.Text = Trim(Rs.Fields(2))
    FrmMain.txtComments.Text = Trim(Rs.Fields(4))
    If Rs.Fields(3) > "" Then
        Data = Rs.Fields(3)
    
    Open App.Path & "\tmpfile" For Binary As #1
    Put 1, , Data
    Close
    FrmMain.picMain.Picture = LoadPicture(App.Path & "\tmpfile")
    FrmMain.picBuffer.TOP = FrmMain.picMain.TOP
    FrmMain.picBuffer.ScaleWidth = FrmMain.picMain.ScaleWidth
    FrmMain.picBuffer.ScaleHeight = FrmMain.picMain.ScaleHeight
    StretchSourcePictureFromPicture FrmMain.picMain, FrmMain.picBuffer
    Kill App.Path & "\tmpfile"
    FrmMain.picMain.Picture = LoadPicture
    FrmMain.picBuffer.Visible = True
    Else
    FrmMain.picBuffer.Visible = False
    FrmMain.picMain.Visible = True
    FrmMain.picMain.Cls
    FrmMain.picMain.Print
    FrmMain.picMain.Print
    FrmMain.picMain.Print
    FrmMain.picMain.Print
    FrmMain.picMain.Print
    FrmMain.picMain.Print "           No Picture Available."
    End If
    
    FrmMain.txtLeftStatus.Text = "There are" & Str(Rs.RecordCount) & " files in the database. Showing file " & Str(post) & "."
    FrmMain.txtpost.Text = Str(post)
    Rs.Close: Set Rs = Nothing: Db.Close: Set Db = Nothing
End Sub
