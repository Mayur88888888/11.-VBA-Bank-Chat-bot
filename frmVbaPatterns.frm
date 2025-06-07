VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVbaPatterns 
   Caption         =   "UserForm1"
   ClientHeight    =   5340
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11544
   OleObjectBlob   =   "frmVbaPatterns.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmVbaPatterns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCopy_Click()
    If txtCode.text = "" Then
        MsgBox "No code to copy.", vbExclamation
        Exit Sub
    End If

    On Error GoTo errHandler
    Dim DataObj As New MSForms.DataObject
    DataObj.SetText txtCode.text
    DataObj.PutInClipboard
    MsgBox "Code copied to clipboard!", vbInformation
    Unload Me
    Exit Sub

errHandler:
    MsgBox "Clipboard error: " & Err.Description, vbCritical
End Sub



Private Sub btnSearch_Click()
    Dim path As String, line As String, keyword As String
    Dim fileNum As Integer, block As String
    Dim matches As Collection
    Set matches = New Collection

    keyword = LCase(Trim(txtSearch.text))
    path = ThisWorkbook.path & "\VBA_Pattern_Bank.txt"

    If Dir(path) = "" Then
        MsgBox "VBA_Pattern_Bank.txt not found in workbook folder.", vbExclamation
        Exit Sub
    End If

    fileNum = FreeFile
    Open path For Input As #fileNum

    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        If line Like "[Category:*]" Then
            block = line & vbCrLf
        ElseIf InStr(1, line, "Description:", vbTextCompare) > 0 Then
            block = block & line & vbCrLf
        ElseIf Trim(line) = "" Then
            block = ""
        Else
            block = block & line & vbCrLf
            If InStr(LCase(block), keyword) > 0 Then
                matches.Add block
            End If
        End If
    Loop
    Close #fileNum

    lstResults.Clear
    txtCode.text = ""

    If matches.Count = 0 Then
        MsgBox "No matching pattern found.", vbInformation
    Else
        Dim i As Integer
        For i = 1 To matches.Count
            lstResults.AddItem "Pattern " & i
            lstResults.List(lstResults.ListCount - 1, 1) = matches(i) ' store full code in hidden column
        Next i
    End If
End Sub


Private Sub lstResults_Click()
txtCode.text = ""
    If lstResults.ListIndex >= 0 Then
        
        txtCode.text = lstResults.List(lstResults.ListIndex, 1)
    End If
    'frmVbaPatterns.Hide
    'Unload Me
End Sub

