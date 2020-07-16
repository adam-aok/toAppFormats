VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProjectSheetInsert 
   Caption         =   "Insert Project Page"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6870
   OleObjectBlob   =   "ProjectSheetInsert.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProjectSheetInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddProject_Click()

Dim nArray() As String, mnArray() As Variant, dirArray() As Variant, psArray() As Variant
'Marketing number valid check
Dim mnCheck As Boolean
Dim pnList As String, TotalRows As Long, n As Long, f As Long, folderInfo As String, pArea As String, nameConflict As String
'splits input name into name, project number, marketing number
nArray = Split(FileName.Value, "_")

'array of marketing number inputs
mnArray = Array("", "M0010", "M0020", "M0030", "M0040", "M0050", "M0060", "M0070", "M0080", "M0090", "M0100", "M0110", "M0120", "M0130", "M0140", "M0150", "M0160", "", "", "", "M0200")
dirArray = Array("", "", "", "", "", "Old Template", "Final", "MKG Review", "PA Review", "Tracey Review", "Publish")
psArray = Array("", "010_Healthcare", "020_Senior-Living", "030_Primary-and-Secondary", "040_Higher-Education", "050_Residential", "060_Mixed-Use", "070_Workplace", "080_Hospitality", "090_Government", "100_Retail", "110_Urban-Design", "120_Science-and-Technology", "130_Cultural", "140_Country-Clubs", "150_Religious", "160_Large-Scale", "", "", "", "200_Sports&Exhibition")

'checks if the filename is 3 parts
If UBound(nArray()) <> 2 Then
    MsgBox FileName.Value & " does not follow the format of file name_XXXXX_MXXXX"
    Exit Sub
'checks if filename is empty
End If
If Len(nArray(0)) = 0 Then
    MsgBox "Name is empty"
    Exit Sub
End If
'check if project number is empty
If Len(nArray(1)) = 0 Then
    MsgBox "Project number is empty."
    Exit Sub
End If
'check if marketing number is empty
If Len(nArray(2)) = 0 Then
    MsgBox "Marketing number is empty"
    Exit Sub
End If
'check if folder selection was made
If (Folder.Value = Null) Then
    MsgBox "Select folder"
    Exit Sub
End If


'Choose practice area from filename. Pulls practice area folder corresponding
For i = 1 To UBound(mnArray)
    If mnArray(i) = nArray(2) Then
        mnCheck = True
        pArea = psArray(i)
    End If
Next i

'check for valid marketing number
If mnCheck = False Then
    MsgBox "Invalid Marketing Number"
Exit Sub
End If

'project number check section
Dim pnArr As Variant, pN As String, lCheck As Boolean
pnArr = Split(nArray(1), ".")

'checks for 3-part project number
If UBound(pnArr) = 2 Then
    If IsLetter(pnArr(2)) <> False Then
        lCheck = False
    Else
        lCheck = True
    End If
End If

'trims project number if it begins with zero (spreadsheet trims leading zeros by default)
'redundant code now that all project numbers have been fixed
'If (pnArr(0) <> "00000") And lCheck = False Then
'    If Left$(pnArr(0), 1) = "0" And Len(pnArr(0)) <= 5 Then
'          pnArr(0) = CLng(pnArr(0))
'    End If
'End If

'converts project number to string
pN = CStr(pnArr(0))

'checks for format of XXXXX.XX
If Len(nArray(1)) > 5 And UBound(pnArr) > 0 Then
    pN = pN & "." & pnArr(1)
    'checks for format of XXXXX.XX.X
    If Len(nArray(1)) > 8 Then
        pN = pN & "." & pnArr(2)
    End If
End If

'Activate Correct Sheet
Sheet1.Activate

'find total rows of the page
TotalRows = Rows(Rows.Count).End(xlUp).Row

'Find if entry exists in the spreadsheet already row-by-row
For n = 2 To TotalRows
    If Cells(n, 3) = pN And pArea = Cells(n, 1) And pN <> "00000" Then
        'initializes inserted project dialog box
        folderInfo = ""
        If Cells(n, 5) = "Y" Then
            folderInfo = folderInfo & "Old Template: Yes" & vbCrLf
        Else
            folderInfo = folderInfo & "Old Template: No" & vbCrLf
        End If
        If Cells(n, 6) <> "" Then
            folderInfo = folderInfo & "Old Review: " & Cells(n, 6) & vbCrLf
        Else
            folderInfo = folderInfo & "Old Review: No Review" & vbCrLf
        End If
        If Cells(n, 7) = "Y" Then
            folderInfo = folderInfo & "New Template: Yes" & vbCrLf
        End If
        
        'checks if name of updated file is the same as the one in the list
        nameConflict = ""
        If nArray(0) <> Cells(n, 2) Then
            nameConflict = "NAME CONFLICT DETECTED. PLEASE UPDATE"
        End If
        
        'updates row for projects with userform data. elseif updates the folder with a date if folder has changed
        If Folder.Value = "Final" Then
            Cells(n, 6).Value = "Final"
            Cells(n, 7).Value = "Y"
            Cells(n, 10).Value = Format(Now, "mm/dd/yyyy HH:mm:ss")
            Cells(n, 11).Value = hasA4
            Cells(n, 12).Value = vDesc
            Cells(n, 13).Value = vRes
            Cells(n, 14).Value = vSF
        ElseIf Cells(n, 6) <> Folder.Value Then
            Cells(n, 6).Value = Folder.Value
            Cells(n, 10).Value = Format(Now, "mm/dd/yyyy HH:mm:ss")
        End If
        MsgBox "The project number of" & vbCrLf & nArray(0) & vbCrLf & "is already in use by" & vbCrLf & Cells(n, 2) & "--in row " & n & vbCrLf & nameConflict & vbCrLf & "Directory: " & Cells(n, 1) & vbCrLf & folderInfo
        Sheet1.Cells(n, 2).Select
        Exit Sub
    End If
Next n

Dim emptyRow As Long

'Determine EmptyRow
emptyRow = TotalRows + 1

'Transfer Info of New Entry
Cells(emptyRow, 1).Value = pArea
Cells(emptyRow, 2).Value = nArray(0)
Cells(emptyRow, 3).Value = nArray(1)
Cells(emptyRow, 4).Value = nArray(2)
    If Folder.Value = "Final" Then
        Cells(emptyRow, 6).Value = "Final"
        Cells(emptyRow, 7).Value = "Y"
        Cells(emptyRow, 11).Value = hasA4
        Cells(emptyRow, 12).Value = vDesc
        Cells(emptyRow, 13).Value = vRes
        Cells(emptyRow, 14).Value = vSF
    Else
        Cells(emptyRow, 6).Value = Folder.Value
    End If
Cells(emptyRow, 10).Value = Format(Now, "mm/dd/yyyy HH:mm:ss")

MsgBox "Added"

Sheet1.Cells(emptyRow, 2).Select

Call UserForm_Initialize


End Sub
Public Function IsLetter(strValue As Variant) As Boolean
    Dim intPos As Integer
    For intPos = 1 To Len(strValue)
        Select Case Asc(Mid(strValue, intPos, 1))
            Case 65 To 90, 97 To 122
                IsLetter = True
            Case Else
                IsLetter = False
                Exit For
        End Select
    Next
End Function
      
Private Sub UserForm_Initialize()

'Empty FileName
FileName.Value = ""

'Empty FolderBox
Folder.Clear

'Fill Folder
With Folder
    .AddItem "Final"
    .AddItem "MKG"
    .AddItem "PA"
    .AddItem "Director"
    .AddItem "Publish"
End With
Folder.Value = Null

'null A4
hasA4.Value = False

'Initialize vision checkboxes
vDesc.Value = False
vRes.Value = False
vSF.Value = False

FileName.SetFocus

End Sub
