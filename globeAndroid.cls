Sub WriteGlobeAndroid()

Dim FilePath As String
Dim CellData As String
Dim LastCol As Integer
Dim LastRow As Long
Dim i As Integer, j As Integer, c As Integer
Dim q As String
Dim sh1 As Worksheet

Set sh1 = ActiveWorkbook.Sheets("List")
q = Chr(34)
LastCol = 18
With ActiveSheet
    LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
End With
CellData = ""
c = 1
FilePath = Application.DefaultFilePath & "\GlobeAndroid.json"

Open FilePath For Output As #2
Print #2, "["

For i = 2 To LastRow
    CellData = CellData & "{" & vbCrLf
    For j = 1 To LastCol
        If (Trim(sh1.Cells(i, j).Value) <> "") Then
        If (j = 18) Then
            CellData = CellData & vbCrLf & "]," & vbCrLf & q & "projectAddress_2" & q & ": " & q & (Trim(sh1.Cells(i, j).Value)) & q & vbCrLf
        ElseIf (j > 13) Then
            CellData = CellData & "," & vbCrLf & q & "https://pedatabase.blob.core.windows.net/peappv2pimages/" & LCase(Trim(sh1.Cells(i, j).Value)) & "_iphone.jpg" & q
        ElseIf (j = 13) Then
            CellData = CellData & q & "imagesList" & q & ": [" & vbCrLf & q & (Trim(sh1.Cells(i, j).Value)) & "_iphone" & q
            
        ElseIf (j = 11) Then
            CellData = CellData & q & "legacy" & q & ": " & LCase(Trim(sh1.Cells(i, j).Value)) & "," & vbCrLf
        ElseIf (j = 9) Then
            Dim CoordArray() As String
                    CoordArray = Split(Trim(sh1.Cells(i, j).Value), ",")
                    CellData = CellData & q & "latitude" & q & ": " & CoordArray(0) & "," & vbCrLf & q & "longitude" & q & ": " & CoordArray(1) & "," & vbCrLf
        ElseIf (j = 8) Then
            Dim AddArray() As String
                    Dim AddStr As String
                    Dim a As Integer
                    AddStr = ""
                    AddArray = Split(sh1.Cells(i, j).Value, vbLf)
                    For a = 0 To (UBound(AddArray))
                    If a = 0 Then
                            AddStr = AddStr & AddArray(0)
                    Else
                            AddStr = AddStr & "\n" & AddArray(a)
                    End If
                    Next a
                    CellData = CellData & q & "projectAddress" & q & ": " & q & AddStr & q & "," & vbCrLf
        ElseIf (j = 7) Then
                    Dim DescArray() As String
                    Dim ConcDesc As String
                    Dim d As Integer
                    ConcDesc = ""
                    DescArray = Split(sh1.Cells(i, j).Value, vbLf)
                    For d = 0 To (UBound(DescArray))
                        If d = 0 Then
                            ConcDesc = ConcDesc & DescArray(0)
                        Else
                            ConcDesc = ConcDesc & "\n\n" & DescArray(d)
                        End If
                    Next d
                    CellData = CellData & q & "projectInformation" & q & ": " & q & ConcDesc & q & "," & vbCrLf
                ElseIf (j = 4) Then
            CellData = CellData & q & "practiceArea" & q & ": " & q & Trim(sh1.Cells(i, j).Value) & q & "," & vbCrLf
        ElseIf (j = 3) Then
            CellData = CellData & q & "jobName" & q & ": " & q & "\n" & Trim(sh1.Cells(i, j).Value) & q & "," & vbCrLf
        Else
            CellData = CellData & ""
        End If
        ElseIf (j = 13) And (Trim(sh1.Cells(i, j).Value) = "") Then
            CellData = CellData & q & "imagesList" & q & ": ["
        End If
    Next j
    CellData = CellData & "}," & vbCrLf
    If sh1.Cells(i, 10).Value = False Then
        CellData = ""
    ElseIf sh1.Cells(i, 8).Value = "" Then
        CellData = ""
    ElseIf sh1.Cells(i, 11).Value <> True Then
        If sh1.Cells(i, 6).Value = "" Then
            CellData = ""
        ElseIf sh1.Cells(i, 13).Value = "" Then
            CellData = ""
        End If
    End If
    If CellData <> "" Then
    Print #2, CellData
    c = c + 1
    CellData = ""
    End If
Next i
Print #2, vbCrLf & "]"
Close #2
MsgBox ("Done. File " & "GlobeAndroid.json " & "is in: " & Application.DefaultFilePath)

End Sub
