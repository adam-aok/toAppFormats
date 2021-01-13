Sub WriteGlobeFileIpad()

Dim FilePath As String
Dim CellData As String
Dim LastCol As Integer
Dim LastRow As Long
Dim i As Integer, j As Integer, c As Integer
Dim myarray As Variant, headerArray() As Variant, headerResults() As Long
Dim q As String
Dim t As Long, s As Long


Dim sh1 As Worksheet

Set sh1 = ActiveWorkbook.Sheets("Data")
myarray = Array("", "ProjectNumber", "Reference", "jobName", "PracticeArea", "OfficeCode", "BuiltorUnbuilt", "projectInformation", "projectAddress", "projectCoordinates", "InApp", "Legacy", "Services", "ImagePath1", "ImagePath2", "ImagePath3", "ImagePath4", "ImagePath5", "projectAddress2")


headerArray = Array("Project Number", "Published Name", "Practice Area", _
                    "Office", "Built or Unbuilt", "Project Page Description", _
                    "Complete Address", "Latitude", "Longitude", _
                    "In App?", "Legacy", "Services", _
                    "ImagePath1", "ImagePath2", "ImagePath3", _
                    "ImagePath4", "ImagePath5", "projectAddress2")
                    
ReDim headerResults(UBound(headerArray))

For s = 0 To UBound(headerResults)
    For t = 1 To 20
        If (sh1.Cells(1, t).Value = headerArray(s)) Then
            headerResults(s) = t
        End If
    Next t
Next s

LastCol = 18
With ActiveSheet
    LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
End With

q = Chr(34)
CellData = ""
c = 1
'MsgBox (Application.DefaultFilePath)

FilePath = Application.DefaultFilePath & "\GlobeiPad.cs"

'Open FilePath For Output As #2

Dim fsT As Object
Set fsT = CreateObject("ADODB.Stream")
fsT.Type = 2 'Specify stream type - we want To save text/string data.
fsT.Charset = "utf-8" 'Specify charset For the source text data.
fsT.Open 'Open the stream And write binary data To the object
For i = 2 To LastRow
    CellData = CellData & "Globe n" & c & " = new Globe();" & vbCrLf
    For j = 1 To LastCol
        If (Trim(sh1.Cells(i, j).Value) <> "") Then
         If (j = headerResults(17)) Then
                CellData = CellData & "n" & c & "." & "projectAddress2" & " = " & q & (Trim(sh1.Cells(i, j).Value)) & q & ";" & vbCrLf
            ElseIf (j = headerResults(16)) Then
                CellData = CellData & "n" & c & "." & "imagePath5" & " = " & q & ("ProjectImages/") & LCase(Trim(sh1.Cells(i, j).Value)) & "_ipad.jpg" & q & ";" & vbCrLf
            ElseIf (j = headerResults(15)) Then
                CellData = CellData & "n" & c & "." & "ImagePath4" & " = " & q & ("ProjectImages/") & LCase(Trim(sh1.Cells(i, j).Value)) & "_ipad.jpg" & q & ";" & vbCrLf
            ElseIf (j = headerResults(14)) Then
                CellData = CellData & "n" & c & "." & "ImagePath3" & " = " & q & ("ProjectImages/") & LCase(Trim(sh1.Cells(i, j).Value)) & "_ipad.jpg" & q & ";" & vbCrLf
            ElseIf (j = headerResults(13)) Then
                CellData = CellData & "n" & c & "." & "ImagePath2" & " = " & q & ("ProjectImages/") & LCase(Trim(sh1.Cells(i, j).Value)) & "_ipad.jpg" & q & ";" & vbCrLf
            ElseIf (j = headerResults(12)) Then
                CellData = CellData & "n" & c & "." & "ImagePath1" & " = " & q & ("ProjectImages/") & LCase(Trim(sh1.Cells(i, j).Value)) & "_iphone.jpg" & q & ";" & vbCrLf
            ElseIf (j = headerResults(11)) Then
                CellData = CellData & ""
            ElseIf (j = headerResults(10)) Then
                CellData = CellData & "n" & c & "." & "Legacy" & " = " & (LCase(Trim(sh1.Cells(i, j).Value))) & ";" & vbCrLf
            ElseIf (j = headerResults(9)) Then
                CellData = CellData & ""
            ElseIf (j = headerResults(8)) Then
                CellData = CellData & "n" & c & "." & "projectLongitude" & " = " & Trim(sh1.Cells(i, j).Value) & ";" & vbCrLf
           ElseIf (j = headerResults(7)) Then
                CellData = CellData & "n" & c & "." & "projectLatitude" & " = " & Trim(sh1.Cells(i, j).Value) & ";" & vbCrLf
           ElseIf (j = headerResults(6)) Then
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
               CellData = CellData & "n" & c & "." & "projectAddress" & " = " & q & AddStr & q & ";" & vbCrLf
            ElseIf (j = headerResults(5)) Then
                Dim DescArray() As String
                Dim ConcDesc As String
                Dim d As Integer
                ConcDesc = ""
                DescArray = Split(sh1.Cells(i, headerResults(5)).Value, vbLf)
                For d = 0 To (UBound(DescArray))
                    If d = 0 Then
                        ConcDesc = ConcDesc & DescArray(0)
                    Else
                        ConcDesc = ConcDesc & "\n\n" & DescArray(d)
                    End If
                Next d
                CellData = CellData & "n" & c & "." & "projectInformation" & " = " & q & ConcDesc & q & ";" & vbCrLf
            ElseIf (j = headerResults(4)) Then
                CellData = CellData & ""
            ElseIf (j = headerResults(3)) Then
                CellData = CellData & ""
            ElseIf (j = headerResults(2)) Then
                CellData = CellData & "n" & c & "." & "PracticeArea" & " = " & q & Trim(sh1.Cells(i, j).Value) & q & ";" & vbCrLf
            ElseIf (j = headerResults(1)) Then
                CellData = CellData & "n" & c & "." & "jobName" & " = " & q & "\n" & Trim(sh1.Cells(i, j).Value) & q & ";" & vbCrLf
            ElseIf (j = 2) Then
                CellData = CellData & ""
            ElseIf (j = 1) Then
                CellData = CellData & ""
            Else
                'CellData = CellData & "n" & c & "." & "PracticeArea" & " = " & q & Trim(sh1.Cells(i, j).Value) & q & ";" & vbCrLf
            End If
        End If
        
    Next j
    

CellData = CellData & "pjData.Add (" & "n" & c & ");" & vbCrLf
    If sh1.Cells(i, headerResults(9)).Value = False Then
        CellData = ""
    ElseIf sh1.Cells(i, headerResults(6)).Value = "" Then
        CellData = ""
    ElseIf sh1.Cells(i, headerResults(10)).Value <> True Then
        If sh1.Cells(i, headerResults(2)).Value = "" Then
            CellData = ""
        ElseIf sh1.Cells(i, headerResults(1)).Value = "" Then
            CellData = ""
        End If
    End If
    If CellData <> "" Then
    fsT.WriteText CellData
        'Print #2, CellData
    c = c + 1
    CellData = ""
    End If
    
Next i



fsT.SaveToFile FilePath, 2 'Save binary data To disk
'Close #2

MsgBox ("Done. File is in: " & Application.DefaultFilePath)

End Sub

Sub WriteGlobes()

Dim FilePathiPhone As String, FilePathiPad As String
Dim CellData As String
Dim CellDataiPad As String, CellDataiPhone As String
Dim LastCol As Integer, LastRow As Long
Dim i As Integer, j As Integer, c As Integer
Dim q As String
Dim myarray As Variant


Dim sh1 As Worksheet

Set sh1 = ActiveWorkbook.Sheets("List")
myarray = Array("ProjectNumber", "Reference", "jobName", "PracticeArea", "OfficeCode", "BuiltorUnbuilt", "projectInformation", "projectAddress", "projectCoordinates", "InApp", "Legacy", "Services", "ImagePath1", "ImagePath2", "ImagePath3", "ImagePath4", "ImagePath5", "projectAddress2")

LastCol = 17
With ActiveSheet
    LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
End With

q = Chr(34)
CellData = ""
c = 1

FilePathiPhone = Application.DefaultFilePath & "\GlobeiPhone.cs"
FilePathiPad = Application.DefaultFilePath & "\GlobeiPad.cs"

Open FilePathiPhone For Output As #2
Open FilePathiPad For Output As #3

For i = 2 To LastRow
CellDataiPhone = ""
CellDataiPad = ""
CellData = CellData & "Globe n" & c & " = new Globe();" & vbCrLf
    For j = 2 To LastCol
    If (Trim(sh1.Cells(i, j).Value) <> "") Then
        If (j = Range("A1","A20").Find("ProjectNumber") Then
            CellData = CellData & "n" & c & "." & myarray(j) & " = " & q & "\n" & Trim(sh1.Cells(i, j).Value) & q & ";" & vbCrLf
        ElseIf (j = 3) Then
            If (Trim(sh1.Cells(i, j).Value) = "Y") Then
                CellData = CellData & "n" & c & "." & myarray(j) & " = " & "true" & ";" & vbCrLf
            Else
                CellData = CellData & "n" & c & "." & myarray(j) & " = " & "true" & ";" & vbCrLf
            End If
        ElseIf (j = 4) Then
            CellData = CellData & ""
        ElseIf (j = 5) Then
            If (Trim(sh1.Cells(i, j - 1).Value) = "Project Page*") Then
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
                CellData = CellData & "n" & c & "." & myarray(j) & " = " & q & ConcDesc & q & ";" & vbCrLf
            Else
                CellData = CellData & ""
            End If
        ElseIf (j = 6) Then
            CellData = CellData & "n" & c & "." & myarray(j) & " = " & q & UCase(Trim(sh1.Cells(i, j).Value)) & q & ";" & vbCrLf
        ElseIf (j = 7) Then
            CellData = CellData & "n" & c & "." & myarray(j) & " = " & q & Trim(sh1.Cells(i, j).Value) & "\n"
        ElseIf (j = 8) Then
            CellData = CellData & Trim(sh1.Cells(i, j).Value) & "\n"
        ElseIf (j = 9) Then
                Dim CoordArray() As String
                CoordArray = Split(Trim(sh1.Cells(i, j).Value), ",")
                CellData = CellData & "n" & c & "." & "projectLatitude" & " = " & CoordArray(0) & ";" & vbCrLf & "n" & c & "." & "projectLongitude" & " = " & CoordArray(1) & ";" & vbCrLf
        ElseIf (j = 10) Then
            CellData = CellData & Trim(sh1.Cells(i, j).Value)
        ElseIf (j = 11) Then
            CellData = CellData & " " & Trim(sh1.Cells(i, j).Value)
        ElseIf (j = 12) Then
            CellData = CellData & "\n" & Trim(sh1.Cells(i, j).Value) & ";" & vbCrLf
        ElseIf (j = 13) Or (j = 14) Then
            CellData = CellData & "n" & c & "." & "myarray(j)" & " = " & Trim(sh1.Cells(i, j).Value) & ";" & vbCrLf
        ElseIf (j = 15) Then
            CellDataiPhone = CellDataiPhone & "n" & c & "." & myarray(j) & " = " & q & ("ProjectImages/") & LCase(Trim(sh1.Cells(i, j).Value)) & "_iphone.jpg" & q & ";" & vbCrLf
            CellDataiPad = CellDataiPad & "n" & c & "." & myarray(j) & " = " & q & ("ProjectImages/") & LCase(Trim(sh1.Cells(i, j).Value)) & "_ipad.jpg" & q & ";" & vbCrLf
        ElseIf (j > 15) Then
            CellDataiPhone = CellDataiPhone & "n" & c & "." & myarray(j) & " = " & q & ("https://pedatabase.blob.core.windows.net/peappv2pimages/") & LCase(Trim(sh1.Cells(i, j).Value)) & "_iphone.jpg" & q & ";" & vbCrLf
            CellDataiPad = CellDataiPad & "n" & c & "." & myarray(j) & " = " & q & ("ProjectImages/") & LCase(Trim(sh1.Cells(i, j).Value)) & "_ipad.jpg" & q & ";" & vbCrLf
        End If
    End If
        
Next j
    

CellDataiPhone = CellData & CellDataiPhone & "pjData.Add (" & "n" & c & ");" & vbCrLf
CellDataiPad = CellData & CellDataiPad & "pjData.Add (" & "n" & c & ");" & vbCrLf
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
    Print #2, CellDataiPhone
    Print #3, CellDataiPhone
    c = c + 1
    CellData = ""
    End If
    
Next i

Close #2

MsgBox ("Done. File is in: " & Application.DefaultFilePath)

End Sub
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



'Open FilePath For Output As #2
'Print #2, "["

Dim fsT As Object
Set fsT = CreateObject("ADODB.Stream")
fsT.Type = 2 'Specify stream type - we want To save text/string data.
fsT.Charset = "utf-8" 'Specify charset For the source text data.
fsT.Open 'Open the stream And write binary data To the object
fsT.WriteText "["

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
    fsT.WriteText CellData
        'Print #2, CellData
    c = c + 1
    CellData = ""
    End If
Next i


fsT.WriteText vbCrLf & "]"
fsT.SaveToFile FilePath, 2 'Save binary data To disk
'Close #2

MsgBox ("Done. File " & "GlobeAndroid.json " & "is in: " & Application.DefaultFilePath)

End Sub
