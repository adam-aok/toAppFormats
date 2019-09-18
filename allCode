Sub BuiltorUnbuilt_Click()
Dim sh1 As Worksheet
Dim sh2 As Worksheet
Set sh1 = ActiveWorkbook.Sheets("Sheet1")
Set sh2 = ActiveWorkbook.Sheets("Sheet2")

Dim i As Integer, n As Integer, num As String, five As String, stat As String

i = 2
n = 2

For i = 1 To 2030
num = sh1.Cells(i, 2).Value
stat = sh1.Cells(i, 4).Value
If (stat <> "Built") Then
    For n = 1 To 8187
        five = sh1.Cells(n, 5).Value
        If (InStr(five, num) <> 0) Then
            sh1.Cells(i, 4).Value = "Built"
        End If
    Next n
End If
Next i


End Sub

Sub FixProjectNum()

Dim c As Integer, n As Integer, s As Integer, pnum As String, fmt As String, zero As String, diff As Integer

fmt = "00000"

For c = 2 To 2030
pnum = Cells(c, 2).Value
If (Len(pnum) <> 10) Then
    If (Len(pnum) < 5) Then
        diff = Len(fmt) - Len(pnum)
        zero = Mid(fmt, 1, diff)
        pnum = zero + pnum
    End If
        pnum = pnum + ".00.0"
        Cells(c, 2).Value = pnum
End If
Next c
End Sub
'Works from Vision-downloaded project descriptions to add to the current list in
Sub ProjectDescriptionAdd()

Dim sh1 As Worksheet
Dim sh2 As Worksheet
Set sh1 = ActiveWorkbook.Sheets("List")
Set sh2 = ActiveWorkbook.Sheets("Vision")

Dim i As Integer, n As Integer, pn As String, apn As String, add As String, listRange As Integer, visionRange As Integer, pnumLoc1 As Integer, pnumLoc2 As Integer
'Declares project number column locations for sheets 1 and 2
pnumLoc1 = 1
pnumLoc2 = 1
'List length
listRange = 2030
'Vision list length
visionRange = 1394


'Declares project page description column locations for sheets 1 and 2
Dim pdLoc1 As Integer, pdLoc2 As Integer
pdLoc1 = 7
pdLoc2 = 5



For i = 2 To visionRange
apn = sh2.Cells(i, pnumLoc2).Value
    For n = 2 To listRange
        pn = sh1.Cells(n, pnumLoc2).Value
        If (InStr(pn, apn) <> 0) Then
            sh1.Cells(n, pdLoc1).Value = sh2.Cells(i, pdLoc2).Value
        End If
    Next n
Next i

End Sub
'Adds the published names collected from the vision list in sheet 2
Sub ProjectPublishedNameAdd()

Dim sh1 As Worksheet
Dim sh2 As Worksheet
Set sh1 = ActiveWorkbook.Sheets("List")
Set sh2 = ActiveWorkbook.Sheets("Vision")

Dim i As Integer, n As Integer, pn As String, apn As String, add As String, listRange As Integer, visionRange As Integer, pnumLoc1 As Integer, pnumLoc2 As Integer
'Declares project number column locations for sheets 1 and 2
pnumLoc1 = 1
pnumLoc2 = 1
'List length
listRange = 2030
'Vision list length
visionRange = 1394



'Declares published name column locations for sheets 1 and 2
Dim pubnameLoc1 As Integer, pubnameLoc2 As Integer, j As Integer
pubnameLoc1 = 3
pubnameLoc2 = 2


For i = 2 To visionRange
apn = sh2.Cells(i, pnumLoc2).Value
    For n = 2 To listRange
        pn = sh1.Cells(n, pnumLoc1).Value
        If (InStr(pn, apn) <> 0) Then
            sh1.Cells(n, pubnameLoc1).Value = sh2.Cells(i, pubnameLoc2).Value
        End If
    Next n
Next i

End Sub

'Adds the practice area to the List from the vision sheet
Sub PracticeAreaAdd()

Dim sh1 As Worksheet
Dim sh2 As Worksheet
Set sh1 = ActiveWorkbook.Sheets("List")
Set sh2 = ActiveWorkbook.Sheets("Vision")

Dim i As Integer, n As Integer, pn As String, apn As String, add As String, listRange As Integer, visionRange As Integer, pnumLoc1 As Integer, pnumLoc2 As Integer
'Declares project number column locations for sheets 1 and 2
pnumLoc1 = 1
pnumLoc2 = 1
'List length
listRange = 2030
'Vision list length
visionRange = 1394



'Declares practice area column locations for sheets 1 and 2
Dim pAreaLoc1 As Integer, pAreaLoc2 As Integer
pAreaLoc1 = 4
pAreaLoc2 = 3



For i = 2 To visionRange
apn = sh2.Cells(i, pnumLoc2).Value
    For n = 2 To listRange
        pn = sh1.Cells(n, pnumLoc1).Value
        If (InStr(pn, apn) <> 0) Then
            sh1.Cells(n, pAreaLoc1).Value = sh2.Cells(i, pAreaLoc2).Value
        End If
    Next n
Next i

End Sub

'Adds the office code to the list from the project sheet
Sub OfficeCodeAdd()

Dim sh1 As Worksheet
Dim sh2 As Worksheet
Set sh1 = ActiveWorkbook.Sheets("List")
Set sh2 = ActiveWorkbook.Sheets("Vision")

Dim i As Integer, n As Integer, pn As String, apn As String, add As String, listRange As Integer, visionRange As Integer, pnumLoc1 As Integer, pnumLoc2 As Integer
'Declares project number column locations for sheets 1 and 2
pnumLoc1 = 1
pnumLoc2 = 1
'List length
listRange = 2030
'Vision list length
visionRange = 1394



'Declares practice area column locations for sheets 1 and 2
Dim codeLoc1 As Integer, codeLoc2 As Integer
codeLoc1 = 5
codeLoc2 = 4



For i = 2 To visionRange
apn = sh2.Cells(i, pnumLoc2).Value
    For n = 2 To listRange
        pn = sh1.Cells(n, pnumLoc1).Value
        If (InStr(pn, apn) <> 0) Then
            sh1.Cells(n, codeLoc1).Value = sh2.Cells(i, codeLoc2).Value
        End If
    Next n
Next i

End Sub

'Adds the office code to the list from the project sheet
Sub NoCheck()

Dim sh1 As Worksheet
Dim sh2 As Worksheet
Set sh1 = ActiveWorkbook.Sheets("List")
Set sh2 = ActiveWorkbook.Sheets("NoList")

Dim i As Integer, n As Integer, pn As String, apn As String, add As String, listRange As Integer, visionRange As Integer, pnameLoc1 As Integer, pnameLoc2 As Integer
'Declares project name column locations for sheets 1 and 3
pnameLoc1 = 2
pnameLoc2 = 1
'List length
listRange = 2030
'Vision list length
visionRange = 809



'Declares in app vs not in app
Dim codeLoc1 As Integer, codeLoc2 As Integer
codeLoc1 = 10



For i = 2 To listRange
apn = sh1.Cells(i, pnameLoc1).Value
    For n = 1 To visionRange
        pn = sh2.Cells(n, pnameLoc2).Value
        If (InStr(pn, apn) <> 0) Then
            sh1.Cells(i, codeLoc1).Value = "No"
            sh2.Cells(n, pnameLoc2).Value = ""
        End If
    Next n
Next i

End Sub


'Checks for legacy projects - pin or pin + description
Sub PinCheck()

Dim sh1 As Worksheet
'Dim sh2 As Worksheet
Set sh1 = ActiveWorkbook.Sheets("List")
'Set sh2 = ActiveWorkbook.Sheets("Justpin")

Dim i As Integer, n As Integer, pn As String, apn As String, add As String, listRange As Integer, visionRange As Integer, pnameLoc1 As Integer, pnameLoc2 As Integer
'Declares project name column locations for sheets 1 and 3
pnameLoc1 = 2
pnameLoc2 = 1
'List length
listRange = 2030
'Vision list length
visionRange = 2080



'Declares in app vs not in app
Dim codeLoc1 As Integer, codeLoc2 As Integer
codeLoc1 = 11



For i = 2 To listRange
apn = sh1.Cells(i, pnameLoc1).Value
    'For n = 1 To visionRange
        'pn = sh2.Cells(n, pnameLoc2).Value
        'If (InStr(pn, apn) <> 0) Then
            'sh1.Cells(i, codeLoc1).Value = "Yes"
        'End If
        If (sh1.Cells(i, codeLoc1).Value = "") Then
            sh1.Cells(i, codeLoc1).Value = "No"
        End If
    'Next n
Next i

End Sub

'Adds Service Types for Each Project based on an image list
Sub ServiceTypeAdd()

Dim sh1 As Worksheet
Dim sh2 As Worksheet
Set sh1 = ActiveWorkbook.Sheets("List")
Set sh2 = ActiveWorkbook.Sheets("ServiceType")

Dim i As Integer, n As Integer, pn As String, apn As String, add As String, listRange As Integer, visionRange As Integer, pnumLoc1 As Integer, pnumLoc2 As Integer
'Declares project number column locations for sheets 1 and 2
pnumLoc1 = 1
pnumLoc2 = 1
'List length
listRange = 2030
'Vision list length
visionRange = 10354



'Declares service type column locations for sheets 1 and 2
Dim serviceLoc1 As Integer, serviceLoc2 As Integer
serviceLoc1 = 12
serviceLoc2 = 7



For i = 2 To listRange
    For n = 2 To visionRange
        pn = sh1.Cells(i, pnumLoc1).Value
        apn = sh2.Cells(n, pnumLoc2).Value
        If (InStr(pn, apn) <> 0) Then
            'sh1.Cells(i, serviceLoc1).Value = ""
            If (sh1.Cells(i, serviceLoc1).Value = "") Then
                sh1.Cells(i, serviceLoc1).Value = sh2.Cells(n, serviceLoc2).Value
            Else
                sh1.Cells(i, serviceLoc1).Value = sh1.Cells(i, serviceLoc1).Value & ", " & sh2.Cells(n, serviceLoc2).Value
            End If
        End If
    Next n
Next i

End Sub

Sub PublishedNameNew()

Dim sh1 As Worksheet
Dim sh2 As Worksheet
Set sh1 = ActiveWorkbook.Sheets("List")
Set sh2 = ActiveWorkbook.Sheets("UpdateSheet")

Dim i As Integer, n As Integer, pn As String, apn As String, add As String, listRange As Integer, visionRange As Integer, pnumLoc1 As Integer, pnumLoc2 As Integer
'Declares project number column locations for sheets 1 and 2
pnumLoc1 = 1
pnumLoc2 = 1
'List length
listRange = 2030
'Vision list length
visionRange = 855



'Declares published name column locations for sheets 1 and 2
Dim pubnameLoc1 As Integer, pubnameLoc2 As Integer, j As Integer
pubnameLoc1 = 3
pubnameLoc2 = 2


For i = 2 To listRange
    For n = 2 To visionRange
        pn = sh1.Cells(i, pnumLoc1).Value
        apn = sh2.Cells(n, pnumLoc2).Value
        If (InStr(pn, apn) <> 0) Then
            sh1.Cells(i, pubnameLoc1).Value = sh2.Cells(n, pubnameLoc2).Value
        End If
    Next n
Next i

End Sub

Sub PracticeAreaUpdate()

Dim sh1 As Worksheet
Dim sh2 As Worksheet
Set sh1 = ActiveWorkbook.Sheets("List")
Set sh2 = ActiveWorkbook.Sheets("UpdateSheet")

Dim i As Integer, n As Integer, pn As String, apn As String, add As String, listRange As Integer, visionRange As Integer, pnumLoc1 As Integer, pnumLoc2 As Integer
'Declares project number column locations for sheets 1 and 2
pnumLoc1 = 1
pnumLoc2 = 1
'List length
listRange = 2030
'Vision list length
visionRange = 855



'Declares published name column locations for sheets 1 and 2
Dim pAreaLoc1 As Integer, pAreaLoc2 As Integer, j As Integer
pAreaLoc1 = 4
pAreaLoc2 = 5


For i = 2 To listRange
    For n = 2 To visionRange
        pn = sh1.Cells(i, pnumLoc1).Value
        apn = sh2.Cells(n, pnumLoc2).Value
        If (InStr(pn, apn) <> 0) Then
            sh1.Cells(i, pAreaLoc1).Value = sh2.Cells(n, pAreaLoc2).Value
        End If
    Next n
Next i

End Sub

Sub ImagePaths()

Dim sh1 As Worksheet
Dim sh2 As Worksheet
Set sh1 = ActiveWorkbook.Sheets("List")
Set sh2 = ActiveWorkbook.Sheets("NewProjects")

Dim i As Integer, n As Integer, pn As String, apn As String, add As String, listRange As Integer, visionRange As Integer, pnumLoc1 As Integer, pnumLoc2 As Integer
'Declares project number column locations for sheets 1 and 2
pnumLoc1 = 1
pnumLoc2 = 2
'List length
listRange = 2030
'Vision list length
visionRange = 281



'Declares published name column locations for sheets 1 and 2
Dim ImageLoc1 As Integer, ImageLoc2 As Integer, j As Integer
ImageLoc1 = 13
ImageLoc2 = 3


For i = 2 To listRange
    ImageLoc1 = 13
    For n = 2 To visionRange
        pn = sh1.Cells(i, pnumLoc1).Value
        apn = sh2.Cells(n, pnumLoc2).Value
        If (sh2.Cells(n, 1) = "") Then
            sh2.Cells(n, pnumLoc2).Value = sh2.Cells(n - 1, pnumLoc2).Value
        End If
        If (InStr(pn, apn) <> 0) Then
            If (sh2.Cells(n, pnumLoc2).Value = sh2.Cells(n - 1, pnumLoc2).Value) Then
                ImageLoc1 = ImageLoc1 + 1
            End If
            If sh2.Cells(n, ImageLoc2).Value <> "" Then
                sh1.Cells(i, ImageLoc1).Value = sh2.Cells(n, ImageLoc2).Value
                sh2.Cells(n, ImageLoc2).Value = ""
            End If
        End If
    Next n
Next i

End Sub


Sub CompleteAddressPull()

Dim sh1 As Worksheet
Dim sh2 As Worksheet
Set sh1 = ActiveWorkbook.Sheets("List")
Set sh2 = ActiveWorkbook.Sheets("UpdateSheet")

Dim i As Integer, n As Integer, pn As String, apn As String, add As String, listRange As Integer, visionRange As Integer, pnumLoc1 As Integer, pnumLoc2 As Integer
'Declares project number column locations for sheets 1 and 2
pnumLoc1 = 1
pnumLoc2 = 1
'List length
listRange = 2032
'Vision list length
visionRange = 811



'Declares published name column locations for sheets 1 and 2
Dim AddressNameLoc1 As Integer, Add1Loc2 As Integer, Add2Loc2 As Integer, CityLoc2 As Integer, StateLoc2 As Integer, ZipLoc2 As Integer, CountryLoc2 As Integer
AddressNameLoc1 = 8
Add1Loc2 = 6
Add2Loc2 = 7
CityLoc2 = 8
StateLoc2 = 9
ZipLoc2 = 10
CountryLoc2 = 11


For i = 2 To listRange
    For n = 2 To visionRange
        pn = sh1.Cells(i, pnumLoc1).Value
        apn = sh2.Cells(n, pnumLoc2).Value
        If (InStr(pn, apn) <> 0) Then
            sh1.Cells(i, AddressNameLoc1).Value = sh2.Cells(n, Add1Loc2).Value
            If (sh2.Cells(n, Add2Loc2).Value <> "") Then
                sh1.Cells(i, AddressNameLoc1).Value = sh1.Cells(i, AddressNameLoc1).Value & ", " & sh2.Cells(n, Add2Loc2).Value
            End If
            sh1.Cells(i, AddressNameLoc1).Value = sh1.Cells(i, AddressNameLoc1).Value & vbCrLf & sh2.Cells(n, CityLoc2).Value & ", " & sh2.Cells(n, StateLoc2).Value & " " & sh2.Cells(n, ZipLoc2).Value & vbCrLf & sh2.Cells(n, CountryLoc2).Value
            
        End If
    Next n
Next i

End Sub

Sub GeoCode()
Dim GeoCode As Worksheet
Set GeoCode = ActiveWorkbook.Sheets("List")
Dim i As Integer


For i = 2 To 2032
'GeoCode.Range(Cells(i, 9), Cells(i, 9)).Formula = Ò=MyGeoCode(HÓ & i & Ò)Ó
If GeoCode.Cells(i, 8).Value <> "" Then
GeoCode.Cells(i, 9).Formula = "=MyGeoCode(H" & i & ")"

'Application.Wait (Now + TimeValue(Ò0:00:02Ó))
Application.Wait (Now + TimeValue("0:00:02"))

GeoCode.Cells(i, 9) = GeoCode.Cells(i, 9).Value
End If
Next

End Sub

Sub GeoCode1()
Dim GeoCode1 As Worksheet
Set GeoCode1 = ActiveWorkbook.Sheets("List")
Dim i As Integer


For i = 2 To 2032
'GeoCode.Range(Cells(i, 9), Cells(i, 9)).Formula = Ò=MyGeoCode(HÓ & i & Ò)Ó
If GeoCode1.Cells(i, 9).Value = "Not Found (try again, you may have done too many too fast)" Then

GeoCode1.Cells(i, 9).Formula = "=MyGeoCode(H" & i & ")"
'If GeoCode1.Cells(i, 8).Value <> "" Then
'GeoCode1.Cells(i, 9).Formula = "=MyGeoCode(H" & i & ")"
Application.Wait (Now + TimeValue("0:00:03"))
GeoCode1.Cells(i, 9) = GeoCode1.Cells(i, 9).Value

End If
Next

End Sub

Sub WriteGlobeFileOld()

Dim FilePath As String
Dim CellData As String
Dim LastCol As Integer
Dim LastRow As Integer
Dim i As Integer, j As Integer, c As Integer
Dim myarray As Variant
Dim q As String


Dim sh1 As Worksheet

Set sh1 = ActiveWorkbook.Sheets("List")
myarray = Array("ProjectNumber", "Reference", "jobName", "PracticeArea", "OfficeCode", "BuiltorUnbuilt", "projectInformation", "projectAddress", "projectCoordinates", "InApp", "Legacy", "Services", "ImagePath1", "ImagePath2", "ImagePath3", "ImagePath4", "ImagePath5", "projectAddress2")

LastCol = 18
'sh1.UsedRange.SpecialCells(x1CellTypeLastCell).Column
LastRow = 2032
'sh1.UsedRange.SpecialCells(x1CellTypeLastCell).Row

q = Chr(34)
CellData = ""
c = 1
'MsgBox (Application.DefaultFilePath)

FilePath = Application.DefaultFilePath & "\Globe1.txt"

Open FilePath For Output As #2

For i = 2 To LastRow
    CellData = CellData & "Globe n" & c & " = new Globe();" & vbCrLf
    For j = 1 To LastCol
                
        If (Trim(sh1.Cells(i, j).Value) <> "") Then
            If (j > 12) Then
                CellData = CellData & "n" & c & "." & myarray(j - 1) & " = " & q & ("ProjectImages/") & (Trim(sh1.Cells(i, j).Value)) & "_ipad.jpg" & q & ";" & vbCrLf
            ElseIf (j = 11) Then
                CellData = CellData & "n" & c & "." & myarray(j - 1) & " = " & (Trim(sh1.Cells(i, j).Value)) & ";" & vbCrLf
            ElseIf (j = 10) Then
                CellData = CellData & ""
            ElseIf (j = 9) Then
                Dim CoordArray() As String
                CoordArray = Split(Trim(sh1.Cells(i, j).Value), ",")
                CellData = CellData & "n" & c & "." & "projectLatitude" & " = " & CoordArray(0) & ";" & vbCrLf & "n" & c & "." & "projectLongitude" & " = " & CoordArray(1) & ";" & vbCrLf
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
                CellData = CellData & "n" & c & "." & myarray(j - 1) & " = " & q & ConcDesc & q & ";" & vbCrLf
            ElseIf (j = 8) Then
                Dim AddArray() As String
                Dim AddStr As String
                Dim a As Integer
                AddStr = ""
                AddArray = Split(sh1.Cells(i, j).Value, vbCrLf)
                For a = 0 To (UBound(AddArray))
                    If a = 0 Then
                        AddStr = AddStr & AddArray(0)
                    Else
                        AddStr = AddStr & "\n\n & AddArray(a)"
                    End If
                Next a
                CellData = CellData & "n" & c & "." & myarray(j - 1) & " = " & q & AddStr & q & ";" & vbCrLf
            ElseIf (j = 3) Then
                CellData = CellData & "n" & c & "." & myarray(j - 1) & " = " & q & "\n" & Trim(sh1.Cells(i, j).Value) & q & ";" & vbCrLf
            ElseIf (j = 2) Then
                CellData = CellData & ""
            Else
                CellData = CellData & "n" & c & "." & myarray(j - 1) & " = " & q & Trim(sh1.Cells(i, j).Value) & q & ";" & vbCrLf
            End If
        End If
        
    Next j
    
    CellData = CellData & "pjData.Add (" & "n" & c & ");" & vbCrLf
If sh1.Cells(i, 11).Value = True Then
        If sh1.Cells(i, 8).Value <> "" Then
            Print #2, CellData
            c = c + 1
            CellData = ""
        End If
    Else
        If sh1.Cells(i, 10).Value = "No" Then
            CellData = ""
        ElseIf sh1.Cells(i, 13).Value = "" Then
            CellData = ""
        Else
        Print #2, CellData
        c = c + 1
        CellData = ""
    End If
    End If
    
Next i

Close #2

MsgBox ("Done. File is in: " & Application.DefaultFilePath)

End Sub


Sub WriteGlobeFileIpad()

Dim FilePath As String
Dim CellData As String
Dim LastCol As Integer
Dim LastRow As Long
Dim i As Integer, j As Integer, c As Integer
Dim myarray As Variant
Dim q As String


Dim sh1 As Worksheet

Set sh1 = ActiveWorkbook.Sheets("List")
myarray = Array("ProjectNumber", "Reference", "jobName", "PracticeArea", "OfficeCode", "BuiltorUnbuilt", "projectInformation", "projectAddress", "projectCoordinates", "InApp", "Legacy", "Services", "ImagePath1", "ImagePath2", "ImagePath3", "ImagePath4", "ImagePath5", "projectAddress2")

LastCol = 18
'sh1.UsedRange.SpecialCells(x1CellTypeLastCell).Column
With ActiveSheet
    LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
End With
'sh1.UsedRange.SpecialCells(x1CellTypeLastCell).Row

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
            If (j = 18) Then
                CellData = CellData & "n" & c & "." & myarray(j - 1) & " = " & q & (Trim(sh1.Cells(i, j).Value)) & q & ";" & vbCrLf
            ElseIf (j > 12) Then
                CellData = CellData & "n" & c & "." & myarray(j - 1) & " = " & q & ("ProjectImages/") & LCase(Trim(sh1.Cells(i, j).Value)) & "_ipad.jpg" & q & ";" & vbCrLf
            ElseIf (j = 12) Then
                CellData = CellData & ""
            ElseIf (j = 11) Then
                CellData = CellData & "n" & c & "." & myarray(j - 1) & " = " & (LCase(Trim(sh1.Cells(i, j).Value))) & ";" & vbCrLf
            ElseIf (j = 10) Then
                CellData = CellData & ""
            ElseIf (j = 9) Then
                Dim CoordArray() As String
                CoordArray = Split(Trim(sh1.Cells(i, j).Value), ",")
                CellData = CellData & "n" & c & "." & "projectLatitude" & " = " & CoordArray(0) & ";" & vbCrLf & "n" & c & "." & "projectLongitude" & " = " & CoordArray(1) & ";" & vbCrLf
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
               CellData = CellData & "n" & c & "." & myarray(j - 1) & " = " & q & AddStr & q & ";" & vbCrLf

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
                CellData = CellData & "n" & c & "." & myarray(j - 1) & " = " & q & ConcDesc & q & ";" & vbCrLf
            ElseIf (j = 6) Then
                CellData = CellData & ""
            ElseIf (j = 5) Then
                CellData = CellData & ""
            ElseIf (j = 3) Then
                CellData = CellData & "n" & c & "." & myarray(j - 1) & " = " & q & "\n" & Trim(sh1.Cells(i, j).Value) & q & ";" & vbCrLf
            ElseIf (j = 2) Then
                CellData = CellData & ""
            ElseIf (j = 1) Then
                CellData = CellData & ""
            Else
                CellData = CellData & "n" & c & "." & myarray(j - 1) & " = " & q & Trim(sh1.Cells(i, j).Value) & q & ";" & vbCrLf
            End If
        End If
        
    Next j
    
CellData = CellData & "pjData.Add (" & "n" & c & ");" & vbCrLf
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

fsT.SaveToFile FilePath, 2 'Save binary data To disk
'Close #2

MsgBox ("Done. File is in: " & Application.DefaultFilePath)

End Sub
Sub WriteGlobeiPhone()

Dim FilePathiPhone As String
Dim FilePath
Dim CellData As String
Dim LastCol As Integer
Dim LastRow As Long
Dim i As Integer, j As Integer, c As Integer
Dim myarray As Variant
Dim q As String


Dim sh1 As Worksheet

Set sh1 = ActiveWorkbook.Sheets("List")
myarray = Array("ProjectNumber", "Reference", "jobName", "PracticeArea", "OfficeCode", "BuiltorUnbuilt", "projectInformation", "projectAddress", "projectCoordinates", "InApp", "Legacy", "Services", "ImagePath1", "ImagePath2", "ImagePath3", "ImagePath4", "ImagePath5", "projectAddress2")

LastCol = 18
With ActiveSheet
    LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
End With

q = Chr(34)
CellData = ""
c = 1
'MsgBox (Application.DefaultFilePath)

FilePath = Application.DefaultFilePath & "\GlobeiPhone.cs"

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
         If (j = 18) Then
                CellData = CellData & "n" & c & "." & myarray(j - 1) & " = " & q & (Trim(sh1.Cells(i, j).Value)) & q & ";" & vbCrLf
            ElseIf (j > 13) Then
                CellData = CellData & "n" & c & "." & myarray(j - 1) & " = " & q & ("https://pedatabase.blob.core.windows.net/peappv2pimages/") & LCase(Trim(sh1.Cells(i, j).Value)) & "_iphone.jpg" & q & ";" & vbCrLf
            ElseIf (j = 13) Then
                CellData = CellData & "n" & c & "." & myarray(j - 1) & " = " & q & ("ProjectImages/") & LCase(Trim(sh1.Cells(i, j).Value)) & "_iphone.jpg" & q & ";" & vbCrLf
            ElseIf (j = 12) Then
                CellData = CellData & ""
            ElseIf (j = 11) Then
                CellData = CellData & "n" & c & "." & myarray(j - 1) & " = " & (LCase(Trim(sh1.Cells(i, j).Value))) & ";" & vbCrLf
            ElseIf (j = 10) Then
                CellData = CellData & ""
            ElseIf (j = 9) Then
                Dim CoordArray() As String
                CoordArray = Split(Trim(sh1.Cells(i, j).Value), ",")
                CellData = CellData & "n" & c & "." & "projectLatitude" & " = " & CoordArray(0) & ";" & vbCrLf & "n" & c & "." & "projectLongitude" & " = " & CoordArray(1) & ";" & vbCrLf
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
               CellData = CellData & "n" & c & "." & myarray(j - 1) & " = " & q & AddStr & q & ";" & vbCrLf

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
                CellData = CellData & "n" & c & "." & myarray(j - 1) & " = " & q & ConcDesc & q & ";" & vbCrLf
            ElseIf (j = 6) Then
                CellData = CellData & ""
            ElseIf (j = 5) Then
                CellData = CellData & ""
            ElseIf (j = 3) Then
                CellData = CellData & "n" & c & "." & myarray(j - 1) & " = " & q & "\n" & Trim(sh1.Cells(i, j).Value) & q & ";" & vbCrLf
            ElseIf (j = 2) Then
                CellData = CellData & ""
            ElseIf (j = 1) Then
                CellData = CellData & ""
            Else
                CellData = CellData & "n" & c & "." & myarray(j - 1) & " = " & q & Trim(sh1.Cells(i, j).Value) & q & ";" & vbCrLf
            End If
        End If
        
    Next j
    

CellData = CellData & "pjData.Add (" & "n" & c & ");" & vbCrLf
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

'Close #2
fsT.SaveToFile FilePath, 2 'Save binary data To disk

MsgBox ("Done. File is in: " & Application.DefaultFilePath)

End Sub
Sub ProjectDescriptionCheck()

Dim sh1 As Worksheet
Dim sh2 As Worksheet
Set sh1 = ActiveWorkbook.Sheets("List")
Set sh2 = ActiveWorkbook.Sheets("Vision")

Dim i As Integer, n As Integer, pn As String, apn As String, add As String, listRange As Integer, visionRange As Integer, pnumLoc1 As Integer, pnumLoc2 As Integer
'Declares project number column locations for sheets 1 and 2
pnumLoc1 = 1
pnumLoc2 = 1
'List length
listRange = 2044
'Vision list length
visionRange = 880



'Declares published name column locations for sheets 1 and 2
Dim descLoc1 As Integer, descLoc2 As Integer, desc1 As String, desc2 As String
descLoc1 = 7
descLoc2 = 6


For i = 2 To listRange
    For n = 2 To visionRange
        pn = sh1.Cells(i, pnumLoc1).Value
        apn = sh2.Cells(n, pnumLoc2).Value
        If (InStr(pn, apn) <> 0) Then
            desc1 = sh1.Cells(i, descLoc1).Value
            desc2 = sh2.Cells(n, descLoc2).Value
            If (InStr(desc1, desc2) <> 0) Then
                sh2.Cells(n, descLoc2 + 1).Value = "True"
            Else
                sh2.Cells(n, descLoc2 + 1).Value = "False"
                sh2.Cells(n, descLoc2 + 2).Value = desc1
            End If
        End If
    Next n
Next i

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
        If (j = 2) Then
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
