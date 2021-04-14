VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WELDclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private p_WELDcollection As Collection
Private p_SampleField As WELDfields
Private detailNumber, weldnumber, guide, JOINTtype, weldtype, BEVELtype, weldsize, GRINDING, weldANGLE, NDT, Wbookletnumber As Integer


'Initialize this class with it's collection:
Private Sub Class_Initialize()
    Set p_WELDcollection = New Collection
End Sub

'Allow for adding SampleWELDfields:
Public Sub AddField(field As WELDfields, Optional key As String)
    Set p_SampleField = field
    If key <> "" Then p_WELDcollection.Add field, key Else p_WELDcollection.Add field
End Sub

'Expose the collection:
Public Property Get WELDcollection() As Collection
    Set WELDcollection = p_WELDcollection
End Property

Sub ghgh()
Dim WELD As New WELDfields
WELD
End Sub
Public Function getMANUALquantity()
getMANUALquantity = manualcheckquantity
End Function
Public Function getFAILquantity()
Dim element As Variant
Dim FAILquantity As Integer
For Each element In WELDcollection
    If element.STATUSweldtype = "NO" Or element.STATUSweldsize = "NO" Or element.STATUSweldNDT = "NO" _
    Or element.STATUSweldjointtype = "NO" Or element.STATUSweldANGLE = "NO" Or element.STATUSweldfinish = "NO" _
    Or element.STATUSweldbeveltype = "NO" Or element.STATUSweldbooklet = "NO" Then
        FAILquantity = FAILquantity + 1
    End If
Next element
getFAILquantity = FAILquantity
End Function

Public Function getresult()
On Error Resume Next
Dim i As Integer
For i = 1 To 6
    If ThisWorkbook.Sheets("resultWELD-" & i) Is Nothing Then
        ThisWorkbook.Sheets.Add(After:=Sheets(ThisWorkbook.Sheets.Count)).name = "resultWELD-" & i
        Exit For
    End If
Next i
On Error GoTo 0
ThisWorkbook.Sheets("resultWELD-" & i).Activate
newsheet = ActiveSheet.name
toprow = 2
firstcol = 2
If toprow <> 0 And firstcol <> 0 Then
    WELDtoptablefill
    WELDcontenttablefill
    createbuttonmain
    createbuttonup
    highlight
End If

ThisWorkbook.Sheets("mainVIEW").Activate
End Function
Sub WELDtoptablefill()
Dim WELDtabletop As Range
Set WELDtabletop = Cells(toprow, firstcol)
With WELDtabletop
detailNumber = 1
weldnumber = 2
guide = 3
JOINTtype = 4
weldtype = 5
BEVELtype = 6
weldsize = 7
GRINDING = 8
weldANGLE = 9
NDT = 10
Wbookletnumber = 11
    .Cells(1, detailNumber) = "detail NUMBER"
    .Cells(1, weldnumber) = "weld NUMBER"
    .Cells(1, guide) = "GUIDE"
    .Cells(1, JOINTtype) = "JOINT TYPE"
    .Cells(1, weldtype) = "weld TYPE"
    .Cells(1, BEVELtype) = "BEVEL TYPE"
    .Cells(1, weldsize) = "weld SIZE"
    .Cells(1, GRINDING) = "GRINDING"
    .Cells(1, weldANGLE) = "weld ANGLE"
    .Cells(1, NDT) = "NDT class"
    .Cells(1, Wbookletnumber) = "NUMBER in Welding booklet"
    
End With
End Sub
Sub WELDcontenttablefill()
Dim lrow As Integer
Dim firstcolumn, secondcolumnumn As Integer
Dim element As WELDfields
lrow = WELDcollection.Count
On Error Resume Next
Dim arr() As String
ReDim arr(1 To lrow, 1 To 11)
rown = 1
On Error Resume Next
For Each element In WELDcollection
    bar ("creating a WELD table  " & Round(rown / WELDcollection.Count / 2 * 100, 0) & "%")
    arr(rown, detailNumber) = element.detailsNUM
    arr(rown, 2) = element.weldnumber
    arr(rown, 3) = element.GUID
    arr(rown, 4) = element.weldjointtype
    arr(rown, 5) = element.weldtype
    arr(rown, 6) = element.weldbeveltype
    arr(rown, 7) = element.weldsize
    arr(rown, 8) = element.weldfinish
    arr(rown, 9) = element.weldANGLE
    arr(rown, 10) = element.weldNDT
    arr(rown, 11) = element.weldbooklet
    rown = rown + 1
Next
With ActiveSheet
    Dim WELDtable As Range
    Set WELDtable = Cells(toprow + 1, firstcol)
    Set WELDtable = WELDtable.Resize(lrow, 11)
    WELDtable = arr
    .Columns.EntireColumn.AutoFit
    .Cells(toprow, 1).EntireRow.RowHeight = 25
End With
On Error GoTo 0
End Sub
Function highlight()
lrow = Cells(toprow, firstcol).End(xlDown).row
Dim rng As Range
Set rng = Range(Cells(toprow + 1, firstcol), Cells(lrow, 11))
For i = 1 To rng.Rows.Count
    bar ("highlighting a WELD table  " & Round(i / rng.Rows.Count / 2 * 100, 0) + 50 & "%")
    If WELDcollection(rng.Cells(i, 2)).STATUSweldjointtype = "NO" Or WELDcollection(rng.Cells(i, 2)).STATUSweldjointtype = "" Then
        Cells(i + toprow, 3 + firstcol).Interior.ColorIndex = 22
        With ThisWorkbook.ActiveSheet.Cells(i + toprow, JOINTtype + firstcol - 1)
            .Interior.ColorIndex = 44
            .Font.ColorIndex = 3
            .Font.Bold = True
        End With
    End If
    If WELDcollection(rng.Cells(i, 2)).STATUSweldtype = "NO" Or WELDcollection(rng.Cells(i, 2)).STATUSweldtype = "" Then
        With ThisWorkbook.ActiveSheet.Cells(i + toprow, weldtype + firstcol - 1)
            .Interior.ColorIndex = 44
            .Font.ColorIndex = 3
            .Font.Bold = True
        End With
    End If
    If WELDcollection(rng.Cells(i, 2)).STATUSweldbeveltype = "NO" Or WELDcollection(rng.Cells(i, 2)).STATUSweldbeveltype = "" Then
        With ThisWorkbook.ActiveSheet.Cells(i + toprow, BEVELtype + firstcol - 1)
            .Interior.ColorIndex = 44
            .Font.ColorIndex = 3
            .Font.Bold = True
        End With
    End If
    If WELDcollection(rng.Cells(i, 2)).STATUSweldsize = "NO" Or WELDcollection(rng.Cells(i, 2)).STATUSweldsize = "" Then
        With ThisWorkbook.ActiveSheet.Cells(i + toprow, weldsize + firstcol - 1)
            .Interior.ColorIndex = 44
            .Font.ColorIndex = 3
            .Font.Bold = True
        End With
    End If
    If WELDcollection(rng.Cells(i, 2)).STATUSweldfinish = "NO" Or WELDcollection(rng.Cells(i, 2)).STATUSweldfinish = "" Then
        With ThisWorkbook.ActiveSheet.Cells(i + toprow, GRINDING + firstcol - 1)
            .Interior.ColorIndex = 44
            .Font.ColorIndex = 3
            .Font.Bold = True
        End With
    End If
    If WELDcollection(rng.Cells(i, 2)).STATUSweldANGLE = "NO" Or WELDcollection(rng.Cells(i, 2)).STATUSweldANGLE = "" Then
        With ThisWorkbook.ActiveSheet.Cells(i + toprow, weldANGLE + firstcol - 1)
            .Interior.ColorIndex = 44
            .Font.ColorIndex = 3
            .Font.Bold = True
        End With
    End If
    If WELDcollection(rng.Cells(i, 2)).STATUSweldNDT = "NO" Or WELDcollection(rng.Cells(i, 2)).STATUSweldNDT = "" Then
        With ThisWorkbook.ActiveSheet.Cells(i + toprow, NDT + firstcol - 1)
            .Interior.ColorIndex = 44
            .Font.ColorIndex = 3
            .Font.Bold = True
        End With
    End If
    If WELDcollection(rng.Cells(i, 2)).STATUSweldbooklet = "NO" Or WELDcollection(rng.Cells(i, 2)).STATUSweldbooklet = "" Then
        With ThisWorkbook.ActiveSheet.Cells(i + toprow, Wbookletnumber + firstcol - 1)
            .Interior.ColorIndex = 44
            .Font.ColorIndex = 3
            .Font.Bold = True
        End With
    End If
Next i
End Function
