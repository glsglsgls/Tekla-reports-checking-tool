VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WELDFILLETclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private p_WELDFILLETcollection As Collection
Private p_SampleField As WELDfields
Private detailNumber, WELDFILLETnumber, guide, JOINTtype, WELDFILLETtype, BEVELtype, WELDFILLETsize, GRINDING, WELDFILLETANGLE, NDT, Wbookletnumber As Integer


'Initialize this class with it's collection:
Private Sub Class_Initialize()
    Set p_WELDFILLETcollection = New Collection
End Sub

'Allow for adding SampleWELDfields:
Public Sub AddField(field As WELDfields, Optional key As String)
    Set p_SampleField = field
    If key <> "" Then p_WELDFILLETcollection.Add field, key Else p_WELDFILLETcollection.Add field
End Sub

'Expose the collection:
Public Property Get WELDFILLETcollection() As Collection
    Set WELDFILLETcollection = p_WELDFILLETcollection
End Property

Sub ghgh()
Dim WELDFILLET As New WELDfields
WELDFILLET
End Sub
Public Function getFAILquantity()
Dim element As Variant
Dim FAILquantity As Integer
For Each element In WELDFILLETcollection
    If element.STATUSweldsize = "NO" Then
        FAILquantity = FAILquantity + 1
    End If
Next element
getFAILquantity = FAILquantity
End Function

Public Function getresult()
On Error Resume Next
Dim i As Integer
For i = 1 To 6
    If ThisWorkbook.Sheets("result FILLET WELD-" & i) Is Nothing Then
        ThisWorkbook.Sheets.Add(After:=Sheets(ThisWorkbook.Sheets.Count)).name = "result FILLET WELD-" & i
        Exit For
    End If
Next i
On Error GoTo 0
ThisWorkbook.Sheets("result FILLET WELD-" & i).Activate
newsheet = ActiveSheet.name
toprow = 2
firstcol = 2
If toprow <> 0 And firstcol <> 0 Then
    WELDFILLETtoptablefill
    WELDFILLETcontenttablefill
    createbuttonmain
    createbuttonup
    highlight
End If

ThisWorkbook.Sheets("mainVIEW").Activate
End Function
Sub WELDFILLETtoptablefill()
Dim WELDFILLETtabletop As Range
Set WELDFILLETtabletop = Cells(toprow, firstcol)
With WELDFILLETtabletop
    .Cells(1, 1) = "details NUMBERS"
    .Cells(1, 2) = "weld NUMBER"
    .Cells(1, 3) = "GUIDE"
    .Cells(1, 4) = "THICKNESS of elements"
    .Cells(1, 5) = "WELD SIZE (less then 0.7 x min THCK)"
End With
With Range(Cells(toprow - 1, firstcol), Cells(toprow - 1, firstcol + 4))
    .Merge
    .value = "ATTENTION! This report shows only those welds, which size is less than 0,7 x min thickness!"
    .Font.Bold = True
    .Font.Italic = True
    .HorizontalAlignment = xlCenter
End With
End Sub
Sub WELDFILLETcontenttablefill()
Dim lrow As Integer
Dim firstcolumn, secondcolumnumn As Integer
Dim element As WELDfields
lrow = WELDFILLETcollection.Count
On Error Resume Next
Dim arr() As String
ReDim arr(1 To lrow, 1 To 11)
rown = 1
On Error Resume Next
For Each element In WELDFILLETcollection
    bar ("creating a WELDFILLET table  " & Round(rown / WELDFILLETcollection.Count / 2 * 100, 0) & "%")
    arr(rown, 1) = element.detailsNUM
    arr(rown, 2) = element.weldnumber
    arr(rown, 3) = element.GUID
    arr(rown, 4) = element.weldtype
    arr(rown, 5) = element.weldsize
    rown = rown + 1
Next
With ActiveSheet
    Dim WELDFILLETtable As Range
    Set WELDFILLETtable = Cells(toprow + 1, firstcol)
    Set WELDFILLETtable = WELDFILLETtable.Resize(lrow, 5)
    WELDFILLETtable = arr
    .Columns.EntireColumn.AutoFit
    .Cells(toprow, 1).EntireRow.RowHeight = 25
End With
On Error GoTo 0
End Sub
Function highlight()
lrow = Cells(toprow, firstcol).End(xlDown).row
Dim rng As Range
Set rng = Range(Cells(toprow + 1, firstcol), Cells(lrow, 5))
For i = 1 To rng.Rows.Count
    bar ("highlighting a WELDFILLET table  " & Round(i / rng.Rows.Count / 2 * 100, 0) + 50 & "%")
    If WELDFILLETcollection(rng.Cells(i, 3)).STATUSweldsize = "NO" Then
        With ThisWorkbook.ActiveSheet.Cells(i + toprow, 5 + firstcol - 1)
            .Interior.ColorIndex = 44
            .Font.ColorIndex = 3
            .Font.Bold = True
        End With
    End If
Next i
End Function


