VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TRANSclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private p_TRANScollection As Collection
Private p_SampleField As TRANSfields

'Initialize this class with it's collection:
Private Sub Class_Initialize()
    Set p_TRANScollection = New Collection
End Sub

'Allow for adding SampleTRANSfields:
Public Sub AddField(field As TRANSfields, Optional key As String)
    Set p_SampleField = field
    If key <> "" Then p_TRANScollection.Add field, key Else p_TRANScollection.Add field
End Sub

'Expose the collection:
Public Property Get TRANScollection() As Collection
    Set TRANScollection = p_TRANScollection
End Property
Public Function getMANUALquantity()
getMANUALquantity = manualcheckquantity
End Function
Public Function getFAILquantity()
Dim element As Variant
Dim FAILquantity As Integer
For Each element In TRANScollection
    If element.STATUSlength = "FAIL" Or element.STATUSwidth = "FAIL" Or element.STATUSthickness = "FAIL" _
    Or element.STATUSmaterial = "FAIL" Or element.STATUStoughness = "FAIL" Then FAILquantity = FAILquantity + 1
Next element
getFAILquantity = FAILquantity
End Function


Public Function getresult()
On Error Resume Next
Dim i As Integer
For i = 1 To 6
    If ThisWorkbook.Sheets("resultTRANS-" & i) Is Nothing Then
        ThisWorkbook.Sheets.Add(After:=Sheets(ThisWorkbook.Sheets.Count)).name = "resultTRANS-" & i
        Exit For
    End If
Next i
On Error GoTo 0
ThisWorkbook.Sheets("resultTRANS-" & i).Activate
newsheet = ActiveSheet.name
toprow = 2
firstcol = 2
If toprow <> 0 And firstcol <> 0 Then
    TRANStoptablefill
    TRANScontenttablefill
    createbuttonmain
    createbuttonup
End If

ThisWorkbook.Sheets("mainVIEW").Activate
End Function

Sub TRANStoptablefill()
Dim TRANStabletop As Range
Set TRANStabletop = Cells(toprow, firstcol)
With TRANStabletop
    .Cells(1, 1) = "Detail"
    .Cells(1, 2) = "Node"
    .Cells(1, 3) = "Tag"
    .Cells(1, 4) = "Type"
    .Cells(1, 5) = "GUIDE"
    .Cells(1, 6) = "Tekla WIDTH"
    .Cells(1, 7) = "KM WIDTH"
    .Cells(1, 8) = "Tekla LENGTH"
    .Cells(1, 9) = "KM LENGTH"
    .Cells(1, 10) = "Tekla THICKNESS"
    .Cells(1, 11) = "KM THICKNESS"
    .Cells(1, 12) = "Tekla MATERIAL"
    .Cells(1, 13) = "KM MATERIAL"
    .Cells(1, 14) = "Tekla TOUGHNESS"
    .Cells(1, 15) = "KM TOUGHNESS"
End With
End Sub
Sub TRANScontenttablefill()
Dim lrow As Integer
Dim TRANStable As Range
Dim firstcolumn, secondcolumnumn, indicator As Integer
Set TRANStable = Cells(toprow, firstcol)
Set TRANStable = TRANStable.Resize(TRANScollection.Count) 'numcolumns)
Dim element As TRANSfields
With TRANStable
lrow = 2
On Error Resume Next
number = 1
manualcheckquantity = 0
For Each element In TRANScollection
indicator = 0
    bar ("creating a TRANSITION element table  " & Round(number / TRANScollection.Count * 100, 0) & "%")
    Dim el() As String
    el() = Split(element.detail, "-")
    If el(4) = "F" Then el(4) = "FLANGE PLATE"
    If el(4) = "W" Then el(4) = "WEB PLATE"
    .Cells(lrow, 1) = el(0)
    .Cells(lrow, 2) = el(1)
    .Cells(lrow, 3) = el(2) & "-" & el(3)
    .Cells(lrow, 4) = el(4)
    .Cells(lrow, 5) = element.guide
firstcolumn = 6
secondcolumn = 7
    'The row below requered only for a moment, we dont want to implement SIS identification
    'PLEASE remove this row!
    If element.width = "" Then GoTo skipitbecauseSIS
    .Cells(lrow, firstcolumn) = element.width
    .Cells(lrow, secondcolumn) = baseTRANS.TRANS_BASEcollection(element.detail).width
        If element.STATUSwidth = "FAIL" Then
        .Cells(lrow, firstcolumn).Interior.ColorIndex = 44
        .Cells(lrow, secondcolumn).Interior.ColorIndex = 44
        .Cells(lrow, firstcolumn).Font.ColorIndex = 3
        .Cells(lrow, firstcolumn).Font.Bold = True
        End If
skipitbecauseSIS:
        If .Cells(lrow, secondcolumn) = "" Then
            .Cells(lrow, secondcolumn).Interior.ColorIndex = 22
            indicator = indicator + 1
        End If
firstcolumn = 8
secondcolumn = 9
    .Cells(lrow, firstcolumn) = element.length
    .Cells(lrow, secondcolumn) = baseTRANS.TRANS_BASEcollection(element.detail).length
        If element.STATUSlength = "FAIL" Then
        .Cells(lrow, firstcolumn).Interior.ColorIndex = 44
        .Cells(lrow, secondcolumn).Interior.ColorIndex = 44
        .Cells(lrow, firstcolumn).Font.ColorIndex = 3
        .Cells(lrow, firstcolumn).Font.Bold = True
        End If
        If .Cells(lrow, secondcolumn) = "" Then
            .Cells(lrow, secondcolumn).Interior.ColorIndex = 22
            indicator = indicator + 1
    End If
firstcolumn = 10
secondcolumn = 11
    .Cells(lrow, firstcolumn) = element.thickness
    .Cells(lrow, secondcolumn) = baseTRANS.TRANS_BASEcollection(element.detail).thickness
        If element.STATUSthickness = "FAIL" Then
        .Cells(lrow, firstcolumn).Interior.ColorIndex = 44
        .Cells(lrow, secondcolumn).Interior.ColorIndex = 44
        .Cells(lrow, firstcolumn).Font.ColorIndex = 3
        .Cells(lrow, firstcolumn).Font.Bold = True
        End If
        If .Cells(lrow, secondcolumn) = "" Then
            .Cells(lrow, secondcolumn).Interior.ColorIndex = 22
            indicator = indicator + 1
    End If
firstcolumn = 12
secondcolumn = 13
    .Cells(lrow, firstcolumn) = element.material
    .Cells(lrow, secondcolumn) = baseTRANS.TRANS_BASEcollection(element.detail).material
    If element.STATUSmaterial = "FAIL" Then
        .Range(Cells(lrow, firstcolumn), Cells(lrow, secondcolumn)).Interior.ColorIndex = 44
        For k = 1 To Len(.Cells(lrow, firstcolumn))
            If Mid$(.Cells(lrow, firstcolumn).value, k, 1) <> Mid$(.Cells(lrow, secondcolumn).value, k, 1) Then
                .Cells(lrow, firstcolumn).Characters(start:=k, length:=1).Font.ColorIndex = 3
                .Cells(lrow, firstcolumn).Characters(start:=k, length:=1).Font.Bold = True
                .Cells(lrow, secondcolumn).Characters(start:=k, length:=1).Font.ColorIndex = 3
                .Cells(lrow, secondcolumn).Characters(start:=k, length:=1).Font.Bold = True
            End If
        Next k
    End If
    If .Cells(lrow, secondcolumn) = "" Then
            .Cells(lrow, secondcolumn).Interior.ColorIndex = 22
            indicator = indicator + 1
    End If
firstcolumn = 14
secondcolumn = 15
    .Cells(lrow, firstcolumn) = element.toughness
    .Cells(lrow, secondcolumn) = baseTRANS.TRANS_BASEcollection(element.detail).toughness
        If element.STATUStoughness = "FAIL" Then
        .Cells(lrow, firstcolumn).Interior.ColorIndex = 44
        .Cells(lrow, secondcolumn).Interior.ColorIndex = 44
        .Cells(lrow, firstcolumn).Font.ColorIndex = 3
        .Cells(lrow, firstcolumn).Font.Bold = True
        End If
    If .Cells(lrow, secondcolumn) = "" Then
            .Cells(lrow, secondcolumn).Interior.ColorIndex = 22
            indicator = indicator + 1
    End If
    lrow = lrow + 1
    number = number + 1
    If indicator > 0 Then manualcheckquantity = manualcheckquantity + 1
Next
End With
With ActiveSheet
    .Columns.EntireColumn.AutoFit
    .Cells(toprow, 1).EntireRow.RowHeight = 25
End With
On Error GoTo 0
End Sub


