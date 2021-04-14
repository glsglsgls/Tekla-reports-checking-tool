VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SECONDARYFRclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_SECONDARYFRcollection As Collection
Private p_SampleField As PGfields

'Initialize this class with it's collection:
Private Sub Class_Initialize()
    Set p_SECONDARYFRcollection = New Collection
End Sub

'Allow for adding SamplePGfields:
Public Sub AddField(field As PGfields, Optional key As String)
    Set p_SampleField = field
    If key <> "" Then p_SECONDARYFRcollection.Add field, key Else p_SECONDARYFRcollection.Add field
End Sub

'Expose the collection:
Public Property Get SECONDARYFRcollection() As Collection
    Set SECONDARYFRcollection = p_SECONDARYFRcollection
End Property
Public Function getMANUALquantity()
getMANUALquantity = manualcheckquantity
End Function
Public Function getFAILquantity()
Dim element As Variant
Dim FAILquantity As Integer
For Each element In SECONDARYFRcollection
    If element.STATUSprofile = "FAIL" Or element.STATUSmaterial = "FAIL" Or element.STATUStoughness = "FAIL" Then FAILquantity = FAILquantity + 1
Next element
getFAILquantity = FAILquantity
End Function

Public Function getresult()
On Error Resume Next
Dim i As Integer
For i = 1 To 6
    If ThisWorkbook.Sheets("resultSecondaryFRAMING-" & i) Is Nothing Then
        ThisWorkbook.Sheets.Add(After:=Sheets(ThisWorkbook.Sheets.Count)).name = "resultSecondaryFRAMING-" & i
        Exit For
    End If
Next i
On Error GoTo 0
ThisWorkbook.Sheets("resultSECONDARYFRAMING-" & i).Activate
newsheet = ActiveSheet.name
toprow = 2
firstcol = 2
If toprow <> 0 And firstcol <> 0 Then
    SECONDARYFRtoptablefill
    SECONDARYFRcontenttablefill
    createbuttonmain
    createbuttonup
End If
ThisWorkbook.Sheets("mainVIEW").Activate
End Function
Sub SECONDARYFRtoptablefill()
Dim SECONDARYFRtabletop As Range
Set SECONDARYFRtabletop = Cells(toprow, firstcol)
With SECONDARYFRtabletop
    .Cells(1, 1) = "Item"
    .Cells(1, 2) = "GUIDE"
    .Cells(1, 3) = "Tekla DESIGNATION"
    .Cells(1, 4) = "KM DESIGNATION"
    .Cells(1, 5) = "Tekla MATERIAL"
    .Cells(1, 6) = "KM MATERIAL"
    .Cells(1, 7) = "Tekla TOUGHNESS"
    .Cells(1, 8) = "KM TOUGHNESS"
    .Cells(1, 9) = "LENGTH<12m"
End With
End Sub
Sub SECONDARYFRcontenttablefill()
Dim lrow As Integer
Dim SECONDARYFRtable As Range
Dim firstcolumn, secondcolumn, indicator, j, k As Integer
Set SECONDARYFRtable = Cells(toprow, firstcol)
Set SECONDARYFRtable = SECONDARYFRtable.Resize(SECONDARYFRcollection.Count) 'numcolumns)
Dim element As PGfields
With SECONDARYFRtable
lrow = 2
number = 1
manualcheckquantity = 0
On Error Resume Next
For Each element In SECONDARYFRcollection
indicator = 0

    bar ("creating a SECONDARYFR table  " & Round(number / SECONDARYFRcollection.Count * 100, 0) & "%")
    .Cells(lrow, 1) = element.item
    .Cells(lrow, 2) = element.assembly
firstcolumn = 3
secondcolumn = 4
    .Cells(lrow, firstcolumn) = element.designation
    .Cells(lrow, secondcolumn) = basePG.PG_BASEcollection(element.item).designation
    If element.STATUSprofile = "FAIL" Then
        .Range(Cells(lrow, firstcolumn), Cells(lrow, secondcolumn)).Interior.ColorIndex = 44
        Dim a() As String
        Dim b() As String
        a() = Split(Replace(.Cells(lrow, firstcolumn), "OD ", "OD x "), " x ")
        b() = Split(Replace(.Cells(lrow, secondcolumn), "OD ", "OD x "), " x ")
        For j = LBound(a()) To UBound(a())
            If a(j) <> b(j) Then
                For k = 1 To Len(.Cells(lrow, firstcolumn))
                    If Mid$(.Cells(lrow, firstcolumn).value, k, Len(a(j))) = a(j) Then
                        .Cells(lrow, firstcolumn).Characters(start:=k, length:=Len(a(j))).Font.ColorIndex = 3
                        .Cells(lrow, firstcolumn).Characters(start:=k, length:=Len(a(j))).Font.Bold = True
                    End If
                Next k
                For k = 1 To Len(.Cells(lrow, secondcolumn))
                    If Mid$(.Cells(lrow, secondcolumn).value, k, Len(b(j))) = b(j) Then
                        .Cells(lrow, secondcolumn).Characters(start:=k, length:=Len(b(j))).Font.ColorIndex = 3
                        .Cells(lrow, secondcolumn).Characters(start:=k, length:=Len(b(j))).Font.Bold = True
                    End If
                Next k
            End If
        Next j
    End If
    If .Cells(lrow, secondcolumn) = "" Then
        .Cells(lrow, secondcolumn).Interior.ColorIndex = 22
        indicator = indicator + 1
    End If
firstcolumn = 5
secondcolumn = 6
    .Cells(lrow, firstcolumn) = replaceMinus6(element.material)
    .Cells(lrow, secondcolumn) = basePG.PG_BASEcollection(element.item).material
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
firstcolumn = 7
secondcolumn = 8
    .Cells(lrow, firstcolumn) = element.toughness
    .Cells(lrow, secondcolumn) = basePG.PG_BASEcollection(element.item).toughness
    If element.STATUStoughness = "FAIL" Then
        .Range(Cells(lrow, firstcolumn), Cells(lrow, secondcolumn)).Interior.ColorIndex = 44
        .Range(Cells(lrow, firstcolumn), Cells(lrow, secondcolumn)).Font.ColorIndex = 3
        .Range(Cells(lrow, firstcolumn), Cells(lrow, secondcolumn)).Font.Bold = True
    End If
        If .Cells(lrow, secondcolumn) = "" Then
            .Cells(lrow, secondcolumn).Interior.ColorIndex = 22
            indicator = indicator + 1
    End If

firstcolumn = 9
    .Cells(lrow, firstcolumn) = element.STATUSlength
    If element.STATUSlength = "FAIL" Then
        .Cells(lrow, firstcolumn).Interior.ColorIndex = 44
        .Cells(lrow, firstcolumn).Font.ColorIndex = 3
        .Cells(lrow, firstcolumn).Font.Bold = True
    End If
    If .Cells(lrow, firstcolumn) = "" Then
        .Cells(lrow, firstcolumn).Interior.ColorIndex = 22
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


