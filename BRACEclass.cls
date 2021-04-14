VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BRACEclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private p_BRACEcollection As Collection
Private p_SampleField As PGfields

'Initialize this class with it's collection:
Private Sub Class_Initialize()
    Set p_BRACEcollection = New Collection
End Sub

'Allow for adding SamplePGfields:
Public Sub AddField(field As PGfields, Optional key As String)
    Set p_SampleField = field
    If key <> "" Then p_BRACEcollection.Add field, key Else p_BRACEcollection.Add field
End Sub

'Expose the collection:
Public Property Get BRACEcollection() As Collection
    Set BRACEcollection = p_BRACEcollection
End Property
Public Function getMANUALquantity()
getMANUALquantity = manualcheckquantity
End Function
Public Function getFAILquantity()
Dim element As Variant
Dim FAILquantity As Integer
For Each element In BRACEcollection
    If element.STATUSprofile = "FAIL" Or element.STATUSmaterial = "FAIL" Or element.STATUStoughness = "FAIL" Then FAILquantity = FAILquantity + 1
Next element
getFAILquantity = FAILquantity
End Function

Public Function getresult()
On Error Resume Next
Dim i As Integer
For i = 1 To 6
    If ThisWorkbook.Sheets("resultBRACE-" & i) Is Nothing Then
        ThisWorkbook.Sheets.Add(After:=Sheets(ThisWorkbook.Sheets.Count)).name = "resultBRACE-" & i
        Exit For
    End If
Next i
On Error GoTo 0
ThisWorkbook.Sheets("resultBRACE-" & i).Activate
newsheet = ActiveSheet.name
toprow = 2
firstcol = 2
If toprow <> 0 And firstcol <> 0 Then
    BRACEtoptablefill
    BRACEcontenttablefill
    createbuttonmain
    createbuttonup
End If
ThisWorkbook.Sheets("mainVIEW").Activate
End Function
Sub BRACEtoptablefill()
Dim BRACEtabletop As Range
Set BRACEtabletop = Cells(toprow, firstcol)
With BRACEtabletop
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
Sub BRACEcontenttablefill()
Dim lrow As Integer
Dim BRACEtable As Range
Dim firstcolumn, secondcolumnumn, indicator As Integer
Set BRACEtable = Cells(toprow, firstcol)
Set BRACEtable = BRACEtable.Resize(BRACEcollection.Count) 'numcolumns)
Dim element As PGfields
With BRACEtable
lrow = 2
number = 1
manualcheckquantity = 0
On Error Resume Next
For Each element In BRACEcollection
indicator = 0

    bar ("creating a BRACE table  " & Round(number / BRACEcollection.Count * 100, 0) & "%")
    .Cells(lrow, 1) = element.item
    .Cells(lrow, 2) = element.assembly
firstcolumn = 3
secondcolumn = 4
    'The row below requered only for a moment, we dont want to implement SIS identification
    'PLEASE remove this row!
    If element.designation = "Complex" Then GoTo skipbecauseSIS
    .Cells(lrow, firstcolumn) = element.designation
    .Cells(lrow, secondcolumn) = basePG.PG_BASEcollection(element.item).designation
    If element.STATUSprofile = "FAIL" Then
        .Range(Cells(lrow, firstcolumn), Cells(lrow, secondcolumn)).Interior.ColorIndex = 44
        Dim a() As String
        Dim b() As String
        a() = Split(Replace(.Cells(lrow, firstcolumn), "OD ", "OD x "), "x")
        b() = Split(Replace(.Cells(lrow, secondcolumn), "OD ", "OD x "), "x")
        For j = LBound(a()) To UBound(a())
            If a(j) <> b(j) Then
                For k = 1 To Len(.Cells(lrow, firstcolumn))
                    If Mid$(.Cells(lrow, firstcolumn).value, k, Len(a(j))) = a(j) Then
                        .Cells(lrow, firstcolumn).Characters(start:=k, length:=Len(a(j))).Font.ColorIndex = 3
                        .Cells(lrow, firstcolumn).Characters(start:=k, length:=Len(a(j))).Font.Bold = True
                        Exit For
                    End If
                Next k
                For k = 1 To Len(.Cells(lrow, secondcolumn))
                    If Mid$(.Cells(lrow, secondcolumn).value, k, Len(b(j))) = b(j) Then
                        .Cells(lrow, secondcolumn).Characters(start:=k, length:=Len(b(j))).Font.ColorIndex = 3
                        .Cells(lrow, secondcolumn).Characters(start:=k, length:=Len(b(j))).Font.Bold = True
                        Exit For
                    End If
                Next k
            End If
        Next j
        If UBound(a()) = 1 Then
            If a(0) <> b(0) And a(1) = b(0) And a(1) <> b(1) Then
                For k = 1 To Len(.Cells(lrow, firstcolumn))
                    If Mid$(.Cells(lrow, firstcolumn).value, k, Len(a(1))) = a(1) Then
                        .Cells(lrow, firstcolumn).Characters(start:=k, length:=Len(a(1))).Font.ColorIndex = 1
                        .Cells(lrow, firstcolumn).Characters(start:=k, length:=Len(a(1))).Font.Bold = False
                        Exit For
                    End If
                Next k
                For k = 1 To Len(.Cells(lrow, secondcolumn))
                    If Mid$(.Cells(lrow, secondcolumn).value, k, Len(b(0))) = b(0) Then
                        .Cells(lrow, secondcolumn).Characters(start:=k, length:=Len(b(0))).Font.ColorIndex = 1
                        .Cells(lrow, secondcolumn).Characters(start:=k, length:=Len(b(0))).Font.Bold = False
                        Exit For
                    End If
                Next k
            End If
            If a(1) <> b(1) And a(0) = b(1) And a(0) <> b(0) Then
                For k = 1 To Len(.Cells(lrow, firstcolumn))
                    If Mid$(.Cells(lrow, firstcolumn).value, k, Len(a(0))) = a(0) Then
                        .Cells(lrow, firstcolumn).Characters(start:=k, length:=Len(a(0))).Font.ColorIndex = 1
                        .Cells(lrow, firstcolumn).Characters(start:=k, length:=Len(a(0))).Font.Bold = False
                        Exit For
                    End If
                Next k
                For k = 1 To Len(.Cells(lrow, secondcolumn))
                    If Mid$(.Cells(lrow, secondcolumn).value, k, Len(b(1))) = b(1) Then
                        .Cells(lrow, secondcolumn).Characters(start:=k, length:=Len(b(1))).Font.ColorIndex = 1
                        .Cells(lrow, secondcolumn).Characters(start:=k, length:=Len(b(1))).Font.Bold = False
                        Exit For
                    End If
                Next k
            End If
        End If
    End If
skipbecauseSIS:
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

