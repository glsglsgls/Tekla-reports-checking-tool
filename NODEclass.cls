VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NODEclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private p_NODEcollection As Collection
Private p_SampleField As NODEfields

'Initialize this class with it's collection:
Private Sub Class_Initialize()
    Set p_NODEcollection = New Collection
End Sub

'Allow for adding SampleSTIFFfields:
Public Sub AddField(field As NODEfields, Optional key As String)
    Set p_SampleField = field
    If key <> "" Then p_NODEcollection.Add field, key Else p_NODEcollection.Add field
End Sub

'Expose the collection:
Public Property Get NODEcollection() As Collection
    Set NODEcollection = p_NODEcollection
End Property
Public Function getMANUALquantity()
getMANUALquantity = manualcheckquantity
End Function
Public Function getFAILquantity()
Dim element As Variant
Dim FAILquantity As Integer
For Each element In NODEcollection
    If element.STATUSAsize = "FAIL" Or element.STATUSBsize = "FAIL" Or element.STATUSAthickness = "FAIL" _
    Or element.STATUSBthickness = "FAIL" Or element.STATUSAtoughness = "FAIL" Or element.STATUSBtoughness = "FAIL" _
    Or element.STATUSmaterial = "FAIL" Then FAILquantity = FAILquantity + 1
Next element
getFAILquantity = FAILquantity
End Function

Public Function getresult()
On Error Resume Next
Dim i As Integer
For i = 1 To 6
    If ThisWorkbook.Sheets("resultNODE-" & i) Is Nothing Then
        ThisWorkbook.Sheets.Add(After:=Sheets(ThisWorkbook.Sheets.Count)).name = "resultNODE-" & i
        Exit For
    End If
Next i
On Error GoTo 0
ThisWorkbook.Sheets("resultNODE-" & i).Activate
newsheet = ActiveSheet.name
toprow = 2
firstcol = 2
If toprow <> 0 And firstcol <> 0 Then
    NODEtoptablefill
    NODEcontenttablefill
    createbuttonmain
    createbuttonup
End If

ThisWorkbook.Sheets("mainVIEW").Activate
End Function

Sub NODEtoptablefill()
Dim NODEtabletop As Range
Set NODEtabletop = Cells(toprow, firstcol)
With NODEtabletop
    .Cells(1, 1) = "Detail"
    .Cells(1, 2) = "Tag"
    .Cells(1, 3) = "Type"
    .Cells(1, 4) = "Deck"
    .Cells(1, 5) = "GUIDE"
    .Cells(1, 6) = "Tekla SIZE A"
    .Cells(1, 7) = "KM SIZE A"
    .Cells(1, 8) = "Tekla SIZE B"
    .Cells(1, 9) = "KM SIZE B"
    .Cells(1, 10) = "Tekla TOUGHNESS A"
    .Cells(1, 11) = "KM TOUGHNESS A"
    .Cells(1, 12) = "Tekla TOUGHNESS B"
    .Cells(1, 13) = "KM TOUGHNESS B"
    .Cells(1, 14) = "Tekla THICKNESS A"
    .Cells(1, 15) = "KM THICKNESS A"
    .Cells(1, 16) = "Tekla THICKNESS B"
    .Cells(1, 17) = "KM THICKNESS B"
    .Cells(1, 18) = "Tekla MATERIAL"
    .Cells(1, 19) = "KM MATERIAL"
End With
End Sub
Sub NODEcontenttablefill()
Dim lrow As Integer
Dim NODEtable As Range
Dim firstcolumn, secondcolumnumn, indicator As Integer
Set NODEtable = Cells(toprow, firstcol)
Set NODEtable = NODEtable.Resize(NODEcollection.Count) 'numcolumns)
Dim element As NODEfields
With NODEtable
lrow = 2
On Error Resume Next
number = 1
manualcheckquantity = 0
For Each element In NODEcollection
indicator = 0

    bar ("creating a NODE table  " & Round(number / NODEcollection.Count * 100, 0) & "%")
    Dim el() As String
    el() = Split(element.detail, "-")
    If el(2) = "st" Then el(2) = "STAR PLATE"
    If el(2) = "od" Then el(2) = "OD. TUBE"
    If el(2) = "webins" Then el(2) = "WEB INSERT"
    If el(2) = "webins2" Then el(2) = "WEB INSERT 'T2'"
    If el(2) = "webins3" Then el(2) = "WEB INSERT 'T3'"
    If el(2) = "intst" Then el(2) = "INTERMEDIATE STAR PLATE"
    el(3) = "DECK " & el(3)
    .Cells(lrow, 1) = el(0)
    .Cells(lrow, 2) = el(1)
    .Cells(lrow, 3) = el(2)
    .Cells(lrow, 4) = el(3)
    .Cells(lrow, 5) = element.guide
firstcolumn = 6
secondcolumn = 7
    'The row below requered only for a moment, we dont want to implement SIS identification
    'PLEASE remove this row!
    If element.AsizeA = "" Then GoTo skipbecauseSIS
    If element.AsizeB <> "blank" Then .Cells(lrow, firstcolumn) = element.AsizeA & "x" & element.AsizeB Else .Cells(lrow, firstcolumn) = element.AsizeA
    If baseNODE.NODE_BASEcollection(element.detail).AsizeB <> "blank" Then .Cells(lrow, secondcolumn) = _
        baseNODE.NODE_BASEcollection(element.detail).AsizeA & "x" & baseNODE.NODE_BASEcollection(element.detail).AsizeB Else .Cells(lrow, secondcolumn) = baseNODE.NODE_BASEcollection(element.detail).AsizeA
    If element.STATUSAsize = "FAIL" Then
        .Range(Cells(lrow, firstcolumn), Cells(lrow, secondcolumn)).Interior.ColorIndex = 44
        Dim a() As String
        Dim b() As String
        a() = Split(.Cells(lrow, firstcolumn), "x")
        b() = Split(.Cells(lrow, secondcolumn), "x")
        For j = LBound(a()) To UBound(a())
            If Abs(a(j) - b(j)) > 4 Then
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
            If a(0) <> b(0) And Abs(a(1) - b(0)) < 4 And a(1) <> b(1) Then
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
            If a(1) <> b(1) And Abs(a(0) - b(1)) < 4 And a(0) <> b(0) Then
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
firstcolumn = 8
secondcolumn = 9
    If element.BsizeB <> "blank" Then .Cells(lrow, firstcolumn) = element.BsizeA & "x" & element.BsizeB Else .Cells(lrow, firstcolumn) = element.BsizeA
    If baseNODE.NODE_BASEcollection(element.detail).BsizeB <> "blank" Then .Cells(lrow, secondcolumn) = _
        baseNODE.NODE_BASEcollection(element.detail).BsizeA & "x" & baseNODE.NODE_BASEcollection(element.detail).BsizeB Else .Cells(lrow, secondcolumn) = baseNODE.NODE_BASEcollection(element.detail).BsizeA
    If element.STATUSBsize = "FAIL" Then
        .Range(Cells(lrow, firstcolumn), Cells(lrow, secondcolumn)).Interior.ColorIndex = 44
        a() = Split(.Cells(lrow, firstcolumn), "x")
        b() = Split(.Cells(lrow, secondcolumn), "x")
        For j = LBound(a()) To UBound(a())
            If Abs(a(j) - b(j)) > 4 Then
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
            If a(0) <> b(0) And Abs(a(1) - b(0)) < 4 And a(1) <> b(1) Then
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
            If a(1) <> b(1) And Abs(a(0) - b(1)) < 4 And a(0) <> b(0) Then
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
    If .Cells(lrow, secondcolumn) = "" Then
        .Cells(lrow, secondcolumn).Interior.ColorIndex = 22
        indicator = indicator + 1
    End If
firstcolumn = 10
secondcolumn = 11
    .Cells(lrow, firstcolumn) = element.Atoughness
    .Cells(lrow, secondcolumn) = baseNODE.NODE_BASEcollection(element.detail).Atoughness
    If element.STATUSAtoughness = "FAIL" Then
        .Range(Cells(lrow, firstcolumn), Cells(lrow, secondcolumn)).Interior.ColorIndex = 44
        .Range(Cells(lrow, firstcolumn), Cells(lrow, secondcolumn)).Font.ColorIndex = 3
        .Range(Cells(lrow, firstcolumn), Cells(lrow, secondcolumn)).Font.Bold = True
    End If
    If .Cells(lrow, secondcolumn) = "" Then
        .Cells(lrow, secondcolumn).Interior.ColorIndex = 22
        indicator = indicator + 1
    End If
        If .Cells(lrow, secondcolumn) = "" Then
            .Cells(lrow, secondcolumn).Interior.ColorIndex = 22
            indicator = indicator + 1
    End If

firstcolumn = 12
secondcolumn = 13
    .Cells(lrow, firstcolumn) = element.Btoughness
    .Cells(lrow, secondcolumn) = baseNODE.NODE_BASEcollection(element.detail).Btoughness
    If element.STATUSBtoughness = "FAIL" Then
        .Range(Cells(lrow, firstcolumn), Cells(lrow, secondcolumn)).Interior.ColorIndex = 44
        .Range(Cells(lrow, firstcolumn), Cells(lrow, secondcolumn)).Font.ColorIndex = 3
        .Range(Cells(lrow, firstcolumn), Cells(lrow, secondcolumn)).Font.Bold = True
    End If
    If .Cells(lrow, secondcolumn) = "" Then
        .Cells(lrow, secondcolumn).Interior.ColorIndex = 22
        indicator = indicator + 1
    End If
        If .Cells(lrow, secondcolumn) = "" Then
            .Cells(lrow, secondcolumn).Interior.ColorIndex = 22
            indicator = indicator + 1
    End If

firstcolumn = 14
secondcolumn = 15
    .Cells(lrow, firstcolumn) = element.Athickness
    .Cells(lrow, secondcolumn) = baseNODE.NODE_BASEcollection(element.detail).Athickness
    If element.STATUSAthickness = "FAIL" Then
        .Range(Cells(lrow, firstcolumn), Cells(lrow, secondcolumn)).Interior.ColorIndex = 44
        .Range(Cells(lrow, firstcolumn), Cells(lrow, secondcolumn)).Font.ColorIndex = 3
        .Range(Cells(lrow, firstcolumn), Cells(lrow, secondcolumn)).Font.Bold = True
    End If
    If .Cells(lrow, secondcolumn) = "" Then
        .Cells(lrow, secondcolumn).Interior.ColorIndex = 22
        indicator = indicator + 1
    End If
        If .Cells(lrow, secondcolumn) = "" Then
            .Cells(lrow, secondcolumn).Interior.ColorIndex = 22
            indicator = indicator + 1
    End If

firstcolumn = 16
secondcolumn = 17
    .Cells(lrow, firstcolumn) = element.Bthickness
    .Cells(lrow, secondcolumn) = baseNODE.NODE_BASEcollection(element.detail).Bthickness
    If element.STATUSBthickness = "FAIL" Then
        .Range(Cells(lrow, firstcolumn), Cells(lrow, secondcolumn)).Interior.ColorIndex = 44
        .Range(Cells(lrow, firstcolumn), Cells(lrow, secondcolumn)).Font.ColorIndex = 3
        .Range(Cells(lrow, firstcolumn), Cells(lrow, secondcolumn)).Font.Bold = True
    End If
    If .Cells(lrow, secondcolumn) = "" Then
        .Cells(lrow, secondcolumn).Interior.ColorIndex = 22
        indicator = indicator + 1
    End If
        If .Cells(lrow, secondcolumn) = "" Then
            .Cells(lrow, secondcolumn).Interior.ColorIndex = 22
            indicator = indicator + 1
    End If

firstcolumn = 18
secondcolumn = 19
    .Cells(lrow, firstcolumn) = replaceMinus6(element.material)
    .Cells(lrow, secondcolumn) = baseNODE.NODE_BASEcollection(element.detail).material
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

