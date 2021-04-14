VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "STIFFjclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private p_STIFFJcollection As Collection
Private p_SampleField As STIFFfields

'Initialize this class with it's collection:
Private Sub Class_Initialize()
    Set p_STIFFJcollection = New Collection
End Sub

'Allow for adding SampleSTIFFfields:
Public Sub AddField(field As STIFFfields, Optional key As String)
    Set p_SampleField = field
    If key <> "" Then p_STIFFJcollection.Add field, key Else p_STIFFJcollection.Add field
End Sub

'Expose the collection:
Public Property Get STIFFJcollection() As Collection
    Set STIFFJcollection = p_STIFFJcollection
End Property
Public Function getMANUALquantity()
getMANUALquantity = manualcheckquantity
End Function
Public Function getFAILquantity()
Dim element As Variant
Dim FAILquantity As Integer
For Each element In STIFFJcollection
    If element.STATUSsize = "FAIL" Or element.STATUSquantity = "FAIL" Or element.STATUSthickness = "FAIL" _
    Or element.STATUSmaterial = "FAIL" Or element.STATUStoughness = "FAIL" Then FAILquantity = FAILquantity + 1
Next element
getFAILquantity = FAILquantity
End Function


Public Function getresult()
On Error Resume Next
Dim i As Integer
For i = 1 To 6
    If ThisWorkbook.Sheets("result J-elements-" & i) Is Nothing Then
        ThisWorkbook.Sheets.Add(After:=Sheets(ThisWorkbook.Sheets.Count)).name = "result J-elements-" & i
        Exit For
    End If
Next i
On Error GoTo 0
ThisWorkbook.Sheets("result J-elements-" & i).Activate
newsheet = ActiveSheet.name
toprow = 2
firstcol = 2
If toprow <> 0 And firstcol <> 0 Then
    STIFFJtoptablefill
    STIFFJcontenttablefill
    createbuttonmain
    createbuttonup
End If

ThisWorkbook.Sheets("mainVIEW").Activate
End Function
Sub STIFFJtoptablefill()
Dim STIFFJtabletop As Range
Set STIFFJtabletop = Cells(toprow, firstcol)
With STIFFJtabletop
    .Cells(1, 1) = "Detail"
    .Cells(1, 2) = "Tag"
    .Cells(1, 3) = "Type"
    .Cells(1, 4) = "Deck"
    .Cells(1, 5) = "GUIDE"
    .Cells(1, 6) = "Tekla SIZE (AxB)"
    .Cells(1, 7) = "KM SIZE (AxB)"
    .Cells(1, 8) = "Tekla THICKNESS"
    .Cells(1, 9) = "KM THICKNESS"
    .Cells(1, 10) = "Tekla MATERIAL"
    .Cells(1, 11) = "KM MATERIAL"
    .Cells(1, 12) = "Tekla TOUGHNESS"
    .Cells(1, 13) = "KM TOUGHNESS"
End With
End Sub
Sub STIFFJcontenttablefill()
Dim lrow As Integer
Dim STIFFJtable As Range
Dim firstcolumn, secondcolumnumn, indicator As Integer
Set STIFFJtable = Cells(toprow, firstcol)
Set STIFFJtable = STIFFJtable.Resize(STIFFJcollection.Count) 'numcolumns)
Dim element As STIFFfields
With STIFFJtable
lrow = 2
On Error Resume Next
number = 1
manualcheckquantity = 0
For Each element In STIFFJcollection
indicator = 0

    bar ("creating a J-stiffener table  " & Round(number / STIFFJcollection.Count * 100, 0) & "%")
    Dim el() As String
    el() = Split(element.detail, "-")
    If el(2) = "S" Then
        el(2) = "STIFFENER"
    ElseIf el(2) = "G" Then
        el(2) = "GUSSET PLATE"
    End If
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
    If element.sizeA = "" Then GoTo skipitbecauseSIS
    .Cells(lrow, firstcolumn) = element.sizeA & "x" & element.sizeB
    .Cells(lrow, secondcolumn) = baseSTIFFj.STIFFJ_BASEcollection(element.detail).sizeA & "x" & baseSTIFFj.STIFFJ_BASEcollection(element.detail).sizeB
    If element.STATUSsize = "FAIL" Then
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
        
skipitbecauseSIS:
    End If
    If .Cells(lrow, secondcolumn) = "" Then
        .Cells(lrow, secondcolumn).Interior.ColorIndex = 22
        indicator = indicator + 1
    End If
firstcolumn = 8
secondcolumn = 9
    .Cells(lrow, firstcolumn) = element.thickness
    .Cells(lrow, secondcolumn) = baseSTIFFj.STIFFJ_BASEcollection(element.detail).thickness
    If element.STATUSthickness = "FAIL" Then
        .Range(Cells(lrow, firstcolumn), Cells(lrow, secondcolumn)).Interior.ColorIndex = 44
        .Range(Cells(lrow, firstcolumn), Cells(lrow, secondcolumn)).Font.ColorIndex = 3
        .Range(Cells(lrow, firstcolumn), Cells(lrow, secondcolumn)).Font.Bold = True
    End If
    If .Cells(lrow, secondcolumn) = "" Then
        .Cells(lrow, secondcolumn).Interior.ColorIndex = 22
        indicator = indicator + 1
    End If
firstcolumn = 10
secondcolumn = 11
    .Cells(lrow, firstcolumn) = replaceMinus6(element.material)
    .Cells(lrow, secondcolumn) = baseSTIFFj.STIFFJ_BASEcollection(element.detail).material
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
firstcolumn = 12
secondcolumn = 13
    .Cells(lrow, firstcolumn) = element.toughness
    .Cells(lrow, secondcolumn) = baseSTIFFj.STIFFJ_BASEcollection(element.detail).toughness
    If element.STATUStoughness = "FAIL" Then
        .Range(Cells(lrow, firstcolumn), Cells(lrow, secondcolumn)).Interior.ColorIndex = 44
        .Range(Cells(lrow, firstcolumn), Cells(lrow, secondcolumn)).Font.ColorIndex = 3
        .Range(Cells(lrow, firstcolumn), Cells(lrow, secondcolumn)).Font.Bold = True
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
