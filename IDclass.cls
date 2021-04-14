VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private p_IDcollection As Collection
Private p_SampleField As IDfields
Private detailNumber, IDnumber, guide, JOINTtype, IDtype, BEVELtype, IDsize, GRINDING, IDANGLE, NDT, Wbookletnumber As Integer


'Initialize this class with it's collection:
Private Sub Class_Initialize()
    Set p_IDcollection = New Collection
End Sub

'Allow for adding SampleIDfields:
Public Sub AddField(field As IDfields, Optional key As String)
    Set p_SampleField = field
    If key <> "" Then p_IDcollection.Add field, key Else p_IDcollection.Add field
End Sub

'Expose the collection:
Public Property Get IDcollection() As Collection
    Set IDcollection = p_IDcollection
End Property

Sub ghgh()
Dim ID As New IDfields
ID
End Sub
Public Function getFAILquantity()
Dim element As Variant
Dim FAILquantity As Integer
For Each element In IDcollection
    If element.STATUSIDcode = "NO" Then
        FAILquantity = FAILquantity + 1
    End If
Next element
getFAILquantity = FAILquantity
End Function

Public Function getresult()
On Error Resume Next
Dim i As Integer
For i = 1 To 6
    If ThisWorkbook.Sheets("resultID-" & i) Is Nothing Then
        ThisWorkbook.Sheets.Add(After:=Sheets(ThisWorkbook.Sheets.Count)).name = "resultID-" & i
        Exit For
    End If
Next i
On Error GoTo 0
ThisWorkbook.Sheets("resultID-" & i).Activate
newsheet = ActiveSheet.name
toprow = 2
firstcol = 2
If toprow <> 0 And firstcol <> 0 Then
    IDtoptablefill
    IDcontenttablefill
    createbuttonmain
    createbuttonup
    highlight
End If

ThisWorkbook.Sheets("mainVIEW").Activate
End Function
Sub IDtoptablefill()
Dim IDtabletop As Range
Set IDtabletop = Cells(toprow, firstcol)
With IDtabletop
Wbookletnumber = 11
    .Cells(1, 1) = "detail NUMBER"
    .Cells(1, 2) = "Profile"
    .Cells(1, 3) = "Material"
    .Cells(1, 4) = "Toughness"
    .Cells(1, 5) = "ID code"
End With
End Sub
Sub IDcontenttablefill()
Dim lrow As Integer
Dim firstcolumn, secondcolumnumn As Integer
Dim element As IDfields
lrow = IDcollection.Count
On Error Resume Next
Dim arr() As String
ReDim arr(1 To lrow, 1 To 5)
rown = 1
On Error Resume Next
For Each element In IDcollection
    bar ("creating a ID table  " & Round(rown / IDcollection.Count / 2 * 100, 0) & "%")
    arr(rown, 1) = element.detail
    arr(rown, 2) = element.profile
    arr(rown, 3) = element.material
    arr(rown, 4) = element.toughness
    arr(rown, 5) = element.IDcode
    rown = rown + 1
Next
With ActiveSheet
    Dim IDtable As Range
    Set IDtable = Cells(toprow + 1, firstcol)
    Set IDtable = IDtable.Resize(lrow, 5)
    IDtable = arr
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
    bar ("highlighting a ID table  " & Round(i / rng.Rows.Count / 2 * 100, 0) + 50 & "%")
    If IDcollection(rng.Cells(i, 1)).STATUSIDcode = "NO" Then
        'Cells(i + toprow, 4 + firstcol).Interior.ColorIndex = 22
        With ThisWorkbook.ActiveSheet.Cells(i + toprow, 5 + firstcol - 1)
            .Interior.ColorIndex = 44
            .Font.ColorIndex = 3
            .Font.Bold = True
        End With
    End If
Next i
End Function


