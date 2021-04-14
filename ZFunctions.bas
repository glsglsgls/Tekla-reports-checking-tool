Attribute VB_Name = "ZFunctions"
Option Explicit
Function LastRow() As Integer
    Dim rf As Range
    Set rf = ActiveSheet.UsedRange.Find("*", , xlValues, xlWhole, , xlPrevious)
    If Not rf Is Nothing Then LastRow = rf.row Else LastRow = 1
End Function
Function lastcol() As Integer
    Dim rf As Range
    Set rf = ActiveSheet.UsedRange.Find("*", , xlValues, xlWhole, , xlPrevious)
    If Not rf Is Nothing Then lastcol = rf.Column Else lastcol = 1
End Function

Function replaceMinus6(value As String) As String
If value Like "????-?" Then value = Left(value, 4)
value = Replace(value, "C", "C")
value = Replace(value, "Á", "")
value = Replace(value, "KCV-40", "")
value = Replace(value, "-", "")
replaceMinus6 = value
End Function
Sub prepairing(condition As Boolean)
Application.DisplayAlerts = condition
ActiveWorkbook.ActiveSheet.DisplayPageBreaks = condition
Application.AskToUpdateLinks = condition
Application.EnableEvents = condition
ActiveSheet.DisplayPageBreaks = False
If condition = False Then
    Application.ScreenUpdating = False
    Application.Cursor = xlWait
    Application.Calculation = xlCalculationManual
    ThisWorkbook.Sheets("mainVIEW").Unprotect
Else
    Application.Cursor = xlDefault
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    Application.ScreenUpdating = True
    ThisWorkbook.Sheets("mainVIEW").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End If
End Sub
Function prepOFF()
prepairing (False)
End Function
Function prepON()
prepairing (True)
End Function
Function lcol(i, j) As Integer
lcol = Cells(i, j).End(xlToRight).Column
End Function
Function cTN(value As Variant)
If IsNumeric(value) = False Then cTN = 0 Else cTN = value
End Function
Function ifblank(value As Variant)
If value = "" Or value = "-" Or value = " " Then
    ifblank = "blank"
Else
    ifblank = value
End If
End Function
Function IsActiveSheet(ByVal targetSheet As Worksheet) As Boolean
    IsActiveSheet = targetSheet.name = ActiveSheet.name And _
            targetSheet.Parent.name = ActiveWorkbook.name
End Function
Function buttonquantity() As Integer
Dim but As Variant
Dim quantity As Integer
For Each but In ThisWorkbook.Sheets("mainVIEW").Shapes
    quantity = quantity + 1
Next
buttonquantity = quantity
End Function
Function bar(value As String)
If value <> "" Then Application.StatusBar = value Else Application.StatusBar = False
End Function
Function ifcyrilic(value As String)
Dim res As String
res = value
res = Replace(res, "ÄÂÓÒAÂÐ", "")
res = Replace(res, "Á", "B")
res = Replace(res, "Ê", "K")
res = Replace(res, "Ø", "Sh")
res = Replace(res, "_", "")
ifcyrilic = res
End Function
Function createbuttonmain()
    Dim btn As button
    Set btn = ActiveSheet.buttons.Add(0, 16, 48, 43)
    btn.OnAction = "buttonmaingo"
    btn.Caption = "MAIN sheet"
    Cells(40, 1).Select
End Function
Function changelogbutton()
Dim btn As button
Set btn = ThisWorkbook.Sheets("mainVIEW").buttons.Add(50, 12, 60, 25)
With btn
    .OnAction = "showchangelog"
    .Caption = "changelog"
End With
End Function
Sub clearmainview()
ZFunctions.prepOFF
ThisWorkbook.Sheets("mainVIEW").Activate
ThisWorkbook.Sheets("mainVIEW").Cells.Clear
ThisWorkbook.Sheets("mainVIEW").Cells.Interior.Color = RGB(242, 242, 242)
Dim IShape As Shape
For Each IShape In Sheets("mainVIEW").Shapes
    If Right(IShape.OnAction, 5) <> "start" Then IShape.Delete
Next
changelogbutton
ZFunctions.prepON
End Sub
Sub showchangelog()
UserForm1.Show 0
End Sub
Function createbuttonup()
    Dim btn As button
    Dim lrow As Integer
    lrow = LastRow
    If lrow < 40 Then Exit Function
    Set btn = ActiveSheet.buttons.Add(0, lrow * 15 - 35, 48, 45)
    btn.OnAction = "buttonupgo"
    btn.Caption = "go up"
    Cells(40, 1).Select
End Function
Function writeFAILquantity(quantity As String)
With ThisWorkbook.Sheets("mainview")
    If quantity > 1 Then
        .Cells(lastrowinmainsheet(2), 3).value = quantity & " mismatches were found"
    ElseIf quantity = 1 Then
        .Cells(lastrowinmainsheet(2), 3).value = quantity & " mismatch was found"
    ElseIf quantity = 0 Then
        .Cells(lastrowinmainsheet(2), 3).value = "No mismatch"
    End If
End With
End Function
Function writeMANUALquantity(quantity As String)
With ThisWorkbook.Sheets("mainview")
    If quantity <> 0 Then .Cells(lastrowinmainsheet(3) + 1, 3).value = quantity & " elements for manual check" 'to be checked manually"
End With
End Function
Function lastrowinmainsheet(col As Integer) As Integer
lastrowinmainsheet = ThisWorkbook.Sheets("mainVIEW").Cells(Rows.Count, col).End(xlUp).row
End Function
Function borders(row As Integer)
Dim rg As Range
Set rg = Cells(row + 1, 3) 'Range(Cells(row, 2), Cells(row + 1, 2))
rg.borders.LineStyle = xlDouble
rg.borders.ColorIndex = 16
rg.borders(xlInsideVertical).LineStyle = xlNone
rg.borders(xlInsideHorizontal).LineStyle = xlNone
rg.borders(xlEdgeRight).LineStyle = xlNone
rg.borders(xlEdgeTop).LineStyle = xlNone
rg.borders(xlEdgeLeft).LineStyle = xlNone
Cells(row + 1, 2).borders(xlDiagonalDown).LineStyle = xlDouble
End Function
Function setrowheight(name As String)
With Sheets("mainVIEW")
.Rows(lastrowinmainsheet(2) + 2).RowHeight = firstrowheight
.Rows(lastrowinmainsheet(2) + 3).RowHeight = secondrowheight
ActiveSheet.Shapes.AddShape(msoShapeTrapezoid, 97, 110 + (buttonquantity - 1) * buttonheight, 180, 20).Select
    Selection.Placement = xlFreeFloating
    Selection.ShapeRange.ShapeStyle = msoShapeStylePreset62
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Go to"
    Selection.OnAction = "button" & name & "go" & indexsheet
    Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
    Selection.ShapeRange.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
End With
End Function

Function writemybook(val As String)
mybook = Right(val, Len(val) - 16)
mybook = Replace(mybook, ".xlsx", "") & "   "
mybook = Replace(mybook, ".xls", "")
mybook = Replace(mybook, "_", " ")
mybook = Replace(mybook, "profile", "report")
If mybook = "DENT   " Then mybook = "IDENT code report"
End Function

Function highlightredchar(lrow, firstcolumn, secondcolumn)
Dim a() As String
Dim b() As String
a() = Split(.Cells(lrow, firstcolumn), "x")
b() = Split(.Cells(lrow, secondcolumn), "x")
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
If UBound(a()) = 1 Then
    If a(0) <> b(0) And a(1) = b(0) And a(1) <> b(1) Then
        For k = 1 To Len(.Cells(lrow, firstcolumn))
            If Mid$(.Cells(lrow, firstcolumn).value, k, Len(a(1))) = a(1) Then
                .Cells(lrow, firstcolumn).Characters(start:=k, length:=Len(a(1))).Font.ColorIndex = 1
                .Cells(lrow, firstcolumn).Characters(start:=k, length:=Len(a(1))).Font.Bold = False
            End If
        Next k
        For k = 1 To Len(.Cells(lrow, secondcolumn))
            If Mid$(.Cells(lrow, secondcolumn).value, k, Len(b(0))) = b(0) Then
                .Cells(lrow, secondcolumn).Characters(start:=k, length:=Len(b(0))).Font.ColorIndex = 1
                .Cells(lrow, secondcolumn).Characters(start:=k, length:=Len(b(0))).Font.Bold = False
            End If
        Next k
    End If
    If a(1) <> b(1) And a(0) = b(1) And a(0) <> b(0) Then
        For k = 1 To Len(.Cells(lrow, firstcolumn))
            If Mid$(.Cells(lrow, firstcolumn).value, k, Len(a(0))) = a(0) Then
                .Cells(lrow, firstcolumn).Characters(start:=k, length:=Len(a(0))).Font.ColorIndex = 1
                .Cells(lrow, firstcolumn).Characters(start:=k, length:=Len(a(0))).Font.Bold = False
            End If
        Next k
        For k = 1 To Len(.Cells(lrow, secondcolumn))
            If Mid$(.Cells(lrow, secondcolumn).value, k, Len(b(1))) = b(1) Then
                .Cells(lrow, secondcolumn).Characters(start:=k, length:=Len(b(1))).Font.ColorIndex = 1
                .Cells(lrow, secondcolumn).Characters(start:=k, length:=Len(b(1))).Font.Bold = False
            End If
        Next k
    End If
End If
End Function









