Attribute VB_Name = "Methods"
Option Explicit
Dim f As Object

Sub start()
    mainFORM.Show
End Sub
Sub clearalltextboxes()
    Dim tb As Control
    On Error Resume Next
    For Each tb In mainFORM.Controls
        tb.Text = ""
    Next
    On Error GoTo 0
End Sub
Sub clearselection()
    Set f = Nothing
End Sub
Public Sub choosebook()
    path = ""
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .AllowMultiSelect = True
        .Filters.Add "Excel Files", "*.xlsx", 1
        .Filters.Add "Excel 97 Files", "*.xls", 2
        '.Filters.Add "Excel files", "*.xlsx*;*.xls*", 1
        '.Filters.Add "Excel files", "*.xls*;*.xlsx*", 1
        .Show
        On Error Resume Next
        path = .SelectedItems.item(1)
        On Error GoTo 0
        If InStr(path, ".xlsx") = 0 Then
            ThisWorkbook.Sheets("mainVIEW").Activate
            End
        End If
    End With
End Sub
Public Sub openbook(path As String)
    Set pathbook = Workbooks.Open(path, ReadOnly:=False)
End Sub
Public Sub insertFILESinTEXTBOXES()
    Methods.clearalltextboxes
    Dim j As Integer
    Set f = Application.FileDialog(3)
    With f
        .Filters.Add "Excel files", "*.xls*;*.xlsx*", 1
        .Show
        Dim varFile As Variant
        For Each varFile In .SelectedItems
            Dim a() As String
            a() = Split(varFile, "\")
                If Left(a(UBound(a())), 9) = "AWP1_1_11" Then
                    With mainFORM.TBPG
                        If .value = "" Then .value = a(UBound(a())) Else .value = .value & Chr(10) & a(UBound(a()))
                    End With
                    
                ElseIf Left(a(UBound(a())), 9) = "AWP1_1_12" Then
                    With mainFORM.TBBRACE
                        If .value = "" Then .value = a(UBound(a())) Else .value = .value & Chr(10) & a(UBound(a()))
                    End With
                    
                ElseIf Left(a(UBound(a())), 9) = "AWP1_1_13" Then
                    With mainFORM.TBSTIFFSP
                        If .value = "" Then .value = a(UBound(a())) Else .value = .value & Chr(10) & a(UBound(a()))
                    End With
                    
                ElseIf Left(a(UBound(a())), 9) = "AWP1_1_14" Then
                    With mainFORM.TBSTIFFJ
                        If .value = "" Then .value = a(UBound(a())) Else .value = .value & Chr(10) & a(UBound(a()))
                    End With
                    
                ElseIf Left(a(UBound(a())), 9) = "AWP1_1_15" Then
                    With mainFORM.TBNODE
                        If .value = "" Then .value = a(UBound(a())) Else .value = .value & Chr(10) & a(UBound(a()))
                    End With
                    
                ElseIf Left(a(UBound(a())), 9) = "AWP1_1_17" Then
                    With mainFORM.TBTRANS
                        If .value = "" Then .value = a(UBound(a())) Else .value = .value & Chr(10) & a(UBound(a()))
                    End With
                    
                ElseIf Left(a(UBound(a())), 9) = "AWP1_1_18" Then
                    With mainFORM.TBHnode
                        If .value = "" Then .value = a(UBound(a())) Else .value = .value & Chr(10) & a(UBound(a()))
                    End With
                
                'ATTENTION!!! fillet weld will be shown with other welding reports, but it has its own class!!!
                ElseIf Left(a(UBound(a())), 9) = "AWP1_1_19" Or Left(a(UBound(a())), 9) = "AWP1_1_20" Or _
                Left(a(UBound(a())), 9) = "AWP1_1_21" Or Left(a(UBound(a())), 9) = "AWP1_2_13" Or Left(a(UBound(a())), 9) = "AWP1_0_1_" Then
                    With mainFORM.TBWELD
                        If .value = "" Then .value = a(UBound(a())) Else .value = .value & Chr(10) & a(UBound(a()))
                    End With
                    
                ElseIf Left(a(UBound(a())), 9) = "AWP1_2_11" Then
                    With mainFORM.TBSECONDARYFR
                        If .value = "" Then .value = a(UBound(a())) Else .value = .value & Chr(10) & a(UBound(a()))
                    End With
                ElseIf Left(a(UBound(a())), 9) = "AWP1_0_2_" Then
                    With mainFORM.TBIDCD
                        If .value = "" Then .value = a(UBound(a())) Else .value = .value & Chr(10) & a(UBound(a()))
                    End With
                End If
       Next
    End With
End Sub
Sub openandbindworkbooks()

If f Is Nothing Then
    MsgBox "No reports selected", vbInformation, "Information message"
    Exit Sub
End If
If mainFORM.ComboBox1.value = "" Then
    MsgBox "Please choose the module", vbInformation, "Information message"
    Exit Sub
End If

ZFunctions.prepOFF
basesheet = mainFORM.ComboBox1.value
mainFORM.Hide
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    If Left(ws.name, 6) = "result" Then ws.Delete
Next
Dim j As Integer
ThisWorkbook.Sheets("mainVIEW").Activate
ThisWorkbook.Sheets("mainVIEW").Cells.Clear
ThisWorkbook.Sheets("mainVIEW").Cells.Interior.Color = RGB(242, 242, 242)
Dim IShape As Shape
For Each IShape In Sheets("mainVIEW").Shapes
    IShape.Delete
Next
Dim myrow As Variant
For Each myrow In ThisWorkbook.Sheets("mainVIEW").Rows("1:50")
    If myrow.row Mod 2 = 0 Then myrow.RowHeight = 25 Else myrow.RowHeight = 20
Next
ThisWorkbook.Sheets("mainVIEW").Rows(1).RowHeight = 43
firstrowheight = 25
secondrowheight = 20
buttonheight = firstrowheight + secondrowheight - 1

ThisWorkbook.Sheets("mainVIEW").buttons.Add(75, 12, 200, 25).Select
With Selection
    .OnAction = "start"
    .Caption = "PERFORM A NEW CHECK"
End With

With ThisWorkbook.Sheets("mainVIEW").Cells(2, 2)
    .value = "Checking is finished"
    .Font.Size = 24
End With

    Dim varFile As Variant
    Dim number As Integer
    number = 1
    For Each varFile In f.SelectedItems
        bar ("Progress  " & Round(number / f.SelectedItems.Count / 5 * 100, 0) * 5 & "%")
        Dim a() As String
        a() = Split(varFile, "\")
            If Left(a(UBound(a())), 9) = "AWP1_1_11" Then
                openbook (varFile)
                Set PGbook = pathbook
                writemybook (a(UBound(a())))
                Call compare.main
                Set PGbook = Nothing

            ElseIf Left(a(UBound(a())), 9) = "AWP1_1_12" Then
                openbook (varFile)
                Set BRACEbook = pathbook
                writemybook (a(UBound(a())))
                Call compare.main
                Set BRACEbook = Nothing

            ElseIf Left(a(UBound(a())), 9) = "AWP1_1_13" Then
                openbook (varFile)
                Set STIFFSPbook = pathbook
                writemybook (a(UBound(a())))
                Call compare.main
                Set STIFFSPbook = Nothing

            ElseIf Left(a(UBound(a())), 9) = "AWP1_1_14" Then
                openbook (varFile)
                Set STIFFJbook = pathbook
                writemybook (a(UBound(a())))
                Call compare.main
                Set STIFFJbook = Nothing

            ElseIf Left(a(UBound(a())), 9) = "AWP1_1_15" Then
                openbook (varFile)
                Set NODEbook = pathbook
                writemybook (a(UBound(a())))
                Call compare.main
                Set NODEbook = Nothing

            ElseIf Left(a(UBound(a())), 9) = "AWP1_1_17" Then
                openbook (varFile)
                Set TRANSbook = pathbook
                writemybook (a(UBound(a())))
                Call compare.main
                Set TRANSbook = Nothing

            ElseIf Left(a(UBound(a())), 9) = "AWP1_1_18" Then
                openbook (varFile)
                Set HNODEbook = pathbook
                writemybook (a(UBound(a())))
                Call compare.main
                Set HNODEbook = Nothing

            ElseIf Left(a(UBound(a())), 9) = "AWP1_1_19" Or Left(a(UBound(a())), 9) = "AWP1_1_20" Or _
                Left(a(UBound(a())), 9) = "AWP1_1_21" Or Left(a(UBound(a())), 9) = "AWP1_2_13" Then
                openbook (varFile)
                Set WELDbook = pathbook
                writemybook (a(UBound(a())))
                Call compare.main
                Set WELDbook = Nothing

            ElseIf Left(a(UBound(a())), 9) = "AWP1_2_11" Then
                openbook (varFile)
                Set SECONDARYFRbook = pathbook
                writemybook (a(UBound(a())))
                Call compare.main
                Set SECONDARYFRbook = Nothing
            
            ElseIf Left(a(UBound(a())), 9) = "AWP1_0_2_" Then
                openbook (varFile)
                Set IDbook = pathbook
                writemybook (a(UBound(a())))
                Call compare.main
                Set IDbook = Nothing
                
            ElseIf Left(a(UBound(a())), 9) = "AWP1_0_1_" Then
                openbook (varFile)
                Set WELDFILLETbook = pathbook
                writemybook (a(UBound(a())))
                Call compare.main
                Set WELDFILLETbook = Nothing
                
            End If
        number = number + 1
    Next
preparations
ZFunctions.prepON
End
End Sub
Sub preparations()
    Dim ws As Worksheet
    Dim lrow, lcol, firstcol, firstrow, row, guidecolumn As Integer
    Dim blue As Long
    blue = RGB(115, 154, 190)
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.name, 6) = "result" Then
            Dim rng As Range
            Dim rowrange As Range
            Set rng = Sheets(ws.name).UsedRange.Find("*", , xlValues, xlWhole, , xlPrevious)
            lrow = rng.row
            lcol = rng.Column
            firstrow = ws.Cells(1, lcol).End(xlDown).row
            firstcol = ws.Cells(lrow, 1).End(xlToRight).Column
            lcol = ws.Cells(firstrow, Columns.Count).End(xlToLeft).Column
            Dim imei As Integer
            Set rng = ws.Cells(firstrow + 1, firstcol)
            Set rng = rng.Resize(lrow - firstrow, lcol - 1)
            Set rowrange = ws.Cells(firstrow, firstcol)
            Set rowrange = rowrange.Resize(, lcol)
            Dim cell As Range
            For Each cell In rowrange
                If cell.value = "GUIDE" Then
                    guidecolumn = cell.Column
                    Exit For
                ElseIf cell.value = "Assembly pos." Then
                    guidecolumn = cell.Column
                    Exit For
                Else
                    guidecolumn = 5
                End If
            Next
            For Each cell In rng
                If cell.Interior.Pattern = xlNone And cell.row Mod 2 = 0 Then cell.Interior.Color = RGB(240, 245, 250)
                If cell.value = "blank" Then cell.Font.Color = RGB(208, 206, 206)
                                
'!!!!!here you can turn off the "manual check" highlighting. just remove the symbol "'" from the row below, and add this symbol before the next row
                'If cell.value = "blablabla" Then
                If cell.value = "Unique node" Or cell.Interior.ColorIndex = 22 Then
                    With rng.Rows(cell.row - firstrow)
                        .Font.Bold = True
                        .Interior.Color = RGB(208, 206, 206)
                        .Font.Color = RGB(0, 32, 96)
                    End With
                    Set rowrange = rng.Cells(cell.row - firstrow, guidecolumn - firstcol + 2)
                    Set rowrange = rowrange.Resize(, rng.Columns.Count - (guidecolumn - firstcol + 1))
                    With rowrange
                        .Clear
                        .Merge
                        .Interior.Color = RGB(208, 206, 206)
                        .Font.Color = RGB(0, 32, 96)
                        .Font.Bold = True
                        .value = "MANUAL CHECK"
                        .HorizontalAlignment = xlCenter
                        .RowHeight = 25
                        .VerticalAlignment = xlCenter
                    End With
                End If
            Next
            Set rng = ws.Cells(firstrow, firstcol)
            Set rng = rng.Resize(, lcol - 1)
            Dim columnrng As Range
            rng.Rows.Interior.Color = RGB(155, 194, 230)
            rng.Rows.borders(xlEdgeBottom).LineStyle = xlNone
            For Each cell In rng
                If Left(cell.value, 5) = "Tekla" And Left(cell(1, 2), 2) = "KM" Then
                    Set columnrng = ws.Cells(firstrow + 1, cell.Column)
                    Set columnrng = columnrng.Resize(lrow - firstrow, 2)
                    columnrng.borders(xlEdgeLeft).Weight = xlMedium
                    columnrng.borders(xlEdgeLeft).Color = blue 'Index = 16
                    columnrng.borders(xlEdgeRight).Weight = xlMedium
                    columnrng.borders(xlEdgeRight).Color = blue 'Index = 16
                End If
                If cell.value = "GUIDE" Then
                    ws.Columns(cell.Column).ColumnWidth = 9
                    Set columnrng = ws.Cells(firstrow + 1, cell.Column)
                    Set columnrng = columnrng.Resize(lrow - firstrow)
                    columnrng.Font.Color = RGB(142, 169, 212)
                End If
            Next
        End If
        
    Next
        With Range(Cells(4, 3), Cells(LastRow + 2, 3))
        .Interior.ColorIndex = 2
        .Select
        .borders.Weight = xlMedium
        .borders.ColorIndex = 16
        .borders(xlEdgeLeft).ColorIndex = 2
        .borders(xlEdgeTop).ColorIndex = 2
        .borders(xlInsideHorizontal).LineStyle = xlLineStyleNone
    End With
    With Range(Cells(4, 2), Cells(LastRow + 1, 2))
        Cells.VerticalAlignment = xlCenter
        Cells.HorizontalAlignment = xlCenter
        Cells(2, 2).HorizontalAlignment = xlLeft
    End With
    Columns(2).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    For Each cell In Range(Cells(4, 4), Cells(LastRow, 4))
        If cell.row Mod 2 = 0 Then
            cell.VerticalAlignment = xlBottom
            borders (cell.row)
        Else
            cell.VerticalAlignment = xlTop
        End If
    Next
    Columns(2).EntireColumn.ColumnWidth = 3
    Columns(3).ColumnWidth = 35
    'Columns(3).HorizontalAlignment = xlRight
    Dim val As Variant
    For Each val In Range(Cells(3, 3), Cells(LastRow + 5, 3))
        If Len(val.value) > 38 Then val.HorizontalAlignment = xlRight Else val.HorizontalAlignment = xlCenter
    Next
    'Columns(3).EntireColumn.AutoFit
    'If Columns(3).ColumnWidth < 35 Then Columns(3).ColumnWidth = 35
    Columns(4).EntireColumn.AutoFit
    If Columns(4).ColumnWidth < 30 Then Columns(4).ColumnWidth = 30
    Rows(LastRow + 2).RowHeight = 17
    Range("C2:D2").Merge
    Range("C2:D2").HorizontalAlignment = xlCenter
    Cells(40, 1).Select
End Sub






Sub addcode(namesub, sheetname As String)
    Dim subtext, subtop As String
    Dim i, ind As Integer
    subtext = "Sub " & "button" & namesub & vbCrLf & "thisworkbook.sheets(" & sheetname & ").activate" & vbCrLf & "End Sub"
    subtop = "Sub " & "button" & namesub
    With Workbooks(ThisWorkbook.name).VBProject.VBComponents("Module1").CodeModule
        For i = .CountOfLines To 1 Step -1
            If .Lines(i, 1) = subtop Then ind = ind + 1
        Next i
        If ind = 0 Then .InsertLines .CountOfLines + 1, subtext
    End With
End Sub
Function ClearModule()
Dim start As Long
Dim Lines As Long
Dim i As Variant, a As Variant
    With Workbooks(ThisWorkbook.name).VBProject.VBComponents("Module1").CodeModule
        For i = .CountOfLines To 1 Step -1
            If Left(.Lines(i, 1), 10) = "Sub button" Then
                Application.DisplayAlerts = False
                .DeleteLines i, 1
                Application.DisplayAlerts = True
            End If
        Next
    End With
End Function
