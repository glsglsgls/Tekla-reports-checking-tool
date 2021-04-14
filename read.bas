Attribute VB_Name = "read"
Option Explicit
Sub basePGread()
ThisWorkbook.Sheets(basesheet).Activate
Dim lrow, i As Integer
ThisWorkbook.Activate
lrow = Range("A1").End(xlDown).row
Dim basePGinput As Range
Set basePGinput = Range("A2:E" & lrow)
Set basePG = New basePGclass
On Error Resume Next
For i = 1 To basePGinput.Rows.Count
    Dim field As New PGfields
    field.item = basePGinput.Cells(i, 1).value
    field.designation = basePGinput.Cells(i, 2).value
    field.material = basePGinput.Cells(i, 3).value
    field.toughness = basePGinput.Cells(i, 4).value
    basePG.AddField field, field.item
    Set field = Nothing
Next i
On Error GoTo 0
End Sub
Sub baseSTIFFSPread()
ThisWorkbook.Sheets(basesheet).Activate
Dim lrow, i As Integer
ThisWorkbook.Activate
lrow = Range("F1").End(xlDown).row
Dim baseSTIFFSPinput As Range
Set baseSTIFFSPinput = Range("F2:M" & lrow)
Set baseSTIFFsp = New baseSTIFFspclass
On Error Resume Next
For i = 1 To baseSTIFFSPinput.Rows.Count
    Dim field As New STIFFfields
    field.detail = baseSTIFFSPinput.Cells(i, 1).value & "-" & baseSTIFFSPinput.Cells(i, 2).value
    field.quantity = baseSTIFFSPinput.Cells(i, 4).value
    Dim Size() As String
    Size() = Split(baseSTIFFSPinput.Cells(i, 5).value, "x")
    field.sizeA = Size(0)
    field.sizeB = Size(1)
    field.thickness = baseSTIFFSPinput.Cells(i, 6).value
    field.material = baseSTIFFSPinput.Cells(i, 7).value
    field.toughness = baseSTIFFSPinput.Cells(i, 8).value
    baseSTIFFsp.AddField field, field.detail
    Set field = Nothing
Next i
On Error GoTo 0
End Sub
Sub baseSTIFFJread()
ThisWorkbook.Sheets(basesheet).Activate
Dim lrow, lcolLEV, i As Integer
Dim baseLEVELinput As Range
ThisWorkbook.Activate
lrow = Range("O1").End(xlDown).row
Dim baseSTIFFJinput As Range
Set baseSTIFFJinput = Range("O2:V" & lrow)
lcolLEV = Range("Y2").End(xlToRight).Column
Set baseLEVELinput = Range(Cells(1, 25), Cells(3, lcolLEV))
Set baseSTIFFj = New baseSTIFFJclass
On Error Resume Next
For i = 1 To baseSTIFFJinput.Rows.Count
    Dim field As New STIFFfields
    field.detail = baseSTIFFJinput.Cells(i, 1).value & "-" & baseSTIFFJinput.Cells(i, 2).value
    field.sizeA = baseSTIFFJinput.Cells(i, 3).value
    field.sizeB = baseSTIFFJinput.Cells(i, 4).value
    field.thickness = baseSTIFFJinput.Cells(i, 5).value
    field.material = baseSTIFFJinput.Cells(i, 7).value
    field.toughness = baseSTIFFJinput.Cells(i, 8).value
    field.levelDECKA = baseLEVELinput.Cells(2, 1).value
    Dim col As Variant
    For Each col In baseLEVELinput
        If Right(baseSTIFFJinput.Cells(i, 2).value, 1) = baseLEVELinput.Cells(1, Asc(col) - 64) Then
            field.level = baseLEVELinput.Cells(2, Asc(col) - 64).value - field.levelDECKA
            Exit For
        End If
    Next col
    baseSTIFFj.AddField field, field.detail
    Set field = Nothing
Next i
On Error GoTo 0
End Sub
Sub baseNODEread()
ThisWorkbook.Sheets(basesheet).Activate
Dim lrow, lcolLEV, i As Integer
ThisWorkbook.Activate
lrow = Range("AJ1").End(xlDown).row
Dim baseNODEinput As Range
Set baseNODEinput = Range("AJ2:AP" & lrow)
Set baseNODE = New baseNODEclass
On Error Resume Next
For i = 1 To baseNODEinput.Rows.Count
    Dim field As New NODEfields
    field.detail = baseNODEinput.Cells(i, 1).value & "-" & baseNODEinput.Cells(i, 2).value
    Dim a() As String
    If i = 244 Then
    End If
    a() = Split(baseNODEinput.Cells(i, 3).value, " / ")
        field.AsizeA = ifblank(a(0))
        If UBound(a()) = 1 Then
            field.AsizeB = ifblank(a(1))
        Else
            field.AsizeB = "blank"
        End If
    a() = Split(baseNODEinput.Cells(i, 4).value, " / ")
    If UBound(a()) <> -1 Then
        field.BsizeA = ifblank(a(0))
        If UBound(a()) = 1 Then
            field.BsizeB = ifblank(a(1))
        Else
            field.BsizeB = ifblank("blank")
        End If
    Else
        field.BsizeA = "blank"
        field.BsizeB = "blank"
    End If
    a() = Split(baseNODEinput.Cells(i, 5).value, " / ")
        field.Athickness = ifblank(a(0))
        If UBound(a()) = 1 Then field.Bthickness = ifblank(a(1)) Else field.Bthickness = "blank"
    a() = Split(baseNODEinput.Cells(i, 6).value, " / ")
        field.Atoughness = ifblank(a(0))
        If UBound(a()) = 1 Then field.Btoughness = ifblank(a(1)) Else field.Btoughness = "blank"
    field.material = baseNODEinput.Cells(i, 7).value
    baseNODE.AddField field, field.detail
    Set field = Nothing
Next i
On Error GoTo 0
End Sub
Sub baseTRANSread()
ThisWorkbook.Sheets(basesheet).Activate
Dim lrow, i As Integer
ThisWorkbook.Activate
lrow = Range("AR1").End(xlDown).row
Dim baseTRANSinput As Range
Set baseTRANSinput = Range("AR2:BA" & lrow)
Set baseTRANS = New baseTRANSclass
On Error Resume Next
For i = 1 To baseTRANSinput.Rows.Count
    Dim field As New TRANSfields
    field.detail = baseTRANSinput.Cells(i, 1).value & "-" & baseTRANSinput.Cells(i, 2).value & "-" & Left(baseTRANSinput.Cells(i, 3).value, 1)
    field.length = baseTRANSinput.Cells(i, 5).value + baseTRANSinput.Cells(i, 7).value
    field.width = baseTRANSinput.Cells(i, 6).value
    field.thickness = baseTRANSinput.Cells(i, 8).value
    field.material = baseTRANSinput.Cells(i, 9).value
    field.toughness = baseTRANSinput.Cells(i, 10).value
    baseTRANS.AddField field, field.detail
    Set field = Nothing
Next i
On Error GoTo 0
End Sub
Sub baseHNODEread()
ThisWorkbook.Sheets(basesheet).Activate
Dim lrow, lcolLEV, i As Integer
ThisWorkbook.Activate
lrow = Range("BC1").End(xlDown).row
Dim baseHNODEinput As Range
Set baseHNODEinput = Range("BC2:BH" & lrow)
Set baseHNODE = New baseHNODEclass
On Error Resume Next
For i = 1 To baseHNODEinput.Rows.Count
    Dim field As New HNODEfields
    field.detail = baseHNODEinput.Cells(i, 1).value & "-" & baseHNODEinput.Cells(i, 2).value
    Dim a() As String
    a() = Split(baseHNODEinput.Cells(i, 3).value, "x")
        field.sizeA = ifblank(a(0))
        If UBound(a()) = 1 Then
            field.sizeB = ifblank(a(1))
        Else
            field.sizeB = "blank"
        End If
        field.thickness = baseHNODEinput.Cells(i, 4).value
        field.material = baseHNODEinput.Cells(i, 5).value
        field.toughness = baseHNODEinput.Cells(i, 6).value
    baseHNODE.AddField field, field.detail
    Set field = Nothing
Next i
On Error GoTo 0
End Sub


Sub PGread()
Dim lrow, lcol, i As Integer
PGbook.Activate
lrow = LastRow
lcol = lastcol
Dim PGinput As Range
Set PGinput = Range("A3:F" & lrow)
Set PG = New PGclass
On Error Resume Next
For i = 1 To PGinput.Rows.Count
    Dim field As New PGfields
    field.item = PGinput.Cells(i, 1).value
    field.assembly = PGinput.Cells(i, 2).value
    field.designation = ifcyrilic(PGinput.Cells(i, 3).value)
    field.material = replaceMinus6(PGinput.Cells(i, 4).value)
    field.toughness = PGinput.Cells(i, 5).value
    If PGinput.Cells(i, 6).value = "ok" Then field.STATUSlength = "OK" Else field.STATUSlength = "FAIL"
    PG.AddField field
    Set field = Nothing
Next i
On Error GoTo 0
PGbook.Close False
End Sub
Sub BRACEread()
Dim lrow, lcol, i As Integer
BRACEbook.Activate
lrow = LastRow
lcol = lastcol
Dim BRACEinput As Range
Set BRACEinput = Range("A3:E" & lrow)
Set BRACE = New BRACEclass
On Error Resume Next
For i = 1 To BRACEinput.Rows.Count
    Dim field As New PGfields
    field.item = BRACEinput.Cells(i, 1).value
    field.assembly = BRACEinput.Cells(i, 2).value
    field.designation = ifcyrilic(BRACEinput.Cells(i, 3).value)
    field.material = replaceMinus6(BRACEinput.Cells(i, 4).value)
    field.toughness = BRACEinput.Cells(i, 5).value
    If BRACEinput.Cells(i, 6).value = "ok" Then field.STATUSlength = "OK" Else field.STATUSlength = "FAIL"
    BRACE.AddField field
    Set field = Nothing
Next i
On Error GoTo 0
BRACEbook.Close False
End Sub
Sub STIFFSPread()
Dim lrow, lcol, i As Integer
STIFFSPbook.Activate
lrow = LastRow
lcol = lastcol
Dim STIFFSPinput As Range
Set STIFFSPinput = Range("A3:H" & lrow)
Set STIFFsp = New STIFFSPclass
On Error Resume Next
For i = 1 To STIFFSPinput.Rows.Count
    Dim field As New STIFFfields
    field.detail = STIFFSPinput.Cells(i, 1).value & "-" & STIFFSPinput.Cells(i, 2).value
    field.guide = STIFFSPinput.Cells(i, 4).value
    Dim Size() As String
    Size() = Split(STIFFSPinput.Cells(i, 5).value, "x")
    field.sizeA = Size(0)
    field.sizeB = Size(1)
    field.thickness = STIFFSPinput.Cells(i, 6).value
    field.material = STIFFSPinput.Cells(i, 7).value
    field.toughness = STIFFSPinput.Cells(i, 8).value
    STIFFsp.AddField field
    Set field = Nothing
Next i
On Error GoTo 0
STIFFSPbook.Close False
End Sub
Sub STIFFJread()
Dim lrow, lcol, i As Integer
STIFFJbook.Activate
lrow = LastRow
lcol = lastcol
Dim STIFFJinput As Range
Set STIFFJinput = Range("A3:J" & lrow)
Set STIFFj = New STIFFjclass
On Error Resume Next
For i = 1 To STIFFJinput.Rows.Count
    Dim field As New STIFFfields
    Dim Prefix As String
    If Left(STIFFJinput.Cells(i, 1).value, 4) = "GUSS" Then
        Prefix = "G"
    ElseIf Left(STIFFJinput.Cells(i, 1).value, 5) = "STIFF" Then
        Prefix = "S"
    End If
    field.detail = STIFFJinput.Cells(i, 2).value & "-" & STIFFJinput.Cells(i, 3).value & "-" & Prefix
    field.sizeA = STIFFJinput.Cells(i, 4).value
    field.sizeB = STIFFJinput.Cells(i, 5).value
    field.thickness = STIFFJinput.Cells(i, 6).value
    field.material = STIFFJinput.Cells(i, 7).value
    field.toughness = STIFFJinput.Cells(i, 8).value
    field.level = STIFFJinput.Cells(i, 9).value
    field.guide = STIFFJinput.Cells(i, 10).value
    STIFFj.AddField field
    Set field = Nothing
Next i
On Error GoTo 0
STIFFJbook.Close False
End Sub
Sub WELDread()
Dim lrow, i As Integer
On Error GoTo 0
WELDbook.Activate
lrow = Range("A3").End(xlDown).row
Dim WELDinput As Range
Set WELDinput = Range("A3:S" & lrow)
Set WELD = New WELDclass
On Error Resume Next
For i = 1 To WELDinput.Rows.Count
    Dim field As New WELDfields
    field.detailsNUM = WELDinput.Cells(i, 1).value
    field.weldnumber = WELDinput.Cells(i, 2).value
    field.GUID = WELDinput.Cells(i, 3).value
    field.weldtype = WELDinput.Cells(i, 4).value
    field.STATUSweldtype = WELDinput.Cells(i, 5).value
    field.weldsize = WELDinput.Cells(i, 6).value
    field.STATUSweldsize = WELDinput.Cells(i, 7).value
    field.weldANGLE = WELDinput.Cells(i, 8).value
    field.STATUSweldANGLE = WELDinput.Cells(i, 9).value
    field.weldfinish = WELDinput.Cells(i, 10).value
    field.STATUSweldfinish = WELDinput.Cells(i, 11).value
    field.weldbooklet = WELDinput.Cells(i, 12).value
    field.STATUSweldbooklet = WELDinput.Cells(i, 13).value
    field.weldjointtype = WELDinput.Cells(i, 14).value
    field.STATUSweldjointtype = WELDinput.Cells(i, 15).value
    field.weldbeveltype = WELDinput.Cells(i, 16).value
    field.STATUSweldbeveltype = WELDinput.Cells(i, 17).value
    field.weldNDT = WELDinput.Cells(i, 18).value
    field.STATUSweldNDT = WELDinput.Cells(i, 19).value
    WELD.AddField field, CStr(field.weldnumber)
    Set field = Nothing
Next i
On Error GoTo 0
WELDbook.Close False
End Sub

Sub WELDFILLETread()
Dim lrow, i As Integer
On Error GoTo 0
WELDFILLETbook.Activate
lrow = Range("A1").End(xlDown).row
Dim WELDFILLETinput As Range
Set WELDFILLETinput = Range("A2:F" & lrow)
Set WELDFILLET = New WELDFILLETclass
On Error Resume Next
For i = 1 To WELDFILLETinput.Rows.Count
    Dim field As New WELDfields
    field.detailsNUM = WELDFILLETinput.Cells(i, 1).value
    field.weldnumber = WELDFILLETinput.Cells(i, 2).value
    field.GUID = WELDFILLETinput.Cells(i, 3).value
    field.weldtype = WELDFILLETinput.Cells(i, 4).value
    field.weldsize = WELDFILLETinput.Cells(i, 5).value
    field.STATUSweldsize = WELDFILLETinput.Cells(i, 6).value
    WELDFILLET.AddField field, CStr(field.GUID)
    Set field = Nothing
Next i
On Error GoTo 0
WELDFILLETbook.Close False
End Sub


Sub NODEread()
Dim lrow, lcol, i As Integer
NODEbook.Activate
lrow = LastRow
lcol = lastcol
Dim NODEinput As Range
Set NODEinput = Range("A3:J" & lrow)
Set NODE = New NODEclass
On Error Resume Next
For i = 1 To NODEinput.Rows.Count
    Dim field As New NODEfields
    Dim Prefix As String
    If Left(NODEinput.Cells(i, 1).value, 4) = "STAR" Then
        Prefix = "st"
    ElseIf Left(NODEinput.Cells(i, 1).value, 5) = "INTER" Then
        Prefix = "intst"
    ElseIf Left(NODEinput.Cells(i, 1).value, 2) = "OD" Then
        If Len(NODEinput.Cells(i, 4).value) > 6 Then Prefix = "cone" Else Prefix = "od"
    ElseIf NODEinput.Cells(i, 1).value = "WEB_INSERT" Then
        Prefix = "webins"
    ElseIf NODEinput.Cells(i, 1).value = "WEB_INSERT_2" Then
        Prefix = "webins2"
    ElseIf NODEinput.Cells(i, 1).value = "WEB_INSERT_3" Then
        Prefix = "webins3"
    End If
    
    field.detail = NODEinput.Cells(i, 2).value & "-" & NODEinput.Cells(i, 3).value & "-" & Prefix
    Dim a() As String
    a() = Split(NODEinput.Cells(i, 4).value, "/")
        field.AsizeA = ifblank(a(0))
        If UBound(a()) = 1 Then field.AsizeB = a(1) Else field.AsizeB = "blank"
    a() = Split(NODEinput.Cells(i, 5).value, "/")
        field.BsizeA = ifblank(a(0))
        If UBound(a()) = 1 Then field.BsizeB = a(1) Else field.BsizeB = "blank"
    a() = Split(NODEinput.Cells(i, 6).value, "/")
        field.Athickness = ifblank(a(0))
        If UBound(a()) = 1 Then field.Bthickness = a(1) Else field.Bthickness = "blank"
    a() = Split(NODEinput.Cells(i, 7).value, "/")
        field.Atoughness = ifblank(a(0))
        If UBound(a()) = 1 Then field.Btoughness = a(1) Else field.Btoughness = "blank"
    a() = Split(NODEinput.Cells(i, 8).value, "/")
        field.material = ifblank(a(0))
    field.level = NODEinput.Cells(i, 9).value
    field.guide = NODEinput.Cells(i, 10).value
    NODE.AddField field
    Set field = Nothing
Next i
On Error GoTo 0
NODEbook.Close False
End Sub

Sub HBRACEread()
Dim lrow, lcol, i As Integer
HBRACEbook.Activate
lrow = LastRow
lcol = lastcol
Dim HBRACEinput As Range
Set HBRACEinput = Range("A3:J" & lrow)
Set HBRACE = New HBRACEclass
On Error Resume Next
For i = 1 To HBRACEinput.Rows.Count
    Dim field As New STIFFfields
    field.detail = HBRACEinput.Cells(i, 2).value & "-" & HBRACEinput.Cells(i, 3).value & "-" & Left(HBRACEinput.Cells(i, 1).value, 1)
    field.sizeA = HBRACEinput.Cells(i, 4).value
    field.sizeB = HBRACEinput.Cells(i, 5).value
    field.thickness = HBRACEinput.Cells(i, 6).value
    field.material = HBRACEinput.Cells(i, 7).value
    field.toughness = HBRACEinput.Cells(i, 8).value
    field.level = HBRACEinput.Cells(i, 9).value
    field.guide = HBRACEinput.Cells(i, 10).value
    HBRACE.AddField field
    Set field = Nothing
Next i
On Error GoTo 0
HBRACEbook.Close False
End Sub
Sub TRANSread()
Dim lrow, lcol, i As Integer
TRANSbook.Activate
lrow = LastRow
lcol = lastcol
Dim TRANSinput As Range
Set TRANSinput = Range("A3:J" & lrow)
Set TRANS = New TRANSclass
On Error Resume Next
For i = 1 To TRANSinput.Rows.Count
    Dim field As New TRANSfields
    field.detail = TRANSinput.Cells(i, 1).value & "-" & TRANSinput.Cells(i, 2).value & "-" & TRANSinput.Cells(i, 3).value & "-" & Left(TRANSinput.Cells(i, 4).value, 1)
    field.width = TRANSinput.Cells(i, 5).value
    field.length = TRANSinput.Cells(i, 6).value
    field.thickness = TRANSinput.Cells(i, 7).value
    field.toughness = TRANSinput.Cells(i, 8).value
    field.material = TRANSinput.Cells(i, 9).value
    field.guide = TRANSinput.Cells(i, 10).value
    TRANS.AddField field
    Set field = Nothing
Next i
On Error GoTo 0
TRANSbook.Close False
End Sub
Sub HNODEread()
Dim lrow, lcol, i As Integer
HNODEbook.Activate
lrow = LastRow
lcol = lastcol
Dim HNODEinput As Range
Set HNODEinput = Range("A3:I" & lrow)
Set HNODE = New HNODEclass
On Error Resume Next
For i = 1 To HNODEinput.Rows.Count
    Dim field As New HNODEfields
    Dim Prefix As String
    If Left(HNODEinput.Cells(i, 1).value, 5) = "INSER" Then
        Prefix = "insweb"
    ElseIf Left(HNODEinput.Cells(i, 1).value, 5) = "BOTTO" Then
        Prefix = "botfl"
    ElseIf Left(HNODEinput.Cells(i, 1).value, 3) = "TOP" Then
        Prefix = "topfl"
    ElseIf Left(HNODEinput.Cells(i, 1).value, 5) = "STIFF" Then
        Prefix = "stiff"
    End If
    field.detail = HNODEinput.Cells(i, 2).value & "-" & HNODEinput.Cells(i, 3).value & "-" & Prefix
    Dim a() As String
    a() = Split(HNODEinput.Cells(i, 4).value, "x")
    field.sizeA = a(0)
    field.sizeB = a(1)
    field.thickness = HNODEinput.Cells(i, 5).value
    field.toughness = HNODEinput.Cells(i, 6).value
    field.material = HNODEinput.Cells(i, 7).value
    field.level = HNODEinput.Cells(i, 8).value
    field.guide = HNODEinput.Cells(i, 9).value
    HNODE.AddField field
    Set field = Nothing
Next i
On Error GoTo 0
HNODEbook.Close False
End Sub
Sub SECONDARYFRread()
Dim lrow, lcol, i As Integer
SECONDARYFRbook.Activate
lrow = LastRow
lcol = lastcol
Dim SECONDARYFRinput As Range
Set SECONDARYFRinput = Range("B3:E" & lrow)
Set SECONDARYFR = New SECONDARYFRclass
On Error Resume Next
For i = 1 To SECONDARYFRinput.Rows.Count
    Dim field As New PGfields
    field.item = SECONDARYFRinput.Cells(i, 1).value
    field.assembly = SECONDARYFRinput.Cells(i, 2).value
    field.designation = ifcyrilic(SECONDARYFRinput.Cells(i, 3).value)
    field.material = replaceMinus6(SECONDARYFRinput.Cells(i, 4).value)
    field.toughness = SECONDARYFRinput.Cells(i, 5).value
    If SECONDARYFRinput.Cells(i, 6).value = "ok" Then field.STATUSlength = "OK" Else field.STATUSlength = "FAIL"
    SECONDARYFR.AddField field
    Set field = Nothing
Next i
On Error GoTo 0
SECONDARYFRbook.Close False
End Sub
Sub IDread()
Dim lrow, i As Integer
On Error GoTo 0
IDbook.Activate
lrow = Range("A1").End(xlDown).row
Dim IDinput As Range
Set IDinput = Range("A2:G" & lrow)
Set ID = New IDclass
On Error Resume Next
For i = 1 To IDinput.Rows.Count
    Dim field As New IDfields
    field.detail = IDinput.Cells(i, 1).value
    field.profile = IDinput.Cells(i, 3).value
    field.material = IDinput.Cells(i, 4).value
    field.toughness = IDinput.Cells(i, 5).value
    field.IDcode = IDinput.Cells(i, 6).value
    If IDinput.Cells(i, 7).value = "ok" Then field.STATUSIDcode = "OK" Else field.STATUSIDcode = "NO"
    ID.AddField field, CStr(field.detail)
    Set field = Nothing
Next i
On Error GoTo 0
IDbook.Close False
End Sub
