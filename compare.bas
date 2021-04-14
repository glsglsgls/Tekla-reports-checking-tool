Attribute VB_Name = "compare"

Option Explicit
Sub main()
On Error Resume Next
If Not PGbook Is Nothing Then
    basePGread   'read database
    PGread       'read PG
    PGcompare    'compare PG
    PG.getresult
    If IsNumeric(Right(newsheet, 1)) = True Then indexsheet = Right(newsheet, 1) Else indexsheet = "1"
    Dim PGfail, PGmanual As Integer
    PGfail = PG.getFAILquantity
    PGmanual = PG.getMANUALquantity
    With Sheets("mainVIEW")
        .Cells(lastrowinmainsheet(2) + 2, 2).value = mybook
        setrowheight ("PG")
        writeFAILquantity (PGfail)
        writeMANUALquantity (PGmanual)
    End With
End If

If Not BRACEbook Is Nothing Then
    basePGread     'read database
    BRACEread      'read BRACING and COLUMN
    BRACEcompare   'compare BRACING and COLUMN
    BRACE.getresult
    Dim BRACEfail, BRACEmanual As Integer
   indexsheet = 0
    If IsNumeric(Right(newsheet, 1)) = True Then indexsheet = Right(newsheet, 1) Else indexsheet = "1"
    BRACEfail = BRACE.getFAILquantity
    BRACEmanual = BRACE.getMANUALquantity
    With Sheets("mainVIEW")
        .Cells(lastrowinmainsheet(2) + 2, 2).value = mybook
        setrowheight ("BRACE")
        writeFAILquantity (BRACEfail)
        writeMANUALquantity (BRACEmanual)
    End With
End If

If Not STIFFSPbook Is Nothing Then
    baseSTIFFSPread   'read database
    STIFFSPread       'read SP/SK STIFFENERS
    STIFFSPcompare    'compare SP/SK STIFFENERS
    STIFFsp.getresult
   indexsheet = 0
    If IsNumeric(Right(newsheet, 1)) = True Then indexsheet = Right(newsheet, 1) Else indexsheet = "1"
    Dim STIFFSPfail, STIFFSPmanual As Integer
    STIFFSPfail = STIFFsp.getFAILquantity
    STIFFSPmanual = STIFFsp.getMANUALquantity
    With Sheets("mainVIEW")
        .Cells(lastrowinmainsheet(2) + 2, 2).value = mybook
        setrowheight ("STIFFSP")
        writeFAILquantity (STIFFSPfail)
        writeMANUALquantity (STIFFSPmanual)
    End With
End If

If Not STIFFJbook Is Nothing Then
    baseSTIFFJread   'read database
    STIFFJread       'read SP/SK STIFFENERS
    STIFFJcompare    'compare SP/SK STIFFENERS
    STIFFj.getresult
   indexsheet = 0
    If IsNumeric(Right(newsheet, 1)) = True Then indexsheet = Right(newsheet, 1) Else indexsheet = "1"
    Dim STIFFJfail, STIFFJmanual As Integer
    STIFFJfail = STIFFj.getFAILquantity
    STIFFJmanual = STIFFj.getMANUALquantity
    With Sheets("mainVIEW")
        .Cells(lastrowinmainsheet(2) + 2, 2).value = mybook
        setrowheight ("STIFFJ")
        writeFAILquantity (STIFFJfail)
        writeMANUALquantity (STIFFJmanual)
    End With
End If

If Not HBRACEbook Is Nothing Then
    baseSTIFFJread   'read database
    HBRACEread       'read SP/SK STIFFENERS
    HBRACEcompare    'compare SP/SK STIFFENERS
    HBRACE.getresult
   indexsheet = 0
    If IsNumeric(Right(newsheet, 1)) = True Then indexsheet = Right(newsheet, 1) Else indexsheet = "1"
    Dim HBRACEfail, HBRACEmanual As Integer
    HBRACEfail = HBRACE.getFAILquantity
    HBRACEmanual = HBRACE.getMANUALquantity
    With Sheets("mainVIEW")
        .Cells(lastrowinmainsheet(2) + 2, 2).value = mybook
        setrowheight ("HBRACE")
        writeFAILquantity (HBRACEfail)
        writeMANUALquantity (HBRACEmanual)
    End With
End If

If Not NODEbook Is Nothing Then
    baseNODEread    'read database
    NODEread        'read NODE
    NODEcompare     'compare NODE
    NODE.getresult
   indexsheet = 0
    If IsNumeric(Right(newsheet, 1)) = True Then indexsheet = Right(newsheet, 1) Else indexsheet = "1"
    Dim NODEfail, NODEmanual As Integer
    NODEfail = NODE.getFAILquantity
    NODEmanual = NODE.getMANUALquantity
    With Sheets("mainVIEW")
        .Cells(lastrowinmainsheet(2) + 2, 2).value = mybook
        setrowheight ("NODE")
        writeFAILquantity (NODEfail)
        writeMANUALquantity (NODEmanual)
    End With
End If

If Not HNODEbook Is Nothing Then
    baseHNODEread    'read database
    HNODEread        'read HNODE
    HNODEcompare     'compare HNODE
    HNODE.getresult
   indexsheet = 0
    If IsNumeric(Right(newsheet, 1)) = True Then indexsheet = Right(newsheet, 1) Else indexsheet = "1"
    Dim HNODEfail, HNODEmanual As Integer
    HNODEfail = HNODE.getFAILquantity
    HNODEmanual = HNODE.getMANUALquantity
    With Sheets("mainVIEW")
        .Cells(lastrowinmainsheet(2) + 2, 2).value = mybook
        setrowheight ("HNODE")
        writeFAILquantity (HNODEfail)
        writeMANUALquantity (HNODEmanual)
    End With
End If

If Not TRANSbook Is Nothing Then
    baseTRANSread   'read database
    TRANSread       'read TRANSITION profile
    TRANScompare    'compare TRANSITION profile
    TRANS.getresult
   indexsheet = 0
    If IsNumeric(Right(newsheet, 1)) = True Then indexsheet = Right(newsheet, 1) Else indexsheet = "1"
    Dim TRANSfail, TRANSmanual As Integer
    TRANSfail = TRANS.getFAILquantity
    TRANSmanual = TRANS.getMANUALquantity
    With Sheets("mainVIEW")
        .Cells(lastrowinmainsheet(2) + 2, 2).value = mybook
        setrowheight ("TRANS")
        writeFAILquantity (TRANSfail)
        writeMANUALquantity (TRANSmanual)
    End With
End If

If Not WELDbook Is Nothing Then
    WELDread       'read WELD list
    WELD.getresult
        Dim WELDfail As Integer
   indexsheet = 0
    If IsNumeric(Right(newsheet, 1)) = True Then indexsheet = Right(newsheet, 1) Else indexsheet = "1"
    WELDfail = WELD.getFAILquantity
    With Sheets("mainVIEW")
        .Cells(lastrowinmainsheet(2) + 2, 2).value = mybook
        setrowheight ("WELD")
        writeFAILquantity (WELDfail)
    End With
End If

If Not SECONDARYFRbook Is Nothing Then
    basePGread     'read database
    SECONDARYFRread      'read BRACING and COLUMN
    SECONDARYFRcompare   'compare BRACING and COLUMN
    SECONDARYFR.getresult
    Dim SECONDARYFRfail, SECONDARYFRmanual As Integer
   indexsheet = 0
    If IsNumeric(Right(newsheet, 1)) = True Then indexsheet = Right(newsheet, 1) Else indexsheet = "1"
    SECONDARYFRfail = SECONDARYFR.getFAILquantity
    SECONDARYFRmanual = SECONDARYFR.getMANUALquantity
    With Sheets("mainVIEW")
        .Cells(lastrowinmainsheet(2) + 2, 2).value = mybook
        setrowheight ("secondaryFRAMING")
        writeFAILquantity (SECONDARYFRfail)
        writeMANUALquantity (SECONDARYFRmanual)
    End With
End If

If Not IDbook Is Nothing Then
    IDread       'read ID list
    ID.getresult
        Dim IDfail As Integer
   indexsheet = 0
    If IsNumeric(Right(newsheet, 1)) = True Then indexsheet = Right(newsheet, 1) Else indexsheet = "1"
    IDfail = ID.getFAILquantity
    With Sheets("mainVIEW")
        .Cells(lastrowinmainsheet(2) + 2, 2).value = mybook
        setrowheight ("ID")
        writeFAILquantity (IDfail)
    End With
End If


If Not WELDFILLETbook Is Nothing Then
    WELDFILLETread       'read WELDFILLET list
    WELDFILLET.getresult
        Dim WELDFILLETfail As Integer
   indexsheet = 0
    If IsNumeric(Right(newsheet, 1)) = True Then indexsheet = Right(newsheet, 1) Else indexsheet = "1"
    WELDFILLETfail = WELDFILLET.getFAILquantity
    With Sheets("mainVIEW")
        .Cells(lastrowinmainsheet(2) + 2, 2).value = mybook
        setrowheight ("WELDFILLET")
        writeFAILquantity (WELDFILLETfail)
    End With
End If



ThisWorkbook.Sheets("mainVIEW").Activate
End Sub

Sub PGcompare()
Dim element As Variant
For Each element In PG.PGcollection
    If basePG.PG_BASEcollection(element.item).designation = element.designation Then element.STATUSprofile = "OK" Else element.STATUSprofile = "FAIL"
    If basePG.PG_BASEcollection(element.item).material = replaceMinus6(element.material) Then element.STATUSmaterial = "OK" Else element.STATUSmaterial = "FAIL"
    If basePG.PG_BASEcollection(element.item).toughness = element.toughness Then element.STATUStoughness = "OK" Else element.STATUStoughness = "FAIL"
Next element
'see status PGfields in PG collection
End Sub
Sub BRACEcompare()
Dim element As Variant
For Each element In BRACE.BRACEcollection
    If basePG.PG_BASEcollection(element.item).designation = element.designation Then element.STATUSprofile = "OK" Else element.STATUSprofile = "FAIL"
    If basePG.PG_BASEcollection(element.item).material = replaceMinus6(element.material) Then element.STATUSmaterial = "OK" Else element.STATUSmaterial = "FAIL"
    If basePG.PG_BASEcollection(element.item).toughness = element.toughness Then element.STATUStoughness = "OK" Else element.STATUStoughness = "FAIL"
Next element
'see status PGfields in BRACE collection
End Sub
Sub STIFFSPcompare()
Dim element As Variant
For Each element In STIFFsp.STIFFSPcollection
'baseSTIFFSP.STIFFSP_BASEcollection
On Error Resume Next
    If baseSTIFFsp.STIFFSP_BASEcollection(element.detail).quantity = element.quantity Then element.STATUSquantity = "OK" Else element.STATUSquantity = "FAIL"
    If Abs(baseSTIFFsp.STIFFSP_BASEcollection(element.detail).sizeA - element.sizeA) < 4 And Abs(baseSTIFFsp.STIFFSP_BASEcollection(element.detail).sizeB - element.sizeB) < 4 Or _
    Abs(baseSTIFFsp.STIFFSP_BASEcollection(element.detail).sizeA - element.sizeB) < 4 And Abs(baseSTIFFsp.STIFFSP_BASEcollection(element.detail).sizeB - element.sizeA) < 4 Then _
        element.STATUSsize = "OK" Else element.STATUSsize = "FAIL"
    If baseSTIFFsp.STIFFSP_BASEcollection(element.detail).thickness = element.thickness Then element.STATUSthickness = "OK" Else element.STATUSthickness = "FAIL"
    If baseSTIFFsp.STIFFSP_BASEcollection(element.detail).material = replaceMinus6(element.material) Then element.STATUSmaterial = "OK" Else element.STATUSmaterial = "FAIL"
    If baseSTIFFsp.STIFFSP_BASEcollection(element.detail).toughness = element.toughness Then element.STATUStoughness = "OK" Else element.STATUStoughness = "FAIL"
Next element
On Error GoTo 0
'see status PGfields in BRACE collection
End Sub
Sub STIFFJcompare()
Dim element As Variant
Dim lcolLEV As Integer
Dim baseLEVELinput As Range
On Error Resume Next
lcolLEV = Range("Y2").End(xlToRight).Column
Set baseLEVELinput = Range(Cells(1, 25), Cells(2, lcolLEV))
baseSTIFFj.levelA = 0
If baseLEVELinput.Cells(2, 2).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseSTIFFj.LevelB = baseLEVELinput.Cells(2, 2).value - baseLEVELinput.Cells(2, 1).value Else baseSTIFFj.LevelB = ""
If baseLEVELinput.Cells(2, 3).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseSTIFFj.LevelC = baseLEVELinput.Cells(2, 3).value - baseLEVELinput.Cells(2, 1).value Else baseSTIFFj.LevelC = ""
If baseLEVELinput.Cells(2, 4).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseSTIFFj.LevelD = baseLEVELinput.Cells(2, 4).value - baseLEVELinput.Cells(2, 1).value Else baseSTIFFj.LevelD = ""
If baseLEVELinput.Cells(2, 5).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseSTIFFj.LevelE = baseLEVELinput.Cells(2, 5).value - baseLEVELinput.Cells(2, 1).value Else baseSTIFFj.LevelE = ""
If baseLEVELinput.Cells(2, 6).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseSTIFFj.LevelF = baseLEVELinput.Cells(2, 6).value - baseLEVELinput.Cells(2, 1).value Else baseSTIFFj.LevelF = ""
If baseLEVELinput.Cells(2, 7).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseSTIFFj.LevelG = baseLEVELinput.Cells(2, 7).value - baseLEVELinput.Cells(2, 1).value Else baseSTIFFj.LevelG = ""
If baseLEVELinput.Cells(2, 8).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseSTIFFj.LevelH = baseLEVELinput.Cells(2, 8).value - baseLEVELinput.Cells(2, 1).value Else baseSTIFFj.LevelH = ""
If baseLEVELinput.Cells(2, 9).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseSTIFFj.LevelI = baseLEVELinput.Cells(2, 9).value - baseLEVELinput.Cells(2, 1).value Else baseSTIFFj.LevelI = ""
If baseLEVELinput.Cells(2, 10).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseSTIFFj.LevelJ = baseLEVELinput.Cells(2, 10).value - baseLEVELinput.Cells(2, 1).value Else baseSTIFFj.LevelJ = ""
For Each element In STIFFj.STIFFJcollection
    If Abs(element.level - baseSTIFFj.levelA) < 3 Or Abs(element.level - baseLEVELinput.Cells(2, 1).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 1)
    ElseIf Abs(element.level - baseSTIFFj.LevelB) < 3 Or Abs(element.level - baseLEVELinput.Cells(2, 2).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 2)
    ElseIf Abs(element.level - baseSTIFFj.LevelC) < 3 Or Abs(element.level - baseLEVELinput.Cells(2, 3).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 3)
    ElseIf Abs(element.level - baseSTIFFj.LevelD) < 3 Or Abs(element.level - baseLEVELinput.Cells(2, 4).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 4)
    ElseIf Abs(element.level - baseSTIFFj.LevelE) < 3 Or Abs(element.level - baseLEVELinput.Cells(2, 5).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 5)
    ElseIf Abs(element.level - baseSTIFFj.LevelF) < 3 Or Abs(element.level - baseLEVELinput.Cells(2, 6).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 6)
    ElseIf Abs(element.level - baseSTIFFj.LevelG) < 3 Or Abs(element.level - baseLEVELinput.Cells(2, 7).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 7)
    ElseIf Abs(element.level - baseSTIFFj.LevelH) < 3 Or Abs(element.level - baseLEVELinput.Cells(2, 8).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 8)
    ElseIf Abs(element.level - baseSTIFFj.LevelI) < 3 Or Abs(element.level - baseLEVELinput.Cells(2, 9).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 9)
    ElseIf Abs(element.level - baseSTIFFj.LevelJ) < 3 Or Abs(element.level - baseLEVELinput.Cells(2, 10).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 10)
    End If
    If Abs(baseSTIFFj.STIFFJ_BASEcollection(element.detail).sizeA - element.sizeA) < 4 And Abs(baseSTIFFj.STIFFJ_BASEcollection(element.detail).sizeB - element.sizeB) < 4 Or _
    Abs(baseSTIFFj.STIFFJ_BASEcollection(element.detail).sizeA - element.sizeB) < 4 And Abs(baseSTIFFj.STIFFJ_BASEcollection(element.detail).sizeB - element.sizeA) < 4 Then _
        element.STATUSsize = "OK" Else element.STATUSsize = "FAIL"
    If baseSTIFFj.STIFFJ_BASEcollection(element.detail).thickness = element.thickness Then element.STATUSthickness = "OK" Else element.STATUSthickness = "FAIL"
    If baseSTIFFj.STIFFJ_BASEcollection(element.detail).material = replaceMinus6(element.material) Then element.STATUSmaterial = "OK" Else element.STATUSmaterial = "FAIL"
    If baseSTIFFj.STIFFJ_BASEcollection(element.detail).toughness = element.toughness Then element.STATUStoughness = "OK" Else element.STATUStoughness = "FAIL"
Next element
'If there are some items in the report, which not included in database
End Sub

Sub NODEcompare()
Dim element As Variant
Dim lcolLEV As Integer
Dim baseLEVELinput, baseMIDLEVELinput As Range
On Error Resume Next
lcolLEV = Range("Y2").End(xlToRight).Column
Set baseLEVELinput = Range(Cells(1, 25), Cells(2, lcolLEV))
baseNODE.levelA = 0
If baseLEVELinput.Cells(2, 2).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseNODE.LevelB = baseLEVELinput.Cells(2, 2).value - baseLEVELinput.Cells(2, 1).value Else baseNODE.LevelB = ""
If baseLEVELinput.Cells(2, 3).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseNODE.LevelC = baseLEVELinput.Cells(2, 3).value - baseLEVELinput.Cells(2, 1).value Else baseNODE.LevelC = ""
If baseLEVELinput.Cells(2, 4).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseNODE.LevelD = baseLEVELinput.Cells(2, 4).value - baseLEVELinput.Cells(2, 1).value Else baseNODE.LevelD = ""
If baseLEVELinput.Cells(2, 5).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseNODE.LevelE = baseLEVELinput.Cells(2, 5).value - baseLEVELinput.Cells(2, 1).value Else baseNODE.LevelE = ""
If baseLEVELinput.Cells(2, 6).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseNODE.LevelF = baseLEVELinput.Cells(2, 6).value - baseLEVELinput.Cells(2, 1).value Else baseNODE.LevelF = ""
If baseLEVELinput.Cells(2, 7).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseNODE.LevelG = baseLEVELinput.Cells(2, 7).value - baseLEVELinput.Cells(2, 1).value Else baseNODE.LevelG = ""
If baseLEVELinput.Cells(2, 8).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseNODE.LevelH = baseLEVELinput.Cells(2, 8).value - baseLEVELinput.Cells(2, 1).value Else baseNODE.LevelH = ""
If baseLEVELinput.Cells(2, 9).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseNODE.LevelI = baseLEVELinput.Cells(2, 9).value - baseLEVELinput.Cells(2, 1).value Else baseNODE.LevelI = ""
If baseLEVELinput.Cells(2, 10).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseNODE.LevelJ = baseLEVELinput.Cells(2, 10).value - baseLEVELinput.Cells(2, 1).value Else baseNODE.LevelJ = ""
lcolLEV = Range("Y5").End(xlToRight).Column
Set baseMIDLEVELinput = Range(Cells(4, 25), Cells(5, lcolLEV))
If baseMIDLEVELinput.Cells(2, 1).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseNODE.MIDlevel1 = baseMIDLEVELinput.Cells(2, 1).value - baseLEVELinput.Cells(2, 1).value Else baseNODE.MIDlevel1 = ""
If baseMIDLEVELinput.Cells(2, 2).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseNODE.MIDlevel2 = baseMIDLEVELinput.Cells(2, 2).value - baseLEVELinput.Cells(2, 1).value Else baseNODE.MIDlevel2 = ""
If baseMIDLEVELinput.Cells(2, 3).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseNODE.MIDlevel3 = baseMIDLEVELinput.Cells(2, 3).value - baseLEVELinput.Cells(2, 1).value Else baseNODE.MIDlevel3 = ""
If baseMIDLEVELinput.Cells(2, 4).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseNODE.MIDlevel4 = baseMIDLEVELinput.Cells(2, 4).value - baseLEVELinput.Cells(2, 1).value Else baseNODE.MIDlevel4 = ""
If baseMIDLEVELinput.Cells(2, 5).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseNODE.MIDlevel5 = baseMIDLEVELinput.Cells(2, 5).value - baseLEVELinput.Cells(2, 1).value Else baseNODE.MIDlevel5 = ""
If baseMIDLEVELinput.Cells(2, 6).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseNODE.MIDlevel6 = baseMIDLEVELinput.Cells(2, 6).value - baseLEVELinput.Cells(2, 1).value Else baseNODE.MIDlevel6 = ""
For Each element In NODE.NODEcollection
    If Abs(element.level - baseNODE.levelA) < 3 Or Abs(element.level - baseLEVELinput.Cells(2, 1).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 1)
    ElseIf Abs(element.level - baseNODE.LevelB) < 3 Or Abs(element.level - baseLEVELinput.Cells(2, 2).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 2)
    ElseIf Abs(element.level - baseNODE.LevelC) < 3 Or Abs(element.level - baseLEVELinput.Cells(2, 3).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 3)
    ElseIf Abs(element.level - baseNODE.LevelD) < 3 Or Abs(element.level - baseLEVELinput.Cells(2, 4).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 4)
    ElseIf Abs(element.level - baseNODE.LevelE) < 3 Or Abs(element.level - baseLEVELinput.Cells(2, 5).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 5)
    ElseIf Abs(element.level - baseNODE.LevelF) < 3 Or Abs(element.level - baseLEVELinput.Cells(2, 6).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 6)
    ElseIf Abs(element.level - baseNODE.LevelG) < 3 Or Abs(element.level - baseLEVELinput.Cells(2, 7).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 7)
    ElseIf Abs(element.level - baseNODE.LevelH) < 3 Or Abs(element.level - baseLEVELinput.Cells(2, 8).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 8)
    ElseIf Abs(element.level - baseNODE.LevelI) < 3 Or Abs(element.level - baseLEVELinput.Cells(2, 9).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 9)
    ElseIf Abs(element.level - baseNODE.LevelJ) < 3 Or Abs(element.level - baseLEVELinput.Cells(2, 10).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 10)
    ElseIf Abs(element.level - baseNODE.MIDlevel1) < 3 Or Abs(element.level - baseMIDLEVELinput.Cells(2, 1).value) < 3 Then
        element.detail = element.detail & "-" & baseMIDLEVELinput.Cells(1, 1)
    ElseIf Abs(element.level - baseNODE.MIDlevel2) < 3 Or Abs(element.level - baseMIDLEVELinput.Cells(2, 2).value) < 3 Then
        element.detail = element.detail & "-" & baseMIDLEVELinput.Cells(1, 2)
    ElseIf Abs(element.level - baseNODE.MIDlevel3) < 3 Or Abs(element.level - baseMIDLEVELinput.Cells(2, 3).value) < 3 Then
        element.detail = element.detail & "-" & baseMIDLEVELinput.Cells(1, 3)
    ElseIf Abs(element.level - baseNODE.MIDlevel4) < 3 Or Abs(element.level - baseMIDLEVELinput.Cells(2, 4).value) < 3 Then
        element.detail = element.detail & "-" & baseMIDLEVELinput.Cells(1, 4)
    ElseIf Abs(element.level - baseNODE.MIDlevel5) < 3 Or Abs(element.level - baseMIDLEVELinput.Cells(2, 5).value) < 3 Then
        element.detail = element.detail & "-" & baseMIDLEVELinput.Cells(1, 5)
    ElseIf Abs(element.level - baseNODE.MIDlevel6) < 3 Or Abs(element.level - baseMIDLEVELinput.Cells(2, 6).value) < 3 Then
        element.detail = element.detail & "-" & baseMIDLEVELinput.Cells(1, 6)
    End If
    If element.AsizeB <> "blank" And element.BsizeB <> "blank" And baseNODE.NODE_BASEcollection(element.detail).AsizeB <> "blank" And baseNODE.NODE_BASEcollection(element.detail).BsizeB <> "blank" Then
        If Abs(cTN(baseNODE.NODE_BASEcollection(element.detail).AsizeA) - cTN(element.AsizeA)) < 4 And Abs(cTN(baseNODE.NODE_BASEcollection(element.detail).AsizeB) - cTN(element.AsizeB)) < 4 Or _
            Abs(cTN(baseNODE.NODE_BASEcollection(element.detail).AsizeA) - cTN(element.AsizeB)) < 4 And Abs(cTN(baseNODE.NODE_BASEcollection(element.detail).AsizeB) - cTN(element.AsizeA)) < 4 Then
            element.STATUSAsize = "OK"
        Else
            element.STATUSAsize = "FAIL"
        End If
        If Abs(cTN(baseNODE.NODE_BASEcollection(element.detail).BsizeA) - cTN(element.BsizeA)) < 4 And Abs(cTN(baseNODE.NODE_BASEcollection(element.detail).BsizeB) - cTN(element.BsizeB)) < 4 Or _
            Abs(cTN(baseNODE.NODE_BASEcollection(element.detail).BsizeA) - cTN(element.BsizeB)) < 4 And Abs(cTN(baseNODE.NODE_BASEcollection(element.detail).BsizeB) - cTN(element.BsizeA)) < 4 Then _
            element.STATUSBsize = "OK" Else element.STATUSBsize = "FAIL"
    Else
        'If Abs(cTN(baseNODE.NODE_BASEcollection(element.detail).AsizeA) - cTN(element.AsizeA)) < 4 Or _
        '   Abs(cTN(baseNODE.NODE_BASEcollection(element.detail).AsizeA) - cTN(element.BsizeA)) < 4 And Abs(cTN(baseNODE.NODE_BASEcollection(element.detail).BsizeA) - cTN(element.AsizeA)) < 4 Then
        '    element.STATUSAsize = "OK"
        'Else
        '    element.STATUSAsize = "FAIL"
        'End If
        If Abs(cTN(baseNODE.NODE_BASEcollection(element.detail).AsizeA) - cTN(element.AsizeA)) < 4 Or _
           Abs(cTN(baseNODE.NODE_BASEcollection(element.detail).AsizeA) - cTN(element.BsizeA)) < 4 Then
            element.STATUSAsize = "OK"
        Else
            element.STATUSAsize = "FAIL"
        End If
        If Abs(cTN(baseNODE.NODE_BASEcollection(element.detail).BsizeA) - cTN(element.BsizeA)) < 4 Or _
           Abs(cTN(baseNODE.NODE_BASEcollection(element.detail).BsizeA) - cTN(element.AsizeA)) < 4 Then
            element.STATUSBsize = "OK"
        Else
            element.STATUSBsize = "FAIL"
        End If
    End If
    If cTN(baseNODE.NODE_BASEcollection(element.detail).Athickness) = cTN(element.Athickness) Then element.STATUSAthickness = "OK" Else element.STATUSAthickness = "FAIL"
    If cTN(baseNODE.NODE_BASEcollection(element.detail).Bthickness) = cTN(element.Bthickness) Then element.STATUSBthickness = "OK" Else element.STATUSBthickness = "FAIL"
    If baseNODE.NODE_BASEcollection(element.detail).Atoughness = element.Atoughness Then element.STATUSAtoughness = "OK" Else element.STATUSAtoughness = "FAIL"
    If baseNODE.NODE_BASEcollection(element.detail).Btoughness = element.Btoughness Then element.STATUSBtoughness = "OK" Else element.STATUSBtoughness = "FAIL"
    If baseNODE.NODE_BASEcollection(element.detail).material = replaceMinus6(element.material) Then element.STATUSmaterial = "OK" Else element.STATUSmaterial = "FAIL"
Next element
End Sub


Sub HBRACEcompare()
Dim element As Variant
Dim lcolLEV As Integer
Dim baseLEVELinput As Range
On Error Resume Next
lcolLEV = Range("Y2").End(xlToRight).Column
Set baseLEVELinput = Range(Cells(1, 25), Cells(2, lcolLEV))
baseSTIFFj.levelA = 0
If baseLEVELinput.Cells(2, 2).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseSTIFFj.LevelB = baseLEVELinput.Cells(2, 2).value - baseLEVELinput.Cells(2, 1).value Else baseSTIFFj.LevelB = ""
If baseLEVELinput.Cells(2, 3).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseSTIFFj.LevelC = baseLEVELinput.Cells(2, 3).value - baseLEVELinput.Cells(2, 1).value Else baseSTIFFj.LevelC = ""
If baseLEVELinput.Cells(2, 4).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseSTIFFj.LevelD = baseLEVELinput.Cells(2, 4).value - baseLEVELinput.Cells(2, 1).value Else baseSTIFFj.LevelD = ""
If baseLEVELinput.Cells(2, 5).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseSTIFFj.LevelE = baseLEVELinput.Cells(2, 5).value - baseLEVELinput.Cells(2, 1).value Else baseSTIFFj.LevelE = ""
If baseLEVELinput.Cells(2, 6).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseSTIFFj.LevelF = baseLEVELinput.Cells(2, 6).value - baseLEVELinput.Cells(2, 1).value Else baseSTIFFj.LevelF = ""
If baseLEVELinput.Cells(2, 7).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseSTIFFj.LevelG = baseLEVELinput.Cells(2, 7).value - baseLEVELinput.Cells(2, 1).value Else baseSTIFFj.LevelG = ""
If baseLEVELinput.Cells(2, 8).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseSTIFFj.LevelH = baseLEVELinput.Cells(2, 8).value - baseLEVELinput.Cells(2, 1).value Else baseSTIFFj.LevelH = ""
If baseLEVELinput.Cells(2, 9).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseSTIFFj.LevelI = baseLEVELinput.Cells(2, 9).value - baseLEVELinput.Cells(2, 1).value Else baseSTIFFj.LevelI = ""
If baseLEVELinput.Cells(2, 10).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseSTIFFj.LevelJ = baseLEVELinput.Cells(2, 10).value - baseLEVELinput.Cells(2, 1).value Else baseSTIFFj.LevelJ = ""
For Each element In HBRACE.HBRACEcollection
    If Abs(element.level - baseSTIFFj.levelA) < 1.5 Or Abs(element.level - baseLEVELinput.Cells(2, 1).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 1)
    ElseIf Abs(element.level - baseSTIFFj.LevelB) < 2 Or Abs(element.level - baseLEVELinput.Cells(2, 2).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 2)
    ElseIf Abs(element.level - baseSTIFFj.LevelC) < 2 Or Abs(element.level - baseLEVELinput.Cells(2, 3).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 3)
    ElseIf Abs(element.level - baseSTIFFj.LevelD) < 2 Or Abs(element.level - baseLEVELinput.Cells(2, 4).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 4)
    ElseIf Abs(element.level - baseSTIFFj.LevelE) < 2 Or Abs(element.level - baseLEVELinput.Cells(2, 5).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 5)
    ElseIf Abs(element.level - baseSTIFFj.LevelF) < 2 Or Abs(element.level - baseLEVELinput.Cells(2, 6).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 6)
    ElseIf Abs(element.level - baseSTIFFj.LevelG) < 2 Or Abs(element.level - baseLEVELinput.Cells(2, 7).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 7)
    ElseIf Abs(element.level - baseSTIFFj.LevelH) < 2 Or Abs(element.level - baseLEVELinput.Cells(2, 8).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 8)
    ElseIf Abs(element.level - baseSTIFFj.LevelI) < 2 Or Abs(element.level - baseLEVELinput.Cells(2, 9).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 9)
    ElseIf Abs(element.level - baseSTIFFj.LevelJ) < 2 Or Abs(element.level - baseLEVELinput.Cells(2, 10).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 10)
    End If
    If Abs(baseSTIFFj.STIFFJ_BASEcollection(element.detail).sizeA - element.sizeA) < 4 And Abs(baseSTIFFj.STIFFJ_BASEcollection(element.detail).sizeB - element.sizeB) < 4 Or _
    Abs(baseSTIFFj.STIFFJ_BASEcollection(element.detail).sizeA - element.sizeB) < 4 And Abs(baseSTIFFj.STIFFJ_BASEcollection(element.detail).sizeB - element.sizeA) < 4 Then _
        element.STATUSsize = "OK" Else element.STATUSsize = "FAIL"
    If baseSTIFFj.STIFFJ_BASEcollection(element.detail).thickness = element.thickness Then element.STATUSthickness = "OK" Else element.STATUSthickness = "FAIL"
    If baseSTIFFj.STIFFJ_BASEcollection(element.detail).material = replaceMinus6(element.material) Then element.STATUSmaterial = "OK" Else element.STATUSmaterial = "FAIL"
    If baseSTIFFj.STIFFJ_BASEcollection(element.detail).toughness = element.toughness Then element.STATUStoughness = "OK" Else element.STATUStoughness = "FAIL"
Next element
'If there are some items in the report, which not included in database
End Sub

Sub TRANScompare()
Dim element As Variant
For Each element In TRANS.TRANScollection
'baseTRANS.TRANS_BASEcollection
    If Abs(baseTRANS.TRANS_BASEcollection(element.detail).length - element.length) < 4 And Abs(baseTRANS.TRANS_BASEcollection(element.detail).width - element.width) < 4 _
       Or Abs(baseTRANS.TRANS_BASEcollection(element.detail).width - element.length) < 4 And Abs(baseTRANS.TRANS_BASEcollection(element.detail).length - element.width) < 4 Then
        element.STATUSlength = "OK"
        element.STATUSwidth = "OK"
    Else
        element.STATUSlength = "FAIL"
        element.STATUSwidth = "FAIL"
    End If
    'If baseTRANS.TRANS_BASEcollection(element.detail).width = element.width Then element.STATUSwidth = "OK" Else element.STATUSwidth = "FAIL"
    If baseTRANS.TRANS_BASEcollection(element.detail).thickness = element.thickness Then element.STATUSthickness = "OK" Else element.STATUSthickness = "FAIL"
    If baseTRANS.TRANS_BASEcollection(element.detail).material = replaceMinus6(element.material) Then element.STATUSmaterial = "OK" Else element.STATUSmaterial = "FAIL"
    If baseTRANS.TRANS_BASEcollection(element.detail).toughness = element.toughness Then element.STATUStoughness = "OK" Else element.STATUStoughness = "FAIL"
Next element
'see status PGfields in BRACE collection
End Sub

Sub HNODEcompare()
Dim element As Variant
Dim lcolLEV As Integer
Dim baseLEVELinput, baseMIDLEVELinput As Range
On Error Resume Next
lcolLEV = Range("Y2").End(xlToRight).Column
Set baseLEVELinput = Range(Cells(1, 25), Cells(2, lcolLEV))
baseHNODE.levelA = 0
If baseLEVELinput.Cells(2, 2).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseHNODE.LevelB = baseLEVELinput.Cells(2, 2).value - baseLEVELinput.Cells(2, 1).value Else baseHNODE.LevelB = ""
If baseLEVELinput.Cells(2, 3).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseHNODE.LevelC = baseLEVELinput.Cells(2, 3).value - baseLEVELinput.Cells(2, 1).value Else baseHNODE.LevelC = ""
If baseLEVELinput.Cells(2, 4).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseHNODE.LevelD = baseLEVELinput.Cells(2, 4).value - baseLEVELinput.Cells(2, 1).value Else baseHNODE.LevelD = ""
If baseLEVELinput.Cells(2, 5).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseHNODE.LevelE = baseLEVELinput.Cells(2, 5).value - baseLEVELinput.Cells(2, 1).value Else baseHNODE.LevelE = ""
If baseLEVELinput.Cells(2, 6).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseHNODE.LevelF = baseLEVELinput.Cells(2, 6).value - baseLEVELinput.Cells(2, 1).value Else baseHNODE.LevelF = ""
If baseLEVELinput.Cells(2, 7).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseHNODE.LevelG = baseLEVELinput.Cells(2, 7).value - baseLEVELinput.Cells(2, 1).value Else baseHNODE.LevelG = ""
If baseLEVELinput.Cells(2, 8).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseHNODE.LevelH = baseLEVELinput.Cells(2, 8).value - baseLEVELinput.Cells(2, 1).value Else baseHNODE.LevelH = ""
If baseLEVELinput.Cells(2, 9).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseHNODE.LevelI = baseLEVELinput.Cells(2, 9).value - baseLEVELinput.Cells(2, 1).value Else baseHNODE.LevelI = ""
If baseLEVELinput.Cells(2, 10).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseHNODE.LevelJ = baseLEVELinput.Cells(2, 10).value - baseLEVELinput.Cells(2, 1).value Else baseHNODE.LevelJ = ""
lcolLEV = Range("Y5").End(xlToRight).Column
Set baseMIDLEVELinput = Range(Cells(4, 25), Cells(5, lcolLEV))
If baseMIDLEVELinput.Cells(2, 1).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseHNODE.MIDlevel1 = baseMIDLEVELinput.Cells(2, 1).value - baseLEVELinput.Cells(2, 1).value Else baseHNODE.MIDlevel1 = ""
If baseMIDLEVELinput.Cells(2, 2).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseHNODE.MIDlevel2 = baseMIDLEVELinput.Cells(2, 2).value - baseLEVELinput.Cells(2, 1).value Else baseHNODE.MIDlevel2 = ""
If baseMIDLEVELinput.Cells(2, 3).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseHNODE.MIDlevel3 = baseMIDLEVELinput.Cells(2, 3).value - baseLEVELinput.Cells(2, 1).value Else baseHNODE.MIDlevel3 = ""
If baseMIDLEVELinput.Cells(2, 4).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseHNODE.MIDlevel4 = baseMIDLEVELinput.Cells(2, 4).value - baseLEVELinput.Cells(2, 1).value Else baseHNODE.MIDlevel4 = ""
If baseMIDLEVELinput.Cells(2, 5).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseHNODE.MIDlevel5 = baseMIDLEVELinput.Cells(2, 5).value - baseLEVELinput.Cells(2, 1).value Else baseHNODE.MIDlevel5 = ""
If baseMIDLEVELinput.Cells(2, 6).value - baseLEVELinput.Cells(2, 1).value > 0 Then baseHNODE.MIDlevel6 = baseMIDLEVELinput.Cells(2, 6).value - baseLEVELinput.Cells(2, 1).value Else baseHNODE.MIDlevel6 = ""
For Each element In HNODE.HNODEcollection
    If Abs(element.level - baseHNODE.levelA) < 1.5 Or Abs(element.level - baseLEVELinput.Cells(2, 1).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 1)
    ElseIf Abs(element.level - baseHNODE.LevelB) < 2 Or Abs(element.level - baseLEVELinput.Cells(2, 2).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 2)
    ElseIf Abs(element.level - baseHNODE.LevelC) < 2 Or Abs(element.level - baseLEVELinput.Cells(2, 3).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 3)
    ElseIf Abs(element.level - baseHNODE.LevelD) < 2 Or Abs(element.level - baseLEVELinput.Cells(2, 4).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 4)
    ElseIf Abs(element.level - baseHNODE.LevelE) < 2 Or Abs(element.level - baseLEVELinput.Cells(2, 5).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 5)
    ElseIf Abs(element.level - baseHNODE.LevelF) < 2 Or Abs(element.level - baseLEVELinput.Cells(2, 6).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 6)
    ElseIf Abs(element.level - baseHNODE.LevelG) < 2 Or Abs(element.level - baseLEVELinput.Cells(2, 7).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 7)
    ElseIf Abs(element.level - baseHNODE.LevelH) < 2 Or Abs(element.level - baseLEVELinput.Cells(2, 8).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 8)
    ElseIf Abs(element.level - baseHNODE.LevelI) < 2 Or Abs(element.level - baseLEVELinput.Cells(2, 9).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 9)
    ElseIf Abs(element.level - baseHNODE.LevelJ) < 2 Or Abs(element.level - baseLEVELinput.Cells(2, 10).value) < 3 Then
        element.detail = element.detail & "-" & baseLEVELinput.Cells(1, 10)
    ElseIf Abs(element.level - baseHNODE.MIDlevel1) < 2 Or Abs(element.level - baseMIDLEVELinput.Cells(2, 1).value) < 3 Then
        element.detail = element.detail & "-" & baseMIDLEVELinput.Cells(1, 1)
    ElseIf Abs(element.level - baseHNODE.MIDlevel2) < 2 Or Abs(element.level - baseMIDLEVELinput.Cells(2, 2).value) < 3 Then
        element.detail = element.detail & "-" & baseMIDLEVELinput.Cells(1, 2)
    ElseIf Abs(element.level - baseHNODE.MIDlevel3) < 2 Or Abs(element.level - baseMIDLEVELinput.Cells(2, 3).value) < 3 Then
        element.detail = element.detail & "-" & baseMIDLEVELinput.Cells(1, 3)
    ElseIf Abs(element.level - baseHNODE.MIDlevel4) < 2 Or Abs(element.level - baseMIDLEVELinput.Cells(2, 4).value) < 3 Then
        element.detail = element.detail & "-" & baseMIDLEVELinput.Cells(1, 4)
    ElseIf Abs(element.level - baseHNODE.MIDlevel5) < 2 Or Abs(element.level - baseMIDLEVELinput.Cells(2, 5).value) < 3 Then
        element.detail = element.detail & "-" & baseMIDLEVELinput.Cells(1, 5)
    ElseIf Abs(element.level - baseHNODE.MIDlevel6) < 2 Or Abs(element.level - baseMIDLEVELinput.Cells(2, 6).value) < 3 Then
        element.detail = element.detail & "-" & baseMIDLEVELinput.Cells(1, 6)
    End If
    If Abs(baseHNODE.HNODE_BASEcollection(element.detail).sizeA - element.sizeA) < 4 And Abs(baseHNODE.HNODE_BASEcollection(element.detail).sizeB - element.sizeB) < 4 Or _
    Abs(baseHNODE.HNODE_BASEcollection(element.detail).sizeA - element.sizeB) < 4 And Abs(baseHNODE.HNODE_BASEcollection(element.detail).sizeB - element.sizeA) < 4 Then
        element.STATUSsize = "OK"
    Else
        element.STATUSsize = "FAIL"
    End If
    If cTN(baseHNODE.HNODE_BASEcollection(element.detail).thickness) = element.thickness Then element.STATUSthickness = "OK" Else element.STATUSthickness = "FAIL"
    If baseHNODE.HNODE_BASEcollection(element.detail).toughness = element.toughness Then element.STATUStoughness = "OK" Else element.STATUStoughness = "FAIL"
    If baseHNODE.HNODE_BASEcollection(element.detail).material = replaceMinus6(element.material) Then element.STATUSmaterial = "OK" Else element.STATUSmaterial = "FAIL"
Next element
End Sub
Sub SECONDARYFRcompare()
Dim element As Variant
For Each element In SECONDARYFR.SECONDARYFRcollection
    If basePG.PG_BASEcollection(element.item).designation = element.designation Then element.STATUSprofile = "OK" Else element.STATUSprofile = "FAIL"
    If basePG.PG_BASEcollection(element.item).material = replaceMinus6(element.material) Then element.STATUSmaterial = "OK" Else element.STATUSmaterial = "FAIL"
    If basePG.PG_BASEcollection(element.item).toughness = element.toughness Then element.STATUStoughness = "OK" Else element.STATUStoughness = "FAIL"
Next element
End Sub

