VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "baseSTIFFJclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private p_STIFFJ_BASEcollection As Collection
Private p_SampleField As STIFFfields
Private AAlevelA, AAlevelB, AAlevelC, AAlevelD, AAlevelE, AAlevelF, AAlevelG, AAlevelH, AAlevelI, AAlevelJ As String
'Initialize this class with it's collection:
Private Sub Class_Initialize()
    Set p_STIFFJ_BASEcollection = New Collection
End Sub

'Allow for adding SampleSTIFFfields:
Public Sub AddField(field As STIFFfields, Optional key As String)
    Set p_SampleField = field
    If key <> "" Then p_STIFFJ_BASEcollection.Add field, key Else p_STIFFJ_BASEcollection.Add field
End Sub

'Expose the collection:
Public Property Get STIFFJ_BASEcollection() As Collection
    Set STIFFJ_BASEcollection = p_STIFFJ_BASEcollection
End Property

Property Get levelA() As String
levelA = AAlevelA
End Property
Property Let levelA(value As String)
AAlevelA = value
End Property
Property Get LevelB() As String
LevelB = AAlevelB
End Property
Property Let LevelB(value As String)
AAlevelB = value
End Property
Property Get LevelC() As String
LevelC = AAlevelC
End Property
Property Let LevelC(value As String)
AAlevelC = value
End Property
Property Get LevelD() As String
LevelD = AAlevelD
End Property
Property Let LevelD(value As String)
AAlevelD = value
End Property
Property Get LevelE() As String
LevelE = AAlevelE
End Property
Property Let LevelE(value As String)
AAlevelE = value
End Property
Property Get LevelF() As String
LevelF = AAlevelF
End Property
Property Let LevelF(value As String)
AAlevelF = value
End Property
Property Get LevelG() As String
LevelG = AAlevelG
End Property
Property Let LevelG(value As String)
AAlevelG = value
End Property
Property Get LevelH() As String
LevelH = AAlevelH
End Property
Property Let LevelH(value As String)
AAlevelH = value
End Property
Property Get LevelI() As String
LevelI = AAlevelI
End Property
Property Let LevelI(value As String)
AAlevelI = value
End Property
Property Get LevelJ() As String
LevelJ = AAlevelJ
End Property
Property Let LevelJ(value As String)
AAlevelJ = value
End Property
