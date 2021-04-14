VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NODEfields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private AAdetail, AAguide, AAAsizeA, AAAsizeB, AABsizeA, AABsizeB, AAAtoughness, AABtoughness, AABthickness, AAAthickness, AAmaterial, AAlevel As String
Private AASTATUSAsize, AASTATUSBsize, AASTATUSAtoughness, AASTATUSBtoughness, AASTATUSAthickness, AASTATUSBthickness, AASTATUSmaterial As String

Property Get detail() As String
detail = AAdetail
End Property
Property Let detail(value As String)
AAdetail = value
End Property

Property Get AsizeA() As String
AsizeA = AAAsizeA
End Property
Property Let AsizeA(value As String)
AAAsizeA = value
End Property
Property Get AsizeB() As String
AsizeB = AAAsizeB
End Property
Property Let AsizeB(value As String)
AAAsizeB = value
End Property
Property Get guide() As String
guide = AAguide
End Property
Property Let guide(value As String)
AAguide = value
End Property
Property Get BsizeA() As String
BsizeA = AABsizeA
End Property
Property Let BsizeA(value As String)
AABsizeA = value
End Property
Property Get BsizeB() As String
BsizeB = AABsizeB
End Property
Property Let BsizeB(value As String)
AABsizeB = value
End Property

Property Get Atoughness() As String
Atoughness = AAAtoughness
End Property
Property Let Atoughness(value As String)
AAAtoughness = value
End Property
Property Get Btoughness() As String
Btoughness = AABtoughness
End Property
Property Let Btoughness(value As String)
AABtoughness = value
End Property

Property Get Athickness() As String
Athickness = AAAthickness
End Property
Property Let Athickness(value As String)
AAAthickness = value
End Property
Property Get Bthickness() As String
Bthickness = AABthickness
End Property
Property Let Bthickness(value As String)
AABthickness = value
End Property

Property Get material() As String
material = AAmaterial
End Property
Property Let material(value As String)
AAmaterial = value
End Property

Property Get level() As String
level = AAlevel
End Property
Property Let level(value As String)
AAlevel = value
End Property
Property Get STATUSAsize() As String
STATUSAsize = AASTATUSAsize
End Property
Property Let STATUSAsize(value As String)
AASTATUSAsize = value
End Property
Property Get STATUSBsize() As String
STATUSBsize = AASTATUSBsize
End Property
Property Let STATUSBsize(value As String)
AASTATUSBsize = value
End Property
Property Get STATUSAtoughness() As String
STATUSAtoughness = AASTATUSAtoughness
End Property
Property Let STATUSAtoughness(value As String)
AASTATUSAtoughness = value
End Property
Property Get STATUSBtoughness() As String
STATUSBtoughness = AASTATUSBtoughness
End Property
Property Let STATUSBtoughness(value As String)
AASTATUSBtoughness = value
End Property
Property Get STATUSAthickness() As String
STATUSAthickness = AASTATUSAthickness
End Property
Property Let STATUSAthickness(value As String)
AASTATUSAthickness = value
End Property
Property Get STATUSBthickness() As String
STATUSBthickness = AASTATUSBthickness
End Property
Property Let STATUSBthickness(value As String)
AASTATUSBthickness = value
End Property
Property Get STATUSmaterial() As String
STATUSmaterial = AASTATUSmaterial
End Property
Property Let STATUSmaterial(value As String)
AASTATUSmaterial = value
End Property
