VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HNODEfields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private AAdetail, AAguide, AAsizeA, AAsizeB, AAtoughness, AAthickness, AAmaterial, AAlevel As String
Private AASTATUSsize, AASTATUStoughness, AASTATUSthickness, AASTATUSmaterial As String

Property Get detail() As String
detail = AAdetail
End Property
Property Let detail(value As String)
AAdetail = value
End Property

Property Get sizeA() As String
sizeA = AAsizeA
End Property
Property Let sizeA(value As String)
AAsizeA = value
End Property
Property Get sizeB() As String
sizeB = AAsizeB
End Property
Property Let sizeB(value As String)
AAsizeB = value
End Property
Property Get guide() As String
guide = AAguide
End Property
Property Let guide(value As String)
AAguide = value
End Property

Property Get toughness() As String
toughness = AAtoughness
End Property
Property Let toughness(value As String)
AAtoughness = value
End Property

Property Get thickness() As String
thickness = AAthickness
End Property
Property Let thickness(value As String)
AAthickness = value
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
Property Get STATUSsize() As String
STATUSsize = AASTATUSsize
End Property
Property Let STATUSsize(value As String)
AASTATUSsize = value
End Property

Property Get STATUStoughness() As String
STATUStoughness = AASTATUStoughness
End Property
Property Let STATUStoughness(value As String)
AASTATUStoughness = value
End Property

Property Get STATUSthickness() As String
STATUSthickness = AASTATUSthickness
End Property
Property Let STATUSthickness(value As String)
AASTATUSthickness = value
End Property
Property Get STATUSmaterial() As String
STATUSmaterial = AASTATUSmaterial
End Property
Property Let STATUSmaterial(value As String)
AASTATUSmaterial = value
End Property

