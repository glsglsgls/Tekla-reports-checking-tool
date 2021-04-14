VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TRANSfields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private AAdetail, AAguide, AAwidth, AAlength, AAtoughness, AAthickness, AAmaterial As String
Private AASTATUSwidth, AASTATUSlength, AASTATUStoughness, AASTATUSthickness, AASTATUSmaterial As String

Property Get detail() As String
detail = AAdetail
End Property
Property Let detail(value As String)
AAdetail = value
End Property
Property Get width() As String
width = AAwidth
End Property
Property Let width(value As String)
AAwidth = value
End Property
Property Get length() As String
length = AAlength
End Property
Property Let length(value As String)
AAlength = value
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
Property Get toughness() As String
toughness = AAtoughness
End Property
Property Let toughness(value As String)
AAtoughness = value
End Property
Property Get guide() As String
guide = AAguide
End Property
Property Let guide(value As String)
AAguide = value
End Property


Property Get STATUSwidth() As String
STATUSwidth = AASTATUSwidth
End Property
Property Let STATUSwidth(value As String)
AASTATUSwidth = value
End Property
Property Get STATUSlength() As String
STATUSlength = AASTATUSlength
End Property
Property Let STATUSlength(value As String)
AASTATUSlength = value
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
Property Get STATUStoughness() As String
STATUStoughness = AASTATUStoughness
End Property
Property Let STATUStoughness(value As String)
AASTATUStoughness = value
End Property


