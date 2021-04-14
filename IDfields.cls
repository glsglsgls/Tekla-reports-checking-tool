VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDfields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private AAdetail, AAprofile, AAmaterial, AAtoughness, AAIDcode, AASTATUSIDcode

Property Get detail() As String
detail = AAdetail
End Property
Property Let detail(value As String)
AAdetail = value
End Property
Property Get profile() As String
profile = AAprofile
End Property
Property Let profile(value As String)
AAprofile = value
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
Property Get IDcode() As String
IDcode = AAIDcode
End Property
Property Let IDcode(value As String)
AAIDcode = value
End Property
Property Get STATUSIDcode() As String
STATUSIDcode = AASTATUSIDcode
End Property
Property Let STATUSIDcode(value As String)
AASTATUSIDcode = value
End Property

