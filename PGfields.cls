VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PGfields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private AAitem, AAdes, AAassembly, AAmaterial, AAtoughness, AASTATUSprofile, AASTATUSmaterial, AASTATUStoughness, AASTATUSlength As String
Property Get item() As String
item = AAitem
End Property
Property Let item(value As String)
AAitem = value
End Property
Property Get designation() As String
designation = AAdes
End Property
Property Let designation(value As String)
AAdes = value
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
Property Get assembly() As String
assembly = AAassembly
End Property
Property Let assembly(value As String)
AAassembly = value
End Property
Property Get STATUSprofile() As String
STATUSprofile = AASTATUSprofile
End Property
Property Let STATUSprofile(value As String)
AASTATUSprofile = value
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
Property Get STATUSlength() As String
STATUSlength = AASTATUSlength
End Property
Property Let STATUSlength(value As String)
AASTATUSlength = value
End Property
