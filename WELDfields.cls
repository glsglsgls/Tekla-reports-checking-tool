VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WELDfields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private AAweldnumber, AAweldbooklet, AAdetailsNUM, AAGUID, AAweldtype, AAweldsize, AAweldfinish, AAweldjointtype, AAweldbeveltype, AAweldNDT, AAweldANGLE, AASTATUSweldtype, AASTATUSweldsize, AASTATUSweldfinish, AASTATUSweldjointtype, AASTATUSweldbeveltype, AASTATUSweldNDT, AASTATUSweldbooklet, AASTATUSweldANGLE

Property Get weldnumber() As String
weldnumber = AAweldnumber
End Property
Property Let weldnumber(value As String)
AAweldnumber = value
End Property
Property Get detailsNUM() As String
detailsNUM = AAdetailsNUM
End Property
Property Let detailsNUM(value As String)
AAdetailsNUM = value
End Property
Property Get GUID() As String
GUID = AAGUID
End Property
Property Let GUID(value As String)
AAGUID = value
End Property
Property Get weldbooklet() As String
weldbooklet = AAweldbooklet
End Property
Property Let weldbooklet(value As String)
AAweldbooklet = value
End Property
Property Get weldtype() As String
weldtype = AAweldtype
End Property
Property Let weldtype(value As String)
AAweldtype = value
End Property
Property Get weldsize() As String
weldsize = AAweldsize
End Property
Property Let weldsize(value As String)
AAweldsize = value
End Property
Property Get weldfinish() As String
weldfinish = AAweldfinish
End Property
Property Let weldfinish(value As String)
AAweldfinish = value
End Property
Property Get weldjointtype() As String
weldjointtype = AAweldjointtype
End Property
Property Let weldjointtype(value As String)
AAweldjointtype = value
End Property
Property Get weldbeveltype() As String
weldbeveltype = AAweldbeveltype
End Property
Property Let weldbeveltype(value As String)
AAweldbeveltype = value
End Property
Property Get weldNDT() As String
weldNDT = AAweldNDT
End Property
Property Let weldNDT(value As String)
AAweldNDT = value
End Property
Property Get weldANGLE() As String
weldANGLE = AAweldANGLE
End Property
Property Let weldANGLE(value As String)
AAweldANGLE = value
End Property

Property Get STATUSweldtype() As String
STATUSweldtype = AASTATUSweldtype
End Property
Property Let STATUSweldtype(value As String)
AASTATUSweldtype = value
End Property
Property Get STATUSweldsize() As String
STATUSweldsize = AASTATUSweldsize
End Property
Property Let STATUSweldsize(value As String)
AASTATUSweldsize = value
End Property
Property Get STATUSweldfinish() As String
STATUSweldfinish = AASTATUSweldfinish
End Property
Property Let STATUSweldfinish(value As String)
AASTATUSweldfinish = value
End Property
Property Get STATUSweldjointtype() As String
STATUSweldjointtype = AASTATUSweldjointtype
End Property
Property Let STATUSweldjointtype(value As String)
AASTATUSweldjointtype = value
End Property
Property Get STATUSweldbeveltype() As String
STATUSweldbeveltype = AASTATUSweldbeveltype
End Property
Property Let STATUSweldbeveltype(value As String)
AASTATUSweldbeveltype = value
End Property
Property Get STATUSweldNDT() As String
STATUSweldNDT = AASTATUSweldNDT
End Property
Property Let STATUSweldNDT(value As String)
AASTATUSweldNDT = value
End Property
Property Get STATUSweldbooklet() As String
STATUSweldbooklet = AASTATUSweldbooklet
End Property
Property Let STATUSweldbooklet(value As String)
AASTATUSweldbooklet = value
End Property
Property Get STATUSweldANGLE() As String
STATUSweldANGLE = AASTATUSweldANGLE
End Property
Property Let STATUSweldANGLE(value As String)
AASTATUSweldANGLE = value
End Property
