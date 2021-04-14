VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "baseSTIFFspclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private p_STIFFsp_BASEcollection As Collection
Private p_SampleField As STIFFfields

'Initialize this class with it's collection:
Private Sub Class_Initialize()
    Set p_STIFFsp_BASEcollection = New Collection
End Sub

'Allow for adding SampleSTIFFfields:
Public Sub AddField(field As STIFFfields, Optional key As String)
    Set p_SampleField = field
    If key <> "" Then p_STIFFsp_BASEcollection.Add field, key Else p_STIFFsp_BASEcollection.Add field
End Sub

'Expose the collection:
Public Property Get STIFFSP_BASEcollection() As Collection
    Set STIFFSP_BASEcollection = p_STIFFsp_BASEcollection
End Property

