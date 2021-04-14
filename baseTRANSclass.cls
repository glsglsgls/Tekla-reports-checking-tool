VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "baseTRANSclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private p_TRANS_BASEcollection As Collection
Private p_SampleField As TRANSfields
'Initialize this class with it's collection:
Private Sub Class_Initialize()
    Set p_TRANS_BASEcollection = New Collection
End Sub

'Allow for adding SampleTRANSfields:
Public Sub AddField(field As TRANSfields, Optional key As String)
    Set p_SampleField = field
    If key <> "" Then p_TRANS_BASEcollection.Add field, key Else p_TRANS_BASEcollection.Add field
End Sub

'Expose the collection:
Public Property Get TRANS_BASEcollection() As Collection
    Set TRANS_BASEcollection = p_TRANS_BASEcollection
End Property
