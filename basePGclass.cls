VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "basePGclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private p_PG_BASEcollection As Collection
Private p_SampleField As PGfields

'Initialize this class with it's collection:
Private Sub Class_Initialize()
    Set p_PG_BASEcollection = New Collection
End Sub

'Allow for adding SamplePGfields:
Public Sub AddField(field As PGfields, Optional key As String)
    Set p_SampleField = field
    If key <> "" Then p_PG_BASEcollection.Add field, key Else p_PG_BASEcollection.Add field
End Sub

'Expose the collection:
Public Property Get PG_BASEcollection() As Collection
    Set PG_BASEcollection = p_PG_BASEcollection
End Property
