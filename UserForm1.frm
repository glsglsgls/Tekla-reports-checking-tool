VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Changelog window"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10005
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
UserForm1.Hide
End Sub


Private Sub LB11_Click()

End Sub

Private Sub LB12_Click()

End Sub

Private Sub UserForm_Deactivate()
End
End Sub

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Top = 240 + Application.Top
    Me.Left = 40 + Application.Left
With UserForm1.LB10
    .AddItem "first release"
End With

With UserForm1.LB11
    .AddItem "Added        -  added module 3TMP5"
    .AddItem "Added        -  added module 2TMP1"
    .AddItem "Changed     -  minor bugs fixed"
End With

With UserForm1.LB12
    .AddItem "INFORMATION  -  SIS has not been implemented!"
    .AddItem "INFORMATION  -  if SIS was applied to the item, 'manual check' will be shown"
    .AddItem ""
    .AddItem "ADDED               -  added module 3TMS5"
    .AddItem "ADDED               -  added module 3TMP2"
End With


End Sub
