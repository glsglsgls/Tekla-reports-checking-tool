VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MAINform 
   Caption         =   "Report dialogbox"
   ClientHeight    =   7290
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11415
   OleObjectBlob   =   "MAINform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MAINform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub STARTbutton_Click()
    ZFunctions.prepON
    Call Methods.openandbindworkbooks
End Sub

Private Sub SELECTbutton_Click()
    Call Methods.clearselection
    Call Methods.insertFILESinTEXTBOXES
End Sub

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Top = 240 + Application.Top
    Me.Left = 40 + Application.Left
    Dim sheet As Worksheet
    For Each sheet In ThisWorkbook.Sheets
        If sheet.name Like "?TM??" Or sheet.name Like "?tm??" Then
            mainFORM.ComboBox1.AddItem sheet.name
        End If
    Next sheet
End Sub

Private Sub UserForm_Terminate()
    ZFunctions.prepON
    End
End Sub
