VERSION 5.00
Object = "{38EE5CE1-4B62-11D3-854F-00A0C9C898E7}#1.0#0"; "MSWEBDVD.DLL"
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   7875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9660
   LinkTopic       =   "Form2"
   ScaleHeight     =   7875
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWEBDVDLibCtl.MSWebDVD DVD1 
      Height          =   4695
      Left            =   1800
      TabIndex        =   0
      Top             =   1560
      Width           =   6375
      _cx             =   11245
      _cy             =   8281
      DisableAutoMouseProcessing=   0   'False
      BackColor       =   1048592
      EnableResetOnStop=   0   'False
      ColorKey        =   1048592
      WindowlessActivation=   0   'False
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Width = Screen.Width
Me.Height = Screen.Height - Form1.Height
Me.Top = 0
Me.Left = 0
DVD1.Left = 10
DVD1.Top = 10
DVD1.Width = Me.Width - 20
DVD1.Height = Me.Height - 20

End Sub
