VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Language"
   ClientHeight    =   540
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3210
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   540
   ScaleWidth      =   3210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   300
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
selec = Combo1.ListIndex + 1
Form2.DVD1.CurrentSubpictureStream = selec
Debug.Print Combo1.ListIndex + 1
Unload Me
End Sub


Private Sub Form_Load()
Me.Left = Form1.Left + ((Form1.Width / 2) - Me.Width / 2)
Me.Top = (Form1.Top + Form1.Height) - Me.Height
selec = Form2.DVD1.CurrentSubpictureStream
For a = 1 To SubCount
Combo1.AddItem (SubTitleLang(a))
Next a
Combo1.ListIndex = selec - 1
End Sub
