VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3525
   LinkTopic       =   "Form1"
   ScaleHeight     =   1380
   ScaleWidth      =   3525
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   255
      TabIndex        =   1
      Top             =   660
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GET IP"
      Height          =   375
      Left            =   690
      TabIndex        =   0
      Top             =   120
      Width           =   1020
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2760
      Top             =   285
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = Winsock1.LocalIP
End Sub
