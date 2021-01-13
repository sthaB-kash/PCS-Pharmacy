VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = "001"
Dim y, m, v As String
y = Str(Format(Date, "yy"))
y = Format(y, "-000")
m = Str(Format(Date, "mm"))
m = Format(m, "-00")
v = Format(Val(Text1.Text) + 1, "0000")
Text1.Text = v + m + y
End Sub
