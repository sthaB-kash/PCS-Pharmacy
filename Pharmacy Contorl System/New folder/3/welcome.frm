VERSION 5.00
Begin VB.Form welcome 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   Caption         =   "Dim bikash As Variant"
   ClientHeight    =   5310
   ClientLeft      =   6405
   ClientTop       =   3195
   ClientWidth     =   8595
   ForeColor       =   &H8000000A&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   Begin VB.Timer loading 
      Interval        =   1
      Left            =   6720
      Top             =   4920
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer4 
      Interval        =   1200
      Left            =   7080
      Top             =   4920
   End
   Begin VB.Timer Timer3 
      Interval        =   1010
      Left            =   7440
      Top             =   4920
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   7800
      Top             =   4920
   End
   Begin VB.Timer Timer1 
      Interval        =   6000
      Left            =   8160
      Top             =   4920
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      Height          =   255
      Index           =   10
      Left            =   -240
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      Height          =   255
      Index           =   9
      Left            =   -240
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      Height          =   255
      Index           =   8
      Left            =   -240
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      Height          =   255
      Index           =   7
      Left            =   -240
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      Height          =   255
      Index           =   6
      Left            =   240
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      Height          =   255
      Index           =   5
      Left            =   480
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      Height          =   255
      Index           =   4
      Left            =   240
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      Height          =   255
      Index           =   3
      Left            =   240
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      Height          =   255
      Index           =   2
      Left            =   360
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      Height          =   255
      Index           =   1
      Left            =   240
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      Height          =   255
      Index           =   0
      Left            =   240
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   7680
      Picture         =   "welcome.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   8160
      Picture         =   "welcome.frx":0368
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Health Is Our Priority"
      BeginProperty Font 
         Name            =   "Blackadder ITC"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   2040
      TabIndex        =   3
      Top             =   3960
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PHRMACY CONTROL SYSTEM"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   18
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   360
      Left            =   1320
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   5940
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TO"
      BeginProperty Font 
         Name            =   "Castellar"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   585
      Left            =   4035
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME "
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   585
      Left            =   3120
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   2520
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   5295
      Left            =   0
      Picture         =   "welcome.frx":098E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Enabled = True
'Image1.Visible = True
'Image2.Visible = True
'Text1.Text = "1"
'Text1.Visible = True
Dim i As Integer
'For i = 1 To 10
'Load Shape1(i)
Shape1(0).Left = 0 - Shape1(0).Width
'Shape1(i).Left = Shape1(i - 1).Left - Shape1(0).Width
'Shape1(i).Visible = True
'Next
Dim bikash As Variant
bikash = 2
  
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X >= 7650 And Y <= 345 Then
    Image1.Visible = True
    Image2.Visible = True
Else 'if x>
    Image1.Visible = False
    Image2.Visible = False
End If

End Sub

Private Sub Image1_Click()
End
End Sub

Private Sub Image2_Click()
welcome.WindowState = 1
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Text1.Text = Str(X)
'Text2.Text = Str(Y)
If X >= 7650 And Y <= 345 Then
    Image1.Visible = True
    Image2.Visible = True
Else 'if x>
    Image1.Visible = False
    Image2.Visible = False
End If
End Sub


Private Sub loading_Timer()
For i = 0 To 6
    Shape1(i).Left = Shape1(i).Left + (bikash * 2)
    bikash = bikash + 20

    If Shape1(i).Left > 4870 Then
    For j = 0 To 6
        Shape1(i).Left = Shape1(i).Left - bikash - 80
    Next
    End If
Next
'Shape1(0).Left = Shape1(0).Left + 10
End Sub

Private Sub Timer1_Timer()
welcome.Hide
loginform.Show
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
loading.Enabled = False
End Sub

Private Sub Timer2_Timer()
'Text1.Text = Str(1 + Val(Text1.Text))
If Label1.Left > 3129 Then
    Label1.Left = Label1.Left - 600
End If

If Label2.Left < 4035 Then
    Label2.Left = Label2.Left + 200
End If

If Label3.Left > 1320 Then
    Label3.Left = Label3.Left - 580
End If

If Label4.Top > 3960 Then
    Label4.Top = Label4.Top - 200
End If
End Sub

Private Sub Timer3_Timer()
Image3.Picture = LoadPicture("E:\Project 3\program_rx_3000.jpg")
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label1.Left = welcome.ScaleWidth
Label2.Left = welcome.ScaleLeft - Label2.Width
Label3.Left = welcome.ScaleWidth
Label4.Top = welcome.ScaleHeight
Timer3.Enabled = False
 
End Sub

Private Sub Timer4_Timer()
Timer2.Enabled = False
If Not Label1.Left = 3130 Then
    Label1.Left = 3130
End If

If Not Label2.Left = 4035 Then
    Label2.Left = 4035
End If

If Not Label3.Left = 1320 Then
    Label3.Left = 1320
End If
Timer4.Enabled = False
Label4.Left = (welcome.ScaleWidth / 2) - (Label4.Width / 2)
End Sub


