VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form loginform 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PHARMACY_CONTROL_SYSTEM"
   ClientHeight    =   5460
   ClientLeft      =   6135
   ClientTop       =   3585
   ClientWidth     =   8205
   FillColor       =   &H00E0E0E0&
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   8205
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   8160
      Top             =   2520
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\BCA\3rd SEMESTER\PROJECT-3\PHARMACY_CONTROL_SYSTEM.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\BCA\3rd SEMESTER\PROJECT-3\PHARMACY_CONTROL_SYSTEM.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Login"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtPw 
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   3480
      PasswordChar    =   "*"
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   3240
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox txtUser 
      Height          =   435
      Left            =   3480
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2520
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.ComboBox selectUser 
      Height          =   435
      ItemData        =   "loginfirm.frx":0000
      Left            =   3480
      List            =   "loginfirm.frx":0002
      TabIndex        =   1
      Text            =   "Select_User"
      Top             =   1680
      Width           =   3735
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton cmdLogin 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Login"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   2640
      Picture         =   "loginfirm.frx":0004
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   2640
      Picture         =   "loginfirm.frx":0677
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   2640
      Picture         =   "loginfirm.frx":0F41
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lbl_forgotPw 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forgot Password?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   4560
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   465
      Left            =   600
      TabIndex        =   6
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UserName"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   585
      Left            =   600
      TabIndex        =   5
      Top             =   2520
      Width           =   1965
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login As"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B&&Y PHARMACY"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   825
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   6090
   End
End
Attribute VB_Name = "loginform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CN As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim username As String
Dim pw As String

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdLogin_Click()
'If txtPw.Tag = txtPw.Text Then
Dim Response As Integer
 
    If txtPw.Text = "admin" Then
        loginform.Hide
        Admin.Show
        MsgBox "Successfully Logged-In", vbInformation, "Access Granted"
    ElseIf txtPw.Text = "staff" Then
        Staff.Show
        loginform.Hide
        MsgBox "Successfully Logged-In", vbInformation, "Access Granted"
    Else
        Response = MsgBox("Incorrect password", vbRetryCancel + vbCritical, "Access Denied")
        If Response = vbRetry Then
            txtPw.SelStart = 0
            txtPw.SelLength = Len(txtPw.Text)
        Else
             End
        End If
        'MsgBox "INVALID PASSWORD...", vbCritical, "Access DENIED"
    End If
'End If
End Sub

Private Sub Form_Load()
CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\BCA\3rd SEMESTER\PROJECT-3\PHARMACY_CONTROL_SYSTEM.mdb;Persist Security Info=False"
Rs.Open "select * from login", CN, adOpenDynamic, adLockPessimistic

selectUser.AddItem "Administrator"
selectUser.AddItem "Staff"
'selectUser.SetFocus
cmdLogin.Enabled = False
End Sub

Private Sub selectUser_Click()
If selectUser.Text = "Administrator" Then
    txtUser.Tag = "Admin"
     'pw = "rs!password"
     txtUser.Visible = True
     txtUser.SetFocus
     txtPw.Visible = False
     
ElseIf selectUser.Text = "Staff" Then
    txtUser.Tag = "Staff"
    txtUser.Visible = True
    txtUser.SetFocus
    txtPw.Visible = False
    
End If
txtUser.Visible = True
txtUser.Text = ""
End Sub

 
Private Sub txtPw_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdLogin_Click
End If
End Sub

Private Sub txtUser_Change()
If txtUser.Text = "Admin" And selectUser.Text = "Administrator" Then
    txtPw.Visible = True
    txtPw.Text = ""
    txtPw.Tag = "admin"
    txtPw.SetFocus
    lbl_forgotPw.Visible = True
    cmdLogin.Enabled = True
    
ElseIf txtUser.Text = "Staff" And selectUser.Text = "Staff" Then
    txtPw.Visible = True
    txtPw.Text = ""
    txtPw.Tag = "staff"
    txtPw.SetFocus
    lbl_forgotPw.Visible = True
    cmdLogin.Enabled = True
End If
End Sub
