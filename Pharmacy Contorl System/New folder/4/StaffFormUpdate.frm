VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form StaffFormUpdate 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7875
   ClientLeft      =   5760
   ClientTop       =   1275
   ClientWidth     =   9435
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H0000FF00&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox GN 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   2640
      TabIndex        =   11
      Text            =   "name"
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox BN 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   2640
      TabIndex        =   10
      Text            =   "name"
      Top             =   1680
      Width           =   3375
   End
   Begin VB.TextBox Address 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   1920
      TabIndex        =   9
      Text            =   "name"
      Top             =   3720
      Width           =   4095
   End
   Begin VB.TextBox Contact 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   1920
      TabIndex        =   8
      Text            =   "name"
      Top             =   4440
      Width           =   4095
   End
   Begin VB.TextBox Email 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   1920
      TabIndex        =   7
      Text            =   "name"
      Top             =   5160
      Width           =   4095
   End
   Begin VB.ComboBox QTY 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7200
      TabIndex        =   6
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Price 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   7200
      TabIndex        =   5
      Text            =   "name"
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FF80FF&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6960
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   1920
      TabIndex        =   2
      Text            =   "name"
      Top             =   3000
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   7680
      TabIndex        =   1
      Text            =   "name"
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7200
      TabIndex        =   0
      Top             =   1680
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker MFD 
      Height          =   375
      Left            =   7200
      TabIndex        =   4
      Top             =   3960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16761024
      CalendarTitleBackColor=   -2147483638
      Format          =   118816769
      CurrentDate     =   43595
      MaxDate         =   58806
      MinDate         =   36161
   End
   Begin MSComCtl2.DTPicker EXP 
      Height          =   375
      Left            =   7200
      TabIndex        =   13
      Top             =   4680
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16761024
      CalendarTitleBackColor=   -2147483638
      Format          =   118816769
      CurrentDate     =   43595
      MaxDate         =   58806
      MinDate         =   36161
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<<Updating Records>>"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   1800
      TabIndex        =   29
      Top             =   240
      Width           =   4080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Generic Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   240
      TabIndex        =   28
      Top             =   960
      Width           =   2205
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   6240
      TabIndex        =   27
      Top             =   3240
      Width           =   825
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   240
      TabIndex        =   26
      Top             =   3000
      Width           =   885
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   240
      TabIndex        =   25
      Top             =   5160
      Width           =   915
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QTY"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   6240
      TabIndex        =   24
      Top             =   2400
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   240
      TabIndex        =   23
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   240
      TabIndex        =   22
      Top             =   3720
      Width           =   1245
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company's Details:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   240
      TabIndex        =   21
      Top             =   2400
      Width           =   2985
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Brand Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   240
      TabIndex        =   20
      Top             =   1680
      Width           =   1950
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EXP"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   6240
      TabIndex        =   19
      Top             =   4680
      Width           =   705
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MFD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   6240
      TabIndex        =   18
      Top             =   3960
      Width           =   825
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last updated on:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1800
      TabIndex        =   17
      Top             =   6360
      Width           =   1515
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inserted on:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2880
      TabIndex        =   16
      Top             =   5880
      Width           =   1065
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Batch no"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   6240
      TabIndex        =   15
      Top             =   960
      Width           =   1365
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   390
      Left            =   6240
      TabIndex        =   14
      Top             =   1680
      Width           =   765
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   5  'Dash-Dot-Dot
      BorderWidth     =   5
      DrawMode        =   4  'Mask Not Pen
      FillColor       =   &H00FF0000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   7695
      Left            =   70
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   9255
   End
End
Attribute VB_Name = "StaffFormUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
Staff.Show
Staff.Enabled = True
Staff.DataGrid1.Enabled = True
FormUpdate.Hide
End Sub

Private Sub cmdOk_Click()
Staff.Enabled = True
Staff.Show
Staff.DataGrid1.Enabled = True
'FormUpdate.Show 0
'Staff.Enabled = False

Staff.Text1.Text = GN.Text
Staff.Text2.Text = BN.Text
Staff.Text3.Text = Text1.Text
Staff.Text4.Text = Address.Text
Staff.Text5.Text = Contact.Text
Staff.Text6.Text = Email.Text
Staff.Text7.Text = QTY.Text
Staff.Text8.Text = Price.Text
Staff.Text9.Text = MFD.Value
Staff.Text10.Text = EXP.Value
Staff.Text13.Text = Now
Staff.Text14.Text = "Staff"
Staff.Text15.Text = Text2.Text
Staff.Text16.Text = Combo1.Text
Staff.Adodc1.Recordset.Update
MsgBox "Updated Successfully..", vbInformation
Unload FormUpdate
FormUpdate.Hide
'Staff.Adodc1.Recordset.Bookmark = Staff.DataGrid1.SelBookmarks
'Staff.Adodc1.Refresh
'Staff.DataGrid1.Refresh
Staff.DataGrid1.AllowUpdate = True
Staff.GridUpdate.SetFocus
'showStaff
'Staff.GridReturn.SetFocus
End Sub
Sub display()
GN.Text = Staff.Rs.Fields("MedicineName")
BN.Text = Staff.Rs.Fields("BrandName")
'Name.Text = Staff.Rs.Fields("Mfdname")
Address.Text = Staff.Rs.Fields("address")
Contact.Text = Staff.Rs.Fields("contact")
Email.Text = Staff.Rs.Fields("email")
QTY.Text = Staff.Rs.Fields("qty")
Price.Text = Staff.Rs.Fields("Price")
MFD.Value = Staff.Rs.Fields("mfd")
EXP.Value = Staff.Rs.Fields("exp")
If Staff.Rs.Fields("DOU") = "" Then
    Label16.Caption = Label16.Caption & "None " & "By:none"
End If
 
Label13.Caption = Label13.Caption + Str(Staff.Rs.Fields("DOE"))
Label13.Caption = Label13.Caption + "by" + Staff.Rs.Fields("InsertedBy")
 
End Sub

Private Sub Form_Load()
Label1.Left = FormUpdate.Width / 2 - Label1.Width / 2
'Label16.Left = FormUpdate.Width / 2 - Label16.Width / 2
'cmdOk.Left = FormUpdate.Width / 2 - cmdOk.Width / 2
'Label13.Left = FormUpdate.Width / 2 - Label13.Width / 2
Combo1.AddItem "Tablet"
Combo1.AddItem "Powder"
Combo1.AddItem "Liquid"
Combo1.AddItem "Capsule"

Dim i As Integer
For i = 1 To 100
 QTY.AddItem i
Next
End Sub

Private Sub Form_LostFocus()
'cmdOk.SetFocus
Beep
End Sub


