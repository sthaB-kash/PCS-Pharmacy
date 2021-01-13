VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormUpdate 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7995
   ClientLeft      =   6405
   ClientTop       =   2235
   ClientWidth     =   9495
   DrawMode        =   6  'Mask Pen Not
   DrawWidth       =   5
   FillColor       =   &H00FF0000&
   FillStyle       =   6  'Cross
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "FormUpdate.frx":0000
   ScaleHeight     =   7995
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFC0C0&
      Height          =   450
      Left            =   7320
      TabIndex        =   29
      Top             =   1800
      Width           =   1815
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
      Left            =   7800
      TabIndex        =   28
      Text            =   "name"
      Top             =   1080
      Width           =   1335
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
      Left            =   2040
      TabIndex        =   4
      Text            =   "name"
      Top             =   3120
      Width           =   4095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7080
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker MFD 
      Height          =   375
      Left            =   7320
      TabIndex        =   23
      Top             =   4080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16761024
      CalendarTitleBackColor=   -2147483638
      Format          =   119013377
      CurrentDate     =   43595
      MaxDate         =   58806
      MinDate         =   36161
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
      Left            =   7320
      TabIndex        =   22
      Text            =   "name"
      Top             =   3360
      Width           =   1815
   End
   Begin VB.ComboBox QTY 
      BackColor       =   &H00FFC0C0&
      Height          =   450
      Left            =   7320
      TabIndex        =   21
      Top             =   2520
      Width           =   1815
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
      Left            =   2040
      TabIndex        =   20
      Text            =   "name"
      Top             =   5280
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
      Left            =   2040
      TabIndex        =   19
      Text            =   "name"
      Top             =   4560
      Width           =   4095
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
      Left            =   2040
      TabIndex        =   5
      Text            =   "name"
      Top             =   3840
      Width           =   4095
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
      Left            =   2760
      TabIndex        =   3
      Text            =   "name"
      Top             =   1800
      Width           =   3375
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
      Left            =   2760
      TabIndex        =   2
      Text            =   "name"
      Top             =   1080
      Width           =   3375
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7080
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker EXP 
      Height          =   375
      Left            =   7320
      TabIndex        =   24
      Top             =   4800
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16761024
      CalendarTitleBackColor=   -2147483638
      Format          =   119013377
      CurrentDate     =   43595
      MaxDate         =   58806
      MinDate         =   36161
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
      Left            =   6360
      TabIndex        =   27
      Top             =   1800
      Width           =   765
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
      Left            =   6360
      TabIndex        =   26
      Top             =   1080
      Width           =   1365
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
      Left            =   3000
      TabIndex        =   18
      Top             =   6000
      Width           =   1065
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
      Left            =   1920
      TabIndex        =   17
      Top             =   6480
      Width           =   1515
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
      Left            =   6360
      TabIndex        =   16
      Top             =   4080
      Width           =   825
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
      Left            =   6360
      TabIndex        =   15
      Top             =   4800
      Width           =   705
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
      Left            =   360
      TabIndex        =   14
      Top             =   1800
      Width           =   1950
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
      Left            =   360
      TabIndex        =   13
      Top             =   2520
      Width           =   2985
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
      Left            =   360
      TabIndex        =   12
      Top             =   3840
      Width           =   1245
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
      Left            =   360
      TabIndex        =   11
      Top             =   4560
      Width           =   1215
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
      Left            =   6360
      TabIndex        =   10
      Top             =   2520
      Width           =   765
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
      Left            =   360
      TabIndex        =   9
      Top             =   5280
      Width           =   915
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
      Left            =   360
      TabIndex        =   8
      Top             =   3120
      Width           =   885
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
      Left            =   6360
      TabIndex        =   7
      Top             =   3360
      Width           =   825
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
      Left            =   360
      TabIndex        =   6
      Top             =   1080
      Width           =   2205
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
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   4080
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   5  'Dash-Dot-Dot
      BorderWidth     =   5
      DrawMode        =   4  'Mask Not Pen
      FillColor       =   &H00FF0000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   7695
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   9255
   End
End
Attribute VB_Name = "FormUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
Admin.Show
Admin.Enabled = True
Admin.DataGrid1.Enabled = True
FormUpdate.Hide
End Sub

Private Sub cmdOk_Click()
Admin.Enabled = True
Admin.Show
Admin.DataGrid1.Enabled = True
'FormUpdate.Show 0
'Admin.Enabled = False

Admin.Text1.Text = GN.Text
Admin.Text2.Text = BN.Text
Admin.Text3.Text = Text1.Text
Admin.Text4.Text = Address.Text
Admin.Text5.Text = Contact.Text
Admin.Text6.Text = Email.Text
Admin.Text7.Text = QTY.Text
Admin.Text8.Text = Price.Text
Admin.Text9.Text = MFD.value
Admin.Text10.Text = EXP.value
Admin.Text13.Text = Now
Admin.Text14.Text = "Admin"
Admin.Text15.Text = Text2.Text
Admin.Text16.Text = Combo1.Text
Admin.Adodc1.Recordset.Update
MsgBox "Updated Successfully..", vbInformation
Unload FormUpdate
FormUpdate.Hide
'Admin.Adodc1.Recordset.Bookmark = Admin.DataGrid1.SelBookmarks
'Admin.Adodc1.Refresh
'Admin.DataGrid1.Refresh
Admin.DataGrid1.AllowUpdate = True
Admin.GridUpdate.SetFocus
'showAdmin
'Admin.GridReturn.SetFocus
End Sub
Sub display()
GN.Text = Admin.Rs.Fields("MedicineName")
BN.Text = Admin.Rs.Fields("BrandName")
'Name.Text = Admin.Rs.Fields("Mfdname")
Address.Text = Admin.Rs.Fields("address")
Contact.Text = Admin.Rs.Fields("contact")
Email.Text = Admin.Rs.Fields("email")
QTY.Text = Admin.Rs.Fields("qty")
Price.Text = Admin.Rs.Fields("Price")
MFD.value = Admin.Rs.Fields("mfd")
EXP.value = Admin.Rs.Fields("exp")
If Admin.Rs.Fields("DOU") = "" Then
    Label16.Caption = Label16.Caption & "None " & "By:none"
End If
 
Label13.Caption = Label13.Caption + Str(Admin.Rs.Fields("DOE"))
Label13.Caption = Label13.Caption + "by" + Admin.Rs.Fields("InsertedBy")
 
End Sub

Private Sub Form_Load()
Label1.Left = FormUpdate.Width / 2 - Label1.Width / 2
'Label16.Left = FormUpdate.Width / 2 - Label16.Width / 2
'cmdOk.Left = FormUpdate.Width / 2 - cmdOk.Width / 2
'Label13.Left = FormUpdate.Width / 2 - Label13.Width / 2
Combo1.AddItem "Tablet"
Combo1.AddItem "Liquid"
Combo1.AddItem "Surgical Equipment"
Combo1.AddItem "Capsule"
Combo1.AddItem "Harbal Product"
Combo1.AddItem "Topical Medicines"
Combo1.AddItem "Powder"
Combo1.AddItem "Antibiotics"
Combo1.AddItem "Antidepressants"
Combo1.AddItem "Anxiety Medications"
Combo1.AddItem "Antifungals"
Combo1.AddItem "Antiviral Drugs"
Combo1.AddItem "Antiparasitic Drugs"
Combo1.AddItem "Sleeping Pills"
Combo1.AddItem "Pain Killer"
Combo1.AddItem "Opioids"
Combo1.AddItem "Medications for Hypertension"
Combo1.AddItem "Medications for diabetes"

Dim i As Integer
For i = 1 To 100
 QTY.AddItem i
Next
End Sub

Private Sub Form_LostFocus()
'cmdOk.SetFocus
Beep
End Sub

