VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Staff 
   BackColor       =   &H00404000&
   Caption         =   "FRONT-LINE_STAFF"
   ClientHeight    =   8145
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16755
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Staff.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   16755
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton GridUpdate 
      BackColor       =   &H0000C000&
      Caption         =   "Update"
      Height          =   615
      Left            =   5655
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   1095
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton GridDelete 
      BackColor       =   &H000000FF&
      Caption         =   "Delete"
      Height          =   615
      Left            =   3855
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   1095
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   19800
      Top             =   2520
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Staff.frx":038A
      Height          =   2175
      Left            =   135
      Negotiate       =   -1  'True
      TabIndex        =   5
      Top             =   2535
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   8454143
      ForeColor       =   16711680
      HeadLines       =   1
      RowHeight       =   26
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Details of Medicine"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text15 
      DataField       =   "Batch"
      DataSource      =   "Adodc1"
      Height          =   480
      Left            =   5295
      TabIndex        =   67
      Text            =   "Text15"
      Top             =   3730
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text16 
      DataField       =   "Type"
      DataSource      =   "Adodc1"
      Height          =   480
      Left            =   5415
      TabIndex        =   66
      Text            =   "Text16"
      Top             =   3130
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text14 
      DataField       =   "UpdatedBy"
      DataSource      =   "Adodc1"
      Height          =   600
      Left            =   5295
      TabIndex        =   65
      Text            =   "Text14"
      Top             =   4330
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdAddAnother 
      Caption         =   "Add &Another"
      Height          =   495
      Left            =   8895
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   7810
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFC0&
      Caption         =   "E&xit"
      Height          =   615
      Left            =   18495
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   10335
      Width           =   1815
   End
   Begin VB.CommandButton cmdInsert 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Insert &New"
      Height          =   600
      Left            =   2055
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   1095
      Width           =   1815
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16095
      Picture         =   "Staff.frx":039F
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   375
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H0000FF00&
      Caption         =   "Up&date"
      Height          =   615
      Left            =   5655
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   1095
      Width           =   1815
   End
   Begin VB.TextBox searchBox 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   480
      Left            =   12135
      TabIndex        =   56
      Text            =   "Enter the name of the medicine to search"
      Top             =   375
      Width           =   3855
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H000000FF&
      Caption         =   "Delete"
      Height          =   615
      Left            =   3855
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   1095
      Width           =   1815
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   "&Add"
      Height          =   495
      Left            =   9255
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   7810
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear All"
      Height          =   495
      Left            =   11295
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   7810
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtBName 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8055
      TabIndex        =   51
      Top             =   2530
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox txtCName 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8055
      TabIndex        =   50
      Top             =   3850
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox txtAddress 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8055
      TabIndex        =   49
      Top             =   4450
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox txtContact 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8055
      TabIndex        =   48
      Top             =   5050
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox txtEmail 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8055
      TabIndex        =   47
      Top             =   5650
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.ComboBox qty 
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
      Left            =   14055
      TabIndex        =   44
      Text            =   "Combo1"
      Top             =   3130
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtPrice 
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
      Left            =   14055
      TabIndex        =   43
      Text            =   "Text1"
      Top             =   3730
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   7095
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   7810
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton GridReturn 
      BackColor       =   &H0000FF00&
      Default         =   -1  'True
      Height          =   615
      Left            =   9255
      Picture         =   "Staff.frx":04F1
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   1095
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   20415
      Top             =   2170
   End
   Begin VB.CommandButton cmdView 
      BackColor       =   &H000080FF&
      Caption         =   "&View one by one"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7455
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   1095
      Width           =   1815
   End
   Begin VB.CommandButton cmdViewReturn 
      BackColor       =   &H00800000&
      Height          =   495
      Left            =   735
      Picture         =   "Staff.frx":0DBB
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   135
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox GName 
      DataField       =   "MedicineName"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7695
      TabIndex        =   36
      Text            =   "name of medicine"
      Top             =   2290
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Insertedby 
      DataField       =   "InsertedBy"
      DataSource      =   "Adodc3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   14655
      TabIndex        =   35
      Text            =   "name of medicine"
      Top             =   2290
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Contact 
      DataField       =   "contact"
      DataSource      =   "Adodc3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7695
      TabIndex        =   34
      Text            =   "name of medicine"
      Top             =   7330
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox DOE 
      DataField       =   "DOE"
      DataSource      =   "Adodc3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   14655
      TabIndex        =   33
      Text            =   "name of medicine"
      Top             =   3490
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox email 
      DataField       =   "email"
      DataSource      =   "Adodc3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7695
      TabIndex        =   32
      Text            =   "name of medicine"
      Top             =   8410
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Price 
      DataField       =   "Price"
      DataSource      =   "Adodc3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   11655
      TabIndex        =   31
      Text            =   "name of medicine"
      Top             =   6250
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Address 
      DataField       =   "address"
      DataSource      =   "Adodc3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7695
      ScrollBars      =   1  'Horizontal
      TabIndex        =   30
      Text            =   "name of medicine"
      Top             =   6250
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox CName 
      DataField       =   "Mfdname"
      DataSource      =   "Adodc3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7695
      TabIndex        =   29
      Text            =   "name of medicine"
      Top             =   5170
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox BName 
      DataField       =   "BrandName"
      DataSource      =   "Adodc3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7695
      TabIndex        =   28
      Text            =   "name of medicine"
      Top             =   3490
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdViewUpdate 
      BackColor       =   &H0000FF00&
      Caption         =   "&Update"
      Height          =   495
      Left            =   8295
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   9490
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdViewDelete 
      BackColor       =   &H000000FF&
      Caption         =   "&Delete"
      Height          =   495
      Left            =   10575
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   9490
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdViewSave 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Save"
      Height          =   495
      Left            =   7695
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   9610
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdViewCancel 
      BackColor       =   &H00FFFFC0&
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   9675
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   9610
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdViewClear 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Clear &All"
      Height          =   495
      Left            =   11775
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   9610
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox DOU 
      DataField       =   "DOU"
      DataSource      =   "Adodc3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   14655
      TabIndex        =   22
      Text            =   "-----------------"
      Top             =   4570
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox updatedBy 
      DataField       =   "UpdatedBy"
      DataSource      =   "Adodc3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   14655
      TabIndex        =   21
      Text            =   "-----------------"
      Top             =   5650
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox sort 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   405
      ItemData        =   "Staff.frx":11FD
      Left            =   12135
      List            =   "Staff.frx":1219
      Sorted          =   -1  'True
      TabIndex        =   20
      Text            =   "Display All Items"
      Top             =   1335
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      DataField       =   "MedicineName"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   5295
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   4330
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text2 
      DataField       =   "BrandName"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   5295
      TabIndex        =   17
      Text            =   "Text2"
      Top             =   4090
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      DataField       =   "Mfdname"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   5175
      TabIndex        =   16
      Text            =   "Text3"
      Top             =   4090
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      DataField       =   "address"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   5295
      TabIndex        =   15
      Text            =   "Text4"
      Top             =   3970
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      DataField       =   "contact"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   5295
      TabIndex        =   14
      Text            =   "Text5"
      Top             =   3970
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      DataField       =   "email"
      DataSource      =   "Adodc1"
      Height          =   480
      Left            =   5295
      TabIndex        =   13
      Text            =   "Text6"
      Top             =   4450
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text7 
      DataField       =   "qty"
      DataSource      =   "Adodc1"
      Height          =   480
      Left            =   5295
      TabIndex        =   12
      Text            =   "Text7"
      Top             =   4450
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      DataField       =   "Price"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   5415
      TabIndex        =   11
      Text            =   "Text8"
      Top             =   4090
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      DataField       =   "mfd"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   5535
      TabIndex        =   10
      Text            =   "Text9"
      Top             =   4210
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      DataField       =   "exp"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   5535
      TabIndex        =   9
      Text            =   "Text10"
      Top             =   4210
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text11 
      DataField       =   "DOE"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   5415
      TabIndex        =   8
      Text            =   "Text11"
      Top             =   4330
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text12 
      DataField       =   "InsertedBy"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   5415
      TabIndex        =   7
      Text            =   "Text12"
      Top             =   4330
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text13 
      DataField       =   "DOU"
      DataSource      =   "Adodc1"
      Height          =   480
      Left            =   5295
      TabIndex        =   6
      Text            =   "Text13"
      Top             =   4330
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox MType 
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
      Left            =   14055
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   2530
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtBatch 
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
      Left            =   14055
      TabIndex        =   3
      Top             =   1930
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox Quantity 
      DataField       =   "qty"
      DataSource      =   "Adodc3"
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
      Left            =   11655
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   5170
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox batchNo 
      DataField       =   "Batch"
      DataSource      =   "Adodc3"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   11655
      TabIndex        =   1
      Text            =   "name of medicine"
      Top             =   2290
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox ComboType 
      DataField       =   "qty"
      DataSource      =   "Adodc3"
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
      Left            =   11655
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   3490
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   615
      Left            =   1935
      Top             =   9495
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   18495
      Top             =   9135
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
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
      RecordSource    =   "medicine"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataListLib.DataList DataList1 
      Bindings        =   "Staff.frx":12AF
      DataSource      =   " "
      Height          =   810
      Left            =   2055
      TabIndex        =   19
      Top             =   8175
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1429
      _Version        =   393216
      BackColor       =   4210816
      ForeColor       =   65280
      ListField       =   "MedicineName"
      BoundColumn     =   " "
   End
   Begin MSComCtl2.DTPicker Exp 
      Height          =   375
      Left            =   14055
      TabIndex        =   45
      Top             =   4905
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Format          =   118423553
      CurrentDate     =   43589
   End
   Begin MSComCtl2.DTPicker MFD 
      Height          =   375
      Left            =   14055
      TabIndex        =   46
      Top             =   4335
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Format          =   118423553
      CurrentDate     =   43589
   End
   Begin MSComCtl2.DTPicker Dexp 
      DataField       =   "exp"
      DataSource      =   "Adodc3"
      Height          =   375
      Left            =   11655
      TabIndex        =   63
      Top             =   8415
      Visible         =   0   'False
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
      Format          =   118423553
      CurrentDate     =   43589
   End
   Begin MSComCtl2.DTPicker Dmfd 
      DataField       =   "mfd"
      DataSource      =   "Adodc3"
      Height          =   375
      Left            =   11655
      TabIndex        =   64
      Top             =   7335
      Visible         =   0   'False
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
      Format          =   118423553
      CurrentDate     =   43589
   End
   Begin VB.TextBox txtGName 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8055
      TabIndex        =   52
      Top             =   1930
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Frame FrameSorting 
      BackColor       =   &H00FF8080&
      Caption         =   "Select for sorting"
      Height          =   975
      Left            =   12120
      TabIndex        =   124
      Top             =   360
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton cmdSwitch_User 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Switch &User"
      Height          =   600
      Left            =   9255
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   1095
      Width           =   1815
   End
   Begin VB.Frame frameStaff 
      BackColor       =   &H0000FF00&
      Caption         =   "FRONT-LINE-STAFF"
      ForeColor       =   &H8000000B&
      Height          =   1815
      Left            =   1680
      TabIndex        =   125
      Top             =   360
      Width           =   9855
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   16920
      TabIndex        =   122
      Top             =   1995
      Width           =   675
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   19080
      TabIndex        =   121
      Top             =   975
      Width           =   660
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wednesday"
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Left            =   16800
      TabIndex        =   120
      Top             =   510
      Width           =   1530
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Left            =   19080
      TabIndex        =   119
      Top             =   600
      Width           =   645
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Left            =   19080
      TabIndex        =   118
      Top             =   1800
      Width           =   645
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   17535
      TabIndex        =   117
      Top             =   1740
      Width           =   90
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   17955
      TabIndex        =   116
      Top             =   1350
      Width           =   90
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   17175
      TabIndex        =   115
      Top             =   1575
      Width           =   60
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   17325
      TabIndex        =   114
      Top             =   1080
      Width           =   120
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   17175
      TabIndex        =   113
      Top             =   1200
      Width           =   120
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   17760
      TabIndex        =   112
      Top             =   1080
      Width           =   60
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   17895
      TabIndex        =   111
      Top             =   1185
      Width           =   60
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   17895
      TabIndex        =   110
      Top             =   1575
      Width           =   60
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   17760
      TabIndex        =   109
      Top             =   1710
      Width           =   60
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   17130
      TabIndex        =   108
      Top             =   1350
      Width           =   90
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   17340
      TabIndex        =   107
      Top             =   1710
      Width           =   60
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   17505
      TabIndex        =   106
      Top             =   960
      Width           =   180
   End
   Begin VB.Shape Shape3 
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   17535
      Shape           =   3  'Circle
      Top             =   1395
      Width           =   120
   End
   Begin VB.Line Line1 
      X1              =   17580
      X2              =   17580
      Y1              =   1445
      Y2              =   1125
   End
   Begin VB.Line Line2 
      X1              =   17580
      X2              =   17780
      Y1              =   1435
      Y2              =   1215
   End
   Begin VB.Line Line3 
      X1              =   17580
      X2              =   17955
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label msg_no 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   225
      Left            =   975
      TabIndex        =   105
      Top             =   135
      Width           =   210
   End
   Begin VB.Shape msgCircle 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   960
      Shape           =   3  'Circle
      Top             =   120
      Width           =   255
   End
   Begin VB.Image messageBox 
      Height          =   615
      Left            =   375
      Picture         =   "Staff.frx":12C4
      Stretch         =   -1  'True
      Top             =   255
      Width           =   735
   End
   Begin VB.Image back_image 
      Height          =   3645
      Left            =   -2640
      Picture         =   "Staff.frx":3535
      Stretch         =   -1  'True
      Top             =   7215
      Width           =   2625
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company's Details"
      Height          =   360
      Left            =   5895
      TabIndex        =   104
      Top             =   3135
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   360
      Left            =   5895
      TabIndex        =   103
      Top             =   3855
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   360
      Left            =   5895
      TabIndex        =   102
      Top             =   4455
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact "
      Height          =   360
      Left            =   5895
      TabIndex        =   101
      Top             =   5055
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail"
      Height          =   360
      Left            =   5895
      TabIndex        =   100
      Top             =   5655
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QTY"
      Height          =   360
      Left            =   12735
      TabIndex        =   99
      Top             =   3135
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   360
      Left            =   12735
      TabIndex        =   98
      Top             =   3780
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EXP."
      Height          =   360
      Left            =   12735
      TabIndex        =   97
      Top             =   4905
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MFD"
      Height          =   360
      Left            =   12735
      TabIndex        =   96
      Top             =   4335
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgReturn 
      Height          =   495
      Left            =   7455
      Picture         =   "Staff.frx":1E776
      Stretch         =   -1  'True
      Top             =   7815
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label viewMedicine 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List of Medicines"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2895
      TabIndex        =   95
      Top             =   855
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Label lblDetails 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Details::"
      Height          =   360
      Left            =   7695
      TabIndex        =   94
      Top             =   1095
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblUnderline 
      BackColor       =   &H80000007&
      Caption         =   "Label16"
      Height          =   30
      Left            =   7695
      TabIndex        =   93
      Top             =   1410
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Generic Name"
      Height          =   360
      Index           =   15
      Left            =   7695
      TabIndex        =   92
      Top             =   1935
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Brand Name"
      Height          =   360
      Index           =   16
      Left            =   7695
      TabIndex        =   91
      Top             =   3135
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Company's Details"
      Height          =   360
      Index           =   17
      Left            =   7695
      TabIndex        =   90
      Top             =   4215
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   360
      Index           =   18
      Left            =   7695
      TabIndex        =   89
      Top             =   4815
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Address"
      Height          =   360
      Index           =   19
      Left            =   7695
      TabIndex        =   88
      Top             =   5895
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Contact"
      Height          =   360
      Index           =   20
      Left            =   7695
      TabIndex        =   87
      Top             =   6975
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "E-Mail"
      Height          =   360
      Index           =   21
      Left            =   7695
      TabIndex        =   86
      Top             =   8055
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "QTY"
      Height          =   360
      Index           =   22
      Left            =   11655
      TabIndex        =   85
      Top             =   4815
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Price (Rs.)"
      Height          =   360
      Index           =   23
      Left            =   11655
      TabIndex        =   84
      Top             =   5895
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "MFD"
      Height          =   360
      Index           =   24
      Left            =   11655
      TabIndex        =   83
      Top             =   6975
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "EXP Date"
      Height          =   360
      Index           =   25
      Left            =   11655
      TabIndex        =   82
      Top             =   8055
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Date of Entry"
      Height          =   360
      Index           =   26
      Left            =   14655
      TabIndex        =   81
      Top             =   3135
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Inserted By"
      Height          =   360
      Index           =   27
      Left            =   14655
      TabIndex        =   80
      Top             =   1935
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label SelectMedicine 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Select A Medicine"
      Height          =   360
      Left            =   2895
      TabIndex        =   79
      Top             =   135
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Last updated on"
      Height          =   360
      Index           =   28
      Left            =   14655
      TabIndex        =   78
      Top             =   4215
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Updated By"
      Height          =   360
      Index           =   29
      Left            =   14655
      TabIndex        =   77
      Top             =   5295
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Notification 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Notification"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   180
      Left            =   375
      TabIndex        =   76
      Top             =   855
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter The Following Details of Medicine"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   6855
      TabIndex        =   75
      Top             =   615
      Visible         =   0   'False
      Width           =   9075
   End
   Begin VB.Label label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Generic Name"
      Height          =   360
      Left            =   5895
      TabIndex        =   74
      Top             =   1935
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Brand Name"
      Height          =   360
      Left            =   5895
      TabIndex        =   73
      Top             =   2535
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   360
      Left            =   12735
      TabIndex        =   72
      Top             =   2535
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Batch no"
      Height          =   360
      Left            =   12735
      TabIndex        =   71
      Top             =   1935
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   23160
      TabIndex        =   70
      Top             =   615
      Width           =   180
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Type"
      Height          =   360
      Index           =   31
      Left            =   11655
      TabIndex        =   69
      Top             =   3135
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Batch no."
      Height          =   360
      Index           =   30
      Left            =   11655
      TabIndex        =   68
      Top             =   1935
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Left            =   16680
      TabIndex        =   123
      Top             =   360
      Width           =   3615
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   1005
      Left            =   17040
      Shape           =   3  'Circle
      Top             =   960
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      DrawMode        =   5  'Not Copy Pen
      FillStyle       =   7  'Diagonal Cross
      Height          =   11535
      Left            =   -225
      Top             =   10455
      Width           =   20505
   End
End
Attribute VB_Name = "Staff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 

Public CN As New ADODB.Connection
Public Rs As New ADODB.Recordset

Private Sub cmdADD_Click()
Adodc1.Refresh
If txtGName.Text = "" Or txtBName.Text = "" Or txtCName.Text = "" Or txtAddress.Text = "" Or txtContact.Text = "" Or txtEmail.Text = "" Or txtPrice.Text = "" Or QTY.Text = "" Then
    MsgBox "Please Enter all info.", vbCritical, "SUGGESTION"
Else
Rs.AddNew
'Call cmdClear_Click
Rs.Fields("SN").Value = Adodc1.Recordset.RecordCount + 1
Rs.Fields("MedicineName").Value = txtGName.Text
Rs.Fields("BrandName").Value = txtBName.Text
Rs.Fields("MfdName").Value = txtCName.Text
Rs.Fields("address").Value = txtAddress.Text
Rs.Fields("contact").Value = txtContact.Text
Rs.Fields("email").Value = txtEmail.Text
Rs.Fields("Batch").Value = txtBatch.Text
Rs.Fields("Type").Value = MType.Text
Rs.Fields("qty").Value = Val(QTY.Text)
Rs.Fields("Price").Value = Val(txtPrice.Text)
Rs.Fields("mfd").Value = MFD.Value
Rs.Fields("exp").Value = EXP.Value
Rs.Fields("DOE").Value = Now
Rs.Fields("InsertedBy").Value = "staff"
Rs.Update
MsgBox "Successfully Saved", vbInformation, "SAVED"
'Rs.Close

txtGName.Enabled = False
txtBName.Enabled = False
txtCName.Enabled = False
txtAddress.Enabled = False
txtContact.Enabled = False
txtEmail.Enabled = False
txtBatch.Enabled = False
MType.Enabled = False
txtPrice.Enabled = False
QTY.Enabled = False
MFD.Enabled = False
EXP.Enabled = False


cmdCancel.Visible = False
imgReturn.Visible = True
cmdADD.Visible = False
cmdClear.Enabled = False
cmdAddAnother.Visible = True

End If

Adodc1.Refresh

End Sub

Private Sub cmdAddAnother_Click()
cmdCancel.Visible = True
cmdClear.Value = True
cmdADD.Visible = True
cmdAddAnother.Visible = False
imgReturn.Visible = False
Call cmdClear_Click

txtGName.Enabled = True
txtBName.Enabled = True
txtCName.Enabled = True
txtAddress.Enabled = True
txtContact.Enabled = True
txtEmail.Enabled = True
txtBatch.Enabled = True
MType.Enabled = True
txtPrice.Enabled = True
QTY.Enabled = True
MFD.Enabled = True
EXP.Enabled = True

Adodc1.Refresh

End Sub

Private Sub cmdCancel_Click()
'Call Form_Load
cmdCancel.Visible = False
cmdADD.Visible = False
cmdClear.Visible = False

back_image.Visible = True
DataGrid1.Visible = True
Shape1.Visible = True
cmdInsert.Visible = True
cmdSearch.Value = True
cmdUpdate.Visible = True
cmdSwitch_User.Visible = True
cmdExit.Visible = True
'cmdSort.Visible = True
cmdView.Visible = True
cmdDelete.Visible = True
messageBox.Visible = True
msgCircle.Visible = True
msg_no.Visible = True
Notification.Visible = True
searchBox.Visible = True
cmdSearch.Visible = True
frameStaff.Visible = True
sort.Visible = True


Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label16.Visible = False
Label17.Visible = False

txtGName.Visible = False
txtBName.Visible = False
txtCName.Visible = False
txtAddress.Visible = False
txtContact.Visible = False
txtEmail.Visible = False
txtBatch.Visible = False
MType.Visible = False
txtPrice.Visible = False
QTY.Visible = False
MFD.Visible = False
EXP.Visible = False


txtGName.Enabled = True
txtBName.Enabled = True
txtCName.Enabled = True
txtAddress.Enabled = True
txtContact.Enabled = True
txtEmail.Enabled = True
MType.Enabled = True
txtBatch.Enabled = True
txtPrice.Enabled = True
QTY.Enabled = True
MFD.Enabled = True
EXP.Enabled = True




Call cmdClear_Click
'DataGrid1.Refresh
'Rs.Delete adAffectCurrent
'Rs.Update
'Rs.Close
'Rs.Open "select * from medicine", CN, adOpenDynamic, adLockPessimistic
Adodc1.Refresh
Staff.BackColor = &H404000

End Sub

Private Sub cmdClear_Click()
txtGName.Text = ""
txtBName.Text = ""
txtCName.Text = ""
txtAddress.Text = ""
txtContact.Text = ""
txtEmail.Text = ""
txtBatch.Text = ""
MType.Text = ""
txtPrice.Text = ""
QTY.Text = ""
MFD.Value = Date
EXP.Value = Date
'txtGName.SetFocus
End Sub

Private Sub cmdDelete_Click()
Call cmdView_Click
SelectMedicine.Visible = True
SelectMedicine.Caption = "Please select one for deletion"
SelectMedicine.Top = viewMedicine.Top
SelectMedicine.Left = viewMedicine.Left
viewMedicine.Visible = False
cmdViewUpdate.Visible = False
'Dim L1 As Integer ', L2 As Integer
'L1 = cmdViewUpdate.Left
'L2 = cmdViewDelete.Left
cmdViewDelete.Left = 10720 ' cmdViewUpdate.Left + cmdViewUpdate.Width

End Sub

Private Sub cmdExit_Click()
End
End Sub

 

Private Sub cmdInsert_Click()
MType.AddItem "Tablet"
MType.AddItem "Liquid"
MType.AddItem "Powder"
MType.AddItem "Capsule"
back_image.Visible = False
'Shape1.FillColor = &H4040&
Shape1.Visible = False
Staff.BackColor = &HFF8080
DataGrid1.Visible = False
cmdInsert.Visible = False
cmdSearch.Value = False
cmdUpdate.Visible = False
cmdSwitch_User.Visible = False
cmdExit.Visible = False
'cmdSort.Visible = False
cmdView.Visible = False
cmdDelete.Visible = False
messageBox.Visible = False
msgCircle.Visible = False
msg_no.Visible = False
Notification.Visible = False
searchBox.Visible = False
cmdSearch.Visible = False
frameStaff.Visible = False
sort.Visible = False

cmdCancel.Visible = True
cmdClear.Visible = True
cmdADD.Visible = True
cmdADD.Caption = "Save"

Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = True
Label10.Visible = True
Label11.Visible = True
Label12.Visible = True
Label16.Visible = True
Label17.Visible = True

txtGName.Visible = True
txtBName.Visible = True
txtCName.Visible = True
txtAddress.Visible = True
txtContact.Visible = True
txtEmail.Visible = True
txtBatch.Visible = True
MType.Visible = True
txtPrice.Visible = True
QTY.Visible = True
MFD.Visible = True
EXP.Visible = True

Call cmdClear_Click
'Rs.AddNew

End Sub

Private Sub cmdSwitch_User_Click()
Staff.Hide
Load loginform
loginform.Show
loginform.txtPw.Visible = False
loginform.txtUser.Visible = False
loginform.selectUser.Text = "Select_User"
loginform.lbl_forgotPw.Visible = False
loginform.cmdLogin.Enabled = False

End Sub

 

Private Sub cmdUpdate_Click()
Call cmdView_Click
SelectMedicine.Visible = True
SelectMedicine.Caption = "Please Select Medicine"
SelectMedicine.Top = viewMedicine.Top
SelectMedicine.Left = viewMedicine.Left
viewMedicine.Visible = False
cmdViewDelete.Visible = False
cmdViewUpdate.Left = 6840 'cmdViewDelete.Left - cmdViewUpdate.Width

End Sub

Private Sub cmdView_Click()
'staff.BackColor = &H80FF80
DataList1.Visible = True
DataList1.Top = 1400
DataList1.Left = 2000
DataList1.Width = 3989
DataList1.Height = 9700
'DataGrid2.Visible = True
Shape1.Visible = False

hidestaff
cmdViewReturn.Visible = True
Staff.BackColor = &H80FF80      '&HC0FFC0

'initialize combotype
ComboType.AddItem "Tablet"
ComboType.AddItem "Liquid"
ComboType.AddItem "Powder"
ComboType.AddItem "Capsule"

'initialize Quantity
For i = 1 To 100
    Quantity.AddItem i
Next
VIEW (1)
'SetDataSource


End Sub
' set data source
Sub SetDataSource()

    GName.DataSource = Adodc1
    BName.DataSource = Adodc1
    BName.DataSource = Adodc1
    batchNo.DataSource = Adodc1
    ComboType.DataSource = Adodc1
    CName.DataSource = Adodc1
    Address.DataSource = Adodc1
    Contact.DataSource = Adodc1
    Email.DataSource = Adodc1
    Quantity.DataSource = Adodc1
    Price.DataSource = Adodc1
    Dmfd.DataSource = Adodc1
    Dexp.DataSource = Adodc1
    DOE.DataSource = Adodc1
    Insertedby.DataSource = Adodc1
    DOU.DataSource = Adodc1
    updatedBy.DataSource = Adodc1

End Sub
Function VIEW(ByVal num As Integer)
Dim i As Integer
If num = 1 Then
    viewMedicine.Visible = True
    lblDetails.Visible = True
    lblUnderline.Visible = True
    
    'show labels
    'Dim i As Integer
    For i = 15 To 31
        Label(i).Visible = True
        Label(i).BackStyle = 0
    Next
    'show texboxes
    GName.Visible = True
    BName.Visible = True
    batchNo.Visible = True
    ComboType.Visible = True
    CName.Visible = True
    Address.Visible = True
    Contact.Visible = True
    Email.Visible = True
    Quantity.Visible = True
    Price.Visible = True
    Dmfd.Visible = True
    Dexp.Visible = True
    DOE.Visible = True
    Insertedby.Visible = True
    DOU.Visible = True
    updatedBy.Visible = True
    If DOU.Text = "" Then
        DOU.Text = "----------"
        updatedBy.Text = "----------"
    End If
    
    
    'disable texboxes
    GName.Enabled = False
    BName.Enabled = False
    batchNo.Enabled = False
    ComboType.Enabled = False
    CName.Enabled = False
    Address.Enabled = False
    Contact.Enabled = False
    Quantity.Enabled = False
    Email.Enabled = False
    Price.Enabled = False
    Dmfd.Enabled = False
    Dexp.Enabled = False
    DOE.Enabled = False
    Insertedby.Enabled = False
    updatedBy.Enabled = False
    DOU.Enabled = False
    
    
    'diaplay cammands
    cmdViewUpdate.Visible = True
    cmdViewDelete.Visible = True
    
    
Else
    'hide labels
    For i = 15 To 31
        Label(i).Visible = False
    Next
    lblDetails.Visible = False
    lblUnderline.Visible = False
    
    'hide texboxes
    GName.Visible = False
    BName.Visible = False
    batchNo.Visible = False
    ComboType.Visible = False
    CName.Visible = False
    Address.Visible = False
    Contact.Visible = False
    Email.Visible = False
    Quantity.Visible = False
    Price.Visible = False
    Dmfd.Visible = False
    Dexp.Visible = False
    DOE.Visible = False
    Insertedby.Visible = False
    DOU.Visible = False
    updatedBy.Visible = False
    
    
    'hide cammands
    cmdViewUpdate.Visible = False
    cmdViewDelete.Visible = False
    cmdViewSave.Visible = False
    cmdViewCancel.Visible = False
    cmdViewClear.Visible = False
    
    
End If
End Function
Sub hidestaff()
back_image.Visible = False
DataGrid1.Visible = False
cmdInsert.Visible = False
cmdSearch.Value = False
cmdUpdate.Visible = False
cmdSwitch_User.Visible = False
cmdExit.Visible = False
'cmdSort.Visible = False
cmdView.Visible = False
cmdDelete.Visible = False
messageBox.Visible = False
sort.Visible = False
msgCircle.Visible = False
msg_no.Visible = False
searchBox.Visible = False
cmdSearch.Visible = False
frameStaff.Visible = False
GridDelete.Visible = False
GridUpdate.Visible = False
GridReturn.Visible = False


End Sub

Private Sub cmdViewCancel_Click()

cmdViewUpdate.Visible = True
cmdViewDelete.Visible = True
cmdViewSave.Visible = False
cmdViewCancel.Visible = False
cmdViewClear.Visible = False

'disable all the texboxes
GName.Enabled = False
BName.Enabled = False
batchNo.Enabled = False
ComboType.Enabled = False
CName.Enabled = False
Address.Enabled = False
Contact.Enabled = False
Email.Enabled = False
Quantity.Enabled = False
Price.Enabled = False
Dmfd.Enabled = False
Dexp.Enabled = False


'refresh database
Adodc3.Refresh
Adodc1.Refresh
DataList1.Refresh
DataGrid1.Refresh
End Sub

Private Sub cmdViewClear_Click()
GName.Text = ""
BName.Text = ""
batchNo.Text = ""
ComboType.Text = ""
CName.Text = ""
Address.Text = ""
Contact.Text = ""
Email.Text = ""
Quantity.Text = ""
Price.Text = ""
Dmfd.Value = Date
Dexp.Value = Date

End Sub

Private Sub cmdViewReturn_Click()
DataList1.Visible = False
cmdViewReturn.Visible = False
SelectMedicine.Visible = False
Shape1.Visible = True

showstaff
Staff.BackColor = &H404000
VIEW (0)
viewMedicine.Visible = False

'Rs.Close
'Rs.Open "select * form medicine", CN, adOpenDynamic, adLockPessimistic
Adodc1.Refresh
'Adodc3.Refresh
DataGrid1.Refresh
DataList1.Refresh
cmdViewDelete.Left = 9720 'cmdViewUpdate.Left + cmdViewUpdate.Width
cmdViewUpdate.Left = 7440
'Call GridReturn_Click
Quantity.Clear
ComboType.Clear
End Sub
Public Sub showstaff()
back_image.Visible = True
DataGrid1.Visible = True
cmdInsert.Visible = True
cmdSearch.Value = True
cmdUpdate.Visible = True
cmdSwitch_User.Visible = True
cmdExit.Visible = True
'cmdSort.Visible = True
cmdView.Visible = True
cmdDelete.Visible = True
messageBox.Visible = True
msgCircle.Visible = True
msg_no.Visible = True
Notification.Visible = True
searchBox.Visible = True
cmdSearch.Visible = True
frameStaff.Visible = True
sort.Visible = True
End Sub

Private Sub cmdViewSave_Click()
'Rs.Fields("SN").Value = Adodc1.Recordset.RecordCount + 1
Rs.Fields("MedicineName").Value = GName.Text
Rs.Fields("BrandName").Value = BName.Text
Rs.Fields("MfdName").Value = CName.Text
Rs.Fields("address").Value = Address.Text
Rs.Fields("contact").Value = Contact.Text
Rs.Fields("email").Value = Email.Text
Rs.Fields("Batch").Value = batchNo.Text
Rs.Fields("Type").Value = ComboType.Text
Rs.Fields("qty").Value = Val(Quantity.Text)
Rs.Fields("Price").Value = Val(Price.Text)
Rs.Fields("mfd").Value = Dmfd.Value
Rs.Fields("exp").Value = Dexp.Value
Rs.Fields("DOU").Value = Now
Rs.Fields("UpdatedBy").Value = "staff"
Rs.Update

MsgBox "Saved Successfully.", vbInformation
cmdViewUpdate.Visible = True
cmdViewDelete.Visible = True
cmdViewSave.Visible = False
cmdViewCancel.Visible = False
cmdViewClear.Visible = False

'disable all  the textboxes
GName.Enabled = False
BName.Enabled = False
batchNo.Enabled = False
ComboType.Enabled = False
CName.Enabled = False
Address.Enabled = False
Contact.Enabled = False
Email.Enabled = False
Quantity.Enabled = False
Price.Enabled = False
Dmfd.Enabled = False
Dexp.Enabled = False

End Sub

Private Sub cmdViewUpdate_Click()
cmdViewUpdate.Visible = False
cmdViewDelete.Visible = False
cmdViewSave.Visible = True
cmdViewCancel.Visible = True
cmdViewClear.Visible = True


'enable the textboxes for edit
GName.Enabled = True
BName.Enabled = True
batchNo.Enabled = True
ComboType.Enabled = True
CName.Enabled = True
Address.Enabled = True
Contact.Enabled = True
Email.Enabled = True
Quantity.Enabled = True
Price.Enabled = True
Dmfd.Enabled = True
Dexp.Enabled = True
'Insertedby.Enabled = True

End Sub
 

Private Sub DataGrid1_Click()
GridUpdate.Visible = True
GridDelete.Visible = True
GridReturn.Visible = True
cmdInsert.Enabled = False
cmdUpdate.Enabled = False
'cmdSort.Enabled = False
cmdView.Enabled = False
cmdDelete.Enabled = False
cmdSwitch_User.Enabled = False
'cmdSearch.Enabled = False
'frameStaff.Enabled = False
'searchBox.Enabled = False


'Exchange (1)
'GridReturn.SetFocus
'GridReturn.Default = True

Text1.Visible = True
Text2.Visible = True
Text3.Visible = True
Text4.Visible = True
Text5.Visible = True
Text6.Visible = True
Text7.Visible = True
Text8.Visible = True
Text9.Visible = True
Text10.Visible = True
Text11.Visible = True
Text12.Visible = True
Text13.Visible = True
Text14.Visible = True
Text15.Visible = True
Text16.Visible = True

End Sub



Private Sub DataList1_Click()
'Rs.Close
'Rs.Open "select * from medicine", CN, adOpenDynamic, adLockPessimistic
Adodc1.Recordset.Bookmark = DataList1.SelectedItem
End Sub


Private Sub DataList2_Click()
Adodc1.Recordset.Bookmark = DataList2.SelectedItem
End Sub

Private Sub Form_Load()
DataGrid1.Height = 7740
DataGrid1.Width = 20155
'DataGrid1.Left = 2869
'DataGrid1.Top = 2188 '1680

For i = 1 To 100
    QTY.AddItem i
Next

CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\BCA\3rd SEMESTER\PROJECT-3\PHARMACY_CONTROL_SYSTEM.mdb;Persist Security Info=False"
Rs.Open "select * from medicine", CN, adOpenDynamic, adLockPessimistic

Shape1.Top = 0
Shape1.Left = 0
 
Shape1.Visible = True

End Sub

 
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Text1.Text = X
'Text2.Text = Y

End Sub

Private Sub GridDelete_Click()
'Call cmdDelete_Click

confirm = MsgBox("Would you like to delete the selected item?", vbYesNo + vbQuestion, "Confirmation")
If confirm = vbYes Then
    Adodc1.Recordset.Delete 'rs.Delete 'adAffectCurrent
    'Adodc1.Refresh
    'DataGrid1.Refresh
    'adodc3.Refresh
    MsgBox "The selected item is deleted.", vbInformation
    If Not Rs.EOF Then
        Rs.MoveNext
    ElseIf Not Rs.BOF Then
        Rs.MovePrevious
    End If
    'Rs.Close
    'Rs.Open "select * from medicine", CN, adOpenDynamic, adLockPessimistic
Else
    MsgBox "Selected item is not deleted.", vbInformation, "Message"
End If
End Sub

Private Sub GridReturn_Click()
GridDelete.Visible = False
GridUpdate.Visible = False
GridReturn.Visible = False
cmdInsert.Enabled = True
cmdUpdate.Enabled = True
'cmdSort.Enabled = True
cmdView.Enabled = True
cmdDelete.Enabled = True
cmdSwitch_User.Enabled = True
cmdSearch.Enabled = True
frameStaff.Enabled = True
searchBox.Enabled = True

'Exchange (0)
Rs.Update
Adodc1.Refresh

Text1.Visible = False
Text2.Visible = False
Text3.Visible = False
Text4.Visible = False
Text5.Visible = False
Text6.Visible = False
Text7.Visible = False
Text8.Visible = False
Text9.Visible = False
Text10.Visible = False
Text11.Visible = False
Text12.Visible = False
Text13.Visible = False
Text14.Visible = False
Text15.Visible = fasle
Text16.Visible = fasle

End Sub
Sub Exchange(a As Integer)
'Dim T As Integer ', L As Integer
'T = cmdExit.Top
'L = cmdExit.Left
If a = 1 Then
cmdExit.Top = 6720 'GridReturn.Top
'cmdExit.Left = GridReturn.Left
'GridReturn.Left = L
GridReturn.Top = 6000

Else
cmdExit.Top = 6000
GridReturn.Top = 6720
End If
End Sub

Private Sub GridUpdate_Click()
'Call cmdUpdate_Click
StaffFormUpdate.Show 0
Staff.Enabled = False
'Rs.MoveNext
StaffFormUpdate.GN.Text = Text1.Text '   Rs.Fields("MedicineName")
StaffFormUpdate.BN.Text = Text2.Text '.Rs.Fields("BrandName")
'Name.Text = Text3.Text     'staff.Rs.Fields("Mfdname")
StaffFormUpdate.Text1.Text = Text3.Text
StaffFormUpdate.Address.Text = Text4.Text ' staff.Rs.Fields("address")
StaffFormUpdate.Contact.Text = Text5.Text ' staff.Rs.Fields("contact")
StaffFormUpdate.Email.Text = Text6.Text ' staff.Rs.Fields("email")
StaffFormUpdate.QTY.Text = Text7.Text ' staff.Rs.Fields("qty")
StaffFormUpdate.Price.Text = Text8.Text '     staff.Rs.Fields("Price")
StaffFormUpdate.MFD.Value = Text9.Text ' staff.Rs.Fields("mfd")
StaffFormUpdate.EXP.Value = Text10.Text ' staff.Rs.Fields("exp")
StaffFormUpdate.Text2.Text = Text15.Text
StaffFormUpdate.Combo1.Text = Text16.Text
StaffFormUpdate.Label13.Caption = "Inserted on: "
StaffFormUpdate.Label16.Caption = "Updated on: "
If Text13.Text = "" Then
    StaffFormUpdate.Label16.Caption = StaffFormUpdate.Label16.Caption & "None " & "    By: none"
Else
    StaffFormUpdate.Label16.Caption = StaffFormUpdate.Label16.Caption + Text13.Text + "  By: " + Text14.Text
End If
    
StaffFormUpdate.Label13.Caption = StaffFormUpdate.Label13.Caption + " " + Text11.Text
StaffFormUpdate.Label13.Caption = StaffFormUpdate.Label13.Caption + "   By: " + Text12.Text

StaffFormUpdate.Label16.Left = StaffFormUpdate.Width / 2 - StaffFormUpdate.Label16.Width / 2
StaffFormUpdate.Label13.Left = StaffFormUpdate.Width / 2 - StaffFormUpdate.Label13.Width / 2

'back_image.Top = 0
'back_image.Left = 0
'back_image.Width = staff.Width
'back_image.Height = staff.Height
'hidestaff
'back_image.Visible = True
'DataGrid1.Enabled = False

End Sub

Private Sub imgReturn_Click()
'Rs.Open "select * from medicine", CN, adOpenDynamic, adLockPessimistic
imgReturn.Visible = False
cmdAddAnother.Visible = False
cmdClear.Enabled = True
Call cmdCancel_Click
'DataGrid1.EditActive = True
'DataGrid1.Refresh
'DataGrid2.Visible = True
'DataGrid2.Height = 8689
'DataGrid2.Width = 16867
'DataGrid2.Left = 2969
'DataGrid2.Top = 1680
'DataGrid1.Visible = False
End Sub

Private Sub messageBox_Click()
Notification.Visible = False
msg_no.Caption = "0"
End Sub

Private Sub messageBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X > 339 And X < 1100 And Y > 175 And Y < 860 Then
    Notification.Visible = True
Else
    Notification.Visible = False
End If
End Sub

Private Sub Notification_Click()
'If X > 339 And X < 1100 And Y > 175 And Y < 860 Then
 '   Notification.Visible = True
'Else
  '  Notification.Visible = False
'End If
End Sub

Private Sub searchBox_GotFocus()
cmdSearch.Default = True
searchBox.Text = ""
searchBox.Font.Size = 13
searchBox.FontName = "Times New Roman"
searchBox.ForeColor = vbBlue
End Sub

Private Sub searchBox_LostFocus()
searchBox.Text = "Enter the name of the medicine"
searchBox.FontSize = 10
searchBox.ForeColor = &H8000000A
End Sub

Private Sub Timer1_Timer()
Static i As Integer
Label13.Caption = Time
Label14.Caption = Format(Date, "d")
Label19.Caption = Format(Date, "mmmm")
Label20.Caption = Format(Date, "yyyy")
Label15.Caption = Format(Date, "dddd")
'Label14.Left = Label13.Left - Label14.Width - 120
'Label15.Left = Label13.Left - ((Label14.Width - 120) / 2)
'Text15.Text = Line1.X2
'Text16.Text = Line1.Y2
'If i < 16 Then
'Line1.X2 < 17720 And Line1.Y2 < 1200 Then
'Line1.X2 = Line1.X2 + 24.7
'Line1.Y2 = Line1.Y2 + 21.3
'i = i + 1

'ElseIf i > 15 And i < 31 Then 'Line3.Y2 > 1190 And Line3.Y2 < 1600 And Line3.X2 > 17245 Then
 '   Line1.X2 = Line1.X2 - 23
  '  Line1.Y2 = Line1.Y2 + 21.7
   ' i = i + 1
     
'ElseIf i > 30 And i < 60 Then
 '   Line1.X2 = Line1.X2 - 25.3
  '  Line1.Y2 = Line1.Y2 - 30
   ' i = i + 1
    
'ElseIf i > 89 And i < 120 Then
  '  Line1.X2 = Line1.X2 + 2.3
   ' Line1.Y2 = Line1.Y2 - 5
    'i = i + 1

'End If
'If i = 31 Then
'Timer1.Enabled = False
'End If
End Sub

