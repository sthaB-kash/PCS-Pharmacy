VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Admin 
   BackColor       =   &H00800000&
   Caption         =   "ADMINISTRATOR"
   ClientHeight    =   7365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15900
   FillColor       =   &H00FF0000&
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Admin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   15900
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Admin.frx":08CA
      Height          =   615
      Left            =   2760
      Negotiate       =   -1  'True
      TabIndex        =   141
      Top             =   2880
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   49152
      ForeColor       =   16777215
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   615
      Left            =   2640
      Top             =   3720
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      RecordSource    =   "select * from medicine"
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   1800
      Top             =   9360
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame BillingFrame 
      BackColor       =   &H80000010&
      Caption         =   "Billing"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   5040
      TabIndex        =   151
      Top             =   1080
      Visible         =   0   'False
      Width           =   7935
      Begin VB.TextBox txtQty9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   4920
         TabIndex        =   193
         Text            =   "Text19"
         Top             =   5520
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtMedicine9 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   480
         TabIndex        =   192
         Text            =   "Text18"
         Top             =   5520
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.TextBox Text17 
         Height          =   495
         Left            =   840
         TabIndex        =   183
         Text            =   "Text17"
         Top             =   6360
         Width           =   1455
      End
      Begin VB.TextBox txtMedicine8 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   480
         TabIndex        =   182
         Text            =   "Text41"
         Top             =   5160
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.TextBox txtMedicine7 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   480
         TabIndex        =   181
         Text            =   "Text40"
         Top             =   4800
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.TextBox txtMedicine6 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   480
         TabIndex        =   180
         Text            =   "Text39"
         Top             =   4440
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.TextBox txtMedicine5 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   480
         TabIndex        =   179
         Text            =   "Text38"
         Top             =   4080
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.TextBox txtMedicine4 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   480
         TabIndex        =   178
         Text            =   "Text37"
         Top             =   3720
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.TextBox txtMedicine3 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   480
         TabIndex        =   177
         Text            =   "Text36"
         Top             =   3360
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.TextBox txtMedicine2 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   480
         TabIndex        =   176
         Text            =   "Text35"
         Top             =   3000
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.TextBox txtMedicine1 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   480
         TabIndex        =   175
         Text            =   "Text34"
         Top             =   2640
         Width           =   3855
      End
      Begin VB.TextBox txtQty8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4680
         TabIndex        =   172
         Text            =   "Text24"
         Top             =   5160
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtQty7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4680
         TabIndex        =   171
         Text            =   "Text23"
         Top             =   4800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtQty6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4680
         TabIndex        =   170
         Text            =   "Text22"
         Top             =   4440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtQty5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4680
         TabIndex        =   169
         Text            =   "Text21"
         Top             =   4080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtQty4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4680
         TabIndex        =   168
         Text            =   "Text20"
         Top             =   3720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtQty3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4680
         TabIndex        =   167
         Text            =   "Text19"
         Top             =   3360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtQty2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4680
         TabIndex        =   166
         Text            =   "Text18"
         Top             =   3000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtQty1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   4680
         TabIndex        =   165
         Text            =   "Text17"
         Top             =   2640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Timer Timer3 
         Interval        =   999
         Left            =   7200
         Top             =   840
      End
      Begin VB.CommandButton Billing_Cancel 
         BackColor       =   &H000080FF&
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
         Height          =   480
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   155
         Top             =   6960
         Width           =   1335
      End
      Begin VB.CommandButton Billing_Print 
         BackColor       =   &H0000C000&
         Caption         =   "&Print"
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
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   154
         Top             =   6960
         Width           =   1335
      End
      Begin VB.Label Price9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Label27"
         Height          =   225
         Left            =   6675
         TabIndex        =   194
         Top             =   5520
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Shape Shape13 
         BorderStyle     =   3  'Dot
         Height          =   15
         Left            =   480
         Top             =   5400
         Visible         =   0   'False
         Width           =   6975
      End
      Begin VB.Shape Shape12 
         BorderStyle     =   3  'Dot
         Height          =   15
         Left            =   480
         Top             =   5040
         Visible         =   0   'False
         Width           =   6975
      End
      Begin VB.Shape Shape11 
         BorderStyle     =   3  'Dot
         Height          =   15
         Left            =   480
         Top             =   4680
         Visible         =   0   'False
         Width           =   6975
      End
      Begin VB.Shape Shape10 
         BorderStyle     =   3  'Dot
         Height          =   15
         Left            =   480
         Top             =   4320
         Visible         =   0   'False
         Width           =   6975
      End
      Begin VB.Shape Shape9 
         BorderStyle     =   3  'Dot
         Height          =   15
         Left            =   480
         Top             =   3960
         Visible         =   0   'False
         Width           =   6975
      End
      Begin VB.Shape Shape8 
         BorderStyle     =   3  'Dot
         Height          =   15
         Left            =   480
         Top             =   3600
         Visible         =   0   'False
         Width           =   6975
      End
      Begin VB.Shape Shape7 
         BorderStyle     =   3  'Dot
         Height          =   15
         Left            =   480
         Top             =   3240
         Visible         =   0   'False
         Width           =   6975
      End
      Begin VB.Shape Shape6 
         BorderStyle     =   3  'Dot
         Height          =   15
         Left            =   480
         Top             =   2880
         Width           =   6975
      End
      Begin VB.Label Price8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label68"
         Height          =   225
         Left            =   6675
         TabIndex        =   191
         Top             =   5160
         Width           =   660
      End
      Begin VB.Label Price7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Label34"
         Height          =   225
         Left            =   6675
         TabIndex        =   190
         Top             =   4800
         Width           =   660
      End
      Begin VB.Label Price6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Label33"
         Height          =   225
         Left            =   6675
         TabIndex        =   189
         Top             =   4440
         Width           =   660
      End
      Begin VB.Label Price5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label31"
         Height          =   225
         Left            =   6675
         TabIndex        =   188
         Top             =   4080
         Width           =   660
      End
      Begin VB.Label Price4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Label30"
         Height          =   225
         Left            =   6675
         TabIndex        =   187
         Top             =   3720
         Width           =   660
      End
      Begin VB.Label Price3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Label29"
         Height          =   225
         Left            =   6675
         TabIndex        =   186
         Top             =   3360
         Width           =   660
      End
      Begin VB.Label Price2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Label28"
         Height          =   225
         Left            =   6675
         TabIndex        =   185
         Top             =   3000
         Width           =   660
      End
      Begin VB.Label Price1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "Label27"
         Height          =   225
         Left            =   6675
         TabIndex        =   184
         Top             =   2640
         Width           =   660
      End
      Begin VB.Shape Shape5 
         BorderStyle     =   6  'Inside Solid
         Height          =   15
         Left            =   480
         Top             =   5870
         Width           =   6975
      End
      Begin VB.Shape Shape4 
         BorderStyle     =   3  'Dot
         Height          =   3680
         Left            =   5760
         Top             =   2200
         Width           =   15
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   3  'Dot
         Height          =   3620
         Left            =   4560
         Top             =   2186
         Width           =   15
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   6  'Inside Solid
         Height          =   15
         Left            =   398
         Top             =   2520
         Width           =   7095
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bill no:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   480
         TabIndex        =   163
         Top             =   1680
         Width           =   870
      End
      Begin VB.Label bTime 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "time"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6930
         TabIndex        =   162
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label bDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "date"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6960
         TabIndex        =   161
         Top             =   1680
         Width           =   525
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grant Total:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4080
         TabIndex        =   160
         Top             =   5880
         Width           =   1680
      End
      Begin VB.Label Amount 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2300.00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6240
         TabIndex        =   159
         Top             =   5880
         Width           =   1065
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price/Unit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6120
         TabIndex        =   158
         Top             =   2160
         Width           =   1365
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QTY"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4800
         TabIndex        =   157
         Top             =   2160
         Width           =   660
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Medicine Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   480
         TabIndex        =   156
         Top             =   2160
         Width           =   2085
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balaju-16, KTM"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3330
         TabIndex        =   153
         Top             =   960
         Width           =   1605
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B&&Y PHARMACY"
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
         Left            =   2310
         TabIndex        =   152
         Top             =   480
         Width           =   3435
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4065
         Left            =   355
         TabIndex        =   164
         Top             =   2160
         Width           =   7215
      End
   End
   Begin VB.CommandButton cmdBilling 
      BackColor       =   &H00FF00FF&
      Caption         =   "&Billing"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   18240
      Style           =   1  'Graphical
      TabIndex        =   150
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton GridReturn 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   18120
      Picture         =   "Admin.frx":08DF
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   9720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFC0&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   18120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9720
      Width           =   1815
   End
   Begin VB.CommandButton cmdSwitch_User 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Switch &User"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   18240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8880
      Width           =   1815
   End
   Begin VB.TextBox txtmessage 
      BackColor       =   &H00004080&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   142
      Top             =   120
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1560
      Top             =   480
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "SEARCH DETAILS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   7455
      Left            =   6960
      TabIndex        =   111
      Top             =   2040
      Visible         =   0   'False
      Width           =   8415
      Begin VB.CommandButton FrameOk 
         BackColor       =   &H0000FFFF&
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
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   140
         Top             =   6720
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.Label Label61 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label61"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   6120
         TabIndex        =   139
         Top             =   3290
         Width           =   885
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label60"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   6240
         TabIndex        =   138
         Top             =   2570
         Width           =   885
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EXP."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5280
         TabIndex        =   137
         Top             =   3240
         Width           =   720
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MFD."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5280
         TabIndex        =   136
         Top             =   2520
         Width           =   795
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label57"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   6240
         TabIndex        =   135
         Top             =   1970
         Width           =   885
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label56"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1200
         TabIndex        =   134
         Top             =   1970
         Width           =   885
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label55"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1560
         TabIndex        =   133
         Top             =   4370
         Width           =   885
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label54"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1920
         TabIndex        =   132
         Top             =   4970
         Width           =   885
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label53"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1920
         TabIndex        =   131
         Top             =   3759
         Width           =   885
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label52"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1680
         TabIndex        =   130
         Top             =   3170
         Width           =   885
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label51"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   6240
         TabIndex        =   129
         Top             =   1370
         Width           =   885
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label50"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   6720
         TabIndex        =   128
         Top             =   770
         Width           =   885
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label49"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2280
         TabIndex        =   127
         Top             =   1370
         Width           =   885
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label48"
         DataField       =   " "
         DataSource      =   "Adodc2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2520
         TabIndex        =   126
         Top             =   760
         Width           =   885
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5280
         TabIndex        =   125
         Top             =   4920
         Width           =   75
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5280
         TabIndex        =   124
         Top             =   1920
         Width           =   810
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   123
         Top             =   4920
         Width           =   1185
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   122
         Top             =   4320
         Width           =   885
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inserted on: "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   315
         Left            =   3000
         TabIndex        =   121
         Top             =   5520
         Width           =   1395
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   120
         Top             =   3120
         Width           =   900
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Updated on: "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   315
         Left            =   3120
         TabIndex        =   119
         Top             =   6000
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Batch no."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5280
         TabIndex        =   118
         Top             =   720
         Width           =   1290
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5280
         TabIndex        =   117
         Top             =   1320
         Width           =   765
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QTY:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   116
         Top             =   1920
         Width           =   765
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   115
         Top             =   3720
         Width           =   1170
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company's Details:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   114
         Top             =   2520
         Width           =   2595
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brand Name:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   113
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Generic Name:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   112
         Top             =   720
         Width           =   2025
      End
   End
   Begin VB.ComboBox ComboType 
      DataField       =   "qty"
      DataSource      =   "Adodc1"
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
      Left            =   10800
      TabIndex        =   105
      Text            =   "Combo1"
      Top             =   3600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox batchNo 
      DataField       =   "Batch"
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
      Left            =   10800
      TabIndex        =   104
      Text            =   "name of medicine"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox Quantity 
      DataField       =   "qty"
      DataSource      =   "Adodc1"
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
      Left            =   10800
      TabIndex        =   103
      Text            =   "Combo1"
      Top             =   5280
      Visible         =   0   'False
      Width           =   1815
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
      Left            =   13200
      TabIndex        =   98
      Top             =   2040
      Visible         =   0   'False
      Width           =   1935
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
      Left            =   13200
      TabIndex        =   97
      Text            =   "Combo1"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Admin.frx":11A9
      Height          =   615
      Left            =   2760
      Negotiate       =   -1  'True
      TabIndex        =   10
      Top             =   2160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   49152
      ForeColor       =   16777215
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
   Begin VB.TextBox Text13 
      DataField       =   "DOU"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4440
      TabIndex        =   93
      Text            =   "Text13"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text12 
      DataField       =   "InsertedBy"
      DataSource      =   "Adodc1"
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
      Left            =   4560
      TabIndex        =   92
      Text            =   "Text12"
      Top             =   4440
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text11 
      DataField       =   "DOE"
      DataSource      =   "Adodc1"
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
      Left            =   4560
      TabIndex        =   91
      Text            =   "Text11"
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text10 
      DataField       =   "exp"
      DataSource      =   "Adodc1"
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
      Left            =   4680
      TabIndex        =   90
      Text            =   "Text10"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      DataField       =   "mfd"
      DataSource      =   "Adodc1"
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
      Left            =   4680
      TabIndex        =   89
      Text            =   "Text9"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      DataField       =   "Price"
      DataSource      =   "Adodc1"
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
      Left            =   4560
      TabIndex        =   88
      Text            =   "Text8"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      DataField       =   "qty"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4440
      TabIndex        =   87
      Text            =   "Text7"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      DataField       =   "email"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4440
      TabIndex        =   86
      Text            =   "Text6"
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text5 
      DataField       =   "contact"
      DataSource      =   "Adodc1"
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
      Left            =   4440
      TabIndex        =   85
      Text            =   "Text5"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      DataField       =   "address"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   84
      Text            =   "Text4"
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      DataField       =   "Mfdname"
      DataSource      =   "Adodc1"
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
      Left            =   4320
      TabIndex        =   83
      Text            =   "Text3"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      DataField       =   "BrandName"
      DataSource      =   "Adodc1"
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
      Left            =   4440
      TabIndex        =   82
      Text            =   "Text2"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      DataField       =   "MedicineName"
      DataSource      =   "Adodc1"
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
      Left            =   4440
      TabIndex        =   81
      Text            =   "Text1"
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   15480
      Top             =   9240
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Bindings        =   "Admin.frx":11BE
      DataSource      =   " "
      Height          =   810
      Left            =   1200
      TabIndex        =   80
      Top             =   8280
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1429
      _Version        =   393216
      BackColor       =   4210816
      ForeColor       =   65280
      ListField       =   "MedicineName"
      BoundColumn     =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox sort 
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
      ItemData        =   "Admin.frx":11D3
      Left            =   9360
      List            =   "Admin.frx":11F8
      TabIndex        =   78
      Top             =   1440
      Width           =   4935
   End
   Begin VB.TextBox updatedBy 
      DataField       =   "UpdatedBy"
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
      Left            =   13800
      TabIndex        =   77
      Text            =   "-----------------"
      Top             =   5760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox DOU 
      DataField       =   "DOU"
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
      Left            =   13800
      TabIndex        =   76
      Text            =   "-----------------"
      Top             =   4680
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdViewClear 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Clear &All"
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
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   9720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdViewCancel 
      BackColor       =   &H00FFFFC0&
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
      Left            =   8820
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   9720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdViewSave 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Save"
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   9720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdViewDelete 
      BackColor       =   &H000000FF&
      Caption         =   "&Delete"
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
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   9600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdViewUpdate 
      BackColor       =   &H0000FF00&
      Caption         =   "&Update"
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   9600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox BName 
      DataField       =   "BrandName"
      DataSource      =   "Adodc1"
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
      Left            =   6840
      TabIndex        =   51
      Text            =   "name of medicine"
      Top             =   3600
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox CName 
      DataField       =   "Mfdname"
      DataSource      =   "Adodc1"
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
      Left            =   6840
      TabIndex        =   50
      Text            =   "name of medicine"
      Top             =   5280
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Address 
      DataField       =   "address"
      DataSource      =   "Adodc1"
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
      Left            =   6840
      ScrollBars      =   1  'Horizontal
      TabIndex        =   49
      Text            =   "name of medicine"
      Top             =   6360
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Price 
      DataField       =   "Price"
      DataSource      =   "Adodc1"
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
      Left            =   10800
      TabIndex        =   48
      Text            =   "name of medicine"
      Top             =   6360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox email 
      DataField       =   "email"
      DataSource      =   "Adodc1"
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
      Left            =   6840
      TabIndex        =   47
      Text            =   "name of medicine"
      Top             =   8520
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox DOE 
      DataField       =   "DOE"
      DataSource      =   "Adodc1"
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
      Left            =   13800
      TabIndex        =   46
      Text            =   "name of medicine"
      Top             =   3600
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Contact 
      DataField       =   "contact"
      DataSource      =   "Adodc1"
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
      Left            =   6840
      TabIndex        =   45
      Text            =   "name of medicine"
      Top             =   7440
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Insertedby 
      DataField       =   "InsertedBy"
      DataSource      =   "Adodc1"
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
      Left            =   13800
      TabIndex        =   44
      Text            =   "name of medicine"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1695
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
      Left            =   6840
      TabIndex        =   43
      Text            =   "name of medicine"
      Top             =   2400
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdViewReturn 
      BackColor       =   &H00800000&
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
      Left            =   480
      Picture         =   "Admin.frx":12B7
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdView 
      BackColor       =   &H00C0E0FF&
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
      Left            =   18240
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   19680
      Top             =   2280
   End
   Begin VB.CommandButton GridUpdate 
      BackColor       =   &H0000C000&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   18240
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   4200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton GridDelete 
      BackColor       =   &H000000FF&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   18240
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7920
      Visible         =   0   'False
      Width           =   1695
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
      Left            =   13200
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   13200
      TabIndex        =   29
      Text            =   "Combo1"
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker Exp 
      Height          =   375
      Left            =   13200
      TabIndex        =   32
      Top             =   5010
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
      Format          =   119799809
      CurrentDate     =   43589
   End
   Begin MSComCtl2.DTPicker MFD 
      Height          =   375
      Left            =   13200
      TabIndex        =   31
      Top             =   4440
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
      Format          =   119799809
      CurrentDate     =   43589
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
      Left            =   7200
      TabIndex        =   28
      Top             =   5760
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
      Left            =   7200
      TabIndex        =   27
      Top             =   5160
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
      Left            =   7200
      TabIndex        =   26
      Top             =   4560
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
      Left            =   7200
      TabIndex        =   25
      Top             =   3960
      Visible         =   0   'False
      Width           =   3855
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
      Left            =   7200
      TabIndex        =   24
      Top             =   2640
      Visible         =   0   'False
      Width           =   3855
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
      Left            =   7200
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear All"
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
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   7920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdADD 
      Caption         =   "&Add"
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   7920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H000000FF&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   18240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
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
      Left            =   9600
      TabIndex        =   9
      Text            =   "Enter the name of the medicine to search"
      Top             =   480
      Width           =   3855
   End
   Begin VB.CommandButton cmdReport 
      BackColor       =   &H00FFC0C0&
      Caption         =   "View &Report"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   18240
      Picture         =   "Admin.frx":16F9
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H0000FF00&
      Caption         =   "Up&date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   18240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FFFFC0&
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
      Left            =   13560
      Picture         =   "Admin.frx":1B3B
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   480
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdInsert 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Insert &New"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   18240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdAddAnother 
      Caption         =   "Add &Another"
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   7920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker Dexp 
      DataField       =   "exp"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   10800
      TabIndex        =   106
      Top             =   8520
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
      Format          =   119799809
      CurrentDate     =   43589
   End
   Begin MSComCtl2.DTPicker Dmfd 
      DataField       =   "mfd"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   10800
      TabIndex        =   107
      Top             =   7440
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
      Format          =   119799809
      CurrentDate     =   43589
   End
   Begin VB.TextBox Text14 
      DataField       =   "UpdatedBy"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4440
      TabIndex        =   94
      Text            =   "Text14"
      Top             =   4440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text16 
      DataField       =   "Type"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4560
      TabIndex        =   109
      Text            =   "Text16"
      Top             =   3240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text15 
      DataField       =   "Batch"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4440
      TabIndex        =   108
      Text            =   "Text15"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame search 
      BackColor       =   &H00FF8080&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   9360
      TabIndex        =   143
      Top             =   120
      Width           =   4935
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00004080&
      Caption         =   "Administrator"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   8640
      Left            =   18000
      TabIndex        =   110
      Top             =   2280
      Width           =   2295
      Begin VB.Image Image1 
         Height          =   1080
         Left            =   480
         Picture         =   "Admin.frx":1C8D
         Stretch         =   -1  'True
         Top             =   5280
         Width           =   1200
      End
   End
   Begin VB.Label lbl_no_of_records 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "289"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   2160
      TabIndex        =   174
      Top             =   1635
      Width           =   450
   End
   Begin VB.Label lbl_no_of_records1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Records: "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   150
      TabIndex        =   173
      Top             =   1635
      Width           =   1995
   End
   Begin VB.Label Label67 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "25 May 2019"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   18840
      TabIndex        =   149
      Top             =   1800
      Width           =   1320
   End
   Begin VB.Label Label66 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "SUN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   19440
      TabIndex        =   148
      Top             =   240
      Width           =   705
   End
   Begin VB.Label Label65 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "09:57:34"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   72
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1500
      Left            =   14520
      TabIndex        =   147
      Top             =   240
      Width           =   4995
   End
   Begin VB.Label Label64 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "PM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   19560
      TabIndex        =   146
      Top             =   1320
      Width           =   570
   End
   Begin VB.Label Label63 
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   72
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1980
      Left            =   14430
      TabIndex        =   145
      Top             =   165
      Width           =   5835
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "12:00:00 "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   16635
      TabIndex        =   40
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Batch no."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   30
      Left            =   10800
      TabIndex        =   102
      Top             =   2040
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   31
      Left            =   10800
      TabIndex        =   101
      Top             =   3240
      Visible         =   0   'False
      Width           =   660
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
      Left            =   22300
      TabIndex        =   100
      Top             =   720
      Width           =   180
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   420
      Left            =   19440
      TabIndex        =   99
      Top             =   360
      Width           =   705
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Batch no"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11880
      TabIndex        =   96
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11880
      TabIndex        =   95
      Top             =   2640
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Brand Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5040
      TabIndex        =   14
      Top             =   2640
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Label label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Generic Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5040
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   1920
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
      Left            =   6000
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   9075
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
      Left            =   360
      TabIndex        =   79
      Top             =   840
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Updated By"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   29
      Left            =   13800
      TabIndex        =   75
      Top             =   5400
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Last updated on"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   28
      Left            =   13800
      TabIndex        =   74
      Top             =   4320
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label SelectMedicine 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Select A Medicine"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      TabIndex        =   73
      Top             =   240
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Inserted By"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   27
      Left            =   13800
      TabIndex        =   67
      Top             =   2040
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Date of Entry"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   26
      Left            =   13800
      TabIndex        =   66
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "EXP Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   25
      Left            =   10800
      TabIndex        =   65
      Top             =   8160
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "MFD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   24
      Left            =   10800
      TabIndex        =   64
      Top             =   7080
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Price (Rs.)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   23
      Left            =   10800
      TabIndex        =   63
      Top             =   6000
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "QTY"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   22
      Left            =   10800
      TabIndex        =   62
      Top             =   4920
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "E-Mail"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   21
      Left            =   6840
      TabIndex        =   61
      Top             =   8160
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Contact"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   20
      Left            =   6840
      TabIndex        =   60
      Top             =   7080
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   19
      Left            =   6840
      TabIndex        =   59
      Top             =   6000
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   18
      Left            =   6840
      TabIndex        =   58
      Top             =   4920
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Company's Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   17
      Left            =   6840
      TabIndex        =   57
      Top             =   4320
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Brand Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   16
      Left            =   6840
      TabIndex        =   56
      Top             =   3240
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Generic Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   15
      Left            =   6840
      TabIndex        =   55
      Top             =   2040
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Label lblUnderline 
      BackColor       =   &H80000007&
      Caption         =   "Label16"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   30
      Left            =   6840
      TabIndex        =   54
      Top             =   1515
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblDetails 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Details::"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6840
      TabIndex        =   53
      Top             =   1200
      Visible         =   0   'False
      Width           =   1125
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
      Left            =   2040
      TabIndex        =   52
      Top             =   960
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Label search1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   390
      Left            =   3960
      TabIndex        =   36
      Top             =   600
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Image imgReturn 
      Height          =   495
      Left            =   6600
      Picture         =   "Admin.frx":2557
      Stretch         =   -1  'True
      Top             =   7920
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MFD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11880
      TabIndex        =   23
      Top             =   4440
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EXP."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11880
      TabIndex        =   22
      Top             =   5010
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11880
      TabIndex        =   21
      Top             =   3890
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QTY"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   11880
      TabIndex        =   20
      Top             =   3240
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5040
      TabIndex        =   19
      Top             =   5760
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5040
      TabIndex        =   18
      Top             =   5160
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5040
      TabIndex        =   17
      Top             =   4560
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5040
      TabIndex        =   16
      Top             =   3960
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company's Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5040
      TabIndex        =   15
      Top             =   3240
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.Label msg_no 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H8000000B&
      Height          =   225
      Left            =   1005
      TabIndex        =   11
      Top             =   120
      Width           =   120
   End
   Begin VB.Shape msgCircle 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   950
      Shape           =   3  'Circle
      Top             =   110
      Width           =   255
   End
   Begin VB.Image messageBox 
      Height          =   615
      Left            =   360
      Picture         =   "Admin.frx":2999
      Stretch         =   -1  'True
      Top             =   240
      Width           =   735
   End
   Begin VB.Image back_image 
      Height          =   3645
      Left            =   -2895
      Picture         =   "Admin.frx":4C0A
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   2625
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      DrawMode        =   5  'Not Copy Pen
      FillStyle       =   7  'Diagonal Cross
      Height          =   11535
      Left            =   -1080
      Top             =   10560
      Width           =   20505
   End
   Begin VB.Label Label62 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   14400
      TabIndex        =   144
      Top             =   120
      Width           =   5895
   End
   Begin VB.Image Billing_back_image 
      Height          =   10935
      Left            =   -19080
      Picture         =   "Admin.frx":1FE4B
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   20250
   End
End
Attribute VB_Name = "Admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public CN As New ADODB.Connection
Public Rs As New ADODB.Recordset
Public rs1 As New ADODB.Recordset
Public rs2 As New ADODB.Recordset

Public n As Variant

Sub DigitalTime(T As Integer)
If T = 0 Then
    Label62.Visible = False
    Label63.Visible = False
    Label64.Visible = False
    Label65.Visible = False
    Label66.Visible = False
    Label67.Visible = False
Else
    Label62.Visible = True
    Label63.Visible = True
    Label64.Visible = True
    Label65.Visible = True
    Label66.Visible = True
    Label67.Visible = True
End If
End Sub

Private Sub Billing_Cancel_Click()
BillingFrame.Visible = False
Billing_back_image.Visible = False
DigitalTime (1)
showAdmin
End Sub

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
Rs.Fields("InsertedBy").Value = "Admin"
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

Private Sub cmdBilling_Click()
BillingFrame.Visible = True
Billing_back_image.Visible = True
Billing_back_image.Top = 0
Billing_back_image.Left = 0
Label14.Left = BillingFrame.Width / 2 - Label14.Width / 2
Label15.Left = BillingFrame.Width / 2 - Label15.Width / 2
BillingFrame.Left = Admin.Width / 2 - BillingFrame.Width / 2
hideAdmin
DigitalTime (0)

txtMedicine1.Text = ""
txtQty1.Text = ""
Price1.Caption = ""

txtMedicine2.Text = ""
txtQty2.Text = ""
 Price2.Caption = ""

txtMedicine3.Text = ""
txtQty3.Text = ""
Price3.Caption = ""

txtMedicine4.Text = ""
txtQty4.Text = ""
Price4.Caption = ""

txtMedicine5.Text = ""
txtQty5.Text = ""
Price5.Caption = ""

txtMedicine6.Text = ""
txtQty6.Text = ""
Price6.Caption = ""

txtMedicine7.Text = ""
txtQty7.Text = ""
Price7.Caption = ""

txtMedicine8.Text = ""
txtQty8.Text = ""
Price8.Caption = ""

txtMedicine1.SetFocus

End Sub

Private Sub cmdCancel_Click()
'Call Form_Load
DigitalTime (1)
cmdCancel.Visible = False
cmdADD.Visible = False
cmdClear.Visible = False

back_image.Visible = True
DataGrid1.Visible = True
Shape1.Visible = True
cmdInsert.Visible = True
'cmdSearch.Value = True
cmdUpdate.Visible = True
cmdSwitch_User.Visible = True
cmdExit.Visible = True
cmdReport.Visible = True
cmdBilling.Visible = True
cmdView.Visible = True
cmdDelete.Visible = True
messageBox.Visible = True
msgCircle.Visible = True
msg_no.Visible = True
Notification.Visible = True
searchBox.Visible = True
Frame.Visible = True
cmdSearch.Visible = True
search.Visible = True
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
Admin.BackColor = &HFF0000

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
DigitalTime (0)
back_image.Visible = False
'Shape1.FillColor = &H4040&
Shape1.Visible = False
Admin.BackColor = &HFF8080
DataGrid1.Visible = False
cmdInsert.Visible = False
cmdSearch.Value = False
cmdUpdate.Visible = False
cmdSwitch_User.Visible = False
cmdExit.Visible = False
cmdReport.Visible = False
cmdBilling.Visible = False
cmdView.Visible = False
cmdDelete.Visible = False
messageBox.Visible = False
msgCircle.Visible = False
msg_no.Visible = False
Notification.Visible = False
searchBox.Visible = False
cmdSearch.Visible = False
search.Visible = False
sort.Visible = False
Frame.Visible = False

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

Private Sub cmdSearch_Click()
Rs.Close
Rs.Open "select * from medicine where MedicineName='" + searchBox.Text + "'", CN, adOpenDynamic, adLockPessimistic
If Not Rs.EOF Then
     
    Frame1.Visible = True
    Label48.Caption = Rs!medicinename
    Label49.Caption = Rs!BRANDNAME
    Label50.Caption = Rs!batch
    Label51.Caption = Rs!Type
    Label52.Caption = Rs!Mfdname
    Label53.Caption = Rs!Address
    Label54.Caption = Rs!Contact
    Label55.Caption = Rs!Email
    Label56.Caption = Rs!QTY
    Label57.Caption = Rs!Price
    Label43.Caption = Label43.Caption + Str(Rs!DOE) + " by: "
    Label43.Caption = Label43.Caption + Rs!Insertedby
    Label60.Caption = Rs!MFD
    Label61.Caption = Rs!EXP
    If Not Rs!updatedBy = "" Then
        Label41.Caption = Label41.Caption + Str(Rs!DOU) + Rs!updatedBy
        Label41.Visible = True
    End If
    Label43.Left = Frame1.Width / 2 - Label43.Width / 2
    Label41.Left = Frame1.Width / 2 - Label41.Width / 2
    FrameOk.Default = True
    
    RELOAD
Else
    MsgBox "record not found", vbInformation
End If

End Sub
Sub RELOAD()
Rs.Close
Rs.Open "select * from medicine", CN, adOpenDynamic, adLockPessimistic
End Sub

Private Sub cmdSwitch_User_Click()
Admin.Hide
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
'Admin.BackColor = &H80FF80
DataList1.Visible = True
DataList1.Top = 1400
DataList1.Left = 2000
DataList1.Width = 3989
DataList1.Height = 9700
'DataGrid2.Visible = True
Shape1.Visible = False

hideAdmin
cmdViewReturn.Visible = True
Admin.BackColor = &H80FF80      '&HC0FFC0

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

DigitalTime (0)
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
    'DigitalTime (0)
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
    
    'DigitalTime (1)
    
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
Sub hideAdmin()
back_image.Visible = False
DataGrid1.Visible = False
lbl_no_of_records.Visible = False
lbl_no_of_records1.Visible = False
cmdInsert.Visible = False
cmdSearch.Value = False
cmdUpdate.Visible = False
cmdSwitch_User.Visible = False
cmdExit.Visible = False
cmdReport.Visible = False
cmdBilling.Visible = False
cmdView.Visible = False
cmdDelete.Visible = False
messageBox.Visible = False
sort.Visible = False
msgCircle.Visible = False
msg_no.Visible = False
searchBox.Visible = False
cmdSearch.Visible = False
search.Visible = False
Frame.Visible = False
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

showAdmin
Admin.BackColor = &HFF0000
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
DigitalTime (1)
End Sub
Public Sub showAdmin()
back_image.Visible = True
DataGrid1.Visible = True
lbl_no_of_records.Visible = True
lbl_no_of_records1.Visible = True
cmdInsert.Visible = True
'cmdSearch.Value = True
cmdUpdate.Visible = True
cmdSwitch_User.Visible = True
cmdExit.Visible = True
cmdReport.Visible = True
cmdBilling.Visible = True
cmdView.Visible = True
cmdDelete.Visible = True
messageBox.Visible = True
msgCircle.Visible = True
msg_no.Visible = True
Notification.Visible = True
searchBox.Visible = True
cmdSearch.Visible = True
search.Visible = True
sort.Visible = True
Frame.Visible = True
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
Rs.Fields("UpdatedBy").Value = "Admin"
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
 

Private Sub DataCombo1_Click(Area As Integer)
'Adodc1.Recordset.Bookmark = DataCombo1.SelectedItem
Text17.Text = Rs!QTY
Text25.Text = Rs!Price
End Sub

Private Sub DataCombo2_Click(Area As Integer)
'Adodc1.Recordset .Bookmark = DataCombo2.SelectedItem
Text18.Text = Rs!QTY
Text26.Text = Rs!Price
End Sub

Private Sub DataCombo8_Click(Area As Integer)
Rs.Bookmark = DataCombo1.SelectedItem
Text24.Text = Rs!QTY
Text32.Text = Rs!Price
End Sub

Private Sub DataGrid1_Click()
GridUpdate.Visible = True
GridDelete.Visible = True
GridReturn.Visible = True
cmdInsert.Enabled = False
cmdUpdate.Enabled = False
cmdReport.Enabled = False
cmdBilling.Enabled = False
cmdView.Enabled = False
cmdDelete.Enabled = False
cmdSwitch_User.Enabled = False
'cmdSearch.Enabled = False
'search.Enabled = False
'searchBox.Enabled = False


Exchange (1)
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
'Adodc2.Recordset.Bookmark = DataList2.SelectedItem
Text24.Text = Rs!QTY
Text32.Text = Rs!Price
End Sub

Private Sub Form_Load()
DataGrid1.Height = 8730
DataGrid1.Width = 17696
DataGrid1.Left = 150 '2869
DataGrid1.Top = 2200 '1680

For i = 1 To 100
    QTY.AddItem i
Next

CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\BCA\3rd SEMESTER\PROJECT-3\PHARMACY_CONTROL_SYSTEM.mdb;Persist Security Info=False"
Rs.Open "select * from medicine", CN, adOpenDynamic, adLockPessimistic
'Rs.CursorType = adOpenDynamic
Shape1.Top = 0
Shape1.Left = 0
 
Shape1.Visible = True

sort.Clear
sort.Text = "Select for sorting & display only!!"
sort.AddItem "Display All"
sort.AddItem "Inserted by Admin"
sort.AddItem "Inserted by Staff"
sort.AddItem "Updated by Admin"
sort.AddItem "Updated by Staff"
sort.AddItem "Type:Tablet Only"
sort.AddItem "Type:Powder Only"
sort.AddItem "Type:Capule Only"
sort.AddItem "Type:Liquid Only"
sort.AddItem "Type:Equipment Only"
sort.AddItem "Type:Other"
sort.AddItem "Sort by Brand Name"
sort.AddItem "Sort by Company"
sort.AddItem "Sort by MFD"
sort.AddItem "Sort by EXP"
sort.AddItem "Sort by DOE"
sort.AddItem "Sort by DOU"

'Rs.MoveLast
lbl_no_of_records.Caption = Str(Adodc1.Recordset.RecordCount)
End Sub

 
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Text1.Text = X
'Text2.Text = Y

End Sub

Private Sub FrameOK_Click()
Frame1.Visible = False
cmdSearch.Default = False
cmdView.Default = True
'cmdSearch.Enabled = False
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
cmdReport.Enabled = True
cmdBilling.Enabled = True
cmdView.Enabled = True
cmdDelete.Enabled = True
cmdSwitch_User.Enabled = True
cmdSearch.Enabled = True
search.Enabled = True
searchBox.Enabled = True

Exchange (0)
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
'cmdExit.Top = 6720 'GridReturn.Top
'cmdExit.Left = GridReturn.Left
'GridReturn.Left = L
'GridReturn.Top = 6000
'Frame.Height = 6775 'Frame.Height + GridReturn.Height

Else
'cmdExit.Top = 6000
'GridReturn.Top = 6720
'Frame.Height = 5765 'Frame.Height - GridReturn.Height
End If
End Sub

Private Sub GridUpdate_Click()
'Call cmdUpdate_Click
FormUpdate.Show 0
Admin.Enabled = False
'Rs.MoveNext
FormUpdate.GN.Text = Text1.Text '   Rs.Fields("MedicineName")
FormUpdate.BN.Text = Text2.Text '.Rs.Fields("BrandName")
'Name.Text = Text3.Text     'Admin.Rs.Fields("Mfdname")
FormUpdate.Text1.Text = Text3.Text
FormUpdate.Address.Text = Text4.Text ' Admin.Rs.Fields("address")
FormUpdate.Contact.Text = Text5.Text ' Admin.Rs.Fields("contact")
FormUpdate.Email.Text = Text6.Text ' Admin.Rs.Fields("email")
FormUpdate.QTY.Text = Text7.Text ' Admin.Rs.Fields("qty")
FormUpdate.Price.Text = Text8.Text '     Admin.Rs.Fields("Price")
FormUpdate.MFD.Value = Text9.Text ' Admin.Rs.Fields("mfd")
FormUpdate.EXP.Value = Text10.Text ' Admin.Rs.Fields("exp")
FormUpdate.Text2.Text = Text15.Text
FormUpdate.Combo1.Text = Text16.Text
FormUpdate.Label13.Caption = "Inserted on: "
FormUpdate.Label16.Caption = "Updated on: "
If Text13.Text = "" Then
    FormUpdate.Label16.Caption = FormUpdate.Label16.Caption & "None " & "    By: none"
Else
    FormUpdate.Label16.Caption = FormUpdate.Label16.Caption + Text13.Text + "  By: " + Text14.Text
End If
    
FormUpdate.Label13.Caption = FormUpdate.Label13.Caption + " " + Text11.Text
FormUpdate.Label13.Caption = FormUpdate.Label13.Caption + "   By: " + Text12.Text

FormUpdate.Label16.Left = FormUpdate.Width / 2 - FormUpdate.Label16.Width / 2
FormUpdate.Label13.Left = FormUpdate.Width / 2 - FormUpdate.Label13.Width / 2

'back_image.Top = 0
'back_image.Left = 0
'back_image.Width = Admin.Width
'back_image.Height = Admin.Height
'hideAdmin
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
If txtmessage.Visible = False Then
    txtmessage.Visible = True
Else
    txtmessage.Visible = False
End If
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
cmdSearch.Default = True
End Sub

Private Sub searchBox_LostFocus()
searchBox.Text = "Enter the name of the medicine"
searchBox.FontSize = 10
searchBox.ForeColor = &H8000000A
End Sub

Private Sub sort_Click()
DataGrid2.Height = 8730
DataGrid2.Width = 17696
DataGrid2.Left = 150 '2869
DataGrid2.Top = 2200

 Dim name As Variant
 
If sort.Text = "Inserted by Admin" Then
    DataGrid2.Visible = True
    DataGrid1.Visible = False
    Adodc3.RecordSource = "select *from medicine where InsertedBy='Admin'"
    Adodc3.Refresh
    lbl_no_of_records.Caption = Str(Adodc3.Recordset.RecordCount)

ElseIf sort.Text = "Inserted by Staff" Then
    DataGrid2.Visible = True
    DataGrid1.Visible = False
    Adodc3.RecordSource = "select *from medicine where InsertedBy='Staff'"
    Adodc3.Refresh
    lbl_no_of_records.Caption = Str(Adodc3.Recordset.RecordCount)
    
ElseIf sort.Text = "Display All" Then
    DataGrid1.Visible = True
    DataGrid2.Visible = False
    DataGrid1.Refresh
    Adodc1.Refresh
    lbl_no_of_records.Caption = Str(Adodc3.Recordset.RecordCount)

ElseIf sort.Text = "Updated by Admin" Then
    DataGrid2.Visible = True
    DataGrid1.Visible = False
    Adodc3.RecordSource = "select *from medicine where UpdatedBy='Admin'"
    Adodc3.Refresh
    lbl_no_of_records.Caption = Str(Adodc3.Recordset.RecordCount)
    
ElseIf sort.Text = "Inserted by Staff" Then
    DataGrid2.Visible = True
    DataGrid1.Visible = False
    Adodc3.RecordSource = "select *from medicine where UpdatedBy='Staff'"
    Adodc3.Refresh
    lbl_no_of_records.Caption = Str(Adodc3.Recordset.RecordCount)

ElseIf sort.Text = "Type:Tablet Only" Then
    DataGrid2.Visible = True
    DataGrid1.Visible = False
    Adodc3.RecordSource = "select *from medicine where Type='Tablet'"
    Adodc3.Refresh
    lbl_no_of_records.Caption = Str(Adodc3.Recordset.RecordCount)

ElseIf sort.Text = "Type:Powder Only" Then
    DataGrid2.Visible = True
    DataGrid1.Visible = False
    Adodc3.RecordSource = "select *from medicine where Type='Powder'"
    Adodc3.Refresh
    lbl_no_of_records.Caption = Str(Adodc3.Recordset.RecordCount)

ElseIf sort.Text = "Type:Capsule Only" Then
    DataGrid2.Visible = True
    DataGrid1.Visible = False
    Adodc3.RecordSource = "select *from medicine where Type='Capsule'"
    Adodc3.Refresh
    lbl_no_of_records.Caption = Str(Adodc3.Recordset.RecordCount)
    
ElseIf sort.Text = "Type:Equipment Only" Then
    DataGrid2.Visible = True
    DataGrid1.Visible = False
    Adodc3.RecordSource = "select *from medicine where Type='Equipment'"
    Adodc3.Refresh
    lbl_no_of_records.Caption = Str(Adodc3.Recordset.RecordCount)
    
ElseIf sort.Text = "Type:Other" Then
    DataGrid2.Visible = True
    DataGrid1.Visible = False
    Adodc3.RecordSource = "select *from medicine where Type='Other'"
    Adodc3.Refresh
    lbl_no_of_records.Caption = Str(Adodc3.Recordset.RecordCount)
    
ElseIf sort.Text = "Type:Liquid Only" Then
    DataGrid2.Visible = True
    DataGrid1.Visible = False
    Adodc3.RecordSource = "select *from medicine where Type='Liquid'"
    Adodc3.Refresh
    lbl_no_of_records.Caption = Str(Adodc3.Recordset.RecordCount)
    
ElseIf sort.Text = "Sort by Brand Name" Then
    
    name = InputBox("Enter the brand name", BRANDNAME)
    If name <> "" Then
        Adodc3.RecordSource = "select *from medicine where BrandName='" + name + "'"
        Adodc3.Refresh
        lbl_no_of_records.Caption = Str(Adodc3.Recordset.RecordCount)
        DataGrid1.Visible = False
        DataGrid2.Visible = True
    Else
        MsgBox "Please enter the brand name of the medicine.", vbCritical
    End If

ElseIf sort.Text = "Sort by Company" Then
   
    name = InputBox("Enter the name of the company", COMPANY_NAME)
    If name <> "" Then
        Adodc3.RecordSource = "select *from medicine where Mfdname='" + name + "'"
        Adodc3.Refresh
        lbl_no_of_records.Caption = Str(Adodc3.Recordset.RecordCount)
        DataGrid1.Visible = False
        DataGrid2.Visible = True
    Else
        MsgBox "Please enter the name of the company.", vbCritical
    End If
End If
End Sub

 

 

Private Sub Timer1_Timer()
Static i As Integer
Label13.Caption = Time 'Hour(Time)
Label65.Caption = Format(Time, "HH:MM:SS")
Label64.Caption = Format(Time, "AMPM")
Label66.Caption = UCase(Format(Date, "ddd"))
Label67.Caption = Format(Date, "d") + " " + Format(Date, "mmmm") + " " + Format(Date, "yyyy")

'Label13.Caption = Label13.Caption + ":"
'Label13.Caption = Label13.Caption + Str(Minute(Time))
'Label13.Caption = Label13.Caption + ":" + Str(Second(Time))
'Label14.Caption = Format(Date, "d")
Label19.Caption = Date 'Format(Date, "mmmm")
'Label20.Caption = Format(Date, "yyyy")
'Label15.Caption = Format(Date, "dddd")
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

Private Sub Timer2_Timer()
Adodc1.Refresh
'TxtMessage.Text = ""
Dim d1, d2 As Date
Dim name As String
Dim d As Integer
Do While Not Rs.EOF
    d1 = Rs!MFD
    d2 = Rs!EXP
    name = Rs!medicinename
    d = DateDiff("d", d1, d2)
    If d < 5 Then
        txtmessage.Text = txtmessage.Text + name + ", "
        Beep
        msg_no.Caption = Val(msg_no.Caption) + 1
    End If
    Rs.MoveNext
Loop
Timer2.Enabled = False
Rs.MoveFirst
txtmessage.Text = txtmessage.Text + " have expiry date less than 5 days..                                                                                           "
End Sub

Private Sub Timer3_Timer()
Dim m, Y
bDate.Caption = Format(Date, "d")
m = Format(Date, "mmmm")
Y = Format(Date, "yyyy")
bDate.Caption = bDate.Caption + " " + m + " " + Y
'bDate.Caption = bDate.Caption + y
bTime.Caption = Time
End Sub

Private Sub txtMedicine1_Change()
' n = txtMedicine1.Text
 '   Adodc3.RecordSource = "select *from medicine where MedicineName='" + n + "'"
 '    txtQty1.Text = Rs!QTY
 '   txtPrice.Text = Rs!Price
 '   Adodc3.Refresh
End Sub

Private Sub txtMedicine1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Rs.Close
 Rs.Open "select * from medicine where MedicineName='" + txtMedicine1.Text + "'", CN, adOpenDynamic, adLockPessimistic
 If Not Rs.EOF Then
    txtQty1.Visible = True
    Price1.Visible = True
    txtQty1.Text = Rs!QTY
    Price1.Caption = Rs!Price
    
    txtQty1.SetFocus
    txtQty1.SelStart = 0
    txtQty1.SelLength = Len(txtQty1.Text)

    'Price1.SelStart = 0
    'Price1.SelLength = Len(Price1.Text)
    RELOAD
 Else
    MsgBox "There is no such medicine/equipment"
 End If
End If
End Sub



Private Sub txtMedicine2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Rs.Close
 Rs.Open "select * from medicine where MedicineName='" + txtMedicine2.Text + "'", CN, adOpenDynamic, adLockPessimistic
 If Not Rs.EOF Then
    txtQty2.Visible = True
    Price2.Visible = True
    txtQty2.Text = Rs!QTY
    Price2.Caption = Rs!Price
    txtQty2.SetFocus
    txtQty2.SelStart = 0
    txtQty2.SelLength = Len(txtQty1.Text)
    RELOAD
 Else
    MsgBox "There is no such medicine/equipment"
 End If
End If
End Sub

 

Private Sub txtMedicine3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Rs.Close
 Rs.Open "select * from medicine where MedicineName='" + txtMedicine3.Text + "'", CN, adOpenDynamic, adLockPessimistic
 If Not Rs.EOF Then
    txtQty3.Visible = True
    Price3.Visible = True
    txtQty3.Text = Rs!QTY
    Price3.Caption = Rs!Price
    txtQty3.SetFocus
    txtQty3.SelStart = 0
    txtQty3.SelLength = Len(txtQty1.Text)
    RELOAD
 Else
    MsgBox "There is no such medicine/equipment"
 End If
End If
End Sub

 

Private Sub txtMedicine4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Rs.Close
 Rs.Open "select * from medicine where MedicineName='" + txtMedicine4.Text + "'", CN, adOpenDynamic, adLockPessimistic
 If Not Rs.EOF Then
    txtQty4.Visible = True
    Price4.Visible = True
    txtQty4.Text = Rs!QTY
    Price4.Caption = Rs!Price
    txtQty4.SetFocus
    txtQty4.SelStart = 0
    txtQty4.SelLength = Len(txtQty1.Text)
    RELOAD
 Else
    MsgBox "There is no such medicine/equipment"
 End If
End If
End Sub
 

Private Sub txtMedicine5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Rs.Close
 Rs.Open "select * from medicine where MedicineName='" + txtMedicine5.Text + "'", CN, adOpenDynamic, adLockPessimistic
 If Not Rs.EOF Then
    txtQty5.Visible = True
    Price5.Visible = True
    txtQty5.Text = Rs!QTY
    Price5.Caption = Rs!Price
    txtQty5.SetFocus
    txtQty5.SelStart = 0
    txtQty5.SelLength = Len(txtQty1.Text)
    RELOAD
 Else
    MsgBox "There is no such medicine/equipment"
 End If
End If
End Sub

 

Private Sub txtMedicine6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Rs.Close
 Rs.Open "select * from medicine where MedicineName='" + txtMedicine6.Text + "'", CN, adOpenDynamic, adLockPessimistic
 If Not Rs.EOF Then
    txtQty6.Visible = True
    Price6.Visible = True
    txtQty6.Text = Rs!QTY
    Price6.Caption = Rs!Price
    txtQty6.SetFocus
    txtQty6.SelStart = 0
    txtQty6.SelLength = Len(txtQty1.Text)
    RELOAD
 Else
    MsgBox "There is no such medicine/equipment"
 End If
End If
End Sub

 

Private Sub txtMedicine7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Rs.Close
 Rs.Open "select * from medicine where MedicineName='" + txtMedicine7.Text + "'", CN, adOpenDynamic, adLockPessimistic
 If Not Rs.EOF Then
    txtQty7.Visible = True
    Price7.Visible = True
    txtQty7.Text = Rs!QTY
    Price7.Caption = Rs!Price
    txtQty7.SetFocus
    txtQty7.SelStart = 0
    txtQty7SelLength = Len(txtQty1.Text)
    RELOAD
 Else
    MsgBox "There is no such medicine/equipment"
 End If
End If
End Sub

 

Private Sub txtMedicine8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Rs.Close
 Rs.Open "select * from medicine where MedicineName='" + txtMedicine8.Text + "'", CN, adOpenDynamic, adLockPessimistic
 If Not Rs.EOF Then
    txtQty8.Visible = True
    Price8.Visible = True
    txtQty8.Text = Rs!QTY
    Price8.Caption = Rs!Price
    txtQty8.SetFocus
    txtQty8.SelStart = 0
    txtQty8.SelLength = Len(txtQty1.Text)
    RELOAD
 Else
    MsgBox "There is no such medicine/equipment"
 End If
End If
End Sub

Private Sub txtMedicine9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Rs.Close
 Rs.Open "select * from medicine where MedicineName='" + txtMedicine9.Text + "'", CN, adOpenDynamic, adLockPessimistic
 If Not Rs.EOF Then
    txtQty9.Visible = True
    Price9.Visible = True
    txtQty9.Text = Rs!QTY
    Price9.Caption = Rs!Price
    txtQty9.SetFocus
    txtQty9.SelStart = 0
    txtQty9.SelLength = Len(txtQty1.Text)
    RELOAD
 Else
    MsgBox "There is no such medicine/equipment"
 End If
End If
End Sub

Private Sub txtQty1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMedicine2.Visible = True
    txtMedicine2.SetFocus
    Shape7.Visible = True
    Amount.Caption = Val(Price1.Caption)
End If
End Sub

Private Sub txtQty2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMedicine3.Visible = True
    txtMedicine3.SetFocus
    Shape8.Visible = True
    Amount.Caption = Amount.Caption + Val(Price2.Caption)
End If
End Sub

Private Sub txtQty3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMedicine4.Visible = True
    txtMedicine4.SetFocus
    Shape9.Visible = True
    Amount.Caption = Amount.Caption + Val(Price3.Caption)
End If
End Sub

Private Sub txtQty4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMedicine5.Visible = True
    txtMedicine5.SetFocus
    Shape10.Visible = True
    Amount.Caption = Amount.Caption + Val(Price4.Caption)
End If
End Sub

Private Sub txtQty5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMedicine6.Visible = True
    txtMedicine6.SetFocus
    Shape11.Visible = True
    Amount.Caption = Amount.Caption + Val(Price5.Caption)
End If
End Sub

Private Sub txtQty6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMedicine7.Visible = True
    txtMedicine7.SetFocus
    Shape12.Visible = True
    Amount.Caption = Amount.Caption + Val(Price6.Caption)
End If
End Sub

Private Sub txtQty7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMedicine8.Visible = True
    txtMedicine8.SetFocus
    Shape13.Visible = True
    Amount.Caption = Amount.Caption + Val(Price7.Caption)
End If
End Sub

Private Sub txtQty8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMedicine9.Visible = True
    txtMedicine9.SetFocus
    Amount.Caption = Amount.Caption + Val(Price8.Caption)
End If
End Sub

Private Sub txtQty9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'txtMedicine9.Visible = True
    'txtMedicine9.SetFocus
    Billing_Print.SetFocus
    Amount.Caption = Amount.Caption + Val(Price9.Caption)
End If
End Sub
