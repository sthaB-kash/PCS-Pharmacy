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
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Begin VB.TextBox Text32 
         Height          =   330
         Left            =   5880
         TabIndex        =   188
         Text            =   "Text32"
         Top             =   5160
         Width           =   1575
      End
      Begin VB.TextBox Text31 
         Height          =   330
         Left            =   5880
         TabIndex        =   187
         Text            =   "Text31"
         Top             =   4800
         Width           =   1575
      End
      Begin VB.TextBox Text30 
         Height          =   330
         Left            =   5880
         TabIndex        =   186
         Text            =   "Text30"
         Top             =   4440
         Width           =   1575
      End
      Begin VB.TextBox Text29 
         Height          =   330
         Left            =   5880
         TabIndex        =   185
         Text            =   "Text29"
         Top             =   4080
         Width           =   1575
      End
      Begin VB.TextBox Text28 
         Height          =   330
         Left            =   5880
         TabIndex        =   184
         Text            =   "Text28"
         Top             =   3720
         Width           =   1575
      End
      Begin VB.TextBox Text27 
         Height          =   330
         Left            =   5880
         TabIndex        =   183
         Text            =   "Text27"
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox Text26 
         Height          =   330
         Left            =   5880
         TabIndex        =   182
         Text            =   "Text26"
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox Text25 
         Height          =   330
         Left            =   5880
         TabIndex        =   181
         Text            =   "Text25"
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox Text24 
         Height          =   330
         Left            =   4680
         TabIndex        =   180
         Text            =   "Text24"
         Top             =   5160
         Width           =   975
      End
      Begin VB.TextBox Text23 
         Height          =   330
         Left            =   4680
         TabIndex        =   179
         Text            =   "Text23"
         Top             =   4800
         Width           =   975
      End
      Begin VB.TextBox Text22 
         Height          =   330
         Left            =   4680
         TabIndex        =   178
         Text            =   "Text22"
         Top             =   4440
         Width           =   975
      End
      Begin VB.TextBox Text21 
         Height          =   330
         Left            =   4680
         TabIndex        =   177
         Text            =   "Text21"
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox Text20 
         Height          =   330
         Left            =   4680
         TabIndex        =   176
         Text            =   "Text20"
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox Text19 
         Height          =   330
         Left            =   4680
         TabIndex        =   175
         Text            =   "Text19"
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox Text18 
         Height          =   330
         Left            =   4680
         TabIndex        =   174
         Text            =   "Text18"
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox Text17 
         Height          =   330
         Left            =   4680
         TabIndex        =   173
         Text            =   "Text17"
         Top             =   2640
         Width           =   975
      End
      Begin MSDataListLib.DataCombo DataCombo8 
         Bindings        =   "Admin.frx":08CA
         Height          =   360
         Left            =   360
         TabIndex        =   172
         Top             =   5160
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "MedicineName"
         Text            =   "DataCombo8"
      End
      Begin MSDataListLib.DataCombo DataCombo7 
         Bindings        =   "Admin.frx":08DF
         Height          =   360
         Left            =   360
         TabIndex        =   171
         Top             =   4800
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "MedicineName"
         Text            =   "DataCombo7"
      End
      Begin MSDataListLib.DataCombo DataCombo6 
         Bindings        =   "Admin.frx":08F4
         Height          =   360
         Left            =   360
         TabIndex        =   170
         Top             =   4440
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "MedicineName"
         Text            =   "DataCombo6"
      End
      Begin MSDataListLib.DataCombo DataCombo5 
         Bindings        =   "Admin.frx":0909
         Height          =   360
         Left            =   360
         TabIndex        =   169
         Top             =   4080
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "MedicineName"
         Text            =   "DataCombo5"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo4 
         Bindings        =   "Admin.frx":091E
         Height          =   360
         Left            =   360
         TabIndex        =   168
         Top             =   3720
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "MedicineName"
         Text            =   "DataCombo4"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "Admin.frx":0933
         Height          =   360
         Left            =   360
         TabIndex        =   167
         Top             =   3360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "MedicineName"
         Text            =   "DataCombo3"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "Admin.frx":0948
         Height          =   360
         Left            =   360
         TabIndex        =   166
         Top             =   3000
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "MedicineName"
         Text            =   "DataCombo2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Admin.frx":095D
         Height          =   360
         Left            =   360
         TabIndex        =   165
         Top             =   2640
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "MedicineName"
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Timer Timer3 
         Interval        =   999
         Left            =   7200
         Top             =   840
      End
      Begin VB.CommandButton Billing_Cancel 
         BackColor       =   &H000080FF&
         Caption         =   "&Cancel"
         Default         =   -1  'True
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
      Begin VB.Shape Shape5 
         BorderStyle     =   3  'Dot
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
         BorderStyle     =   3  'Dot
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
      Begin VB.Label Label23 
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
         Caption         =   "Medicine_Name"
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
         Width           =   2175
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
      Top             =   6120
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
      Picture         =   "Admin.frx":0972
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
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   142
      Top             =   240
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1560
      Top             =   480
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1320
      Top             =   9480
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
      Bindings        =   "Admin.frx":123C
      Height          =   615
      Left            =   2760
      Negotiate       =   -1  'True
      TabIndex        =   10
      Top             =   1800
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
      Bindings        =   "Admin.frx":1251
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
      ItemData        =   "Admin.frx":1266
      Left            =   9360
      List            =   "Admin.frx":1282
      Sorted          =   -1  'True
      TabIndex        =   78
      Text            =   "Select for sorting"
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
      Picture         =   "Admin.frx":1318
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
      Format          =   118947841
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
      Format          =   118947841
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
      Picture         =   "Admin.frx":175A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
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
      Picture         =   "Admin.frx":1B9C
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
      Format          =   118947841
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
      Format          =   118947841
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Admin.frx":1CEE
      Height          =   615
      Left            =   2760
      Negotiate       =   -1  'True
      TabIndex        =   141
      Top             =   2640
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
         Picture         =   "Admin.frx":1D03
         Stretch         =   -1  'True
         Top             =   5280
         Width           =   1200
      End
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
      Caption         =   "SAT"
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
      Width           =   675
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
      Caption         =   "AM"
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
      Width           =   600
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
      Left            =   20040
      TabIndex        =   52
      Top             =   1800
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
      Picture         =   "Admin.frx":25CD
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
      Picture         =   "Admin.frx":2A0F
      Stretch         =   -1  'True
      Top             =   240
      Width           =   735
   End
   Begin VB.Image back_image 
      Height          =   3645
      Left            =   -2895
      Picture         =   "Admin.frx":4C80
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
      Picture         =   "Admin.frx":1FEC1
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
If txtGName.Text = "" Or txtBName.Text = "" Or txtCName.Text = "" Or txtAddress.Text = "" Or txtContact.Text = "" Or txtEmail.Text = "" Or txtPrice.Text = "" Or qty.Text = "" Then
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
Rs.Fields("qty").Value = Val(qty.Text)
Rs.Fields("Price").Value = Val(txtPrice.Text)
Rs.Fields("mfd").Value = MFD.Value
Rs.Fields("exp").Value = Exp.Value
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
qty.Enabled = False
MFD.Enabled = False
Exp.Enabled = False


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
qty.Enabled = True
MFD.Enabled = True
Exp.Enabled = True

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
label2.Visible = False
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
qty.Visible = False
MFD.Visible = False
Exp.Visible = False


txtGName.Enabled = True
txtBName.Enabled = True
txtCName.Enabled = True
txtAddress.Enabled = True
txtContact.Enabled = True
txtEmail.Enabled = True
MType.Enabled = True
txtBatch.Enabled = True
txtPrice.Enabled = True
qty.Enabled = True
MFD.Enabled = True
Exp.Enabled = True




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
qty.Text = ""
MFD.Value = Date
Exp.Value = Date
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
label2.Visible = True
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
qty.Visible = True
MFD.Visible = True
Exp.Visible = True

Call cmdClear_Click
'Rs.AddNew

End Sub

Private Sub cmdSearch_Click()
Rs.Close
Rs.Open "select * from medicine where MedicineName='" + searchBox.Text + "'", CN, adOpenDynamic, adLockPessimistic
If Not Rs.EOF Then
     
    Frame1.Visible = True
    Label48.Caption = Rs!medicinename
    Label49.Caption = Rs!BrandName
    Label50.Caption = Rs!batch
    Label51.Caption = Rs!Type
    Label52.Caption = Rs!Mfdname
    Label53.Caption = Rs!Address
    Label54.Caption = Rs!Contact
    Label55.Caption = Rs!email
    Label56.Caption = Rs!qty
    Label57.Caption = Rs!Price
    Label43.Caption = Label43.Caption + Str(Rs!DOE) + " by: "
    Label43.Caption = Label43.Caption + Rs!Insertedby
    Label60.Caption = Rs!MFD
    Label61.Caption = Rs!Exp
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
    email.DataSource = Adodc1
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
    email.Visible = True
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
    email.Enabled = False
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
    email.Visible = False
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
email.Enabled = False
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
email.Text = ""
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
End Sub
Public Sub showAdmin()
back_image.Visible = True
DataGrid1.Visible = True
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
Rs.Fields("email").Value = email.Text
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
email.Enabled = False
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
email.Enabled = True
Quantity.Enabled = True
Price.Enabled = True
Dmfd.Enabled = True
Dexp.Enabled = True
'Insertedby.Enabled = True

End Sub
 

Private Sub DataCombo1_Click(Area As Integer)

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
Adodc1.Recordset.Bookmark = DataList2.SelectedItem
End Sub

Private Sub Form_Load()
DataGrid1.Height = 8730
DataGrid1.Width = 17696
DataGrid1.Left = 150 '2869
DataGrid1.Top = 2200 '1680

For i = 1 To 100
    qty.AddItem i
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
FormUpdate.email.Text = Text6.Text ' Admin.Rs.Fields("email")
FormUpdate.qty.Text = Text7.Text ' Admin.Rs.Fields("qty")
FormUpdate.Price.Text = Text8.Text '     Admin.Rs.Fields("Price")
FormUpdate.MFD.Value = Text9.Text ' Admin.Rs.Fields("mfd")
FormUpdate.Exp.Value = Text10.Text ' Admin.Rs.Fields("exp")
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

Private Sub Timer1_Timer()
Static i As Integer
Label13.Caption = Time 'Hour(Time)
Label65.Caption = Time
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
Dim Name As String
Dim d As Integer
Do While Not Rs.EOF
    d1 = Rs!MFD
    d2 = Rs!Exp
    Name = Rs!medicinename
    d = DateDiff("d", d1, d2)
    If d < 5 Then
        txtmessage.Text = txtmessage.Text + Name + ", "
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
