VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "CMSComm"
   ClientHeight    =   10110
   ClientLeft      =   165
   ClientTop       =   390
   ClientWidth     =   12150
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10110
   ScaleWidth      =   12150
   StartUpPosition =   1  'CenterOwner
   WindowState     =   1  'Minimized
   Begin VB.PictureBox picSCADA 
      Enabled         =   0   'False
      Height          =   6315
      Left            =   0
      Picture         =   "frmMain.frx":0442
      ScaleHeight     =   6255
      ScaleWidth      =   9705
      TabIndex        =   78
      Top             =   0
      Width           =   9765
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   0
         Left            =   1020
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   117
         Top             =   1860
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   2
         Left            =   1020
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   116
         Top             =   5160
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   3
         Left            =   1440
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   115
         Top             =   1860
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   4
         Left            =   1440
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   114
         Top             =   3060
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   5
         Left            =   1440
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   113
         Top             =   5160
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   6
         Left            =   2220
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   112
         Top             =   3060
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   7
         Left            =   2220
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   111
         Top             =   4860
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   8
         Left            =   2220
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   110
         Top             =   5160
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   9
         Left            =   2640
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   109
         Top             =   3060
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   10
         Left            =   2640
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   108
         Top             =   4860
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   11
         Left            =   2640
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   107
         Top             =   5160
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   12
         Left            =   3060
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   106
         Top             =   3060
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   13
         Left            =   3840
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   105
         Top             =   3060
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   14
         Left            =   3840
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   104
         Top             =   5160
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   15
         Left            =   4260
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   103
         Top             =   1860
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   16
         Left            =   4260
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   102
         Top             =   3060
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   17
         Left            =   4260
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   101
         Top             =   4860
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   18
         Left            =   4260
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   100
         Top             =   5160
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   19
         Left            =   5100
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   99
         Top             =   1860
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   20
         Left            =   5100
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   98
         Top             =   3060
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   21
         Left            =   5100
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   97
         Top             =   5160
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   22
         Left            =   5520
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   96
         Top             =   4860
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   23
         Left            =   5940
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   95
         Top             =   4860
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   24
         Left            =   6300
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   94
         Top             =   1860
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   25
         Left            =   6300
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   93
         Top             =   5160
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   26
         Left            =   6720
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   92
         Top             =   3060
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   27
         Left            =   6720
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   91
         Top             =   4860
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   28
         Left            =   6720
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   90
         Top             =   5160
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   29
         Left            =   7140
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   89
         Top             =   5160
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   30
         Left            =   7500
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   88
         Top             =   3060
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   31
         Left            =   7500
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   87
         Top             =   5160
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   32
         Left            =   7920
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   86
         Top             =   4860
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   33
         Left            =   8280
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   85
         Top             =   3060
         Width           =   255
      End
      Begin VB.Timer tmrUpdateSCADA 
         Interval        =   2000
         Left            =   6900
         Top             =   4050
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   1
         Left            =   1020
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   84
         Top             =   3060
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   34
         Left            =   8340
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   83
         Top             =   4860
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   35
         Left            =   8700
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   82
         Top             =   3060
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   36
         Left            =   9120
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   81
         Top             =   1860
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   37
         Left            =   9120
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   80
         Top             =   3060
         Width           =   255
      End
      Begin VB.CheckBox chkSCADAstatus 
         BackColor       =   &H80000009&
         Height          =   225
         Index           =   38
         Left            =   9120
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   79
         Top             =   5160
         Width           =   255
      End
   End
   Begin VB.Timer tmrCOMMFailed 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   2280
   End
   Begin VB.PictureBox picWinsock 
      Align           =   2  'Align Bottom
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1515
      ScaleWidth      =   12090
      TabIndex        =   5
      Top             =   7590
      Width           =   12150
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   7320
         TabIndex        =   24
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   7320
         TabIndex        =   23
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "State:"
         Height          =   255
         Index           =   8
         Left            =   6240
         TabIndex        =   22
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "State:"
         Height          =   255
         Index           =   7
         Left            =   3120
         TabIndex        =   21
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   7320
         TabIndex        =   19
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Local Port:"
         Height          =   255
         Index           =   6
         Left            =   6240
         TabIndex        =   18
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   7320
         TabIndex        =   17
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4200
         TabIndex        =   16
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4200
         TabIndex        =   15
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4200
         TabIndex        =   14
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   13
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   12
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "OPC Port:"
         Height          =   255
         Index           =   5
         Left            =   6240
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "OPC IP:"
         Height          =   255
         Index           =   4
         Left            =   6240
         TabIndex        =   10
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "CMS Port:"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "CMS IP:"
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   8
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Local Port:"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Local Name:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox picButton 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FF8080&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   12090
      TabIndex        =   2
      Top             =   9165
      Width           =   12150
      Begin VB.CommandButton cmdShScada 
         Caption         =   "Show SCADA Status"
         Height          =   375
         Left            =   7260
         TabIndex        =   118
         Top             =   120
         Width           =   1755
      End
      Begin VB.CommandButton cmdShStatus 
         Caption         =   "Show CMS Status"
         Height          =   375
         Left            =   5280
         TabIndex        =   77
         Top             =   120
         Width           =   1635
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear Msg"
         Height          =   375
         Left            =   3720
         TabIndex        =   20
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton cmdConnectOPC 
         Caption         =   "Connect OPC"
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdConnectCMS 
         Caption         =   "Connect CMS"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Timer tmrReListenOPC 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3480
      Top             =   2880
   End
   Begin VB.ListBox lstMsg 
      Height          =   840
      ItemData        =   "frmMain.frx":A7D94
      Left            =   7080
      List            =   "frmMain.frx":A7D96
      TabIndex        =   1
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Timer tmrHBOPC 
      Interval        =   60000
      Left            =   5160
      Top             =   3240
   End
   Begin MSWinsockLib.Winsock sckTagRx 
      Left            =   1680
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Timer tmrUpdateExcel 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4500
      Top             =   2220
   End
   Begin VB.Timer tmrReListenCMS 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3360
      Top             =   2160
   End
   Begin VB.Timer tmrTCPStatus 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2880
      Top             =   3120
   End
   Begin MSWinsockLib.Winsock sckCMSAlarm 
      Left            =   2160
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   9900
      Width           =   12150
      _ExtentX        =   21431
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13838
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1560
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1200
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A7D98
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A7EAA
            Key             =   "ClearScn"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A82FE
            Key             =   "Connect"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A8752
            Key             =   "Disconnect"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A8BA6
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A8FFA
            Key             =   "Green"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A944E
            Key             =   "Red"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A98A2
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A99B4
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A9AC6
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A9BD8
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A9CEA
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A9DFC
            Key             =   "Align Right"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A9F0E
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AA020
            Key             =   "Map Network Drive"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AA132
            Key             =   "Macro"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AA244
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AA356
            Key             =   "Camera"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picCMSstatus 
      Enabled         =   0   'False
      Height          =   5505
      Left            =   840
      Picture         =   "frmMain.frx":AA468
      ScaleHeight     =   5445
      ScaleWidth      =   9675
      TabIndex        =   25
      Top             =   4200
      Width           =   9735
      Begin VB.Timer tmrUpdateStatus 
         Interval        =   2000
         Left            =   7140
         Top             =   4680
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   76
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   75
         Top             =   3660
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   74
         Top             =   3960
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   73
         Top             =   1260
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   72
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   5
         Left            =   1440
         TabIndex        =   71
         Top             =   3660
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   6
         Left            =   1440
         TabIndex        =   70
         Top             =   3960
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   7
         Left            =   1440
         TabIndex        =   69
         Top             =   4560
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   8
         Left            =   1860
         TabIndex        =   68
         Top             =   1860
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   9
         Left            =   2280
         TabIndex        =   67
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   10
         Left            =   2250
         TabIndex        =   66
         Top             =   3660
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   11
         Left            =   2280
         TabIndex        =   65
         Top             =   3960
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   12
         Left            =   2640
         TabIndex        =   64
         Top             =   1260
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   13
         Left            =   2640
         TabIndex        =   63
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   14
         Left            =   2640
         TabIndex        =   62
         Top             =   3660
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   15
         Left            =   2640
         TabIndex        =   61
         Top             =   3960
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   16
         Left            =   2640
         TabIndex        =   60
         Top             =   4560
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   17
         Left            =   3060
         TabIndex        =   59
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   18
         Left            =   3060
         TabIndex        =   58
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   19
         Left            =   3060
         TabIndex        =   57
         Top             =   2460
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   20
         Left            =   3060
         TabIndex        =   56
         Top             =   2760
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   21
         Left            =   3060
         TabIndex        =   55
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   22
         Left            =   3060
         TabIndex        =   54
         Top             =   3660
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   23
         Left            =   3060
         TabIndex        =   53
         Top             =   3960
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   24
         Left            =   3060
         TabIndex        =   52
         Top             =   4260
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   25
         Left            =   3900
         TabIndex        =   51
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   26
         Left            =   3900
         TabIndex        =   50
         Top             =   3660
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   27
         Left            =   3900
         TabIndex        =   49
         Top             =   3960
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   28
         Left            =   4260
         TabIndex        =   48
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   29
         Left            =   4260
         TabIndex        =   47
         Top             =   3660
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   30
         Left            =   4260
         TabIndex        =   46
         Top             =   3960
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   31
         Left            =   4260
         TabIndex        =   45
         Top             =   4560
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   32
         Left            =   4680
         TabIndex        =   44
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   33
         Left            =   5460
         TabIndex        =   43
         Top             =   1260
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   34
         Left            =   6300
         TabIndex        =   42
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   35
         Left            =   6300
         TabIndex        =   41
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   36
         Left            =   6300
         TabIndex        =   40
         Top             =   4260
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   37
         Left            =   6720
         TabIndex        =   39
         Top             =   1260
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   38
         Left            =   6660
         TabIndex        =   38
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   39
         Left            =   6660
         TabIndex        =   37
         Top             =   3660
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   40
         Left            =   6660
         TabIndex        =   36
         Top             =   3960
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   41
         Left            =   7080
         TabIndex        =   35
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   42
         Left            =   7080
         TabIndex        =   34
         Top             =   3660
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   43
         Left            =   7080
         TabIndex        =   33
         Top             =   3960
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   44
         Left            =   7500
         TabIndex        =   32
         Top             =   1260
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   45
         Left            =   8340
         TabIndex        =   31
         Top             =   1260
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   46
         Left            =   8700
         TabIndex        =   30
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   47
         Left            =   9120
         TabIndex        =   29
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   48
         Left            =   9120
         TabIndex        =   28
         Top             =   3660
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   49
         Left            =   9120
         TabIndex        =   27
         Top             =   3960
         Width           =   255
      End
      Begin VB.CheckBox chkCMSstatus 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   50
         Left            =   9120
         TabIndex        =   26
         Top             =   4560
         Width           =   255
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      Index           =   1
      X1              =   120
      X2              =   9840
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   9960
      Y1              =   10
      Y2              =   10
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewPacketReceive 
         Caption         =   "&CMS Raw Packet Rx"
      End
      Begin VB.Menu mnuViewDecodedPacket 
         Caption         =   "&CMS Decoded Packet"
      End
      Begin VB.Menu mnuViewSCADARawPacket 
         Caption         =   "SCADA Raw Packet Rx"
      End
      Begin VB.Menu mnuViewba1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewShowMessage 
         Caption         =   "&Show Message"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "&Search For Help On..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private exclApp As Excel.Application
Private exclBook As Excel.Workbook
Private exclSheet As Excel.Worksheet
Private exclOpen As Boolean

Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private p_NoAlarmRead As Integer
Private p_NoSCADAAlarmRead As Integer
Private p_FlagUpdateExcel As Boolean

Private Sub WhenCriticalError()
  
  Screen.MousePointer = vbDefault
  Me.cmdClear.Enabled = False
  Me.cmdConnectCMS.Enabled = False
  Me.cmdConnectOPC.Enabled = False
  Me.cmdShScada.Enabled = False
  Me.cmdShStatus.Enabled = False
  Me.mnuView.Enabled = False
  
  If Me.sckCMSAlarm.State <> sckClosed Then Me.sckCMSAlarm.Close
  If Me.sckTagRx.State <> sckClosed Then Me.sckTagRx.Close
  Close_ExcelFile
  Close_DDEOPC
  
End Sub
Private Sub Proc_AlarmRx()
'process alarm receive from CMS
'g_DefAlarm(MaxSite * MaxSubSys, 4)  'Default alarm read from .ini file
                                     '1 site, 2 subsys, 3 status, 4 excel row
'g_CMSAlarm(MaxSite * MaxSubSys, 3)  'alarm read from CMS Server
'g_NoSite         'no of site from .ini file
'g_NoSubSys       'no of subsys from .ini file
'p_NoAlarmRead    'actual value from CMS
'g_No_CMSAlm      'max define value form CMS
Dim i, j As Integer

  If p_NoAlarmRead > 0 Then
    For i = 1 To p_NoAlarmRead
      For j = 1 To g_No_CMSAlm
        If g_CMSAlarm(i, 1) = g_DefAlarm(j, 1) Then
          If g_CMSAlarm(i, 2) = g_DefAlarm(j, 2) Then
            g_DefAlarm(j, 3) = g_CMSAlarm(i, 3)
            Exit For
          End If
        End If
      Next j
    Next i
  End If
  
End Sub
Private Sub Close_DDEOPC()
'Close DDE OPC Server if opened
Dim winHwnd As Long
Dim RetVal As Long

  'stop calling DDE from the program instead of calling it from system
  '28-12-2002
  Exit Sub
  
  winHwnd = FindWindow(vbNullString, "CSMDDE.dde - ICONICS DDE OPC Server")
  If winHwnd <> 0 Then
    RetVal = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
    If RetVal = 0 Then
      DispMsg "Failed to close DDE OPC."
    End If
  End If

End Sub
  

Private Sub Close_ExcelFile()
  
  If exclOpen Then
    exclBook.Close (False)
    
    Set exclSheet = Nothing
    Set exclBook = Nothing
    Set exclSheet = Nothing
  
    Me.tmrUpdateExcel.Enabled = False
  End If
  
End Sub

Private Sub Open_ExcelFile(ByVal na$)

  Set exclApp = New Excel.Application
  Set exclBook = exclApp.Workbooks.Open(na$)
  Set exclSheet = exclBook.ActiveSheet
  DispMsg ("Opened Excel file " & na$)
  Me.tmrUpdateExcel.Enabled = True
  exclOpen = True
End Sub

Private Sub Open_DDEOPC()
Dim winHwnd As Long
Dim RetVal As Long

  Exit Sub
  
  winHwnd = FindWindow(vbNullString, "CSMDDE.dde - ICONICS DDE OPC Server")
  If winHwnd <> 0 Then
    DispMsg "DDEOPC opened"
  Else
    MsgBox ("Please start DDEOPC first and then restart the program!")
  End If


  '28-12-2001
  'Exit Sub
  'RetVal = Shell("C:\Program Files\ICONICS\DDEOPC Server\DDEOPC.exe", 1)   ' Run DDE OPC.
End Sub


Public Sub DispMsg(ByVal s$)
Dim a$

  If Not Me.mnuViewShowMessage.Checked Then Exit Sub
  
  If lstMsg.ListCount > 4096 Then lstMsg.Clear
  
  a$ = Format(Now, "YYYYMMDD hh:mm:ss")
  Call Me.lstMsg.AddItem(a$ & " -- " & s$)
  
  'select last index
  If Me.lstMsg.ListIndex < 0 Then
    Me.lstMsg.AddItem (a$ & " -- " & "")
    Me.lstMsg.ListIndex = 0
  Else
    Me.lstMsg.ListIndex = Me.lstMsg.ListCount - 1
  End If

End Sub

Public Sub UpdateAlarmFromSCADA(ByVal ss$, ByVal count As Integer)
'update the SCADA alarm array
'for local display only
Dim a$, b$
Dim i As Integer
Dim j As Integer

  a$ = Mid$(ss$, 24, Len(ss$) - 27)     'extract ^STATUS YYYYMMDD hhmmss and ending 3 char
  
  i = 1
  Do While Len(a$) >= 7
    b$ = Mid$(a$, 1, 7)
    a$ = LTrim$(Mid(a$, 8))
    g_SCADAAlarm(i, 1) = Mid$(b$, 1, 3)
    g_SCADAAlarm(i, 2) = Mid$(b$, 4, 3)
    Select Case Mid$(b$, 7, 1)
    Case "U"
      g_SCADAAlarm(i, 3) = "0"
    Case "M"
      g_SCADAAlarm(i, 3) = "1"
    Case "D"
      g_SCADAAlarm(i, 3) = "1"
    Case "?"
      g_SCADAAlarm(i, 3) = "1"
    End Select
    i = i + 1
  Loop
  p_NoSCADAAlarmRead = i - 1
    
    
  DispMsg "Packet received from Alarm Server:"
  DispMsg "    No of alarm read = " & Str(p_NoSCADAAlarmRead)
  DispMsg "    Packet counter   = " & Str(count)

  
End Sub
Public Sub UpdateAlarm(ByVal ss$, ByVal count As Integer)
'update the alarm array
'call by sckCMSAlarm, after receiving packet from CMS
'
Dim a$, b$
Dim i As Integer
Dim j As Integer

  'DispMsg ss$
  a$ = Mid$(ss$, 24, Len(ss$) - 27)     'extract ^STATUS YYYYMMDD hhmmss and ending 3 char
  
  'For i = 1 To MaxSite * MaxSubSys
  '  g_Alarm(i) = ""
  'Next
  
  i = 1
  Do While Len(a$) >= 7
    b$ = Mid$(a$, 1, 7)
    a$ = LTrim$(Mid(a$, 8))
    g_CMSAlarm(i, 1) = Mid$(b$, 1, 3)
    g_CMSAlarm(i, 2) = Mid$(b$, 4, 3)
    Select Case Mid$(b$, 7, 1)
    Case "U"
      g_CMSAlarm(i, 3) = "0"
    Case "M"
      g_CMSAlarm(i, 3) = "1"
    Case "D"
      g_CMSAlarm(i, 3) = "1"
    Case "?"
      g_CMSAlarm(i, 3) = "1"
    End Select
    i = i + 1
  Loop
  
  p_NoAlarmRead = i - 1
    
  DispMsg "Packet received from CMS Server:"
  DispMsg "    No of alarm read = " & Str(p_NoAlarmRead)
  DispMsg "    Packet counter   = " & Str(count)
  
  If Me.mnuViewPacketReceive.Checked Then
    DispMsg "CMS Raw packet:"
    Do While Len(ss$) > 0
      DispMsg Mid(ss$, 1, 80)
      ss$ = Mid(ss$, 81)
    Loop
  End If

  If Me.mnuViewDecodedPacket.Checked Then
    a$ = ""
    'For j = 1 To p_NoAlarmRead
    For j = 1 To p_NoAlarmRead
      a$ = a$ & g_CMSAlarm(j, 1) & g_CMSAlarm(j, 2) & g_CMSAlarm(j, 3) & " "
    Next
    DispMsg "CMS Packet decoded: "
    Do While Len(a$) > 0
      DispMsg Mid(a$, 1, 80)
      a$ = Mid(a$, 81)
    Loop
  End If
      
  p_FlagUpdateExcel = True    'write a flag to indicate need to update Excel file
  
End Sub

Private Sub cmdClear_Click()
  
  Me.lstMsg.Clear

End Sub

Private Sub cmdConnectCMS_Click()
    
  Select Case Me.cmdConnectCMS.Caption
  Case "Connect CMS"
    Me.sckCMSAlarm.Listen
    DispMsg ("Enable CMS client button click")
    Me.cmdConnectCMS.Caption = "Disconnect CMS"
  Case "Disconnect CMS"
    Me.sckCMSAlarm.Close
    DispMsg ("Disable CMS client button click")
    Me.cmdConnectCMS.Caption = "Connect CMS"
  End Select

End Sub

Private Sub cmdConnectOPC_Click()
    
  Select Case Me.cmdConnectOPC.Caption
  Case "Connect OPC"
    Me.sckTagRx.Listen
    DispMsg ("Connect OPC client button click")
    Me.cmdConnectOPC.Caption = "Disconnect OPC"
  Case "Disconnect OPC"
    Me.sckTagRx.Close
    DispMsg ("Disconnect OPC client button click")
    Me.cmdConnectOPC.Caption = "Connect OPC"
  End Select
  
End Sub

Private Sub cmdShScada_Click()
  
  If Me.cmdShScada.Caption = "Show SCADA Status" Then
    Me.cmdShScada.Caption = "Hide SCADA Status"
    Me.cmdShStatus.Caption = "Show CMS Status"
    Me.picSCADA.Visible = True
    Me.picSCADA.ZOrder 0
  Else
    Me.cmdShScada.Caption = "Show SCADA Status"
    Me.picSCADA.Visible = False
    Me.lstMsg.ZOrder 0
  End If
End Sub

Private Sub cmdShStatus_Click()
  
  If Me.cmdShStatus.Caption = "Show CMS Status" Then
    Me.cmdShStatus.Caption = "Hide CMS Status"
    Me.cmdShScada.Caption = "Show SCADA Status"
    Me.picCMSstatus.Visible = True
    Me.picCMSstatus.ZOrder 0
  Else
    Me.cmdShStatus.Caption = "Show CMS Status"
    Me.picCMSstatus.Visible = False
    Me.lstMsg.ZOrder 0
  End If
    
  
End Sub

Private Sub Form_Initialize()

  ReadFromFile Me
  
  If Not g_iniFileError Then
    
    'init command tooltiptext
    Me.cmdClear.ToolTipText = "Clear message screen"
    Me.cmdConnectCMS.ToolTipText = "Connect or Disconnect CMS"
    Me.cmdConnectOPC.ToolTipText = "Connect or Disconnect OPC"
    
    
    'init public variable
    p_FlagUpdateExcel = False
    
    'init TCP winsock for CMS communication
    Me.sckCMSAlarm.Protocol = sckTCPProtocol
    Me.sckCMSAlarm.LocalPort = g_COMM_CMSLocalPortNo
    Me.sbStatusBar.Panels.Item(2).Text = "Local Port: " & Str(g_COMM_CMSLocalPortNo)
    
    'init TCP winsock for Alarm Server communication
    Me.sckTagRx.Protocol = sckTCPProtocol
    Me.sckTagRx.LocalPort = g_COMM_OPCLocalPortNo
    
    'display information
    Me.Label2.Item(0).Caption = Me.sckCMSAlarm.LocalHostName
    
    'CMS socket information
    Me.Label2.Item(1).Caption = g_CMSIP
    Me.Label2.Item(2).Caption = Me.sckCMSAlarm.RemotePort
    Me.Label2.Item(3).Caption = Me.sckCMSAlarm.LocalPort
    Me.Label2.Item(4).Caption = Me.sckCMSAlarm.State
    
    'OPC socket information
    Me.Label2.Item(5).Caption = g_OPCIP
    Me.Label2.Item(6).Caption = Me.sckTagRx.RemotePort
    Me.Label2.Item(7).Caption = Me.sckTagRx.LocalPort
    Me.Label2.Item(8).Caption = Me.sckTagRx.State
    
    'hide status screen
    Me.picCMSstatus.Visible = False
    Me.picSCADA.Visible = False
  
    Screen.MousePointer = vbHourglass
    DoEvents
    
    'init Excel file
    Open_ExcelFile (App.Path & "\CMSExcel.xls")
    
    'init DDE OPC
    'Open_DDEOPC
    
    'enable event timer
    tmrTCPStatus.Enabled = True
  
    'enable heartbeat to OPC
    Me.tmrHBOPC.Enabled = True
    
    'set CMS communication failure timer
    Me.tmrCOMMFailed.Interval = 65535
    Me.tmrCOMMFailed.Enabled = False
    
    'menu setup
    Me.mnuViewDecodedPacket.Checked = False
    Me.mnuViewPacketReceive.Checked = False
    Me.mnuViewSCADARawPacket.Checked = False
    
    Screen.MousePointer = vbDefault
    
    'simulate Connect CMS & OPC Button click
    Call cmdConnectCMS_Click
    Call cmdConnectOPC_Click
  
  Else
    WhenCriticalError
  End If
  
End Sub

Private Sub Form_Load()

  'resume screen size and title
  Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
  Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
  Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
  Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
  Me.Caption = "CMS & SCADA Communicator"
  exclOpen = False
  
  
End Sub

Private Sub Form_Resize()

  On Error GoTo SizeResume
  
  If Me.WindowState <> vbMinimized Then
    Me.Line1(1).X1 = 0
    Me.Line1(1).X2 = Me.ScaleWidth
    Me.Line1(1).Y1 = 0
    Me.Line1(1).Y2 = Me.Line1(1).Y1
    
    Me.Line1(0).X1 = 0
    Me.Line1(0).X2 = Me.ScaleWidth
    Me.Line1(0).Y1 = Me.Line1(1).Y1 + 20
    Me.Line1(0).Y2 = Me.Line1(0).Y1
    
    Me.lstMsg.Top = Me.Line1(1).Y2 + 100
    Me.lstMsg.Left = 0
    Me.lstMsg.Width = Me.ScaleWidth
    'Me.lstMsg.Height = Me.ScaleHeight - Me.sbStatusBar.Height - _
      Me.picButton.ScaleHeight - Me.picWinsock.ScaleHeight - 200
    Me.lstMsg.Height = Me.ScaleHeight - Me.sbStatusBar.Height - _
      Me.picButton.ScaleHeight - Me.picWinsock.ScaleHeight - 200
      
    Me.picCMSstatus.Top = Me.lstMsg.Top
    Me.picCMSstatus.Left = 0
    Me.picCMSstatus.Width = Me.ScaleWidth
    Me.picCMSstatus.Height = Me.ScaleHeight - Me.picButton.Height - _
      Me.picWinsock.Height - Me.sbStatusBar.Height - 200
  
  
    Me.picSCADA.Top = Me.lstMsg.Top
    Me.picSCADA.Left = 0
    Me.picSCADA.Width = Me.ScaleWidth
    Me.picSCADA.Height = Me.ScaleHeight - Me.picButton.Height - _
      Me.picWinsock.Height - Me.sbStatusBar.Height - 200
  End If
  Exit Sub
  
SizeResume:
  Me.Width = 10725
  Me.Height = 7620
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer

  If Me.sckCMSAlarm.State <> sckClosed Then
    Me.sckCMSAlarm.Close
  End If
  
  Close_ExcelFile
  
  Close_DDEOPC
  
  'close all sub forms
  For i = Forms.count - 1 To 1 Step -1
    Unload Forms(i)
  Next
  If Me.WindowState <> vbMinimized Then
    SaveSetting App.Title, "Settings", "MainLeft", Me.Left
    SaveSetting App.Title, "Settings", "MainTop", Me.Top
    SaveSetting App.Title, "Settings", "MainWidth", Me.Width
    SaveSetting App.Title, "Settings", "MainHeight", Me.Height
  End If
End Sub

Private Sub mnuViewDecodedPacket_Click()

  mnuViewDecodedPacket.Checked = Not mnuViewDecodedPacket.Checked
  
End Sub

Private Sub mnuViewPacketReceive_Click()

  mnuViewPacketReceive.Checked = Not mnuViewPacketReceive.Checked
  
End Sub

Private Sub mnuViewSCADARawPacket_Click()
  Me.mnuViewSCADARawPacket.Checked = Not Me.mnuViewSCADARawPacket.Checked
  
End Sub

Private Sub mnuViewShowMessage_Click()

  If Me.mnuViewShowMessage.Checked Then
    Me.DispMsg "Message display was disabled"
    DoEvents
    Me.mnuViewShowMessage.Checked = False
  Else
    Me.mnuViewShowMessage.Checked = True
    Me.DispMsg "Message display enabled"
  End If
  
End Sub


Private Sub sckCMSAlarm_Close()
  
  If Me.sckCMSAlarm.State <> sckClosed Then
    Me.sckCMSAlarm.Close
  End If
  Me.tmrReListenCMS.Enabled = True

End Sub

Private Sub sckCMSAlarm_ConnectionRequest(ByVal requestID As Long)
  
  If Me.sckCMSAlarm.State <> sckClosed Then
    Me.sckCMSAlarm.Close
  End If
  
  Me.sckCMSAlarm.Accept requestID
  DispMsg ("Connect ID is " & Str(requestID))

End Sub

Private Sub sckCMSAlarm_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
Static i
Dim s As String

  On Error GoTo ErrsckCMSAlarm
  
  
  'packet counter
  i = i + 1
  If i > 4096 Then i = 1
  
  'allocate a large enough string buffer and get the data
  strData = String(bytesTotal + 2, Chr$(0))
  
  'get the packet in buffer
  Me.sckCMSAlarm.GetData strData, vbString, bytesTotal
  
  s = Mid$(strData, 1, 7)
  If s <> "^STATUS" Then
    s = "Packet header not found"
    GoTo ReceiveErr
  End If
  
  s = Mid$(strData, Len(strData) - 2, 1)
  If s <> "^" Then
    s = "Packet ending not found"
    GoTo ReceiveErr
  End If
  
  
  UpdateAlarm strData, i
  
  Exit Sub
  
ReceiveErr:
  DispMsg "Data receive error, " & s & ", packet ignored!!"
  Exit Sub
  
ErrsckCMSAlarm:
  DispMsg "Socket sckCMSAlarm error, program stop"
  WhenCriticalError
  
  
End Sub

Private Sub sckCMSAlarm_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'socket error, print the error msg and close the control
  
  DispMsg Description
  
  Call cmdConnectCMS_Click
  
  'Call tbToolBar_ButtonClick(Me.tbToolBar.Buttons.Item(4))

End Sub


Private Sub sckTagRx_Close()

  If Me.sckTagRx.State <> sckClosed Then
    Me.sckTagRx.Close
  End If
  Me.tmrReListenOPC.Enabled = True

End Sub

Private Sub sckTagRx_ConnectionRequest(ByVal requestID As Long)
 
  If Me.sckTagRx.State <> sckClosed Then
    Me.sckTagRx.Close
  End If
  
  Me.sckTagRx.Accept requestID
  DispMsg ("OPC Connect ID is " & Str(requestID))

End Sub

Private Sub sckTagRx_DataArrival(ByVal bytesTotal As Long)
'receive tag info from OPC Server
Dim strData As String
Dim s As String
Static i
  
  On Error GoTo ErrsckTagRx
  
  'packet counter
  i = i + 1
  If i > 4096 Then i = 1
  
  'allocate a large enough string buffer and get the data
  strData = String(bytesTotal + 2, Chr$(0))
  
  'get the packet in buffer
  Me.sckTagRx.GetData strData, vbString, bytesTotal
  
  'display data receive
  If Me.mnuViewSCADARawPacket.Checked Then
    Me.DispMsg "Tags from SCADA Alarm Server: "
    s = strData
    Do While Len(s) > 0
      DispMsg Mid(s, 1, 80)
      s = Mid(s, 81)
    Loop
  Else
    Me.DispMsg "Tags from SCADA Alarm Server received"
  End If
    
  
  'error check
  s = Mid$(strData, 1, 7)
  If s <> "^STATUS" Then
    s = "Packet header not found"
    GoTo ReceiveErr
  End If
  
  s = Mid$(strData, Len(strData) - 2, 1)
  If s <> "^" Then
    s = "Packet ending not found"
    GoTo ReceiveErr
  End If
  
  UpdateAlarmFromSCADA strData, i
  
  'send to CMS
  If Me.sckCMSAlarm.State = sckConnected Then
    Me.sckCMSAlarm.SendData strData
    Me.DispMsg "Tags send to CMS Server."
  End If
  
  Exit Sub
  
ReceiveErr:
  DispMsg "SCADA Data receive error, " & s & ", packet ignored!!"
  Exit Sub
  
ErrsckTagRx:
  DispMsg "Socket sckTagRx error, program stop"
  WhenCriticalError

End Sub

Private Sub sckTagRx_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'socket error, print the error msg and close the control
  DispMsg Description
  
  
End Sub

Private Sub mnuHelpAbout_Click()
  frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
  Dim nRet As Integer

  'if there is no helpfile for this project display a message to the user
  'you can set the HelpFile for your application in the
  'Project Properties dialog
  If Len(App.HelpFile) = 0 Then
    DispMsg ("Unable to display Help Contents. There is no Help associated with this project.")
  Else
    On Error Resume Next
    nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
    If Err Then
      DispMsg Err.Description
    End If
  End If

End Sub

Private Sub mnuHelpContents_Click()
  Dim nRet As Integer

  'if there is no helpfile for this project display a message to the user
  'you can set the HelpFile for your application in the
  'Project Properties dialog
  If Len(App.HelpFile) = 0 Then
    DispMsg "Unable to display Help Contents. There is no Help associated with this project."
  Else
    On Error Resume Next
    nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
    If Err Then
      DispMsg Err.Description
    End If
  End If

End Sub




Private Sub mnuViewOptions_Click()
  frmOptions.Show vbModal, Me
End Sub



Private Sub mnuViewStatusBar_Click()
  mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
  sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
  mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
  
End Sub






Private Sub mnuFileExit_Click()
  'unload the form
  Unload Me

End Sub

Private Sub tmrCOMMFailed_Timer()
'communicatione to CMS false
'write a flag to Excel file
    
  If exclOpen Then
    If exclSheet.Cells(1, 1) <> "1" Then
      exclSheet.Cells(1, 1) = "1"
      DispMsg ("Communication to CMS failed!!")
    End If
  End If

End Sub

Private Sub tmrHBOPC_Timer()
'60 sec heart beat to OPC

  If Me.sckTagRx.State = sckConnected Then
    DispMsg ("Heartbeat sent to OPC")
    Me.sckTagRx.SendData "??OPC"
  End If
  
End Sub



Private Sub tmrReListenCMS_Timer()
'Re-listening the client request after disconnect

  If Me.cmdConnectCMS.Caption = "Connect CMS" Then
    If Me.sckCMSAlarm.State <> sckClosed Then
      Me.sckCMSAlarm.Close
    End If
    If exclOpen Then
      exclSheet.Cells(1, 1) = "1"     'set COMM failed
    End If

    Me.sckCMSAlarm.Listen
    DispMsg ("Re-listening to client request")
    Me.tmrReListenCMS.Enabled = False
    
    Me.cmdConnectCMS.Caption = "Disconnect CMS"
  End If
  
End Sub

Private Sub tmrReListenOPC_Timer()
'Re-listening the client request after disconnect

  'If Me.tbToolBar.Buttons.Item(4).Enabled Then
  'disconnect button enable
    If Me.sckTagRx.State <> sckClosed Then
      Me.sckTagRx.Close
    End If
    Me.sckTagRx.Listen
    DispMsg ("Re-listening to OPC client request")
    Me.tmrReListenOPC.Enabled = False
  'End If
  
End Sub

Private Sub tmrTCPStatus_Timer()
'update Status bar
Dim s As String
Static lastStatusCMS As Integer
Static lastStatusOPC As Integer
  
  'time update
  Me.sbStatusBar.Panels.Item(3).Text = Format(Now, "YYYY/MM/DD HH:MM:SS")
  
  s = "TCP Status: "
  'CMS TCP Status
  If lastStatusCMS = Me.sckCMSAlarm.State Then GoTo OPC_Status
  lastStatusCMS = Me.sckCMSAlarm.State
  Select Case Me.sckCMSAlarm.State
  Case sckClosed
    Me.sbStatusBar.Panels.Item(1).Text = s & "Socket close"
    DispMsg "CMS TCP Socket close"
    Me.Label2.Item(4).Caption = "Close"
    Me.cmdConnectCMS.Caption = "Connect CMS"
  Case sckOpen
    Me.sbStatusBar.Panels.Item(1).Text = s & "Socket open"
    Me.Label2.Item(4).Caption = "Open"
    DispMsg "CMS TCP Socket open"
  Case sckListening
    Me.sbStatusBar.Panels.Item(1).Text = s & "Socket listening"
    Me.Label2.Item(4).Caption = "Listening"
    DispMsg "CMS TCP Socket listening"
  Case sckConnectionPending
    Me.sbStatusBar.Panels.Item(1).Text = s & "Socket pending"
    Me.Label2.Item(4).Caption = "Pending"
    DispMsg "CMS TCP Socket pending"
  Case sckResolvingHost
    Me.sbStatusBar.Panels.Item(1).Text = s & "Socket resolving"
    Me.Label2.Item(4).Caption = "Resolving"
    DispMsg "CMS TCP Socket resolving"
  Case sckHostResolved
    Me.sbStatusBar.Panels.Item(1).Text = s & "Socket resolved"
    Me.Label2.Item(4).Caption = "Resolved"
    DispMsg "CMS TCP Socket resolved"
  Case sckConnecting
    Me.sbStatusBar.Panels.Item(1).Text = s & "Socket connecting"
    Me.Label2.Item(4).Caption = "Connecting"
    DispMsg "CMS TCP Socket connecting"
  Case sckConnected
    Me.sbStatusBar.Panels.Item(1).Text = s & "Socket connected"
    Me.Label2.Item(2).Caption = Me.sckCMSAlarm.RemotePort
    Me.Label2.Item(1).Caption = Me.sckCMSAlarm.RemoteHostIP
    Me.Label2.Item(4).Caption = "Connected"
    DispMsg "CMS TCP Socket connected"
  Case sckClosing
    Me.sbStatusBar.Panels.Item(1).Text = s & "Socket closing"
    Me.Label2.Item(4).Caption = "Closing"
    DispMsg "CMS TCP Socket closing"
  Case sckError
    Me.sbStatusBar.Panels.Item(1).Text = s & "Socket error"
    Me.Label2.Item(4).Caption = "Error"
    DispMsg "CMS TCP Socket error"
  Case Else
    Me.Label2.Item(4).Caption = "Unknown"
    Me.sbStatusBar.Panels.Item(1).Text = s & "CMS Staus unknown: " & Me.sckCMSAlarm.State
  End Select
    
    
OPC_Status:
  'OPC TCP Status
  If lastStatusOPC = Me.sckTagRx.State Then Exit Sub
  lastStatusOPC = Me.sckTagRx.State
  Select Case Me.sckTagRx.State
  Case sckClosed
    Me.Label2.Item(8).Caption = "Close"
    DispMsg "OPC TCP Socket close"
  Case sckOpen
    Me.Label2.Item(8).Caption = "Open"
    DispMsg "OPC TCP Socket open"
  Case sckListening
    Me.Label2.Item(8).Caption = "Listening"
    DispMsg "OPC TCP Socket listening"
  Case sckConnectionPending
    Me.Label2.Item(8).Caption = "Pending"
    DispMsg "OPC TCP Socket pending"
  Case sckResolvingHost
    Me.Label2.Item(8).Caption = "Resolving"
    DispMsg "OPC TCP Socket resolving"
  Case sckHostResolved
    Me.Label2.Item(8).Caption = "Resolved"
    DispMsg "OPC TCP Socket resolved"
  Case sckConnecting
    Me.Label2.Item(8).Caption = "Connecting"
    DispMsg "OPC TCP Socket connecting"
  Case sckConnected
    Me.Label2.Item(6).Caption = Me.sckTagRx.RemotePort
    Me.Label2.Item(5).Caption = Me.sckTagRx.RemoteHostIP
    Me.Label2.Item(8).Caption = "Connected"
    DispMsg "OPC TCP Socket connected"
    Me.tmrReListenOPC.Enabled = False
  Case sckClosing
    Me.Label2.Item(8).Caption = "Closing"
    DispMsg "OPC TCP Socket closing"
  Case sckError
    Me.Label2.Item(8).Caption = "Error"
    DispMsg "OPC TCP Socket error"
  Case Else
    Me.Label2.Item(8).Caption = "Unknown"
    DispMsg "OPC TCP Socket status unknown"
  End Select
    
End Sub



Private Sub tmrUpdateExcel_Timer()
'update excel file
Static a
Dim i As Integer
  
  If (p_FlagUpdateExcel <> True) Then
    'communication to CMS false
    'write a flag to Excel file
    If Me.tmrCOMMFailed.Enabled = False Then
      Me.tmrCOMMFailed.Enabled = True
    End If
    Exit Sub
  End If
  
  'communication to CMS OK
  p_FlagUpdateExcel = False
  Me.tmrCOMMFailed.Enabled = False
  DispMsg ("Update Excel file started")
  Proc_AlarmRx
  
  If exclOpen Then
    exclSheet.Cells(1, 1) = "0" 'reset the communication alarm
    For i = 1 To g_No_CMSAlm
      exclSheet.Cells(g_DefAlarm(i, 4) + 1, 1) = g_DefAlarm(i, 3)
    Next
  End If
  DispMsg ("Update Excel file ended")

End Sub

Private Sub tmrUpdateSCADA_Timer()
Dim i As Integer

  Exit Sub
  
  For i = 1 To p_NoSCADAAlarmRead
    Select Case g_SCADAAlarm(i, 3)
      Case "1"
        Me.chkSCADAstatus.Item(i - 1).Value = Checked
      Case Else
        Me.chkSCADAstatus.Item(i - 1).Value = Unchecked
    End Select
  Next
    
End Sub

Private Sub tmrUpdateStatus_Timer()
Dim i As Integer

  For i = 1 To g_No_CMSAlm
    Select Case g_DefAlarm(i, 3)
      Case "1"
        Me.chkCMSstatus.Item(i - 1).Value = Checked
      Case Else
        Me.chkCMSstatus.Item(i - 1).Value = Unchecked
    End Select
  Next
  
  
End Sub
