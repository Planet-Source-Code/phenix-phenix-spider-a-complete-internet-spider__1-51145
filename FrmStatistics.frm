VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmStatistics 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Phenix Spider :: Statistics"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdToggle 
      Caption         =   "Start Spider"
      Height          =   375
      Left            =   5280
      TabIndex        =   16
      Top             =   840
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   1920
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   0
      Picture         =   "FrmStatistics.frx":0000
      ScaleHeight     =   2355
      ScaleWidth      =   2355
      TabIndex        =   2
      Top             =   0
      Width           =   2415
   End
   Begin MSComctlLib.ListView LstURL 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   2640
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   3836
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "URL"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Scanned"
         Object.Width           =   1588
      EndProperty
   End
   Begin MSComctlLib.ListView LstEmail 
      Height          =   2175
      Left            =   3360
      TabIndex        =   1
      Top             =   2640
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   3836
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   5468
      EndProperty
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Current Page: "
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblPage 
      Height          =   255
      Left            =   3840
      TabIndex        =   13
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Email's Found: "
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblEmail 
      Caption         =   "0"
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "URL's Found: "
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lblUrl 
      Caption         =   "0"
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Pages scanned: "
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblScanned 
      Caption         =   "0"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   840
      Width           =   2655
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   2520
      X2              =   6600
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   2535
      X2              =   6615
      Y1              =   2295
      Y2              =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Email Address:"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "URL's:"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Phenix Spider"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   0
      Width           =   2775
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   2535
      X2              =   6615
      Y1              =   735
      Y2              =   735
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2520
      X2              =   6600
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label2 
      Caption         =   "Version 1.0"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "FrmStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdToggle_Click()
    LogURL FrmMain.TxtUrl
    Toggle
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    FrmStatistics.Visible = False
    FrmMain.Visible = True
End Sub
