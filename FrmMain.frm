VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Phenix Spider"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   6660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdOptions 
      Caption         =   "Options"
      Height          =   375
      Left            =   5160
      TabIndex        =   11
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton CmdView 
      Caption         =   "View Statistics"
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton CmdToggle 
      Caption         =   "Start Spider"
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtMemory 
      Height          =   285
      Left            =   4560
      TabIndex        =   6
      Text            =   "-1"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox TxtUrl 
      Height          =   285
      Left            =   3360
      TabIndex        =   4
      Text            =   "http://www.filesnetwork.com"
      Top             =   820
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   0
      Picture         =   "FrmMain.frx":0000
      ScaleHeight     =   2355
      ScaleWidth      =   2355
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6000
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblScanned 
      Caption         =   "0"
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Pages scanned: "
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   2535
      X2              =   6615
      Y1              =   1575
      Y2              =   1575
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   2520
      X2              =   6600
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "URL's to hold in memory:"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   1230
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Base URL:"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   855
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2520
      X2              =   6600
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   2530
      X2              =   6610
      Y1              =   730
      Y2              =   730
   End
   Begin VB.Label Label2 
      Caption         =   "Version 1.0"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   480
      Width           =   1695
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
      TabIndex        =   1
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'============================================='
' Phenix Spider                               '
'                                             '
' Author: Dominic 'Phenix' Black              '
' Email: phenix@sg15.com                      '
'                                             '
' Feel free to use the code in you appication '
' as long as you give credit to me.           '
'============================================='


Private Sub CmdOptions_Click()
    FrmOptions.Show
    Me.Visible = False
End Sub

Private Sub CmdToggle_Click()
    Toggle
End Sub

Private Sub CmdView_Click()
    FrmStatistics.Show
    Me.Visible = False
End Sub

Private Sub Form_Load()
        Open Replace(FrmOptions.txtSave, "[name]", "search") For Append As #1
        Open Replace(FrmOptions.txtSave, "[name]", "email") For Append As #2
        Open Replace(FrmOptions.txtSave, "[name]", "url") For Append As #3
    
    FrmStatistics.Show
    FrmStatistics.Visible = False
    FrmOptions.Show
    FrmOptions.Visible = False
    
    SpiderOnline = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("Are you sure you wish to quit", vbYesNo) = vbYes Then
        Close #1
        Close #2
        Close #3
        End
    Else
        Cancel = 1
    End If
End Sub
