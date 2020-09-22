VERSION 5.00
Begin VB.Form FrmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Phenix Spider :: Options"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSave 
      Height          =   285
      Left            =   3960
      TabIndex        =   9
      Text            =   "C:\Spider.[name].log"
      Top             =   1920
      Width           =   2655
   End
   Begin VB.CheckBox ChkSave 
      Caption         =   "Save finds to:"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.OptionButton OptNothing 
      Caption         =   "Do nothing"
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   3240
      TabIndex        =   5
      Top             =   1150
      Width           =   3255
   End
   Begin VB.OptionButton OptFind 
      Caption         =   "Find:"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   1200
      Width           =   3495
   End
   Begin VB.OptionButton OptEmail 
      Caption         =   "Log Email Addresses"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   840
      Value           =   -1  'True
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   0
      Picture         =   "FrmOptions.frx":0000
      ScaleHeight     =   2355
      ScaleWidth      =   2355
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "(c) 2004 - Dominic Black"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   2280
      Width           =   2535
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
      TabIndex        =   2
      Top             =   0
      Width           =   2775
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   2530
      X2              =   6610
      Y1              =   730
      Y2              =   730
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
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "FrmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    txtSave = App.Path & "\Spider.[name].log"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    FrmOptions.Visible = False
    FrmMain.Visible = True
End Sub

Private Sub OptEmail_Click()
    FrmStatistics.LstEmail.Visible = True
    FrmStatistics.Label4.Visible = True
    FrmStatistics.Label4.Caption = "Email Address:"
    FrmStatistics.LstEmail.ColumnHeaders(1).Text = "Address"
    FrmStatistics.LstEmail.ListItems.Clear
    FrmStatistics.Label8.Caption = "Email's Found: "
End Sub

Private Sub OptFind_Click()
    FrmStatistics.LstEmail.Visible = True
    FrmStatistics.Label4.Visible = True
    FrmStatistics.Label4.Caption = "Search Results:"
    FrmStatistics.LstEmail.ColumnHeaders(1).Text = "URL"
    FrmStatistics.LstEmail.ListItems.Clear
    FrmStatistics.Label8.Caption = "Found: "
End Sub

Private Sub OptNothing_Click()
    FrmStatistics.LstEmail.Visible = False
    FrmStatistics.Label4.Visible = False
End Sub
