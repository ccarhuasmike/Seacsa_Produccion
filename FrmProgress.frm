VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Frm_Progress 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6375
   ForeColor       =   &H00C0C0C0&
   Icon            =   "FrmProgress.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1170
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   529
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label lbl_progress 
      Caption         =   "Label1"
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
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   5775
   End
End
Attribute VB_Name = "Frm_Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Frm_Progress.Left = 2000
    Frm_Progress.Top = 3000
End Sub

