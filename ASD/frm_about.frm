VERSION 5.00
Begin VB.Form frm_about 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   4575
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frm_about.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3157.746
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   780
      Left            =   240
      Picture         =   "frm_about.frx":0EDA
      ScaleHeight     =   505.68
      ScaleMode       =   0  'User
      ScaleWidth      =   505.68
      TabIndex        =   1
      Top             =   240
      Width           =   780
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   3960
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5337.57
      Y1              =   2236.305
      Y2              =   2236.305
   End
   Begin VB.Label lblDescription 
      ForeColor       =   &H00000000&
      Height          =   1290
      Left            =   1080
      TabIndex        =   2
      Top             =   1800
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "ASD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1440
      Left            =   1080
      TabIndex        =   4
      Top             =   -120
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   1080
      TabIndex        =   5
      Top             =   1440
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Warning: ..."
      ForeColor       =   &H00000000&
      Height          =   1065
      Left            =   240
      TabIndex        =   3
      Top             =   3480
      Width           =   3870
   End
End
Attribute VB_Name = "frm_about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title & "  v" & App.Major & "." & App.Minor & "." & App.Revision
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    
lblDescription.Caption = Replace(App.FileDescription, String(2, Chr(32)), vbCrLf)

lblDisclaimer.Caption = "I accept no responsibility for anything that this prog may do to your system.  Not that i expect it to do anything adverse to your system...  but anyhow.  Send any comments/critisism/suggestions to daone@nerdshack.com"

End Sub


