VERSION 5.00
Object = "{CB157D16-3572-4866-A98B-5CC16264D092}#1.0#0"; "TreeFolder.ocx"
Begin VB.Form frm_options 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save To..."
   ClientHeight    =   3735
   ClientLeft      =   4365
   ClientTop       =   3825
   ClientWidth     =   4455
   Icon            =   "frm_folderbrowse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4455
   Begin MyTreeFolder.Treefolder FolderBrowse 
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4683
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frm_folderbrowse.frx":0EDA
      LabelEdit       =   1
   End
   Begin VB.CommandButton cmd_Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton cmd_OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Select your comic folder:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frm_options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim temp As String

Private Sub cmd_cancel_Click()
frm_comicschooser.txt_save2.Text = temp
Unload Me
End Sub

Private Sub cmd_OK_Click()
Unload Me
End Sub

Private Sub FolderBrowse_PathChanged()
frm_comicschooser.txt_save2.Text = FolderBrowse.Path
End Sub

Private Sub Form_Load()
temp = frm_comicschooser.txt_save2.Text
End Sub
