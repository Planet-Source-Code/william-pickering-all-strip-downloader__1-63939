VERSION 5.00
Begin VB.Form frm_comicschooser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comic Chooser"
   ClientHeight    =   5655
   ClientLeft      =   6165
   ClientTop       =   5280
   ClientWidth     =   5295
   Icon            =   "frm_comicschooser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   5295
   Begin VB.TextBox txt_save2 
      Height          =   285
      Left            =   960
      TabIndex        =   26
      Text            =   "D:\William\Comics\"
      Top             =   4680
      Width           =   3855
   End
   Begin VB.CommandButton cmd_browse 
      Caption         =   "..."
      Height          =   255
      Left            =   4920
      TabIndex        =   25
      Top             =   4680
      Width           =   255
   End
   Begin VB.TextBox txt_filepath 
      Height          =   285
      Index           =   9
      Left            =   2280
      TabIndex        =   24
      Text            =   "For Better or For Worse"
      Top             =   3720
      Width           =   2895
   End
   Begin VB.TextBox txt_filepath 
      Height          =   285
      Index           =   8
      Left            =   2280
      TabIndex        =   23
      Text            =   "Speed Bump"
      Top             =   3360
      Width           =   2895
   End
   Begin VB.TextBox txt_filepath 
      Height          =   285
      Index           =   7
      Left            =   2280
      TabIndex        =   22
      Text            =   "Non Sequitur"
      Top             =   3000
      Width           =   2895
   End
   Begin VB.TextBox txt_filepath 
      Height          =   285
      Index           =   6
      Left            =   2280
      TabIndex        =   21
      Text            =   "Doonesbury"
      Top             =   2640
      Width           =   2895
   End
   Begin VB.TextBox txt_filepath 
      Height          =   285
      Index           =   5
      Left            =   2280
      TabIndex        =   20
      Text            =   "Andy Capp"
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox txt_filepath 
      Height          =   285
      Index           =   4
      Left            =   2280
      TabIndex        =   19
      Text            =   "Wizard of ID"
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox txt_filepath 
      Height          =   285
      Index           =   3
      Left            =   2280
      TabIndex        =   18
      Text            =   "Ginger Meggs"
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox txt_filepath 
      Height          =   285
      Index           =   2
      Left            =   2280
      TabIndex        =   17
      Text            =   "Calvin & Hobbes"
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox txt_filepath 
      Height          =   285
      Index           =   1
      Left            =   2280
      TabIndex        =   16
      Text            =   "Overboard"
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox txt_filepath 
      Height          =   285
      Index           =   0
      Left            =   2280
      TabIndex        =   15
      Text            =   "Garfield"
      Top             =   480
      Width           =   2895
   End
   Begin VB.CheckBox chk_Month 
      Caption         =   "Arrange comics in month folders"
      Height          =   375
      Left            =   2760
      TabIndex        =   14
      Top             =   4200
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox chk_Year 
      Caption         =   "Arrange comics in year folders"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4200
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.OptionButton cpt_comic 
      Caption         =   "For Better or For Worse"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   2055
   End
   Begin VB.OptionButton cpt_comic 
      Caption         =   "Speed Bump"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   2055
   End
   Begin VB.OptionButton cpt_comic 
      Caption         =   "Non Sequitur"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   2055
   End
   Begin VB.OptionButton cpt_comic 
      Caption         =   "Doonesbury"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   2055
   End
   Begin VB.OptionButton cpt_comic 
      Caption         =   "Andy Capp"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton cmd_OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   2415
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   5160
      Width           =   2415
   End
   Begin VB.OptionButton cpt_comic 
      Caption         =   "Garfield"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.OptionButton cpt_comic 
      Caption         =   "Overboard"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.OptionButton cpt_comic 
      Caption         =   "Calvin and Hobbs"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.OptionButton cpt_comic 
      Caption         =   "Ginger Meggs"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
   End
   Begin VB.OptionButton cpt_comic 
      Caption         =   "Wizard of ID"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Save to:"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   4680
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5160
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label2 
      Caption         =   "Select which comic you wish to download and their save-to folders:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5055
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   3
      X1              =   120
      X2              =   5160
      Y1              =   4080
      Y2              =   4080
   End
End
Attribute VB_Name = "frm_comicschooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim x As Long
Public temp As String


Private Sub chk_Year_Click()
If chk_Year.Value = 1 Then
chk_Month.Enabled = True
Else
chk_Month.Enabled = False
End If
End Sub

Private Sub cmd_cancel_Click()
frm_main.ComiC = temp
Unload Me
End Sub

Private Sub cmd_OK_Click()

Garfieldfolder = txt_filepath(0).Text
OverBoardfolder = txt_filepath(1).Text
CalvinHobbesfolder = txt_filepath(2).Text
GingerMeggsfolder = txt_filepath(3).Text
WizardofIDfolder = txt_filepath(4).Text
AndyCappfolder = txt_filepath(5).Text
Doonesburyfolder = txt_filepath(6).Text
NonSequiturfolder = txt_filepath(7).Text
SpeedBumpfolder = txt_filepath(8).Text
ForBetterorForWorsefolder = txt_filepath(9).Text

Unload Me
End Sub

Private Sub cpt_comic_Validate(Index As Integer, Cancel As Boolean)

Select Case Index

Case 0
frm_main.ComiC = Garfield

Case 1
frm_main.ComiC = OverBoard

Case 2
frm_main.ComiC = CalvinHobbes

Case 3
frm_main.ComiC = GingerMeggs

Case 4
frm_main.ComiC = WizardofID

Case 5
frm_main.ComiC = AndyCapp

Case 6
frm_main.ComiC = Doonesbury

Case 7
frm_main.ComiC = NonSequitur

Case 8
frm_main.ComiC = SpeedBump

Case 9
frm_main.ComiC = ForBetterorForWorse

End Select

End Sub

Private Sub Form_Load()
txt_save2.Text = GetINISetting("General", "SaveTo", frm_main.settingspatH, "C:\")
End Sub

Private Sub Form_Unload(Cancel As Integer)
x = WriteINISetting("General", "chkyear", chk_Year.Value, frm_main.settingspatH)
x = WriteINISetting("General", "chkmonth", chk_Month.Value, frm_main.settingspatH)

frm_main.lbl_comic.Caption = "Currently set to download: " & trans(frm_main.ComiC)
frm_main.cmd_do.Caption = "Start downloading '" & trans(frm_main.ComiC) & "'!"

x = WriteINISetting("General", "SaveTo", txt_save2.Text, frm_main.settingspatH)

End Sub
