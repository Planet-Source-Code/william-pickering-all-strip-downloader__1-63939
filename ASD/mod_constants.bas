Attribute VB_Name = "mod_publicstuff"
'comic constants - only in module coz i need them to be public and i cant do it in a form
Public Const Garfield As String = "ga"
Public Const CalvinHobbes As String = "ch"
Public Const OverBoard As String = "ob"
Public Const GingerMeggs As String = "gin"
Public Const WizardofID As String = "crwiz"

Public Const AndyCapp As String = "crcap"
Public Const Doonesbury As String = "db"
Public Const NonSequitur As String = "nq"
Public Const SpeedBump As String = "crspe"
Public Const ForBetterorForWorse As String = "fb"

'file paths...
Public Garfieldfolder As String
Public CalvinHobbesfolder As String
Public OverBoardfolder As String
Public GingerMeggsfolder As String
Public WizardofIDfolder As String
Public AndyCappfolder As String
Public Doonesburyfolder As String
Public NonSequiturfolder As String
Public SpeedBumpfolder As String
Public ForBetterorForWorsefolder As String




Public Function trans(dString As String) As String
Select Case dString

Case Garfield
trans = "Garfield"

Case OverBoard
trans = "Overboard"

Case CalvinHobbes
trans = "Calvin & Hobbes"

Case GingerMeggs
trans = "Ginger Meggs"

Case WizardofID
trans = "Wizard of ID"

Case AndyCapp
trans = "Andy Capp"

Case Doonesbury
trans = "Doonesbury"

Case NonSequitur
trans = "Non Sequitur"

Case SpeedBump
trans = "Speed Bump"

Case ForBetterorForWorse
trans = "For Better or For Worse"

End Select
End Function

