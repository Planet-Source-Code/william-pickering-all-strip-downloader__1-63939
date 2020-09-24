VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frm_main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ASD"
   ClientHeight    =   6855
   ClientLeft      =   1605
   ClientTop       =   2190
   ClientWidth     =   9255
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   9255
   Begin VB.CommandButton cmd_datechange 
      Caption         =   "Set to current date"
      Height          =   255
      Left            =   7440
      TabIndex        =   7
      Top             =   720
      Width           =   1695
   End
   Begin MSComCtl2.MonthView Calendar2 
      Height          =   2370
      Left            =   6480
      TabIndex        =   3
      Top             =   1080
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   20578306
      TitleBackColor  =   16671549
      TrailingForeColor=   -2147483647
      CurrentDate     =   38703
   End
   Begin MSComCtl2.MonthView Calendar 
      Height          =   2370
      Left            =   3600
      TabIndex        =   2
      Top             =   1080
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   20578305
      TitleBackColor  =   16671549
      TrailingForeColor=   -2147483647
      CurrentDate     =   38693
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   480
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   120
   End
   Begin VB.CommandButton cmd_do 
      Caption         =   "Start Downloading 'For Better or For Worse'!"
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   3600
      Width           =   5535
   End
   Begin VB.ListBox lst_log 
      Height          =   3180
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3375
   End
   Begin VB.Image pic_preview 
      Height          =   2685
      Left            =   120
      Picture         =   "Main.frx":0EDA
      Top             =   4080
      Width           =   9000
   End
   Begin VB.Label lbl_comic 
      Caption         =   "Currently set to download: For Better or For Worse"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   9015
   End
   Begin VB.Label Label3 
      Caption         =   "To this day:"
      Height          =   255
      Left            =   6480
      TabIndex        =   5
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Download comics from this day:"
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   720
      Width           =   2655
   End
   Begin VB.Menu mnu_options 
      Caption         =   "Options"
      Begin VB.Menu mnu_comicsfolder 
         Caption         =   "Comics Folder"
      End
      Begin VB.Menu mnu_choosecomic 
         Caption         =   "Comic Options"
      End
   End
   Begin VB.Menu mnu_help 
      Caption         =   "Help"
      Begin VB.Menu mnu_helpfile 
         Caption         =   "Contents"
      End
      Begin VB.Menu mnu_about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'******************************************************************************'
'*  Title: All Strip Downloader                                               *'
'*                                                                            *'
'*  Release Date: 04-01-06                                                    *'
'*                                                                            *'
'*  Description: Download a collection of comics from the internet.           *'
'*               Included is Garfield, Overboard, Calvin & Hobbes,            *'
'*               Ginger Meggs, Andy Capp, Wizard of ID, Doonesbury,           *'
'*               Non Sequitur, Speed Bump, For Better or For Worse.           *'
'*                                                                            *'
'*  Revision Information: First release was sometime in June/July 2005.       *'
'*                        At that time it only downloaded Garfield. I         *'
'*                        Started making other programs, based on the         *'
'*                        same code, that would download other comics.        *'
'*                        This is the compiled version that attempts to       *'
'*                        download all of those (and more) in one program.    *'
'*                                                                            *'
'*                        Since then i have made many changes and             *'
'*                        improvements to bring you what you see before       *'
'*                        you.  I hope you enjoy this release.                *'
'*                                                                            *'
'*  Future: I'm thinking of adding more comics that follow the same sort of   *'
'*          number/letter combinations.  Top of the list at the moment is     *'
'*          Dilbert but I'm having a great deal of difficulty trying to       *'
'*          figure out the pattern.                                           *'
'*                                                                            *'
'*  How YOU can help: if you would like to help then please go to the         *'
'*                    Dilbert website and see if you can figure out the       *'
'*                    combinations they use.  So far i have been able to      *'
'*                    figure out that it has something to do with what day    *'
'*                    of the week the comic was printed.                      *'
'*                                                                            *'
'*                    Also, submit any bugs you find in this software to      *'
'*                    daone[at]nerdshack[dot]com                              *'
'*                                                                            *'
'*  Credits: Many thanks to ComWiz, DaOne, Princess & datacontroller for the  *'
'*           time they put into beta testing this prog and for the help and   *'
'*           the support they provided.                                       *'
'*                                                                            *'
'*  Greetings to: Blak_Deth, Spatz_Naz, Z0ppv3 & ?.                           *'
'*                                                                            *'
'*                                                                            *'
'*                                                                            *'
'*                                                                            *'
'*                                                                            *'
'******************************************************************************'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Option Explicit

Dim MonthString As String
Dim T2dAy As String
Dim FileName As String
Dim FilePath As String
Dim HostPath As String
Dim FileOutPath As String
Dim DAY2DAY As Date

Dim x As Long

Dim temp As String

Public settingspatH As String
Const seTTsfile As String = "settings.ini"

Public ComiC As String


Dim bData() As Byte


Private Sub Calendar_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)

Select Case ComiC

Case "ga"

'make sure we're not gonna download something before Garfield started
If EndDate < "19/06/1978" Then
Cancel = True
End If

'we cant download anything after today...
If EndDate >= Date Then
Cancel = True
End If

'so we dont stuff up the values on the two calendars...
If EndDate >= Calendar2.Value Then
Calendar2.Value = EndDate + 1
End If

'now just 4 the hell of it if they hav the comic we can display it in the picturebox

temp = frm_comicschooser.txt_save2.Text & Garfieldfolder & "\"


If frm_comicschooser.chk_Year.Value = 1 Then
temp = temp & Calendar.Year & "\"

If frm_comicschooser.chk_Month.Value = 1 Then
temp = temp & MonthString & "\"
End If

End If

temp = temp & "\ga" & DateThing & ".gif"

If FileExist(temp) Then
pic_preview.Picture = LoadPicture(temp)
Me.Height = pic_preview.Height + 4950
End If

Case Else

End Select

End Sub

Private Sub calendar2_SelChange(ByVal StartDate As Date, ByVal EndDate As Date, Cancel As Boolean)

Select Case ComiC

Case "ga"

'make sure we're not gonna download something before Garfield started and we need to have a 1 day allowance...
If EndDate < "20/06/1978" Then
Cancel = True
Exit Sub
End If

'we cant download anything after today...
If EndDate > Date Then
Cancel = True
Exit Sub
End If

'so we dont stuff up the values on the two calendars...
If EndDate <= Calendar.Value Then
Calendar.Value = EndDate - 1
Exit Sub
End If

Case Else

End Select

End Sub

Private Sub cmd_browse_Click()
frm_options.Show
End Sub

Private Sub cmd_datechange_Click()
Calendar2.Value = Date
End Sub

Private Sub cmd_do_Click()

On Error GoTo EndOfSub

Do

'configure the date stuff for the filename (make sure everything is 2 digits)
T2dAy = Right(Calendar.Year, 2)

temp = Calendar.Month
If Len(temp) = 1 Then
T2dAy = T2dAy & "0" & Calendar.Month
MonthString = "0" & Calendar.Month
Else
T2dAy = T2dAy & Calendar.Month
MonthString = Calendar.Month
End If

temp = Calendar.Day
If Len(temp) = 1 Then
T2dAy = T2dAy & "0" & Calendar.Day
Else
T2dAy = T2dAy & Calendar.Day
End If

DAY2DAY = Calendar.Value

'if theres no more then dont bother to download...
If DAY2DAY > Date Then
Log "No more comics!"
MsgBox "There is no more comics to download!"
cmd_do.Enabled = True
Exit Sub
End If


'make sure the saving location is correct (this only needs to be done once but oh well...)
If Right(frm_comicschooser.txt_save2.Text, 1) <> "\" Then
frm_comicschooser.txt_save2.Text = frm_comicschooser.txt_save2.Text & "\"
End If

'disable the control so pplz cant stuff it up
cmd_do.Enabled = False


'local stuff
FileName = ComiC & T2dAy & ".gif"

Select Case ComiC

Case Garfield
FilePath = frm_comicschooser.txt_save2.Text & Garfieldfolder & "\"

Case CalvinHobbes
FilePath = frm_comicschooser.txt_save2.Text & CalvinHobbesfolder & "\"

Case OverBoard
FilePath = frm_comicschooser.txt_save2.Text & OverBoardfolder & "\"

Case GingerMeggs
FilePath = frm_comicschooser.txt_save2.Text & GingerMeggsfolder & "\"

Case WizardofID
FilePath = frm_comicschooser.txt_save2.Text & WizardofIDfolder & "\"

Case AndyCapp
FilePath = frm_comicschooser.txt_save2.Text & AndyCappfolder & "\"

Case Doonesbury
FilePath = frm_comicschooser.txt_save2.Text & Doonesburyfolder & "\"

Case NonSequitur
FilePath = frm_comicschooser.txt_save2.Text & NonSequiturfolder & "\"

Case SpeedBump
FilePath = frm_comicschooser.txt_save2.Text & SpeedBumpfolder & "\"

Case ForBetterorForWorse
FilePath = frm_comicschooser.txt_save2.Text & ForBetterorForWorsefolder & "\"

End Select

'sort out the folders stuff
If frm_comicschooser.chk_Year.Value = 1 Then
FilePath = FilePath & Calendar.Year & "\"

If frm_comicschooser.chk_Month.Value = 1 Then
FilePath = FilePath & MonthString & "\"
End If

End If

FileOutPath = FilePath & FileName

'remote stuff
HostPath = "http://images.ucomics.com/comics/" & ComiC & "/" & Calendar.Year & "/" & FileName

'if they already have it...
If FileExist(FileOutPath) = True Then
Log "You already have " & FileName

'elseways we download it
Else

CreateFolder (FilePath & "/")

'download the file
bData() = Inet.OpenURL(HostPath, icByteArray)

'write the file
Open FileOutPath For Binary Access Write As #1
Put #1, , bData()
Close #1

'if it came out a lil to small...  (this makes sure it's at least 8kb)
If FileLen(FileOutPath) < 8000 Then
Kill FileOutPath
Log "Error: " & FileName
Else
Log FileName & " Done!"

'preview it
pic_preview.Picture = LoadPicture(FileOutPath)

'adjust heights of things.  mebe do widths as well but i coudlnt b bothered at the moment.
Me.Height = pic_preview.Height + 4950

End If

End If


'scroll the calendar - just make absolute sure we're not gonna go past the calendar value
If Calendar.Value = Calendar2.Value Then
cmd_do.Enabled = True
Exit Sub
Else
Calendar.Value = Calendar.Value + 1
End If

Me.Refresh

'aah we're finally at the end...  now let's do it all again!
Loop 'Until Calendar.Value = Calendar2.Value 'Next x


cmd_do.Enabled = True

Exit Sub

EndOfSub:
cmd_do.Enabled = True
Close #1
MsgBox "An error has occurred.  Please report it to me" & vbCrLf & Err.Number & vbCrLf & Err.Source & vbCrLf & Err.Description, vbOKOnly, "Error!"

End Sub


Public Function CreateFolder(destDir As String) As Boolean
    
   Dim i As Long
   Dim prevDir As String
    
   On Error Resume Next
    
   For i = Len(destDir) To 1 Step -1
       If Mid(destDir, i, 1) = "\" Then
           prevDir = Left(destDir, i - 1)
           Exit For
       End If
   Next i
    
   If prevDir = "" Then CreateFolder = False: Exit Function
   If Not Len(Dir(prevDir & "\", vbDirectory)) > 0 Then
       If Not CreateFolder(prevDir) Then CreateFolder = False: Exit Function
   End If
    
   On Error GoTo errDirMake
   MkDir destDir
   CreateFolder = True
   Exit Function
    
errDirMake:
   CreateFolder = False

End Function




Private Sub Form_Load()


'so they can't run 2 instances at once...  may get rid of this so they can download multiple comics at once
If App.PrevInstance Then
MsgBox "Another copy of ASD is already running.  The program will now end."
End
End If



'get our settings :)
On Error Resume Next

If Right(App.Path, 1) = "\" Or Right(App.Path, 1) = "/" Then
settingspatH = App.Path & seTTsfile
Else
settingspatH = App.Path & "\" & seTTsfile
End If

If FileExist(settingspatH) = False Then
MsgBox "It would appear that this is the first time that you have run this program.  Please make sure you have the """"Save to"""" directory correct before you start downloading.", vbOKOnly
frm_comicschooser.Show
frm_options.Show
End If

frm_comicschooser.txt_save2.Text = GetINISetting("General", "SaveTo", settingspatH, "C:\")


Calendar.Day = GetINISetting("Date", "Day", settingspatH, "19")

Calendar.Month = GetINISetting("Date", "Month", settingspatH, "6")

Calendar.Year = GetINISetting("Date", "Year", settingspatH, "1978")



Calendar2.Day = GetINISetting("Date2", "Day", settingspatH, "17")

Calendar2.Month = GetINISetting("Date2", "Month", settingspatH, "12")

Calendar2.Year = GetINISetting("Date2", "Year", settingspatH, "2005")


ComiC = GetINISetting("Comic", "comic", settingspatH, "ga")



mod_publicstuff.Garfieldfolder = GetINISetting("Folders", "Garfield", settingspatH, "Garfield")
mod_publicstuff.CalvinHobbesfolder = GetINISetting("Folders", "Calvin&Hobbes", settingspatH, "Calvin & Hobbes")
mod_publicstuff.OverBoardfolder = GetINISetting("Folders", "Overboard", settingspatH, "Overboard")
mod_publicstuff.GingerMeggsfolder = GetINISetting("Folders", "GingerMeggs", settingspatH, "Ginger Meggs")
mod_publicstuff.WizardofIDfolder = GetINISetting("Folders", "WizardofID", settingspatH, "Wizard of ID")
mod_publicstuff.AndyCappfolder = GetINISetting("Folders", "AndyCapp", settingspatH, "Andy Capp")
mod_publicstuff.Doonesburyfolder = GetINISetting("Folders", "Doonesbury", settingspatH, "Doonesbury")
mod_publicstuff.NonSequiturfolder = GetINISetting("Folders", "NonSequitur", settingspatH, "Non Sequitur")
mod_publicstuff.SpeedBumpfolder = GetINISetting("Folders", "SpeedBump", settingspatH, "Speed Bump")
mod_publicstuff.ForBetterorForWorsefolder = GetINISetting("Folders", "ForBetterorForWorse", settingspatH, "For Better or For Worse")

'some weird stuff 2 work on other forms...
frm_comicschooser.chk_Year.Value = GetINISetting("General", "chkyear", frm_main.settingspatH, 1)
frm_comicschooser.chk_Month.Value = GetINISetting("General", "chkmonth", frm_main.settingspatH, 1)


frm_comicschooser.txt_filepath(0).Text = Garfieldfolder
frm_comicschooser.txt_filepath(1).Text = OverBoardfolder
frm_comicschooser.txt_filepath(2).Text = CalvinHobbesfolder
frm_comicschooser.txt_filepath(3).Text = GingerMeggsfolder
frm_comicschooser.txt_filepath(4).Text = WizardofIDfolder
frm_comicschooser.txt_filepath(5).Text = AndyCappfolder
frm_comicschooser.txt_filepath(6).Text = Doonesburyfolder
frm_comicschooser.txt_filepath(7).Text = NonSequiturfolder
frm_comicschooser.txt_filepath(8).Text = SpeedBumpfolder
frm_comicschooser.txt_filepath(9).Text = ForBetterorForWorsefolder


Select Case ComiC

Case Garfield
frm_comicschooser.cpt_comic(0).Value = True

Case OverBoard
frm_comicschooser.cpt_comic(1).Value = True

Case CalvinHobbes
frm_comicschooser.cpt_comic(2).Value = True

Case GingerMeggs
frm_comicschooser.cpt_comic(3).Value = True

Case WizardofID
frm_comicschooser.cpt_comic(4).Value = True

Case AndyCapp
frm_comicschooser.cpt_comic(5).Value = True

Case Doonesbury
frm_comicschooser.cpt_comic(6).Value = True

Case NonSequitur
frm_comicschooser.cpt_comic(7).Value = True

Case SpeedBump
frm_comicschooser.cpt_comic(8).Value = True

Case ForBetterorForWorse
frm_comicschooser.cpt_comic(9).Value = True

End Select

frm_comicschooser.temp = frm_main.ComiC


cmd_do.Caption = "Start downloading '" & trans(ComiC) & "'!"

lbl_comic.Caption = "Currently set to download: " & trans(ComiC)


'set the size of the form and the picturebox
pic_preview.Picture = LoadPicture("D:\William\Comics\Garfield\2001\04\ga010401.gif")
Me.Height = pic_preview.Height + 4950



End Sub

Private Sub Form_Unload(Cancel As Integer)

x = WriteINISetting("Date", "Day", Calendar.Day, settingspatH)

x = WriteINISetting("Date", "Month", Calendar.Month, settingspatH)

x = WriteINISetting("Date", "Year", Calendar.Year, settingspatH)



x = WriteINISetting("Date2", "Day", Calendar2.Day, settingspatH)

x = WriteINISetting("Date2", "Month", Calendar2.Month, settingspatH)

x = WriteINISetting("Date2", "Year", Calendar2.Year, settingspatH)

x = WriteINISetting("Comic", "comic", ComiC, settingspatH)




x = WriteINISetting("Folders", "Garfield", Garfieldfolder, settingspatH)
x = WriteINISetting("Folders", "Calvin&Hobbes", CalvinHobbesfolder, settingspatH)
x = WriteINISetting("Folders", "Overboard", OverBoardfolder, settingspatH)
x = WriteINISetting("Folders", "GingerMeggs", GingerMeggsfolder, settingspatH)
x = WriteINISetting("Folders", "WizardofID", WizardofIDfolder, settingspatH)
x = WriteINISetting("Folders", "AndyCapp", AndyCappfolder, settingspatH)
x = WriteINISetting("Folders", "Doonesbury", Doonesburyfolder, settingspatH)
x = WriteINISetting("Folders", "NonSequitur", NonSequiturfolder, settingspatH)
x = WriteINISetting("Folders", "SpeedBump", SpeedBumpfolder, settingspatH)
x = WriteINISetting("Folders", "ForBetterorForWorse", ForBetterorForWorsefolder, settingspatH)

End

End Sub

Private Function Log(dString As String)
If Trim(dString) = "" Then Exit Function
lst_log.AddItem dString
lst_log.ListIndex = lst_log.ListCount - 1
End Function

Private Sub mnu_about_Click()
frm_about.Show
End Sub

Private Sub mnu_choosecomic_Click()
Load frm_comicschooser
frm_comicschooser.Show
End Sub

Private Sub mnu_comicsfolder_Click()
Load frm_options
frm_options.Show
End Sub

Public Function DateThing() As String

DateThing = Right(Calendar.Year, 2)

temp = Calendar.Month
If Len(temp) = 1 Then
DateThing = DateThing & "0" & Calendar.Month
MonthString = "0" & Calendar.Month
Else
DateThing = DateThing & Calendar.Month
MonthString = Calendar.Month
End If

temp = Calendar.Day
If Len(temp) = 1 Then
DateThing = DateThing & "0" & Calendar.Day
Else
DateThing = DateThing & Calendar.Day
End If

End Function
