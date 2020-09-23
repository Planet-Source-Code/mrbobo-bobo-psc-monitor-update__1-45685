VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmMenu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   780
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   1890
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   52
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   126
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer FlashTimer 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer CheckTimer 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   120
      Top             =   120
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1140
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":0E42
            Key             =   "Settings"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":13DC
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1976
            Key             =   "PSC"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1F10
            Key             =   "Yellow"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":24AA
            Key             =   "Font"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuMenu 
      Caption         =   ""
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'***************Copyright PSST 2003********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive

'This form is just a container for our menus and images and
'provides us with a start form and a base handle for API calls
'You could use a single form but this can cause issues
'with more complex Subclassing than is used here so
'for that reason I like to use a separate form
Option Explicit
'API to show our API generated menu from the Tray
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal lprc As Any) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_LBUTTONDOWN = &H201
Private Const TPM_LEFTALIGN = &H0&
Public SysIcon As TrayClass
Dim TimerCount As Long
'Check for new submissions
Private Sub CheckTimer_Timer()
    'Timer is set to 1 minute
    'CheckFrequency tells us how many minutes
    TimerCount = TimerCount + 1
    If TimerCount >= CheckFrequency Then
        TimerCount = 0
        CheckSubmissions
    End If
End Sub
'Flash the Tray icon if new submission detected
Private Sub FlashTimer_Timer()
    'Timer is set to 350 milliseconds
    SysIcon.IconHandle = IIf(SysIcon.IconHandle <> IL.ListImages("Yellow").Picture, IL.ListImages("Yellow").Picture, IL.ListImages("PSC").Picture)
End Sub

Private Sub Form_Load()
    'We only need one instance thanks
    If App.PrevInstance Then End
    'Retrieve the menu font from registry
    With Me
        .FontBold = CBool(Val(GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "FontBold", "0")))
        .FontItalic = CBool(Val(GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "FontItalic", "0")))
        .FontName = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "FontName", "MS Sans Serif")
        .Fontsize = Val(GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "Fontsize", "8"))
        .ForeColor = Val(GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "ForeColor", "0"))
        .FontStrikethru = CBool(Val(GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "FontStrikethru", "0")))
        .FontUnderline = CBool(Val(GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "FontUnderline", "0")))
    End With
    'Automatically check for new submissions
    AutoCheck = Val(GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "AutoCheck", "1"))
    'How often should we check ?
    CheckFrequency = Val(GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "CheckFrequency", "5"))
    'Where should we navigate to from the link on the menu ?
    LinkAddress = GetSetting("PSST SOFTWARE\" + App.Title, "Settings", "LinkAddress", "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?grpCategories=-1&optSort=DateDescending&txtMaxNumberOfEntriesPerPage=10&blnNewestCode=TRUE&blnResetAllVariables=TRUE&lngWId=1")
    'Enable automatic checking if desired
    CheckTimer.Enabled = CBool(AutoCheck)
    'Load up the tray icon
    Set SysIcon = New TrayClass
    'Set the imagelist to use
    Set IL = frmMenu.ImageList1
    SysIcon.Initialize hwnd, IL.ListImages("PSC").Picture, "PSC Checker"
    SysIcon.ShowIcon
    'Build the menus - see ModMenu
    LoadMenus Me
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Respond to the mouse events of the tray icon
    Dim msgCallBackMessage As Long
    Dim Ttip As String
    Dim P As POINTAPI
    On Error Resume Next
    msgCallBackMessage = x
    DoEvents
    'Stop flashing icon timer
    frmMenu.FlashTimer.Enabled = False
    'Reset the tray icon to default
    If frmMenu.SysIcon.IconHandle <> IL.ListImages("PSC").Picture Then frmMenu.SysIcon.IconHandle = IL.ListImages("PSC").Picture
    Select Case msgCallBackMessage
        'Show the popup menu
        Case WM_RBUTTONDOWN
            DoEvents
            SetForegroundWindow hwnd
            GetCursorPos P
            If MenuPURoot = 0 Then Exit Sub
            SetForegroundWindow hwnd
            TrackPopupMenu MenuPURoot, TPM_LEFTALIGN, P.x, P.y, 0, hwnd, ByVal 0&
            PostMessage hwnd, WM_NULL, 0, 0
        'Launch the Settings dialog
        Case WM_LBUTTONDOWN
            frmSettings.Visible = True
            SetForegroundWindow frmSettings.hwnd
    End Select

End Sub


Private Sub Form_Unload(Cancel As Integer)
    'Save all settings to registry
    Dim frm As Form
    SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "LinkAddress", LinkAddress
    SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "AutoCheck", AutoCheck
    SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "CheckFrequency", CheckFrequency
    With Me
        SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "FontBold", IIf(.FontBold, "1", "0")
        SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "FontItalic", IIf(.FontItalic, "1", "0")
        SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "FontUnderline", IIf(.FontUnderline, "1", "0")
        SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "FontStrikethru", IIf(.FontStrikethru, "1", "0")
        SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "FontName", .FontName
        SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "Fontsize", .Fontsize
        SaveSetting "PSST SOFTWARE\" + App.Title, "Settings", "ForeColor", .ForeColor
    End With
    'Make sure we unload correctly
    For Each frm In Forms
        Unload frm
        Set frm = Nothing
    Next
End Sub

