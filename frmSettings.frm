VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program Settings"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3405
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   204
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   227
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtLink 
      Height          =   315
      Left            =   240
      TabIndex        =   9
      Top             =   2580
      Width           =   2895
   End
   Begin VB.CommandButton cmdCheck 
      Height          =   315
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1920
      Width           =   345
   End
   Begin VB.CommandButton cmdMenuFont 
      Caption         =   "..."
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   345
   End
   Begin VB.ComboBox cboMinutes 
      Height          =   315
      ItemData        =   "frmSettings.frx":000C
      Left            =   1020
      List            =   "frmSettings.frx":0025
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   795
   End
   Begin VB.CheckBox ChMinutes 
      Caption         =   "     Minutes"
      Height          =   255
      Left            =   1380
      TabIndex        =   3
      Top             =   1005
      Width           =   1215
   End
   Begin VB.CheckBox ChCheck 
      Caption         =   "Check for new submissions every..."
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   660
      Value           =   1  'Checked
      Width           =   2835
   End
   Begin VB.CheckBox ChRun 
      Caption         =   "Run at Windows startup"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label lblPSC 
      Caption         =   "PSC Link Address"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2340
      Width           =   1335
   End
   Begin VB.Label lblCheck 
      Caption         =   "Check for new submissions"
      Height          =   195
      Left            =   660
      TabIndex        =   7
      Top             =   1980
      Width           =   1995
   End
   Begin VB.Label lblFont 
      Caption         =   "Change menu font"
      Height          =   255
      Left            =   660
      TabIndex        =   4
      Top             =   1620
      Width           =   1455
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'***************Copyright PSST 2003********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive

'A very simple "Options" form
Option Explicit
Private Sub cboMinutes_Click()
    CheckFrequency = IIf(cboMinutes.ListIndex = 0, 3, cboMinutes.ListIndex * 5)
End Sub

Private Sub ChCheck_Click()
    cboMinutes.Enabled = CBool(ChCheck.Value)
    ChMinutes.Enabled = cboMinutes.Enabled
    AutoCheck = ChCheck.Value
    frmMenu.CheckTimer.Enabled = CBool(AutoCheck)
End Sub

Private Sub cmdCheck_Click()
    CheckSubmissions True
End Sub

Private Sub cmdMenuFont_Click()
    'Brings up the Choose Font dialog - see module CmnDlgFont
    With SelectFont
        'Set initial dialog settings according to frmMenu
        .FontBold = frmMenu.FontBold
        .FontItalic = frmMenu.FontItalic
        .FontName = frmMenu.FontName
        .Fontsize = frmMenu.Fontsize
        .FontColor = frmMenu.ForeColor
        .FontStrikethru = frmMenu.FontStrikethru
        .FontUnderline = frmMenu.FontUnderline
        If Not ShowFont(hwnd) Then Exit Sub
        'Apply the chosen settings
        frmMenu.FontBold = .FontBold
        frmMenu.FontItalic = .FontItalic
        frmMenu.FontName = .FontName
        frmMenu.Fontsize = .Fontsize
        frmMenu.ForeColor = .FontColor
        frmMenu.FontStrikethru = .FontStrikethru
        frmMenu.FontUnderline = .FontUnderline
    End With
    TrashMenus frmMenu

End Sub

Private Sub Form_Load()
    'Load our icons from the imagelist - reduces EXE size
    Me.Icon = IL.ListImages("PSC").Picture
    Set cmdCheck.Picture = IL.ListImages("Check").Picture
    'Adjust controls to current settings
    ChRun.Value = IsRunAtStartUp
    ChCheck.Value = AutoCheck
    cboMinutes.ListIndex = IIf(CheckFrequency = 3, 0, CheckFrequency / 5)
    txtLink.Text = LinkAddress
End Sub
Private Sub ChRun_Click()
    'Write to registry in order to start with Windows
    Dim temp As String
    temp = App.Path
    If Right(temp, 1) <> "\" Then temp = temp + "\"
    If ChRun.Value = 1 Then
        SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", App.Title, temp + App.EXEName + ".exe"
    Else
        SaveSettingString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", App.Title, temp + App.EXEName + ".exe"
        DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", App.Title
    End If
End Sub

Private Sub txtLink_Change()
    'Change the LinkAddress variable accordingly
    LinkAddress = txtLink.Text
End Sub
