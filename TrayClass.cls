VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TrayClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************
'***************Copyright PSST 2003********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive

'Forget where I got this from it was so long ago - but it's pretty standard tray icon stuff
Option Explicit
Private Declare Function GetLastError Lib "Kernel32" () As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const MAX_TIP_LENGTH As Long = 64
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * MAX_TIP_LENGTH
End Type
Private nidTrayIcon As NOTIFYICONDATA
Private bIconDisplayed As Boolean
Private bUpdateOnChange As Boolean
Public Event NIError(ByVal ErrorNumber As Long)
Public Function Initialize(ByVal hwnd As Long, ByVal hIcon As Long, ByVal sTip As String, Optional ByVal uCallbackMessage As Long = WM_MOUSEMOVE) As Long
  With nidTrayIcon
   .cbSize = Len(nidTrayIcon)
   .hIcon = hIcon
   .hwnd = hwnd
   .szTip = Left(sTip, MAX_TIP_LENGTH - 1) & vbNullChar
   .uCallbackMessage = uCallbackMessage
   .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
   .uId = vbNull
  End With
  bIconDisplayed = False
  bUpdateOnChange = True
End Function
Public Function ShowIcon() As Boolean
  If Not bIconDisplayed Then
    ShowIcon = Shell_NotifyIcon(NIM_ADD, nidTrayIcon)
    If ShowIcon = False Then
      RaiseEvent NIError(GetLastError)
    Else
      bIconDisplayed = True
    End If
  End If
End Function
Public Function HideIcon() As Boolean
   If bIconDisplayed Then
    HideIcon = Shell_NotifyIcon(NIM_DELETE, nidTrayIcon)
    If HideIcon = False Then
      RaiseEvent NIError(GetLastError)
    Else
      bIconDisplayed = False
    End If
  End If
End Function
Public Property Let IconHandle(ByVal hIcon As Long)
  nidTrayIcon.hIcon = hIcon
  If bUpdateOnChange Then
    nidTrayIcon.uFlags = NIF_ICON
    Update
    nidTrayIcon.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
  End If
End Property
Public Property Let TipText(ByVal sTip As String)
  nidTrayIcon.szTip = Left(sTip, MAX_TIP_LENGTH - 1) & vbNullChar
  If bUpdateOnChange Then
    nidTrayIcon.uFlags = NIF_TIP
    Update
    nidTrayIcon.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
  End If
End Property
Public Property Let CallbackMessage(ByVal uCallbackMessage As Long)
  nidTrayIcon.uCallbackMessage = uCallbackMessage
  If bUpdateOnChange Then
    nidTrayIcon.uFlags = NIF_MESSAGE
    Update
    nidTrayIcon.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
  End If
End Property
Public Function Update() As Boolean
  If bIconDisplayed Then
    Update = Shell_NotifyIcon(NIM_MODIFY, nidTrayIcon)
    If Update = False Then
      RaiseEvent NIError(GetLastError)
    End If
  End If
End Function
Public Property Get IconHandle() As Long
  IconHandle = nidTrayIcon.hIcon
End Property
Public Property Get TipText() As String
   TipText = Left(nidTrayIcon.szTip, Len(nidTrayIcon.szTip) - 1)
End Property
Public Property Get CallbackMessage() As Long
  CallbackMessage = nidTrayIcon.uCallbackMessage
End Property
Public Property Let UpdateOnChange(bUpdate As Boolean)
  bUpdateOnChange = bUpdate
End Property
Private Property Get UpdateOnChange() As Boolean
  UpdateOnChange = bUpdateOnChange
End Property
Private Sub Class_Terminate()
  HideIcon
End Sub
Public Property Get Visible() As Boolean
  If bIconDisplayed Then
    Visible = True
  End If
End Property
Public Property Let Visible(ByVal bVisible As Boolean)
   If bVisible Then
    ShowIcon
  Else
    HideIcon
  End If
End Property
