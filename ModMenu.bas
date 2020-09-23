Attribute VB_Name = "ModMenu"
'******************************************************************
'***************Copyright PSST 2003********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive

'Mainly menu related API here

Option Explicit
Public Type POINTAPI
    x As Long
    y As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    ItemID As Long
    itemAction As Long
    itemState As Long
    hWndItem As Long
    hDC As Long
    rcItem As RECT
    ItemData As Long
End Type
Private Type MEASUREITEMSTRUCT
    CtlType As Long
    CtlID As Long
    ItemID As Long
    ItemWidth As Long
    ItemHeight As Long
    ItemData As Long
End Type
Private Declare Sub DeleteUrlCacheEntry Lib "wininet.dll" (ByVal lpszUrlName As String)
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CreateMenu Lib "user32" () As Long
Private Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const DT_LEFT = &H0
Private Const DT_NOCLIP = &H100
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4
Private Const ODT_MENU = 1
Private Const ODS_SELECTED = &H1
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_MENU = 4
Private Const COLOR_HIGHLIGHTTEXT = 14
Public Const WM_NULL = &H0
Private Const WM_GETFONT = &H31
Private Const MF_BYPOSITION = &H400&
Private Const MF_POPUP = &H10&
Private Const MF_STRING = &H0&
Private Const MF_OWNERDRAW = &H100&
Private Const MF_SEPARATOR = &H800&
Private Const MF_REMOVE = &H1000&
Public Const GWL_WNDPROC = (-4)
Private Const WM_COMMAND = &H111
Private Const WM_CLOSE = &H10
Private Const WM_DRAWITEM = &H2B
Private Const WM_MEASUREITEM = &H2C
Public gOldProc As Long
Public BoboMenu As Collection
Private ParForm As Form
Public MenuBarBase As Long
Public MenuPURoot As Long
Public ExitHandle As Long
Public ExitSeparatorHandle As Long
Public CheckHandle As Long
Public LinkHandle As Long
Public SettingsHandle As Long
Public IL As ImageList
Public AutoCheck As Long
Public CheckFrequency As Long
Public LinkAddress As String
Public Sub TrashMenus(mForm As Form)
    'Called to remove the menus if we have made a major
    'change that requires WM_MEASUREITEM to be called again
    UnSubClass mForm.hwnd
    RemoveMenu MenuPURoot, 0, MF_BYPOSITION Or MF_REMOVE
    DestroyMenu MenuPURoot
    'Reload the menus
    LoadMenus mForm, True
End Sub
Public Sub LoadMenus(mForm As Form, Optional DontRemove As Boolean)
    Dim BB As BBMenu
    Dim z As Long
    Dim ExistingMenu As Long
    Dim temp As String
    Set BoboMenu = New Collection
    Set ParForm = mForm
    MenuBarBase = GetMenu(ParForm.hwnd) 'frmMenu's menu
    MenuPURoot = CreatePopupMenu 'our new popup menu
    InsertMenu MenuBarBase, GetMenuItemCount(MenuBarBase), MF_POPUP Or MF_STRING Or MF_BYPOSITION, MenuPURoot, ""
    'This is mnuMenu - a VB menu on frmMenu that initialises the menubar on that form
    'we wont be needing it so if we haven't already removed it - remove it now
    If Not DontRemove Then RemoveMenu MenuBarBase, 0, MF_BYPOSITION Or MF_REMOVE
    ExitHandle = CreateMenu 'exit menu
    ExitSeparatorHandle = CreateMenu 'a separator
    CheckHandle = CreateMenu 'check for new submissions menu
    SettingsHandle = CreateMenu 'Program settings menu
    'create a class for each menu - these classes are very small
    'and add each class to a collection
    Set BB = New BBMenu
    With BB
        .Caption = ""
        .Handle = MenuPURoot
        .Image = 0 'no image
        .ParentHandle = MenuBarBase
    End With
    BoboMenu.Add BB, Str(MenuPURoot)
    Set BB = New BBMenu
    With BB
        .Caption = "Navigate to PSC"
        .Handle = LinkHandle
        .Image = IL.ListImages("PSC").Index
        .ParentHandle = MenuPURoot
    End With
    BoboMenu.Add BB, Str(LinkHandle)
    Set BB = New BBMenu
    With BB
        .Caption = "Program Settings"
        .Handle = SettingsHandle
        .Image = IL.ListImages("Settings").Index
        .ParentHandle = MenuPURoot
    End With
    BoboMenu.Add BB, Str(SettingsHandle)
    Set BB = New BBMenu
    With BB
        .Caption = "Check for new submissions"
        .Handle = CheckHandle
        .Image = IL.ListImages("Check").Index
        .ParentHandle = MenuPURoot
    End With
    BoboMenu.Add BB, Str(CheckHandle)
    Set BB = New BBMenu
    With BB
        .Caption = "Exit"
        .Handle = ExitHandle
        .Image = 0
        .ParentHandle = MenuPURoot
    End With
    BoboMenu.Add BB, Str(ExitHandle)
    'Now insert the submenus into our Popup menu
    InsertMenu MenuPURoot, GetMenuItemCount(MenuPURoot), MF_OWNERDRAW Or MF_STRING Or MF_BYPOSITION, LinkHandle, "Navigate to PSC"
    InsertMenu MenuPURoot, GetMenuItemCount(MenuPURoot), MF_OWNERDRAW Or MF_STRING Or MF_BYPOSITION, CheckHandle, "Check for new submissions"
    InsertMenu MenuPURoot, GetMenuItemCount(MenuPURoot), MF_OWNERDRAW Or MF_STRING Or MF_BYPOSITION, SettingsHandle, "Program Settings"
    InsertMenu MenuPURoot, GetMenuItemCount(MenuPURoot), MF_SEPARATOR Or MF_STRING Or MF_BYPOSITION, ExitSeparatorHandle, ""
    InsertMenu MenuPURoot, GetMenuItemCount(MenuPURoot), MF_OWNERDRAW Or MF_STRING Or MF_BYPOSITION, ExitHandle, "Exit"
    SubClass ParForm.hwnd
End Sub
Public Sub SubClass(mhwnd As Long)
    'Start subclassing our form so we can respond to menu activity
    gOldProc& = GetWindowLong(mhwnd, GWL_WNDPROC)
    Call SetWindowLong(mhwnd, GWL_WNDPROC, AddressOf MenuProc)
End Sub
Public Sub UnSubClass(mhwnd As Long)
    'Stop subclassing - this is done automatically when frmMenu unloads
    Call SetWindowLong(mhwnd, GWL_WNDPROC, gOldProc&)
End Sub
Private Function MenuProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim z As Long
    Dim cnt As Long
    Dim MeasureInfo As MEASUREITEMSTRUCT
    Dim DrawInfo As DRAWITEMSTRUCT
    Dim BB As BBMenu
    Dim R As RECT
    Dim IsSelected As Boolean
    Dim hFont As Long
    Dim TopOffset As Long
    Select Case wMsg&
        Case WM_MEASUREITEM 'Menu sizing
            Call CopyMemory(MeasureInfo, ByVal lParam, Len(MeasureInfo))
            If MeasureInfo.CtlType <> ODT_MENU Then Exit Function
            If BoboMenu(Str(MeasureInfo.ItemID)) Is Nothing Then Exit Function
            Set BB = BoboMenu(Str(MeasureInfo.ItemID))
            MeasureInfo.ItemHeight = ParForm.TextHeight(BB.Caption) + 6
            MeasureInfo.ItemWidth = ParForm.TextWidth(BB.Caption) + 36
            Call CopyMemory(ByVal lParam, MeasureInfo, Len(MeasureInfo))
        Case WM_DRAWITEM 'Menu drawing - text and images
            Call CopyMemory(DrawInfo, ByVal lParam, LenB(DrawInfo))
            If DrawInfo.CtlType <> ODT_MENU Then Exit Function
            IsSelected = ((DrawInfo.itemState And ODS_SELECTED) = ODS_SELECTED)
            If BoboMenu(Str(DrawInfo.ItemID)) Is Nothing Then Exit Function
            Set BB = BoboMenu(Str(DrawInfo.ItemID)) 'use the correct class from our collection
            R = DrawInfo.rcItem 'get the rectangle for the specific menu
            'apply menu backcolor according to selection status
            'I've used default menu colors here
            If IsSelected Then
                FillRect DrawInfo.hDC, R, GetSysColorBrush(COLOR_HIGHLIGHT)
            Else
                FillRect DrawInfo.hDC, R, GetSysColorBrush(COLOR_MENU)
            End If
            'Apply frmMenu's font to the menu
            hFont = SendMessage(frmMenu.hwnd, WM_GETFONT, 0, 0)
            hFont = SelectObject(DrawInfo.hDC, hFont)
            'Set the forecolor
            SetTextColor DrawInfo.hDC, IIf(IsSelected, GetSysColor(COLOR_HIGHLIGHTTEXT), ParForm.ForeColor)
            'Print the text leaving room for an icon
            OffsetRect R, 26, 0
            SetBkMode DrawInfo.hDC, 1
            DrawText DrawInfo.hDC, BB.Caption, Len(BB.Caption), R, DT_SINGLELINE Or DT_LEFT Or DT_NOCLIP Or DT_VCENTER
            'Draw the icon
            If BB.Image > 0 Then
                TopOffset = (R.Bottom - R.Top - 16) / 2 'Center the icon vertically
                If TopOffset < 0 Then TopOffset = 0
                IL.ListImages(BB.Image).Draw DrawInfo.hDC, 4, R.Top + TopOffset, 1
            End If
        Case WM_CLOSE
            UnSubClass hwnd
        Case WM_COMMAND
            'respond to clicks on menus
            On Local Error Resume Next
            If BoboMenu(Str(wParam)) Is Nothing Then Exit Function
            Set BB = BoboMenu(Str(wParam))
            If BB Is Nothing Then Exit Function
            Select Case BB.Handle
                Case ExitHandle
                    Unload frmMenu
                Case CheckHandle
                    CheckSubmissions True
                Case SettingsHandle
                    frmSettings.Visible = True
                Case LinkHandle
                    If Len(LinkAddress) > 0 Then
                        ShellExecute ParForm.hwnd, vbNullString, LinkAddress, vbNullString, "c:\", 1
                    Else
                        MsgBox "You need to specify an address in Program Settings for this to work.", vbInformation, "PSC Checker"
                    End If
            End Select
    End Select
    MenuProc = CallWindowProc(gOldProc&, hwnd&, wMsg&, wParam&, lParam&)
End Function
Private Function OneGulp(Src As String) As String
    'read a text file
    On Error Resume Next
    Dim f As Integer, temp As String
    f = FreeFile
    DoEvents
    Open Src For Binary As #f
    temp = String(LOF(f), Chr$(0))
    Get #f, , temp
    Close #f
    OneGulp = temp
End Function

Public Sub CheckSubmissions(Optional ShowMessage As Boolean = False)
    Dim NewCnt As Long
    Dim tmpCnt As Long
    Dim temp As String
    Dim msgString As String
    'stop flashing - reset trayicon to default setting
    frmMenu.FlashTimer.Enabled = False
    frmMenu.SysIcon.TipText = "PSC Checker"
    If frmMenu.SysIcon.IconHandle <> IL.ListImages("PSC").Picture Then frmMenu.SysIcon.IconHandle = IL.ListImages("PSC").Picture
    temp = IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\") & "Source.tmp"
    tmpCnt = Val(GetSetting("PSST SOFTWARE\PSCChecker", "Settings", "Lastcount", ""))
    'Download the ID of the most recent submission - thanks to Ian at PSC for this page
    DeleteUrlCacheEntry "http://www.pscode.com/vb/feeds/LatestCodeId.asp?lngWId=1"
    URLDownloadToFile 0, "http://www.pscode.com/vb/feeds/LatestCodeId.asp?lngWId=1", temp, 0, 0
    'read the source code for the downloaded page
    'It is only a number so is very small in bytes
    NewCnt = Val(OneGulp(temp))
    If NewCnt <> 0 Then
        'Save the results to registry
        SaveSetting "PSST SOFTWARE\PSCChecker", "Settings", "Lastcount", Trim(Str(NewCnt))
    Else
        'If it's zero we must have failed - maybe we're off line
        frmMenu.SysIcon.TipText = "Failed to contact PSC"
        'show a message if not checking automatically via a timer
        If ShowMessage Then MsgBox "Failed to contact PSC", vbCritical, "PSC Checker"
        Exit Sub
    End If
    'If the result is different from last time...
    If NewCnt <> tmpCnt Then
        If ShowMessage Then
            frmMenu.SysIcon.IconHandle = IL.ListImages("PSC").Picture
            If NewCnt < tmpCnt Then
                msgString = "The most recent submission appears to have been removed."
            Else
                msgString = "New submission detected on PSC."
            End If
        Else
            'if checking automatically via a timer
            'flash the icon to inform of new submission
            frmMenu.FlashTimer.Enabled = True
        End If
    Else
        msgString = "No new submissions detected on PSC."
        frmMenu.SysIcon.IconHandle = IL.ListImages("PSC").Picture
    End If
    On Error Resume Next
    Kill temp 'remove the temp file
    'show a message if not checking automatically via a timer
    If ShowMessage Then MsgBox msgString, vbInformation, "PSC Checker"

End Sub
