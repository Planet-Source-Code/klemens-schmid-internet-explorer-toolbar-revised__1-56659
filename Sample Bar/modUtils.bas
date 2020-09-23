Attribute VB_Name = "modUtils"

'*********************************************************************************************
'
' Shell Bands
'
' Declarations module
'
'*********************************************************************************************
'
' Authors:     Eduardo Morcillo
'              Klemens Schmid (seriously changed the code)
' E-Mail:      klemens.schmid@gmx.de
' Web Page:    www.klemid.de
'
' 03/12/2000:  Created by Eduardo Morcillo
' 03/21/2000:  FindIESite now uses the IServiceProvider interface
'              of the band site to get the IE window.
' 12/25/2002:  Seriously modified by Klemens
'
'*********************************************************************************************
Option Explicit

Public Type WINDOWPOS
    hWnd As Long
    hWndInsertAfter As Long
    X As Long
    Y As Long
    cx As Long
    cy As Long
    flags As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type TBBUTTON
   iBitmap As Long
   idCommand As Long
   fsState As Byte
   fsStyle As Byte
   bReserved(0 To 1) As Byte
   dwData As Long
   iString As Long
End Type

Public Type NMTOOLBAR_SHORT
    hdr As NMHDR
    iItem As Long
End Type

Private Type TBBUTTONINFO
   cbSize As Long
   dwMask As Long
   idCommand As Long
   iImage As Long
   fsState As Byte
   fsStyle As Byte
   cx As Integer
   lParam As Long
   pszText As Long
   cchText As Long
End Type

Public Const SW_HIDE = 0
Public Const SW_SHOW = 1
Public Const SW_SHOWNOACTIVATE = 1

Public Const GWL_STYLE = (-16)
Public Const GWL_WNDPROC = (-4)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_ID = (-12)

Public Const WS_VISIBLE = &H10000000
Public Const WS_CHILD = &H40000000
Public Const WS_EX_CLIENTEDGE = &H200&
Public Const WS_TABSTOP = &H10000

Public Const CCS_NORESIZE = &H4&
Public Const CCS_NODIVIDER = &H40&

Public Const RDW_INVALIDATE = &H1
Public Const RDW_UPDATENOW = &H100
Public Const RDW_ERASE = &H4
Public Const RDW_ERASENOW = &H200
Public Const RDW_ALLCHILDREN = &H80

'textbox messages and notifications
Public Const EM_GETSEL = &HB0
Public Const EM_SETSEL = &HB1
Public Const EN_SETFOCUS = &H100
Public Const EN_KILLFOCUS = &H200
Public Const ES_AUTOHSCROLL = &H80&
Public Const CBN_SETFOCUS = &H3
Public Const CBN_KILLFOCUS = &H4

'combobox messages
Public Const CB_SETEDITSEL = &H142
Public Const CB_FINDSTRING = &H14C

Public Const HWND_MESSAGE = -3

Public Const WM_SETFOCUS = &H7
Public Const WM_KILLFOCUS = &H8
Public Const WM_SETTEXT = &HC
Public Const WM_SETFONT = &H30
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CHAR = &H102
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105

Public Const WM_COMMAND = &H111
Public Const WM_USER = &H400
Public Const WM_NOTIFY = &H4E
Public Const WM_LBUTTONDOWN = &H201

Public Const PROP_PREVPROC = "WinProc"
Public Const PROP_OBJECT = "Object"

' Toolbar notification messages:
Public Const TBN_LAST = &H720
Public Const TBN_FIRST = -700&
Public Const TBN_GETBUTTONINFOA = (TBN_FIRST - 0)
Public Const TBN_GETBUTTONINFOW = (TBN_FIRST - 20)
Public Const TBN_BEGINDRAG = (TBN_FIRST - 1)
Public Const TBN_ENDDRAG = (TBN_FIRST - 2)
Public Const TBN_BEGINADJUST = (TBN_FIRST - 3)
Public Const TBN_ENDADJUST = (TBN_FIRST - 4)
Public Const TBN_RESET = (TBN_FIRST - 5)
Public Const TBN_QUERYINSERT = (TBN_FIRST - 6)
Public Const TBN_QUERYDELETE = (TBN_FIRST - 7)
Public Const TBN_TOOLBARCHANGE = (TBN_FIRST - 8)
Public Const TBN_CUSTHELP = (TBN_FIRST - 9)
Public Const TBN_CLOSEUP = (TBN_FIRST - 11)
Public Const TBN_DROPDOWN = (TBN_FIRST - 10)
Public Const TBN_HOTITEMCHANGE = (TBN_FIRST - 13)

' Toolbar and button styles:
Public Const TBSTYLE_BUTTON = &H0
Public Const TBSTYLE_SEP = &H1
Public Const TBSTYLE_CHECK = &H2
Public Const TBSTYLE_GROUP = &H4
Public Const TBSTYLE_CHECKGROUP = (TBSTYLE_GROUP Or TBSTYLE_CHECK)
Public Const TBSTYLE_DROPDOWN = &H8
Public Const TBSTYLE_TOOLTIPS = &H100
Public Const TBSTYLE_WRAPABLE = &H200
Public Const TBSTYLE_ALTDRAG = &H400
Public Const TBSTYLE_FLAT = &H800
Public Const TBSTYLE_LIST = &H1000
Public Const TBSTYLE_AUTOSIZE = &H10         '// automatically calculate the cx of the button
Public Const TBSTYLE_NOPREFIX = &H20         '// if this button should not have accel prefix
Public Const BTNS_WHOLEDROPDOWN = &H80 '??? IE5 only
Public Const TBSTYLE_REGISTERDROP = &H4000&
Public Const TBSTYLE_TRANSPARENT = &H8000&

Public Const BTNS_BUTTON = &H0
Public Const BTNS_SEP = &H1
Public Const BTNS_AUTOSIZE = &H10

Public Const TBSTATE_ENABLED = &H4

Public Enum ECTBToolButtonSyle
    CTBNormal = TBSTYLE_BUTTON
    CTBSeparator = TBSTYLE_SEP
    CTBCheck = TBSTYLE_CHECK
    CTBCheckGroup = TBSTYLE_CHECKGROUP
    CTBDropDown = TBSTYLE_DROPDOWN
    CTBAutoSize = TBSTYLE_AUTOSIZE
    CTBDropDownArrow = BTNS_WHOLEDROPDOWN
End Enum

'Toolbar messages
Public Const TB_GETITEMRECT = (WM_USER + 29)
Public Const TB_GETBUTTON = (WM_USER + 23)
Public Const TB_PRESSBUTTON = (WM_USER + 3)
Public Const TB_BUTTONCOUNT = (WM_USER + 24)
Public Const TB_GETRECT = (WM_USER + 51)
Public Const TB_SETEXTENDEDSTYLE = (WM_USER + 84)
Public Const TB_GETBUTTONINFO = (WM_USER + 65)
Public Const TB_SETBUTTONINFO = (WM_USER + 66)
Public Const TB_SETIMAGELIST = &H400 + 48
Public Const TB_ADDBUTTONSW = &H400 + 68
Public Const TBSTYLE_EX_DRAWDDARROWS = &H1

Public Const ILC_MASK = &H1&
Public Const ILC_COLOR8 = &H8&

'TrackPopupMenu styles
Public Const TPM_CENTERALIGN = &H4&
Public Const TPM_LEFTALIGN = &H0&
Public Const TPM_LEFTBUTTON = &H0&
Public Const TPM_RIGHTALIGN = &H8&
Public Const TPM_RIGHTBUTTON = &H2&
Public Const TPM_TOPALIGN = &H0
Public Const TPM_VCENTERALIGN = &H10
Public Const TPM_BOTTOMALIGN = &H20
Public Const TPM_HORIZONTAL = &H0
Public Const TPM_VERTICAL = &H40
Public Const TPM_NONOTIFY = &H80
Public Const TPM_RETURNCMD = &H100

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetFocus Lib "user32" () As Long

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Function GetLastError Lib "kernel32" () As Long

Public Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As olelib.MSG) As Long
Public Declare Function TranslateMessage Lib "user32" (lpMsg As olelib.MSG) As Long

Public Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Public Declare Function GetWindowID Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const PAGE_EXECUTE_READWRITE = &H40

Public Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd&, Wrect As Any) As Long
Public Declare Function GetWindowPos Lib "user32" Alias "GetWindowPosA" (ByVal hWnd As Long, lParam As Any) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Any) As Long

Public Declare Sub InitCommonControls Lib "comctl32" ()


Public Declare Function ImageList_Create Lib "comctl32" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Public Declare Function ImageList_ReplaceIcon Lib "comctl32" (ByVal himl As Long, ByVal i As Long, ByVal hicon As Long) As Long

Public Declare Function CreateToolbarEx Lib "comctl32" ( _
   ByVal hWnd As Long, _
   ByVal ws As Long, _
   ByVal wID As Long, _
   ByVal nBitmaps As Long, _
   ByVal hBMInst As Long, _
   ByVal wBMID As Long, _
   lpButtons As Any, _
   ByVal iNumButtons As Long, _
   ByVal dxButton As Long, _
   ByVal dyButton As Long, _
   ByVal dxBitmap As Long, _
   ByVal dyBitmap As Long, _
   ByVal uStructSize As Long) As Long

Public Function FindIESite(ByVal BandSite As olelib.IServiceProvider) As IWebBrowserApp
' Returns the explorer window that contains
' the band site
'
' Parameters:
' BandSite    IOleWindow interface of the band site
'
Dim IID_IWebBrowserApp As olelib.UUID
Dim SID_SInternetExplorer As olelib.UUID
  
' Convert IID and SID
' from strings to UUID UDTs
CLSIDFromString IIDSTR_IWebBrowserApp, IID_IWebBrowserApp
CLSIDFromString SIDSTR_SInternetExplorer, SID_SInternetExplorer

' Get the InternetExplorer
' object through IServiceProvider
BandSite.QueryService SID_SInternetExplorer, IID_IWebBrowserApp, FindIESite
         
End Function

Sub Main()

'Initialize common controls
'InitCommonControls
   
End Sub
