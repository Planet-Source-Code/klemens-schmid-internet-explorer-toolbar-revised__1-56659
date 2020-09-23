Attribute VB_Name = "modMenu"
'*********************************************************************************************
' Author:      Klemens Schmid
' E-Mail:      klemens.schmid@gmx.de
' Web Page:    http://www.klemid.de
'
' Description:
' Provides functions to re-use a standard VB menu as PopupMenu in
' another form, e.g. in a Explorer Toolbar.
'
'*********************************************************************************************
Option Explicit

Public Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, _
   ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, _
   ByVal hwnd As Long, lprc As RECT) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long

Global Const MF_BYPOSITION = &H400&

Global Const SM_CXSCREEN = 0     'Width of screen in pixels
Global Const SM_CYSCREEN = 1     'Height of screen in pixels
Global Const SM_CYCAPTION = 4    'Height of form titlebar in pixels
Global Const SM_CXICON = 11      'Width of icon in pixels
Global Const SM_CYICON = 12      'Height of icon in pixels
Global Const SM_CXCURSOR = 13    'Width of mousepointer in pixels
Global Const SM_CYCURSOR = 14    'Height of mousepointer in pixels
Global Const SM_CYMENU = 15      'Height of top menu bar in pixels

Public Function GetSubmenuByCaption(ByVal SourceForm As Form, ByVal Caption As String) As Long
'returns the handle of a submenu with a given caption
Dim pos As Long
Dim lngMenuItems As Long
Dim lngCaptionLength As Long
Dim strCaption As String
Dim hMenu As Long

'get the form's menu
hMenu = GetMenu(SourceForm.hwnd)
If hMenu = 0 Then Exit Function

'Get the count of menu items in this menu.
lngMenuItems = GetMenuItemCount(hMenu)

'Loop through all the items on the menu
For pos = 0 To lngMenuItems - 1
   strCaption = Space(255)
   lngCaptionLength = GetMenuString(hMenu, pos, strCaption, 255, MF_BYPOSITION)
   strCaption = Left$(strCaption, lngCaptionLength)
   If (strCaption = Caption) Then
      'we've got it. Get the handle by pos
      GetSubmenuByCaption = GetSubMenu(hMenu, pos)
      Exit Function
   End If
Next
End Function


