VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmToolbar 
   Caption         =   "Provides the Toolbar controls"
   ClientHeight    =   3000
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTest 
      Caption         =   "&Test1"
      Height          =   315
      Left            =   2760
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtTest 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.ComboBox cboSearch 
      Height          =   315
      ItemData        =   "frmToolbar.frx":0000
      Left            =   120
      List            =   "frmToolbar.frx":000D
      TabIndex        =   1
      Text            =   "Type your search string here"
      Top             =   720
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ToolbarImages 
      Left            =   120
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolbar.frx":0020
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToolbar.frx":05BA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblDescription 
      Caption         =   "This form provides the menus and controls which appear in the search bar plus the code handling the events for those."
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About Klemid's Search Bar"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "Search"
      Begin VB.Menu mnuSite 
         Caption         =   "Site"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************************
' Author:      Klemens Schmid
' E-Mail:      klemid@gmail.de
' Web Page:    http://www.klemid.de
'
' Description:
' This form provides the menus and controls which appear in the
' sample bar plus the code handling the events for those.
'*********************************************************************************************

Option Explicit


Const MODULE_NAME = "frmToolbar"

Private m_IE As InternetExplorer                'the extensions Explorer
Private m_hWndInput As Long                     'window handle
Private m_Toolbar As CToolbar                   'the toolbar keeping our controls

Private Sub cboSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then MsgBox "User pressed Return"
End Sub

'***Add code here for new controls

Public Sub cmdTest_Click()
MsgBox "Textbox contains '" & txtTest & "'"
End Sub

Private Sub mnuAbout_Click()
MsgBox "Sample Explorer Toolbar by Klemens Schmid."
End Sub

Public Sub Init(ByVal IE As InternetExplorer, ByVal Toolbar As CToolbar)
'remember the IE reference
Set m_IE = IE
Set m_Toolbar = Toolbar
'populate the menus and the combobox
'Of course here is plenty of room for improvement but remember it's only a sample
With mnuSite(0)
   .Caption = "&1 Google.de"
   .Tag = "http://www.google.de/search?q=<searchstring>"
End With
Load mnuSite(1)
With mnuSite(1)
   .Caption = "&2 groups.google.com"
   .Tag = "http://groups.google.de/groups?q=<searchstring>"
End With

End Sub

Public Sub SetSearchField(hwnd As Long)
'the toolbar tells us which is the input controls with the search string
m_hWndInput = hwnd
End Sub

Private Sub mnuSite_Click(Index As Integer)
'navigate to the site
Dim strSearch As String
Dim strUrl As String

'get the search string
strSearch = cboSearch.Text
strUrl = mnuSite(Index).Tag
'replace in url
strUrl = Replace(strUrl, "<searchstring>", strSearch)
'do it
m_IE.Navigate strUrl
End Sub
