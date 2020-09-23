Attribute VB_Name = "modLog4VB"
Attribute VB_HelpID = 500
Attribute VB_Description = "Procedure-based API for Log4VB"
'Klemid's Trace API
'by klemens.schmid@gmx.de, 2004

'This code module provides the basic trace capabilities.

'This API is provided as source code. You may modify and redistribute it.
'However the Log4VB viewer receiving and displaying the trace output
'is not provided as source code. It is available as freeware and shareware
'version at www.log4vb.com.

Option Explicit

Public TraceLevel As Integer              'trace level

Public Enum log4vbSeverity
   log4vbInfo
   log4vbWarning
   log4vbError
End Enum

Public Const VERSION = "2.0"

Private Const WM_COPYDATA = &H4A          'ID of the message

Private Type COPYDATASTRUCT               'for data transfer to Trace Log4VB Viewer
    dwData As Long
    cbData As Long
    lpData As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Const VIEWER_CAPTION = " Log4VB"

Private m_AppName As String               'application name
Private m_NestingLevel As Integer         'nesting level (call stack)
Private m_DefaultTraceLevel As Integer    'used when no explicit trace level is specified
Private m_UseSendMessage As Boolean       'use SendMessage rather than PostMessage

Public Sub Init(ByVal AppName As String, ByVal DefaultTraceLevel As Integer, ByVal UseSendMessage As Boolean)
'Description
'  Should be called once at the beginning of the program
'  to initialize the trace.
'  If you want to switch between SendMessage and PostMessage later on
'  you may call Init several times.
'Parameters
'  AppName:             name of the application passed to the trace viewer
'  DefaultTraceLevel:   the trace level use whether not explicitely passed in Log4VB
'  UseSendMessage:      use SendMessage. Advantage: current trace level is passed back

'remember the parameters
m_AppName = AppName
m_DefaultTraceLevel = DefaultTraceLevel
m_UseSendMessage = UseSendMessage

End Sub

Public Sub Log4VB(ByVal Text$, _
               Optional ByVal Module As String = "", _
               Optional ByVal Procedure As String = "", _
               Optional ByVal TraceLevel As Integer = -1, _
               Optional ByVal Severity As log4vbSeverity = log4vbInfo, _
               Optional ByVal User1 As String, _
               Optional ByVal User2 As String)
'Description
'  send a trace message to the Log4VB viewer
'Parameters
'  Text:          description
'  Module:        source module
'  Procedure:     source procedure
'  Severity:      error, warning or info
'  TraceLevel:    Trace granularity. 1=low, 9=high
'  User1:         user-defined parameter
'  User2:         user-defined parameter

Dim strLine$                           'line to be printed
Dim udtData As COPYDATASTRUCT          'structure carrying the net data
Static shWnd As Long                   'to store the trace window handle
Static sTraceInit As Boolean           'to ident if tracing is up
Dim nFile As Integer                   'file handle
Const NO_DATE = ""                     'leave date empty

On Error Resume Next

' check if we are in trace mode
#If Not TRACE_ON > 0 Then
Exit Sub
#Else
'prevent from further execution of sub if trace level is NULL
If TraceLevel = -1 Then TraceLevel = m_DefaultTraceLevel
If TraceLevel <= 0 And sTraceInit = True Then Exit Sub
If TraceLevel > modLog4VB.TraceLevel And sTraceInit = True Then Exit Sub

' set default for component name
If Len(m_AppName) = 0 Then
   m_AppName = GetProjectName
End If
' prepare the string line
strLine = VERSION & vbNullChar & _
          Text & vbNullChar & _
          Procedure & vbNullChar & _
          Module & vbNullChar & _
          App.EXEName & vbNullChar & _
          m_AppName & vbNullChar & _
          User1 & vbNullChar & _
          User2 & vbNullChar & _
          NO_DATE & vbNullChar & _
          Severity & vbNullChar & _
          TraceLevel & vbNullChar & _
          m_NestingLevel & vbNullChar

If Not sTraceInit Then
   ' get the window handle of trace log
   shWnd = FindWindow(vbNullString, VIEWER_CAPTION)
   If shWnd Then
      'do not change this value (1) it identifies the message
      udtData.dwData = 1
      'Send the message to trace log and get the trace level
      modLog4VB.TraceLevel = SendMessage(shWnd, WM_COPYDATA, shWnd, udtData)
      sTraceInit = True
   Else
      'if no tracing is possible
      TraceLevel = 0
   End If
End If

If shWnd Then
   'do not change this value (0) it identifies the message
   udtData.dwData = 0
   'set the length of the message text
   udtData.cbData = LenB(strLine)
   'set the message text
   udtData.lpData = StrPtr(strLine)
   'post the message
   If m_UseSendMessage Then
      'Send message synchronously. Trace level is returned
      modLog4VB.TraceLevel = SendMessage(shWnd, WM_COPYDATA, shWnd, udtData)
   Else
      PostMessage shWnd, WM_COPYDATA, shWnd, udtData
   End If
End If

#End If
End Sub

Private Function GetProjectName() As String
'Description
'  This function figures out the name of our VB project

Static strProjectName As String

If Len(strProjectName) Then
   'not first time
   GetProjectName = strProjectName
   Exit Function
End If

'force error to find out the current source
On Error GoTo Trap
Err.Raise 0

Trap:
strProjectName = Err.Source
GetProjectName = strProjectName

End Function

Public Function VersionString() As String
'assemble the version string
VersionString = App.Major & "." & App.Minor & "." & App.Revision
End Function

