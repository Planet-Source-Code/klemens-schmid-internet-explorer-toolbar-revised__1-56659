VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'defines the interface classes have to implement in order to
'use subclassing

Function WindowProc(ByVal hwnd As Long, _
                    ByVal Msg As Long, _
                    ByVal wParam As Long, _
                    ByVal lParam As Long, _
                    ByRef Handled As Boolean) As Long
'the delegate of the window proc within the class
'Handled indicates whether the message has been handler by the window proc delegate.
'If Handled is false the default window proc is invoked afterwards
End Function

