VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSubclass"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Sub AttachMessage(iwp As ISubclass, ByVal hwnd As Long, ByVal Msg As Long)
Attribute AttachMessage.VB_Description = "Start subclassing for a particular message."
modSubclass.AttachMessage iwp, hwnd, Msg
End Sub

Sub DetachMessage(iwp As ISubclass, ByVal hwnd As Long, ByVal Msg As Long)
Attribute DetachMessage.VB_Description = "Stop subclassing for a particular message."
modSubclass.DetachMessage iwp, hwnd, Msg
End Sub

Public Function CallOldWindowProc( _
   ByVal hwnd As Long, _
   ByVal Msg As Long, _
   ByVal wParam As Long, _
   ByVal lParam As Long _
) As Long
Attribute CallOldWindowProc.VB_Description = "Calls the window procedure for this handle which was installed before you added the subclass."
CallOldWindowProc = modSubclass.CallOldWindowProc(hwnd, Msg, wParam, lParam)
End Function

