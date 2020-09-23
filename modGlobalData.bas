Attribute VB_Name = "modGlobal"
Option Explicit

Public Const PORT_HTTP As Long = 80&
Public Const PORT_SMTP As Long = 25&
Public Const PORT_FTP As Long = 21&
Public Const PORT_NULL As Long = 0&
Public Const R_HOST_NULL As String = vbNullString
Public Const ALL_DATA As Long = -1&
Public Const MINUTE As Long = 60&
'Windows API Stuff
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Public Enum SocketStatus
    scClosed = 0
    scConnecting = 1
    scError = 2
    scConnected = 3
    scListening = 4
    scNotListening = 5
End Enum


