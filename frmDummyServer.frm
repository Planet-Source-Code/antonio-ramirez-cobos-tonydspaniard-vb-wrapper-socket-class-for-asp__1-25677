VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmDummyServer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test Server"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Default         =   -1  'True
      Height          =   495
      Left            =   3015
      TabIndex        =   1
      Top             =   2550
      Width           =   1035
   End
   Begin MSWinsockLib.Winsock wsSSocket 
      Index           =   0
      Left            =   120
      Top             =   2505
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox lstMessages 
      Height          =   2400
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   4005
   End
End
Attribute VB_Name = "frmDummyServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'******************************************************'
'------------------------------------------------------'
' Project: ASPSocket
'
' Module: frmDummyServer
'
' Date: July-28-2001
'
' Programmer: Antonio Ramirez Cobos
'
' Description: Created to test ASPSocket object
'              performance and functionality
'------------------------------------------------------'
'******************************************************'
Private m_SocketIsInUse() As Boolean
Private m_Index As Long

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
'
' make socket a listener socket on port 2000
    wsSSocket(0).LocalPort = 2000
    wsSSocket(0).Listen
    ReDim m_SocketIsInUse(0) As Boolean
End Sub

Private Sub Form_Unload(Cancel As Integer)
'
' Make sure we don't left any loaded sockets
    Dim intJ As Long
    On Error Resume Next
    For intJ = 1 To m_Index
        If m_SocketIsInUse(intJ) Then
            wsSSocket(intJ).Close
            DoEvents
            Unload wsSSocket(intJ)
        End If
    Next
End Sub

Private Sub wsSSocket_Close(Index As Integer)
'
' A client has close connection, unload
' correpondent socket
    Unload wsSSocket(Index)
    m_SocketIsInUse(Index) = False
    lstMessages.AddItem "Client Socket disconnected at " & Time
End Sub

Private Sub wsSSocket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
'
' Another socket on port 2000 is requesting for connection
' we assume [be careful with assumptions on the Net] that
' is our little ASPSocket object calling from default.asp
' page
    Dim lngJ As Long ' Counter
    If Index = 0 Then
        ' Loop through m_SocketIsInuse array
        ' to see if there is any previous index free to use
        For lngJ = 1 To m_Index
            If Not m_SocketIsInUse(lngJ) Then
                Exit For
            End If
        Next
        ' If counter is higher that the current loaded
        ' sockets, then update index and flag array
        If lngJ > m_Index Then
            m_Index = m_Index + 1
            ReDim Preserve m_SocketIsInUse(m_Index)
        End If
        ' Load socket and accept request
        Load wsSSocket(lngJ)
        wsSSocket(lngJ).LocalPort = 0
        wsSSocket(lngJ).Accept requestID
        ' Display message
        lstMessages.AddItem "Connection accepted at " & Time
        m_SocketIsInUse(lngJ) = True
    End If
End Sub

Private Sub wsSSocket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'
'
' Client Socket sent some data.
' Add to message list and Echo back a message to socket
' Note: Use your imagination, can you see the possibilities
'       from an ASP application? Of course... but if you
'       think about something very, very, very interesting
'       please let me know [TONYDSPANIARD@HOTMAIL.COM]
'       My next step is to create another one by using
'       API
Dim strData As String
    wsSSocket(Index).GetData strData, vbString, bytesTotal
    lstMessages.AddItem "Client Socket data: " & strData
    wsSSocket(Index).SendData "You send <b>" & strData & "</b> at server time " & Time
End Sub
