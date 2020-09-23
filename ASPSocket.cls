VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ASPSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'******************************************************'
'------------------------------------------------------'
' Project: ASPSocket
'
' Date: July-28-2001
'
' Programmer: Antonio Ramirez Cobos
'
' Description: Allows the implementation of a client socket
'              to VB and ASP applications. It was designed
'              with ASP programmers in mind, as other type
'              of applications are better to be programmed by
'              using other technologies such as Windows API
'              for Sockets in order to improve VB to the limit.
'              Nevertheless, I have included some event functionality
'              for those wishing to test the object within VB
'
'              In addition, 'cause I consider myself in a continuous learning
'              path with no end on programming, please, if you
'              can improve this object [without using Sockets API]
'              contact me at: *TONYDSPANIARD@HOTMAIL.COM*
'
'              I would be pleased to hear from your opinions,
'              suggestions, and/or recommendations. Also, if you
'              know something I don't know and wish to share it
'              with me, here you'll have your techy pal from Spain
'              that will do exactly the same towards you.
'              Let's build a community of programmers helping
'              programmers without the lucrative desire of fame
'              on the back, let's just do it for *KNOWLEDGE*
'              because that is real power.
'
'              INTELLECTUAL COPYRIGHT STUFF [Is up to you anyway]
'              This code is copyright 2001 Antonio Ramirez Cobos
'              This code may be reused and modified for non-commercial
'              purposes only as long as credit is given to the author
'              in the programmes about box and it's documentation.
'              If you use this code, please email me at:
'              TonyDSpaniard@hotmail.com and let me know what you think
'              and what you are doing with it.
'------------------------------------------------------'
'******************************************************'

'-------- Private Properties
Private m_propRemoteHost As String  '--- Host to connect to
Private m_propRemotePort As Long '------ Port to connect to
Private m_propLocalPort As Long '------- Local port [used mainly with SMTP]
Private m_propReply As Boolean '-------- True if server sent data back
Private m_propStatus As Sockets.SocketStatus '-- Socket status
Private m_bytBuffer() As Byte '--------- Holds server sent data
'------------------------------------------------------'
'******************************************************'

'--------- The *WRAPPED* socket
Private WithEvents Socket As Winsock
Attribute Socket.VB_VarHelpID = -1
'------------------------------------------------------'
'******************************************************'

'--------- Events
Event DataArrival(TotalBytes As Long) '--- Occurs when server sends data back
Event Disconnected() '---- Occurs when server [close] connection
Event Connected() '------- Occurs on connection
'------------------------------------------------------'
'******************************************************'


'******************************************************'
'------------------------------------------------------'
' Properties
'
' RemoteHost: Host to connect to
'
Public Property Get RemoteHost() As String
    RemoteHost = m_propRemoteHost
End Property
Public Property Let RemoteHost(Host As String)
    m_propRemoteHost = Host
End Property
'
'
' RemotePort: Port to connect to
Public Property Get RemotePort() As Long
    RemotePort = m_propRemotePort
End Property
Public Property Let RemotePort(Port As Long)
    m_propRemotePort = Port
End Property
'
'
' LocalPort: Local Port
Public Property Get LocalPort() As Long
    LocalPort = m_propLocalPort
End Property
Public Property Let LocalPort(Port As Long)
    m_propLocalPort = Port
End Property
'
'
' BytesReceived: Total Bytes on buffer received by the server
Public Property Get BytesReceived() As Long
    BytesReceived = UBound(m_bytBuffer)
End Property
'
'
' Connected: True if connected, false otherwise
Public Property Get Connected() As Boolean
    Connected = (m_propStatus = scConnected)
End Property
'
'
' LocalIP: Local IP address [Web Server were DLL sits]
Public Property Get LocalIP() As String
    LocalIP = Socket.LocalIP
End Property
'
'
' LocalHostName: Computer name [Web Server were DLL sits]
Public Property Get LocalHostName() As String
    LocalHostName = Socket.LocalHostName
End Property
'
'
' RemoteHostIP: Remote Host IP address [Server we are contacting]
Public Property Get RemoteHostIP() As String
    If m_propStatus <> scConnected Then Exit Property
    RemoteHostIP = Socket.RemoteHostIP
End Property
'
'
' SocketHandle: Included here for completeness but not very
'               usefull for ASP. It is maily used for API
Public Property Get SocketHandle() As Long
    'This property was designed to be passed to Winsock APIs
    SocketHandle = Socket.SocketHandle
End Property
'
'
' ServerReply: True if data has been sent from the server we're
'              connected
Public Property Get ServerReply() As Boolean
    ServerReply = m_propReply
End Property

'******************************************************'
'------------------------------------------------------'
' Methods
'
' Connect: Tryes to connect to specified remote host and remote port
'          during a specified Time interval [60 sec default]
'
' Params: RHost=Remote Host;
'         RPort= Remote Port;
'         TimeOutInSec=Max. Interval of sec. to try
Public Sub Connect(Optional RHost As String = R_HOST_NULL, Optional RPort As Long = PORT_NULL, Optional TimeOutInSec As Long = Minute)

    Dim Interval As Variant
    
    ' When using the object on an ASP page, the parameters are passed by
    ' value, therefore the IsMissing statement is failing to return
    ' appropriately. A solution is the one presented: setting parameters
    ' to default values and check if the parameters hold different ones.
    ' If the values holded by the parameters are the default ones, we
    ' know the parameters had been omitted.
    If Not RHost = R_HOST_NULL Then m_propRemoteHost = RHost
    If Not RPort = PORT_NULL Then m_propRemotePort = RPort
    
    ' Make sure we are disconnected
    If m_propStatus = scConnected Then
        Disconnect
    End If

    If m_propRemoteHost = "" Then Err.Raise vbObjectError, "ClientSocket:Connect()", "Host must be specified to connect"
    
    ' Check if we are going to send messages through
    ' SMTP port
    If m_propRemotePort = PORT_SMTP Then
        '-- We must set local port to zero otherwise we'll not be
        '   allowed to send more than one message -email-
        If m_propLocalPort <> PORT_NULL Then m_propLocalPort = PORT_NULL
        Socket.LocalPort = PORT_NULL
    End If
    
    ' Try to connect to the server
    m_propStatus = scConnecting
    Socket.Connect m_propRemoteHost, m_propRemotePort
    
   Interval = Time + TimeOutInSec
   
   Do While m_propStatus = scConnecting And Interval > Time
      '-- Loop for maximum interval of a minute to get an answer
      '-- from the server
      DoEvents
   Loop
    
    ' If we didn't connect, raise an error
    If m_propStatus <> scConnected Then
        Socket.Close
        DoEvents
        If Err.Number <> 0 Then Err.Clear
        m_propStatus = scClosed
        Err.Raise vbObjectError, "ClientSocket.Connect()", "Couldn't connect"
    End If
End Sub
'
'
' Disconnect: Close connection to remote host
Public Sub Disconnect()
    ' Close the socket
    If m_propStatus <> scClosed Then
        On Error Resume Next
        Socket.Close
        DoEvents
    End If
    ' Update flag
    m_propStatus = scClosed
    
End Sub
'
'
' Send: Sends data string to remote host
Public Sub Send(Text As String)
    Socket.SendData Text
End Sub
'
'
' SendBinary: Sends binary data to the server
'             Note: Do not use with ASP [sure crash even if
'                   the parameter is variant data type!]
'Public Sub SendBinary(BytArray() As Byte)
'    Socket.SendData BytArray
'End Sub
'
'
' GetData: Extracts and returns specified number of bytes as a string
'          data type from the buffer if any data was sent from remote host
'
' Params: NumberOfBytes=Number of bytes to extract
Public Function GetData(Optional NumberOfBytes As Long = ALL_DATA) As String

    Dim TotalBytes As Long, Contents As String, intJ As Integer
    
    '------ how many bytes requested?
    If NumberOfBytes <> ALL_DATA Then
        TotalBytes = NumberOfBytes
        If TotalBytes > UBound(m_bytBuffer) Then TotalBytes = UBound(m_bytBuffer)
    Else
        '--- none, return the whole lot
        TotalBytes = UBound(m_bytBuffer)
    End If
    '--- If there is nothing to return get out of here
    If TotalBytes < 1 Then
        m_propReply = False
        Exit Function
    End If
    
    '--- Allocate space
    Contents = String(TotalBytes + 1, " ")
    
    '--- Get the bytes
    '--- [pretty cool way to copy from a byte array to a string isn't it?]
    CopyMemory ByVal Contents, m_bytBuffer(0), TotalBytes + 1
    
    '--- Return data extracted
    GetData = Contents
    
    '-- Move data to the beginning of the array [if any]
    '   and resize array
    If TotalBytes = UBound(m_bytBuffer) Then
        '-- We read everything therefore
        '   resize to zero
        m_propReply = False ' End of Server Reply
        ReDim m_bytBuffer(0)
    Else
        CopyMemory m_bytBuffer(0), m_bytBuffer(TotalBytes + 1), UBound(m_bytBuffer) - TotalBytes
        ReDim Preserve m_bytBuffer(UBound(m_bytBuffer) - (TotalBytes + 1))
    End If
End Function
'
'
' GetBinaryData: Fills ByteArray parameter with NumberOfBytes requested. If
'                NumberOfBytes hasn't been set, then ByteArray will be filled
'                filled with all data stored on the buffer.
'                Note: Do not use with ASP [sure crash even if
'                      the parameter is variant data type!]
'Public Sub GetBinaryData(ByteArray() As Byte, Optional NumberOfBytes As Long = ALL_DATA)
'
'    Dim TotalBytes As Long, intJ As Integer
'
'    '-- Find out bytes requested
'    If NumberOfBytes = ALL_DATA Then
'        TotalBytes = UBound(m_bytBuffer)
'    Else
'        TotalBytes = NumberOfBytes
'        ' Reduce TotalBytes if higher than actual buffer size
'        If TotalBytes > UBound(m_bytBuffer) Then TotalBytes = UBound(m_bytBuffer)
'    End If
'    '-- There is no bytes, then get out of here
'    If TotalBytes < 1 Then
'        m_propReply = False
'        Exit Sub
'    End If
'
'    '-- Allocate space
'    ReDim ByteArray(TotalBytes)
'
'    '---- Copy bytes accross
'    CopyMemory ByteArray(0), m_bytBuffer(0), TotalBytes + 1
'
'    '-- Now resize buffer accordingly to the bytes extracted from it
'    If TotalBytes = UBound(m_bytBuffer) Then
'        m_propReply = False ' End of server reply
'        ReDim m_bytBuffer(0)
'    Else
'        '--- Move non-extracted bytes to the front of the buffer
'        '    and resize buffer
'        CopyMemory m_bytBuffer(0), m_bytBuffer(TotalBytes + 1), UBound(m_bytBuffer) - (TotalBytes)
'        ReDim Preserve m_bytBuffer(UBound(m_bytBuffer) - (TotalBytes + 1))
'    End If
'
'End Sub
'
'
' Do_Events: Helping ASP applications to apply this useful
'            VB statement
Public Sub Do_Events()
    DoEvents
End Sub

'******************************************************'
'------------------------------------------------------'
' Object's main events
'
'
Private Sub Class_Initialize()
    Set Socket = New Winsock '[normal init.]
    m_propRemotePort = PORT_HTTP '-- HTTP port by default [80]
    m_propLocalPort = PORT_NULL '--- zero by default [random port selection]
    m_propStatus = scClosed '--------- Closed
    m_propReply = False '----------- Not server reply
    ReDim m_bytBuffer(0) '---------- Init. Buffer that will hold all the data
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    '--- If the socket hasn't been closed, then
    '    close it!
    If Socket.State <> sckClosed Then
        Socket.Close
        DoEvents
    End If
    Set Socket = Nothing
End Sub

' -------------------------------------------------------------- '
' *********************** WINSOCK EVENTS *********************** '
'
' Note: For more info about this events check MSDN library
'
' Socket_Close: Remote host closed the conncection
Private Sub Socket_Close()
    '-- server socket connection closed, close our connection
    '   and raise the event to inform user about it
    m_propStatus = scClosed
    Socket.Close
    DoEvents
    RaiseEvent Disconnected
End Sub
'
'
' Socket_Connect: Connection accepted
Private Sub Socket_Connect()
    '---- connection established, update flag and raise event
    m_propStatus = scConnected
    RaiseEvent Connected
End Sub
'
'
' Socket_DataArrival: Remote host sent some data
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    Dim bytTempBuffer() As Byte, intPos As Integer
    
    '---- Get data arrived
    '     Resize temp buffer to store new incoming bytes of data
    ReDim bytTempBuffer(bytesTotal)
    
    '---- Get the data as binary and store it in temporary
    '     byte array
    Socket.GetData bytTempBuffer, vbByte, bytesTotal
    
    '---- Add data at the buffer's first position *FREE* of data
    If UBound(m_bytBuffer) = 0 Then
        intPos = 0
    Else
        intPos = UBound(m_bytBuffer) + 1
    End If
    
    '---- Now resize buffer holding the whole sent information [past-present]
    '     to hold new bytes
    ReDim Preserve m_bytBuffer(intPos + bytesTotal)
    '---- Copy in one shot [thanks to CopyMemory function]
    CopyMemory m_bytBuffer(intPos), bytTempBuffer(0), bytesTotal
    '--- We had a reply, Communicate this to the user
    m_propReply = True
    '---- RAISE EVENT [If object is been used on a VB application, absurd
    '     as we have Windows API and Winsock control to do the job, but anyway is here]
    RaiseEvent DataArrival(UBound(m_bytBuffer))
End Sub
'
'
' Socket_Error: [Needs definition?]
Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '--- If we are not trying to connect
    '    raise an error returning the error number & description
    If m_propStatus <> scConnecting Then
        m_propStatus = scError
        '--- Raise the event [*POINTLESS* for an ASP Application]
        RaiseEvent Disconnected
        If m_propStatus = scConnected Then Socket.Close
        Err.Raise Number, "ClientSocket:Error()", Description
    Else '- Connecting
        '-- Update flag to stop looping on connection procedure
        m_propStatus = scError
    End If
End Sub
' ******************** END OF WINSOCK EVENTS *********************
