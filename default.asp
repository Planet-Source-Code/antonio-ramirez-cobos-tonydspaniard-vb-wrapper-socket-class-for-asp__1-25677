
<HTML>
<form method="POST">
  <table border="0" width="100%" bgcolor="#0099FF" cellspacing="1" height="167">
    <tr>
      <td width="24%" bgcolor="#FFFFFF" colspan="2" valign="middle" align="center" height="31">
<H1><font size="5" face="Verdana,Arial" color="#0000FF">ASP Socket Test</font></H1>
 </td>
      <td width="80%" valign="top" align="left" bgcolor="#FFFFFF" rowspan="6" height="163">
      <font size="1" face="arial,helvetica"><b>Results<br></b>
<%
'Get form contents
Host=request.form("Host")
Port=request.form("Port")
Data=request.form("Data")

if Host<>"" and Port<>"" and Data<>"" then

Dim ASPSocket 
'Create an instance of the object socket
set ASPSocket = server.CreateObject("Sockets.ASPSocket")

ASPSocket.RemoteHost = cstr(Host)
ASPSocket.RemotePort = clng(Port)

' Attempt to connect. Connect method
' Param 1: Remote Host
' Param 2: Remote Port
' Param 3: Time in Seconds to wait for an answer
' ----------------------------------------------
' You can also use it without any arguments
' but you should set its remote host and remote port
' otherwise they will to the following values:
' Remoteport: 80
' RemoteHost: vbNullString
' TimeInSeconds: 60 sec. 
' As you can see, you *MUST* set at list the remote host
' previous to call the Connect method or when calling the method
' otherwise the connection will fail
' -----------------------------------------------
on error resume next 
ASPSocket.Connect
if err.number <> 0 then 
	Response.Write err.description 
	on error goto 0
else
	' Just trying a few properties
	Response.Write "Local IP: " & ASPSocket.LocalIP & "<BR>"
	Response.Write "Remote Host IP: " & ASPSocket.RemoteHostIP & "<BR>"
	Response.Write "Local Host name: " & ASPSocket.LocalHostName & "<BR>"
	
	' Send your data to the Server
	' Send: 
	' Param 1: Data to send
	ASPSocket.Send cstr(Data)
	'ASPSocket.Do_Events
	if err.number <> 0 then 
		Response.Write err.description 
	else
		do while  ASPSocket.ServerReply 
			' Get everything in a shot
			Response.Write  ASPSocket.GetData()
			ASPSocket.Do_Events 
		loop
	end if
end if
'--- Set socket to nothing, [good programming practice ;-)]
set ASPSocket=nothing
end if
%>
</font></td>
    </tr>
    <tr>
      <td width="24%" bgcolor="#FFCC66" colspan="2" height="21">&nbsp; </td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF" height="25">
<font size="1" face="arial,helvetica">
Data:
</font>
 </td>
<font size="1" face="arial,helvetica">
      <td bgcolor="#FFCC66" height="25"><input type="text" name="Data" size="20" value="Your data"></td>
    </tr>
</font>
    <tr>
      <td nowrap bgcolor="#FFFFFF" height="25">
<font size="1" face="arial,helvetica">
Remote Host [Name or IP]:
</font>
 </td>
<font size="1" face="arial,helvetica">
      <td bgcolor="#FFCC66" height="25"><input type="text" name="Host" size="20" value="localhost"></td>
    </tr>
</font>
    <tr>
      <td bgcolor="#FFFFFF" height="25">
<font size="1" face="arial,helvetica">
Remote Port:
</font>
 </td>
<font size="1" face="arial,helvetica">
      <td bgcolor="#FFCC66" height="25"><input type="text" name="Port" size="20" value="2000"></td>
    </tr>
    <tr>
      <td bgcolor="#FFCC66" colspan="2" height="16">
        <p align="right"><font size="1" face="arial,helvetica"><a href="mailto:TonyDSpaniard@hotmail.com">By
        Antonio Ramirez Cobos</a></font></p>
 </td>
    </tr>
  </table>
  <p><input type="submit" value="Submit"><input type="reset" value="Reset"></p>
</form>
</font>
</HTML>
