<div align="center">

## Winsock Programming


</div>

### Description

This article is meant to explain how to utilize Windows sockets for network transfers in a Visual Basic program. This example code and the explanation were written and tested in Visual Basic 6.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-05-09 08:51:30
**By**             |[Jason Beighel](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jason-beighel.md)
**Level**          |Beginner
**User Rating**    |4.0 (12 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Winsock\_Pr81146592002\.zip](https://github.com/Planet-Source-Code/jason-beighel-winsock-programming__1-34614/archive/master.zip)





### Source Code

<font face="arial, helvetica, sans serif" size="2" color="#004080">
Through Winsock you can do TCP/IP socket connections using the Microsoft Winsock control. This control is the interface between your Visual Basic program and the networking hardware. This component is not used by default and you will need to add it to the toolbar. To do this right-click on the objects toolbar and select components from the dropdown menu. Scroll through the list until you find Microsoft Winsock Control, put a check in the box next to it and click the Apply button. Place one or more of these controls on your form to use them.<BR><Br>
Now before we get into the actual code to use Winsock here is a quick list of what happens during a session between a client and the server. For the purposes of this tutorial I'm calling the computer that waits and accepts the connection the server and the computer that makes the connection request the client (The computer hosting www.planetsourcecode.com is the server your computer contacting it with a web browser is the client).<BR><Br>
1) The Server selects a port and begins to listen on it for connections. There are 65,535 ports available for each IP address the computer has. I recommend using the ports from 1,024 to 65,535 because the others are reserved for specific uses (HTTP traffic uses port 80, FTP uses port 21, etc).<BR><BR>
2) The client makes a connection request to that port on the server. This is the only place where the distinction between the server and the client can be made. After the connection is established both computers are essentially on equal ground as far as abilities.<BR><BR>
3) The server accepts the connection request. After this step the connection is completed. A protocol or some standard should be used to coordinate the communication between the two computers. Using Winsock will not automatically set this for you, you will be responsible for this.<BR><BR>
4) The client or server waits for incoming data while the other sends data. Through the connection only one computer may send data at a time. The way the winsock control handles the data flow this potential problem is masked so you don't have to be overly concerned with it.<BR><BR>
5) The client or server close the connection. Either may close the connection but it is important for both sides to acknowledge that the connection is closed.<BR><BR>
In order to do this you need to create a server program and a client program. We'll start with the server and step 1, listening on a port. In my server program I am going to call the Winsock object wskIn, to mean its a winsock object intended for incoming connections. In order to accept the incoming connections you need to set the Index property to some value, zero is recommended. When the server app gets an incoming connection it will need to assign a winsock object to handle that connection, and the winsock object that is listening for connections should remain listening its role shouldn't change otherwise future connection attempts will be ignored. Don't be too concerned about that I'll explain it more later, for now just trust me and set the Index to 0. Also you will need to set the LocalPort property to the port number you want to listen on. When picking the port to listen on as I said before that the lower ports (0 to 1,024) are reserved for specific uses, so you should pick one of the higher numbers (1,024 to 65,535), the only real restriction is that the port can not already be in use. Once all these two values have been assigned you call the Listen method of the winsock control. After this paragraph of text we now have two lines of code, its gonna be a long tutorial.<BR><BR>
'Remember I am calling my winsock control wskIn<BR>
'wskIn.Index must be set when the control is placed on the form<BR>
'the sub's for it will not be created correctly otherwise<BR>
'Since I set the index to zero we have to include that in every reference to the control<BR>
wskIn(0).LocalPort = 2000 'Remember any number from 0 to 65,535 will work<BR>
wskIn(0).Listen<BR><BR>
Now we have a server that is waiting for incoming connections. Let's make a client that will request a connection. In my client program I am going to call the winsock control wskOut, to mean its a winsock object intended for outgoing connections. On this control we only need to call the Connect method. This method takes two parameters a String that has the address of the computer to connect to and and Integer that has the port number to connect to.<BR><BR>
'Remember I am calling my winsock control wskOut<BR>
wskOut.Connect "MyServer", 2000 'The address and port must match those of the server<BR><BR>
Now the server has to accept the incoming connection. When a connection request gets to the server the listening winsock control (wskIn(0)) will get a ConnectionRequest event and its associated sub will be called. So in that sub we need to accept the connection. As I was saying before is that we can't change the job of the winsock control that is listening for connections. Otherwise we will not receive any new connections and we may lose the current one. So what I suggest is dynamically creating a new winsock control to handle the incoming connections. This is the real reason I wanted the index set to zero, as I Load these new winsock controls I'm going to establish an array out of wskIn. In my examples I am not keeping track of the current loaded winsock controls, just a heads up for those who are going to copy and paste. Once you have your new winsock control to handle the connection you needs to be sure that is is in the closed state, there is no reason it shouldn't be after being created but just to be safe. After that you call the Accept method to actually take the connection. The Accept method uses only one parameter which is the requestID passed to this sub. This requestID is an identifier for this particular connection, this is the only place you will need it so there is no reason to save it. Here is how that will look:<BR><BR>
Private Sub wskIn_ConnectionRequest(Index As Integer, ByVal requestID As Long)<BR>
Load wskIn(1)<BR><BR>
if wsKin(1).State <> sckClosed then<BR>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;wskIn(1).Close<BR>
end if<BR><BR>
wskIn(1).Accept requestID<BR>
End Sub<BR><BR>
Once all this has been completed the winsock control that initiated the connection (wskOut from the client) and the winsock control that accepted the connection (wskIn(1) from the server) will receive Connect events and their associated subs will be called. This is just to let you know that the connection is established and available for use. Now we can begin to send data between the server and client. In the example I chose to send the data in the sub for the Connect event, this isn't required but you must wait until this event comes through before you can send data, you don't need to process the event it just has to occur.<BR><BR>
Now that we have a completed connection we can begin to send data between the computers. In order to send data you only need to use the SendData method of the winsock control. This method takes only a single parameter which is the data to be sent. Once you make the request to send data two events will occur for the winsock control, first the SendProgress event and then the SendComplete event. As with the connection you need to wait until these events occur before you can send more data, you don't need to process them however.<BR><BR>
wskOut.SendData "This will be sent"<BR><BR>
On the receiving end the winsock control that is receiving the data will get a DataArrival event. To retreive the data you must call the GetData method. This method also requires only one parameter which is the variable to store the data in. I was never successful in attempting to retreive data outside of the DataArrival sub but it may be possible. Also if the State parameter of the winsock control is not sckConnected then it will fail on the attempt to receive data. I don't know how you can get a data arrival event after the connection is closed but it seemed to be happening so in the example I am checking for that possibility.<BR><BR>
Private Sub wskIn_DataArrival(Index As Integer, ByVal bytesTotal As Long)<BR>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Dim strData As String<BR><BR>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;If (wskIn(Index).State <> sckConnected) Then<BR>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Exit Sub<BR>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;End If<BR><BR>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;wskIn(Index).GetData strData<BR>
End Sub<BR><BR>
This process can be repeated as much as necessary for any data that you want to send. After you have completed your transfers you need to close the connection. You use the Close method of the winsock control for this. This method does not require any parameters.<BR><BR>
wskIn.Close<BR><BR>
Either end of the connection (server or client) can close the connection and then both sides will receive a Close event to inform you that the connection has been broken.<BR><BR>
There is more to the winsock control but that ought to be enough to get you started. I've enclosed a sample program which is a crude chat application that shows all of this in action.</font>

