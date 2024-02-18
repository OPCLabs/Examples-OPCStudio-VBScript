Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to lock and unlock connections to an OPC UA server. The component attempts to keep the locked
Rem connections open, until unlocked.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Dim endpointDescriptorUrlString: endpointDescriptorUrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
Dim EndpointDescriptor: Set EndpointDescriptor = CreateObject("OpcLabs.EasyOpc.UA.UAEndpointDescriptor")
EndpointDescriptor.UrlString = endpointDescriptorUrlString

' Instantiate the client object.
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")

' Obtain the client connection monitoring service.
Dim ClientConnectionMonitoring: Set ClientConnectionMonitoring = Client.GetServiceByName("OpcLabs.EasyOpc.UA.Services.IEasyUAClientConnectionMonitoring, OpcLabs.EasyOpcUA")
If ClientConnectionMonitoring Is Nothing Then
    WScript.Echo "The client connection monitoring service is not available."
    WScript.Quit
End If

' Obtain the client connection control service.
Dim ClientConnectionControl: Set ClientConnectionControl = Client.GetServiceByName("OpcLabs.EasyOpc.UA.Services.IEasyUAClientConnectionControl, OpcLabs.EasyOpcUA")
If ClientConnectionControl Is Nothing Then
    WScript.Echo "The client connection control service is not available."
    WScript.Quit
End If

' Display the server condition changed events.
WScript.ConnectObject ClientConnectionMonitoring, "ClientConnectionMonitoring_"

WScript.Echo "Reading (1)"
' The first read will cause a connection to the server.
Dim AttributeData1: Set AttributeData1 = Client.Read(endpointDescriptorUrlString, "nsu=http://test.org/UA/Data/ ;i=10853")
WScript.Echo AttributeData1

WScript.Echo "Waiting for 10 seconds..."
' Since the connection is now not used for some time, and it is not locked, it will be closed.
WScript.Sleep 10*1000

WScript.Echo "Locking..."
' Locking the connection causes it to open, if possible.
Dim lockHandle: lockHandle = ClientConnectionControl.LockConnection(EndpointDescriptor)

WScript.Echo "Waiting for 10 seconds..."
' The connection is locked, it will not be closed now.
WScript.Sleep 10*1000

WScript.Echo "Reading (2)"
' The second read, because it closely follows the first one, will reuse the connection that is already open.
Dim AttributeData2: Set AttributeData2 = Client.Read(endpointDescriptorUrlString, "nsu=http://test.org/UA/Data/ ;i=10853")
WScript.Echo AttributeData2

WScript.Echo "Waiting for 10 seconds..."
' The connection is still locked, it will not be closed now.
WScript.Sleep 10*1000

WScript.Echo "Unlocking..."
ClientConnectionControl.UnlockConnection(lockHandle)

WScript.Echo "Waiting for 10 seconds..."
' After some delay, the connection will be closed, because there are no subscriptions to the server and no
' connection locks.
WScript.Sleep 10*1000

WScript.Echo "Finished."



Sub ClientConnectionMonitoring_ServerConditionChanged(Sender, e)
	WScript.Echo e
End Sub


' Example output:
'
'Reading (1)
'"opc.tcp://opcua.demo-this.com:51210/UA/SampleServer" Connecting; Success; Attempt #1
'"opc.tcp://opcua.demo-this.com:51210/UA/SampleServer" Connected; Success
'-1.034588E+18 {Single} @2021-11-15T15:26:39.169 @@2021-11-15T15:26:39.169; Good
'Waiting for 10 seconds...
'"opc.tcp://opcua.demo-this.com:51210/UA/SampleServer" Disconnecting; Success
'"opc.tcp://opcua.demo-this.com:51210/UA/SampleServer" Disconnected(RetrialDelay=Infinite); Success
'Locking
'"opc.tcp://opcua.demo-this.com:51210/UA/SampleServer" Connecting; Success; Attempt #1
'"opc.tcp://opcua.demo-this.com:51210/UA/SampleServer" Connected; Success
'Waiting for 10 seconds...
'Reading (2)
'2.288872E+21 {Single} @2021-11-15T15:26:59.836 @@2021-11-15T15:26:59.836; Good
'Waiting for 10 seconds...
'Unlocking
'Waiting for 10 seconds...
'"opc.tcp://opcua.demo-this.com:51210/UA/SampleServer" Disconnecting; Success
'"opc.tcp://opcua.demo-this.com:51210/UA/SampleServer" Disconnected(RetrialDelay=Infinite); Success
'Finished.

Rem#endregion Example
