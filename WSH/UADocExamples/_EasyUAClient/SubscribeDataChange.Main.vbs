Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to subscribe to changes of a single monitored item and display each change.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

' Instantiate the client object and hook events
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")
WScript.ConnectObject Client, "Client_"

WScript.Echo "Subscribing..."
Client.SubscribeDataChange "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer", "nsu=http://test.org/UA/Data/ ;i=10853", 1000

WScript.Echo "Processing monitored item changed events for 1 minute..."
WScript.Sleep 60*1000



Sub Client_DataChangeNotification(Sender, e)
    ' Display the data
    Dim display: If e.Exception Is Nothing Then display = e.AttributeData Else display = e.ErrorMessageBrief
	WScript.Echo display
End Sub

Rem#endregion Example
