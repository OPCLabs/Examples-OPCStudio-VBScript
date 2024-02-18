Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to change the sampling rate of an existing monitored item subscription.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

' Instantiate the client object and hook events
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")
WScript.ConnectObject Client, "Client_"

WScript.Echo "Subscribing..."
Dim handle: handle = Client.SubscribeDataChange( _
    "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer", _
    "nsu=http://test.org/UA/Data/ ;i=10853", _
    1000)

WScript.Echo "Processing monitored item changed events for 10 seconds..."
WScript.Sleep 10 * 1000

WScript.Echo "Changing subscription..."
Client.ChangeMonitoredItemSubscription handle, 100

WScript.Echo "Processing monitored item changed events for 10 seconds..."
WScript.Sleep 10 * 1000

WScript.Echo "Unsubscribing..."
Client.UnsubscribeAllMonitoredItems

WScript.Echo "Waiting for 5 seconds..."
WScript.Sleep 5 * 1000



Sub Client_DataChangeNotification(Sender, e)
    Dim display: If e.Exception Is Nothing Then display = e.AttributeData Else display = e.ErrorMessageBrief
	WScript.Echo display
End Sub

Rem#endregion Example
