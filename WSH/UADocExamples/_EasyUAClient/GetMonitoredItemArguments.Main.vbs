Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to obtain parameters of certain monitored item subscription.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

' Instantiate the client object and hook events
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")
WScript.ConnectObject Client, "Client_"

WScript.Echo "Subscribing..."

Dim MonitoredItemArguments1: Set MonitoredItemArguments1 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.EasyUAMonitoredItemArguments")
MonitoredItemArguments1.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
MonitoredItemArguments1.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10845"

Dim MonitoredItemArguments2: Set MonitoredItemArguments2 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.EasyUAMonitoredItemArguments")
MonitoredItemArguments2.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
MonitoredItemArguments2.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10853"

Dim MonitoredItemArguments3: Set MonitoredItemArguments3 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.EasyUAMonitoredItemArguments")
MonitoredItemArguments3.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
MonitoredItemArguments3.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10855"

Dim arguments(2)
Set arguments(0) = MonitoredItemArguments1
Set arguments(1) = MonitoredItemArguments2
Set arguments(2) = MonitoredItemArguments3

Dim handleArray: handleArray = Client.SubscribeMultipleMonitoredItems(arguments)

WScript.Echo "Getting monitored item arguments..."
Dim MonitoredItemArguments: Set MonitoredItemArguments = Client.GetMonitoredItemArguments(handleArray(2))

WScript.Echo "NodeDescriptor: " & MonitoredItemArguments.NodeDescriptor
WScript.Echo "SamplingInterval: " & MonitoredItemArguments.MonitoringParameters.SamplingInterval
WScript.Echo "PublishingInterval: " & MonitoredItemArguments.SubscriptionParameters.PublishingInterval

WScript.Echo "Waiting for 5 seconds..."
WScript.Sleep 5 * 1000

WScript.Echo "Unsubscribing..."
Client.UnsubscribeAllMonitoredItems

WScript.Echo "Waiting for 5 seconds..."
WScript.Sleep 5 * 1000



Sub Client_DataChangeNotification(Sender, e)
    ' Your code would do the processing here
End Sub

Rem#endregion Example
