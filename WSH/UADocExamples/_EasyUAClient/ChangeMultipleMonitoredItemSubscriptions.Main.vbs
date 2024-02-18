Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how change the sampling rate of multiple existing monitored item subscriptions.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

' Instantiate the client object and hook events
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")
WScript.ConnectObject Client, "Client_"

WScript.Echo "Subscribing..."
Dim OldMonitoringParameters: Set OldMonitoringParameters = CreateObject("OpcLabs.EasyOpc.UA.UAMonitoringParameters")
OldMonitoringParameters.SamplingInterval = 1000
Dim MonitoredItemArguments1: Set MonitoredItemArguments1 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.EasyUAMonitoredItemArguments")
MonitoredItemArguments1.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
MonitoredItemArguments1.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10845"
MonitoredItemArguments1.MonitoringParameters = OldMonitoringParameters
Dim MonitoredItemArguments2: Set MonitoredItemArguments2 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.EasyUAMonitoredItemArguments")
MonitoredItemArguments2.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
MonitoredItemArguments2.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10853"
MonitoredItemArguments2.MonitoringParameters = OldMonitoringParameters
Dim MonitoredItemArguments3: Set MonitoredItemArguments3 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.EasyUAMonitoredItemArguments")
MonitoredItemArguments3.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
MonitoredItemArguments3.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10855"
MonitoredItemArguments3.MonitoringParameters = OldMonitoringParameters
Dim arguments(2)
Set arguments(0) = MonitoredItemArguments1
Set arguments(1) = MonitoredItemArguments2
Set arguments(2) = MonitoredItemArguments3
Dim handleArray: handleArray = Client.SubscribeMultipleMonitoredItems(arguments)

Dim i: For i = LBound(handleArray) To UBound(handleArray)
    WScript.Echo "handleArray(" & i & "): " & handleArray(i)
Next

WScript.Echo "Processing monitored item changed events for 10 seconds..."
WScript.Sleep 10 * 1000

WScript.Echo "Changing subscriptions..."
Dim NewMonitoringParameters: Set NewMonitoringParameters = CreateObject("OpcLabs.EasyOpc.UA.UAMonitoringParameters")
NewMonitoringParameters.SamplingInterval = 100
Dim SubscriptionChangeArguments1: Set SubscriptionChangeArguments1 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.EasyUASubscriptionChangeArguments")
SubscriptionChangeArguments1.Handle = handleArray(0)
Set SubscriptionChangeArguments1.MonitoringParameters = NewMonitoringParameters
Dim SubscriptionChangeArguments2: Set SubscriptionChangeArguments2 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.EasyUASubscriptionChangeArguments")
SubscriptionChangeArguments2.Handle = handleArray(1)
Set SubscriptionChangeArguments2.MonitoringParameters = NewMonitoringParameters
Dim SubscriptionChangeArguments3: Set SubscriptionChangeArguments3 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.EasyUASubscriptionChangeArguments")
SubscriptionChangeArguments3.Handle = handleArray(2)
Set SubscriptionChangeArguments3.MonitoringParameters = NewMonitoringParameters
Dim subscriptionChangeArguments(2)
Set subscriptionChangeArguments(0) = SubscriptionChangeArguments1
Set subscriptionChangeArguments(1) = SubscriptionChangeArguments2
Set subscriptionChangeArguments(2) = SubscriptionChangeArguments3
Client.ChangeMultipleMonitoredItemSubscriptions subscriptionChangeArguments

WScript.Echo "Processing monitored item changed events for 10 seconds..."
WScript.Sleep 10 * 1000

WScript.Echo "Unsubscribing..."
Client.UnsubscribeAllMonitoredItems

WScript.Echo "Waiting for 5 seconds..."
WScript.Sleep 5 * 1000



Sub Client_DataChangeNotification(Sender, e)
    ' Display the data
    Dim display: If e.Exception Is Nothing Then display = e.AttributeData Else display = e.ErrorMessageBrief
	WScript.Echo e.Arguments.NodeDescriptor & ":" & display
End Sub

Rem#endregion Example
