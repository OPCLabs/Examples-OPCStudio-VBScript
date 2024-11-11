Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to subscribe to changes of a monitored item with data change filter.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Const UADataChangeTrigger_StatusValue = 1

Dim endpointDescriptor: endpointDescriptor = _
    "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    '"http://opcua.demo-this.com:51211/UA/SampleServer"  
    '"https://opcua.demo-this.com:51212/UA/SampleServer/"

' Instantiate the client object and hook events
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")
WScript.ConnectObject Client, "Client_"

' Prepare the arguments.
' Report a notification if either the StatusCode or the value change. 
Dim DataChangeFilter: Set DataChangeFilter = CreateObject("OpcLabs.EasyOpc.UA.UADataChangeFilter")
DataChangeFilter.Trigger = UADataChangeTrigger_StatusValue
'
Dim MonitoringParameters: Set MonitoringParameters = CreateObject("OpcLabs.EasyOpc.UA.UAMonitoringParameters")
Set MonitoringParameters.DataChangeFilter = DataChangeFilter
MonitoringParameters.SamplingInterval = 1000
'
Dim MonitoredItemArguments1: Set MonitoredItemArguments1 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.EasyUAMonitoredItemArguments")
MonitoredItemArguments1.EndpointDescriptor.UrlString = endpointDescriptor
MonitoredItemArguments1.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10853"
MonitoredItemArguments1.MonitoringParameters = MonitoringParameters
'
Dim arguments(0)
Set arguments(0) = MonitoredItemArguments1

WScript.Echo "Subscribing..."
Client.SubscribeMultipleMonitoredItems arguments

WScript.Echo "Processing monitored item changed events for 20 seconds..."
WScript.Sleep 20*1000

WScript.Echo "Unsubscribing..."
Client.UnsubscribeAllMonitoredItems

WScript.Echo "Waiting for 5 seconds..."
WScript.Sleep 5 * 1000



Sub Client_DataChangeNotification(Sender, e)
    ' Display value
    Dim display: If e.Exception Is Nothing Then display = e.AttributeData Else display = "*** Failure: " & e.ErrorMessageBrief
	WScript.Echo display
End Sub

Rem#endregion Example
