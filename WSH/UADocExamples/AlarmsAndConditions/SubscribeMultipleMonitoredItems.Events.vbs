Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to subscribe to multiple events.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Const UAAttributeId_NodeId = 1
Const UAAttributeId_EventNotifier = 12

Const UAFilterOperator_Equals = 1
Const UAFilterOperator_GreaterThanOrEqual = 5

Dim endpointDescriptor
endpointDescriptor = "opc.tcp://opcua.demo-this.com:62544/Quickstarts/AlarmConditionServer"

' Instantiate the client object and hook events
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")
WScript.ConnectObject Client, "Client_"

Dim arguments(1)
Set arguments(0) = CreateMonitoredItemArguments1
Set arguments(1) = CreateMonitoredItemArguments2

WScript.Echo "Subscribing..."
Client.SubscribeMultipleMonitoredItems arguments

WScript.Echo "Processing monitored item changed events for 30 seconds..."
WScript.Sleep 30 * 1000

WScript.Echo "Unsubscribing..."
Client.UnsubscribeAllMonitoredItems

WScript.Echo "Waiting for 5 seconds..."
WScript.Sleep 5 * 1000



Function ObjectTypeIds_BaseEventType
    Dim NodeId: Set NodeId = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
    NodeId.StandardName = "BaseEventType"
    Set ObjectTypeIds_BaseEventType = NodeId
End Function

Function UAFilterElements_SimpleAttribute(TypeId, simpleRelativeBrowsePathString)
    Dim BrowsePathParser: Set BrowsePathParser = CreateObject("OpcLabs.EasyOpc.UA.Navigation.Parsing.UABrowsePathParser")
    Dim Operand: Set Operand = CreateObject("OpcLabs.EasyOpc.UA.Filtering.UASimpleAttributeOperand")
    Set Operand.TypeId.NodeId = TypeId
    Set Operand.QualifiedNames = BrowsePathParser.ParseRelative(simpleRelativeBrowsePathString).ToUAQualifiedNameCollection
    Set UAFilterElements_SimpleAttribute = Operand
End Function

Function UABaseEventObject_Operands_NodeId
    Dim Operand: Set Operand = CreateObject("OpcLabs.EasyOpc.UA.Filtering.UASimpleAttributeOperand")
    Operand.TypeId.NodeId.StandardName = "BaseEventType"
    Operand.AttributeId = UAAttributeId_NodeId
    Set UABaseEventObject_Operands_NodeId = Operand
End Function

Function UABaseEventObject_Operands_EventId
    Set UABaseEventObject_Operands_EventId = UAFilterElements_SimpleAttribute(ObjectTypeIds_BaseEventType, "/EventId")
End Function

Function UABaseEventObject_Operands_EventType
    Set UABaseEventObject_Operands_EventType = UAFilterElements_SimpleAttribute(ObjectTypeIds_BaseEventType, "/EventType")
End Function

Function UABaseEventObject_Operands_SourceNode
    Set UABaseEventObject_Operands_SourceNode = UAFilterElements_SimpleAttribute(ObjectTypeIds_BaseEventType, "/SourceNode")
End Function

Function UABaseEventObject_Operands_SourceName
    Set UABaseEventObject_Operands_SourceName = UAFilterElements_SimpleAttribute(ObjectTypeIds_BaseEventType, "/SourceName")
End Function

Function UABaseEventObject_Operands_Time
    Set UABaseEventObject_Operands_Time = UAFilterElements_SimpleAttribute(ObjectTypeIds_BaseEventType, "/Time")
End Function

Function UABaseEventObject_Operands_ReceiveTime
    Set UABaseEventObject_Operands_ReceiveTime = UAFilterElements_SimpleAttribute(ObjectTypeIds_BaseEventType, "/ReceiveTime")
End Function

Function UABaseEventObject_Operands_LocalTime
    Set UABaseEventObject_Operands_LocalTime = UAFilterElements_SimpleAttribute(ObjectTypeIds_BaseEventType, "/LocalTime")
End Function

Function UABaseEventObject_Operands_Message
    Set UABaseEventObject_Operands_Message = UAFilterElements_SimpleAttribute(ObjectTypeIds_BaseEventType, "/Message")
End Function

Function UABaseEventObject_Operands_Severity
    Set UABaseEventObject_Operands_Severity = UAFilterElements_SimpleAttribute(ObjectTypeIds_BaseEventType, "/Severity")
End Function

Function UABaseEventObject_AllFields
    Dim Fields: Set Fields = CreateObject("OpcLabs.EasyOpc.UA.UAAttributeFieldCollection")

    Fields.Add UABaseEventObject_Operands_NodeId.ToUAAttributeField

    Fields.Add UABaseEventObject_Operands_EventId.ToUAAttributeField
    Fields.Add UABaseEventObject_Operands_EventType.ToUAAttributeField
    Fields.Add UABaseEventObject_Operands_SourceNode.ToUAAttributeField
    Fields.Add UABaseEventObject_Operands_SourceName.ToUAAttributeField
    Fields.Add UABaseEventObject_Operands_Time.ToUAAttributeField
    Fields.Add UABaseEventObject_Operands_ReceiveTime.ToUAAttributeField
    Fields.Add UABaseEventObject_Operands_LocalTime.ToUAAttributeField
    Fields.Add UABaseEventObject_Operands_Message.ToUAAttributeField
    Fields.Add UABaseEventObject_Operands_Severity.ToUAAttributeField

    Set UABaseEventObject_AllFields = Fields
End Function

Function CreateMonitoredItemArguments1
    ' Event filter: The severity is >= 500.
    Dim Operand1: Set Operand1 = UABaseEventObject_Operands_Severity
    Dim Operand2: Set Operand2 = CreateObject("OpcLabs.EasyOpc.UA.Filtering.UALiteralOperand")
    Operand2.Value = 500
    Dim WhereClause: Set WhereClause = CreateObject("OpcLabs.EasyOpc.UA.Filtering.UAContentFilterElement")
    WhereClause.FilterOperator = UAFilterOperator_GreaterThanOrEqual
    WhereClause.FilterOperands.Add Operand1
    WhereClause.FilterOperands.Add Operand2

    Dim EventFilter: Set EventFilter = CreateObject("OpcLabs.EasyOpc.UA.UAEventFilter")
    Set EventFilter.SelectClauses = UABaseEventObject_AllFields
    Set EventFilter.WhereClause = WhereClause

    Dim ServerNodeId: Set ServerNodeId = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
    ServerNodeId.StandardName = "Server"

    Dim MonitoringParameters: Set MonitoringParameters = CreateObject("OpcLabs.EasyOpc.UA.UAMonitoringParameters")
    Set MonitoringParameters.EventFilter = EventFilter
    MonitoringParameters.QueueSize = 1000
    MonitoringParameters.SamplingInterval = 1000

    Dim MonitoredItemArguments: Set MonitoredItemArguments = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.EasyUAMonitoredItemArguments")
    MonitoredItemArguments.AttributeId = UAAttributeId_EventNotifier
    MonitoredItemArguments.EndpointDescriptor.UrlString = endpointDescriptor
    Set MonitoredItemArguments.MonitoringParameters = MonitoringParameters
    Set MonitoredItemArguments.NodeDescriptor.NodeId = ServerNodeId
    MonitoredItemArguments.State = "firstState"

    Set CreateMonitoredItemArguments1 = MonitoredItemArguments
End Function

Function CreateMonitoredItemArguments2
    ' Event filter: The event comes from a specified source node.
    Dim Operand1: Set Operand1 = UABaseEventObject_Operands_SourceNode
    Dim SourceNodeId: Set SourceNodeId = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
    SourceNodeId.ExpandedText = "nsu=http://opcfoundation.org/Quickstarts/AlarmCondition ;ns=2;s=1:Metals/SouthMotor"
    Dim Operand2: Set Operand2 = CreateObject("OpcLabs.EasyOpc.UA.Filtering.UALiteralOperand")
    Set Operand2.Value = SourceNodeId
    Dim WhereClause: Set WhereClause = CreateObject("OpcLabs.EasyOpc.UA.Filtering.UAContentFilterElement")
    WhereClause.FilterOperator = UAFilterOperator_Equals
    WhereClause.FilterOperands.Add Operand1
    WhereClause.FilterOperands.Add Operand2

    Dim EventFilter: Set EventFilter = CreateObject("OpcLabs.EasyOpc.UA.UAEventFilter")
    Set EventFilter.SelectClauses = UABaseEventObject_AllFields
    Set EventFilter.WhereClause = WhereClause

    Dim ServerNodeId: Set ServerNodeId = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
    ServerNodeId.StandardName = "Server"

    Dim MonitoringParameters: Set MonitoringParameters = CreateObject("OpcLabs.EasyOpc.UA.UAMonitoringParameters")
    Set MonitoringParameters.EventFilter = EventFilter
    MonitoringParameters.QueueSize = 1000
    MonitoringParameters.SamplingInterval = 2000

    Dim MonitoredItemArguments: Set MonitoredItemArguments = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.EasyUAMonitoredItemArguments")
    MonitoredItemArguments.AttributeId = UAAttributeId_EventNotifier
    MonitoredItemArguments.EndpointDescriptor.UrlString = endpointDescriptor
    Set MonitoredItemArguments.MonitoringParameters = MonitoringParameters
    Set MonitoredItemArguments.NodeDescriptor.NodeId = ServerNodeId
    MonitoredItemArguments.State = "secondState"

    Set CreateMonitoredItemArguments2 = MonitoredItemArguments
End Function

Sub Client_EventNotification(Sender, e)
    ' Display the event
	WScript.Echo e
End Sub



' Example output (truncated):
'Subscribing...
'Processing monitored item changed events for 30 seconds...
'[firstState] Success
'[secondState] Success
'[firstState] Success; Refresh; RefreshInitiated
'[firstState] Success; Refresh; (10 field results) [EastTank] 500! "The alarm was acknoweledged." @10/14/2019 4:00:13 PM
'[firstState] Success; Refresh; (10 field results) [EastTank] 500! "The alarm was acknoweledged." @10/14/2019 4:00:17 PM
'[firstState] Success; Refresh; (10 field results) [NorthMotor] 500! "The alarm was acknoweledged." @10/14/2019 4:00:02 PM
'[firstState] Success; Refresh; (10 field results) [NorthMotor] 500! "The alarm was acknoweledged." @10/14/2019 4:00:16 PM
'[firstState] Success; Refresh; (10 field results) [SouthMotor] 700! "The alarm was acknoweledged." @10/14/2019 4:00:21 PM
'[firstState] Success; Refresh; (10 field results) [SouthMotor] 500! "The alarm was acknoweledged." @10/14/2019 4:00:03 PM
'[firstState] Success; Refresh; RefreshComplete
'[firstState] Success; (10 field results) [Internal] 500! "Raising Events" @11/8/2019 7:48:08 PM
'[firstState] Success; (10 field results) [Internal] 500! "Events Raised" @11/8/2019 7:48:08 PM
'[secondState] Success; Refresh; RefreshInitiated
'[secondState] Success; Refresh; (10 field results) [SouthMotor] 100! "The dialog was activated" @9/10/2019 8:08:25 PM
'[secondState] Success; Refresh; (10 field results) [SouthMotor] 100! "The alarm is active." @11/8/2019 7:48:07 PM
'[secondState] Success; Refresh; (10 field results) [SouthMotor] 700! "The alarm was acknoweledged." @10/14/2019 4:00:21 PM
'[secondState] Success; Refresh; (10 field results) [SouthMotor] 500! "The alarm was acknoweledged." @10/14/2019 4:00:03 PM
'[secondState] Success; Refresh; (10 field results) [SouthMotor] 100! "The alarm severity has increased." @9/10/2019 8:09:02 PM
'[secondState] Success; Refresh; (10 field results) [SouthMotor] 100! "The alarm severity has increased." @9/10/2019 8:09:59 PM
'[secondState] Success; Refresh; RefreshComplete
'[firstState] Success; (10 field results) [Internal] 500! "Raising Events" @11/8/2019 7:48:09 PM
'[firstState] Success; (10 field results) [Internal] 500! "Events Raised" @11/8/2019 7:48:09 PM
'[firstState] Success; (10 field results) [Internal] 500! "Raising Events" @11/8/2019 7:48:10 PM
'[firstState] Success; (10 field results) [Internal] 500! "Events Raised" @11/8/2019 7:48:10 PM
'[firstState] Success; (10 field results) [Internal] 500! "Raising Events" @11/8/2019 7:48:11 PM
'[firstState] Success; (10 field results) [Internal] 500! "Events Raised" @11/8/2019 7:48:11 PM
'[firstState] Success; (10 field results) [Internal] 500! "Raising Events" @11/8/2019 7:48:12 PM
'[firstState] Success; (10 field results) [Internal] 500! "Events Raised" @11/8/2019 7:48:12 PM
'[firstState] Success; (10 field results) [EastTank] 500! "The alarm severity has increased." @11/8/2019 7:48:13 PM
'[firstState] Success; (10 field results) [Internal] 500! "Raising Events" @11/8/2019 7:48:13 PM
'[firstState] Success; (10 field results) [Internal] 500! "Events Raised" @11/8/2019 7:48:13 PM
'[firstState] Success; (10 field results) [Internal] 500! "Raising Events" @11/8/2019 7:48:14 PM
'[firstState] Success; (10 field results) [Internal] 500! "Events Raised" @11/8/2019 7:48:14 PM
'[firstState] Success; (10 field results) [Internal] 500! "Raising Events" @11/8/2019 7:48:15 PM
'[firstState] Success; (10 field results) [Internal] 500! "Events Raised" @11/8/2019 7:48:15 PM
'[firstState] Success; (10 field results) [Internal] 500! "Raising Events" @11/8/2019 7:48:16 PM
'[firstState] Success; (10 field results) [Internal] 500! "Events Raised" @11/8/2019 7:48:16 PM
'[firstState] Success; (10 field results) [Internal] 500! "Raising Events" @11/8/2019 7:48:17 PM
'[firstState] Success; (10 field results) [Internal] 500! "Events Raised" @11/8/2019 7:48:17 PM
'[firstState] Success; (10 field results) [Internal] 500! "Raising Events" @11/8/2019 7:48:18 PM
'[firstState] Success; (10 field results) [Internal] 500! "Events Raised" @11/8/2019 7:48:18 PM
'[secondState] Success; (10 field results) [SouthMotor] 300! "The alarm severity has increased." @11/8/2019 7:48:18 PM
'...

Rem#endregion Example
