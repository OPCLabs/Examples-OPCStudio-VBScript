Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to acknowledge an OPC UA event.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Const UAAttributeId_NodeId = 1
Const UAAttributeId_EventNotifier = 12

Const UAFilterOperator_Equals = 1

' Define which server we will work with.
Dim EndpointDescriptor: Set EndpointDescriptor = CreateObject("OpcLabs.EasyOpc.UA.UAEndpointDescriptor")
EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:62544/Quickstarts/AlarmConditionServer"

' Instantiate the client objects and hook events
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")
WScript.ConnectObject Client, "Client_"
Dim AlarmsAndConditionsClient: Set AlarmsAndConditionsClient = Client.AsAlarmsAndConditionsClient

'
Dim NodeId
Dim EventId
Dim anEvent: anEvent = False ' Some tools have event objects, but VBScript doesn't, we will use a boolean flag instead.

' Prepare arguments
Dim arguments(0)
Set arguments(0) = CreateMonitoredItemArguments

WScript.Echo "Subscribing..."
Client.SubscribeMultipleMonitoredItems arguments

WScript.Echo "Waiting for an event for 30 seconds..."
Dim endTime: endTime = Now() + 30*(1/24/60/60)
While (Not anEvent) And (Now() < endTime)
    WScript.Sleep 1000
WEnd
If Not anEvent Then
    WScript.Echo "Event not received."
    WScript.Quit
End If

WScript.Echo "Acknowledging an event..."
Dim NodeDescriptor: Set NodeDescriptor = CreateObject("OpcLabs.EasyOpc.UA.UANodeDescriptor")
Set NodeDescriptor.NodeId = NodeId
On Error Resume Next
AlarmsAndConditionsClient.Acknowledge EndpointDescriptor, NodeDescriptor, EventId, "Acknowledged by an automated example code."
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
End If
On Error Goto 0

WScript.Echo "Waiting for 5 seconds..."
WScript.Sleep 5 * 1000

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

Function CreateMonitoredItemArguments
    ' Event filter: Events with specific node ID.
    Dim Operand1: Set Operand1 = UABaseEventObject_Operands_NodeId
    Dim NodeId: Set NodeId = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
    NodeId.ExpandedText = "nsu=http://opcfoundation.org/Quickstarts/AlarmCondition ;ns=2;s=1:Colours/EastTank?Yellow"
    Dim Operand2: Set Operand2 = CreateObject("OpcLabs.EasyOpc.UA.Filtering.UALiteralOperand")
    Set Operand2.Value = NodeId
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
    MonitoringParameters.SamplingInterval = 1000

    Dim MonitoredItemArguments: Set MonitoredItemArguments = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.EasyUAMonitoredItemArguments")
    MonitoredItemArguments.AttributeId = UAAttributeId_EventNotifier
    MonitoredItemArguments.EndpointDescriptor = EndpointDescriptor
    Set MonitoredItemArguments.MonitoringParameters = MonitoringParameters
    Set MonitoredItemArguments.NodeDescriptor.NodeId = ServerNodeId

    Set CreateMonitoredItemArguments = MonitoredItemArguments
End Function

Sub Client_EventNotification(Sender, EventArgs)
    If Not EventArgs.Succeeded Then
        WScript.Echo "*** Failure: " & EventArgs.ErrorMessageBrief
        Exit Sub
    End If

    If Not (EventArgs.EventData Is Nothing) Then
        Dim BaseEventObject: Set BaseEventObject = EventArgs.EventData.BaseEvent
        WScript.Echo BaseEventObject

        ' Make sure we do not catch the event more than once
        If anEvent Then
            Exit Sub
        End If

        Set NodeId = BaseEventObject.NodeId
        EventId = BaseEventObject.EventId

        anEvent = True
    End If
End Sub



' Example output:
'Subscribing...
'Waiting for an event for 30 seconds...
'[EastTank] 100! "The alarm was acknoweledged." @11/9/2019 9:56:23 AM
'Acknowledging an event...
'Waiting for 5 seconds...
'[EastTank] 100! "The alarm was acknoweledged." @11/9/2019 9:56:23 AM
'Unsubscribing...
'Waiting for 5 seconds...

Rem#endregion Example
