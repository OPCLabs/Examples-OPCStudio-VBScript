Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to display all fields of incoming events, or extract specific fields.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Const uaObjectIds_Server = "nsu=http://opcfoundation.org/UA/;i=2253"

Dim endpointDescriptor
endpointDescriptor = "opc.tcp://opcua.demo-this.com:62544/Quickstarts/AlarmConditionServer"

' Instantiate the client object and hook events
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")
WScript.ConnectObject Client, "Client_"

WScript.Echo "Subscribing..."
Client.SubscribeEvent endpointDescriptor, uaObjectIds_Server, 1000

WScript.Echo "Processing event notifications for 30 seconds..."
WScript.Sleep 30*1000

WScript.Echo "Unsubscribing..."
Client.UnsubscribeAllMonitoredItems

WScript.Echo "Waiting for 5 seconds..."
WScript.Sleep 5*1000




Function UAFilterElements_SimpleAttribute(TypeId, simpleRelativeBrowsePathString)
    Dim BrowsePathParser: Set BrowsePathParser = CreateObject("OpcLabs.EasyOpc.UA.Navigation.Parsing.UABrowsePathParser")
    Dim QualifiedNames: Set QualifiedNames = BrowsePathParser.ParseRelative(simpleRelativeBrowsePathString).ToUAQualifiedNAmeCollection

    Dim SimpleAttributeOperand: Set SimpleAttributeOperand = CreateObject("OpcLabs.EasyOpc.UA.Filtering.UASimpleAttributeOperand")
    Set SimpleAttributeOperand.TypeId.NodeId = TypeId
    Set SimpleAttributeOperand.QualifiedNames = QualifiedNames

    Set UAFilterElements_SimpleAttribute = SimpleAttributeOperand
End Function

Function ObjectTypeIds_BaseEventType
    Dim NodeId: Set NodeId = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
    NodeId.StandardName = "BaseEventType"
    Set ObjectTypeIds_BaseEventType = NodeId
End Function

Function UABaseEventObject_Operands_Message
    Set UABaseEventObject_Operands_Message = UAFilterElements_SimpleAttribute(ObjectTypeIds_BaseEventType, "/Message")
End Function

Function UABaseEventObject_Operands_SourceName
    Set UABaseEventObject_Operands_SourceName = UAFilterElements_SimpleAttribute(ObjectTypeIds_BaseEventType, "/SourceName")
End Function

Sub Client_EventNotification(Sender, e)
    WScript.Echo

    ' Display the event
    If e.EventData Is Nothing Then
        WScript.Echo e
        Exit Sub
    End If
    WScript.Echo "All fields:"
    Dim Pair: For Each Pair In e.EventData.FieldResults
        Dim AttributeField: Set AttributeField = Pair.Key
        Dim ValueResult: Set ValueResult = Pair.Value
        WScript.Echo "  " & AttributeField & " -> " & ValueResult
    Next
   
    ' Extracting specific fields using an event type ID and a simple relative path
    WScript.Echo "Source name: " & e.EventData.FieldResults.Item(UABaseEventObject_Operands_SourceName.ToUAAttributeField)
    WScript.Echo "Message: " & e.EventData.FieldResults.Item(UABaseEventObject_Operands_Message.ToUAAttributeField)
End Sub



' Example output (truncated):
'Subscribing...
'Processing event notifications for 30 seconds...
'
'[] Success
'
'[] Success; Refresh; RefreshInitiated
'
'All fields:
'  NodeId="BaseEventType", NodeId -> Success; nsu=http://opcfoundation.org/Quickstarts/AlarmCondition ;ns=2;s=1:Colours/EastTank?OnlineState {OpcLabs.EasyOpc.UA.AddressSpace.UANodeId}
'  NodeId="BaseEventType"/EventId -> Success; [16] {95, 68, 22, 205, 114, ...} {System.Byte[]}
'  NodeId="BaseEventType"/EventType -> Success; DialogConditionType {OpcLabs.EasyOpc.UA.AddressSpace.UANodeId}
'  NodeId="BaseEventType"/SourceNode -> Success; nsu=http://opcfoundation.org/Quickstarts/AlarmCondition ;ns=2;s=1:Colours/EastTank {OpcLabs.EasyOpc.UA.AddressSpace.UANodeId}
'  NodeId="BaseEventType"/SourceName -> Success; EastTank {System.String}
'  NodeId="BaseEventType"/Time -> Success; 9/10/2019 8:08:23 PM {System.DateTime}
'  NodeId="BaseEventType"/ReceiveTime -> Success; 9/10/2019 8:08:23 PM {System.DateTime}
'  NodeId="BaseEventType"/LocalTime -> Success; 00:00, DST {OpcLabs.EasyOpc.UA.UATimeZoneData}
'  NodeId="BaseEventType"/Message -> Success; The dialog was activated {System.String}
'  NodeId="BaseEventType"/Severity -> Success; 100 {System.Int32}
'Source name: Success; EastTank {System.String}
'Message: Success; The dialog was activated {System.String}
'
'All fields:
'  NodeId="BaseEventType", NodeId -> Success; nsu=http://opcfoundation.org/Quickstarts/AlarmCondition ;ns=2;s=1:Colours/EastTank?Red {OpcLabs.EasyOpc.UA.AddressSpace.UANodeId}
'  NodeId="BaseEventType"/EventId -> Success; [16] {124, 156, 219, 54, 120, ...} {System.Byte[]}
'  NodeId="BaseEventType"/EventType -> Success; ExclusiveDeviationAlarmType {OpcLabs.EasyOpc.UA.AddressSpace.UANodeId}
'  NodeId="BaseEventType"/SourceNode -> Success; nsu=http://opcfoundation.org/Quickstarts/AlarmCondition ;ns=2;s=1:Colours/EastTank {OpcLabs.EasyOpc.UA.AddressSpace.UANodeId}
'  NodeId="BaseEventType"/SourceName -> Success; EastTank {System.String}
'  NodeId="BaseEventType"/Time -> Success; 10/14/2019 4:00:13 PM {System.DateTime}
'  NodeId="BaseEventType"/ReceiveTime -> Success; 10/14/2019 4:00:13 PM {System.DateTime}
'  NodeId="BaseEventType"/LocalTime -> Success; 00:00, DST {OpcLabs.EasyOpc.UA.UATimeZoneData}
'  NodeId="BaseEventType"/Message -> Success; The alarm was acknoweledged. {System.String}
'  NodeId="BaseEventType"/Severity -> Success; 500 {System.Int32}
'Source name: Success; EastTank {System.String}
'Message: Success; The alarm was acknoweledged. {System.String}
'
'...

Rem#endregion Example
