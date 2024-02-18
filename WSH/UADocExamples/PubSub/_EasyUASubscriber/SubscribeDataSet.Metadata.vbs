Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to subscribe to dataset messages with RawData field encoding, specifying the metadata necessary
Rem for their decoding directly in the code.
Rem
Rem In order to produce network messages for this example, run the UADemoPublisher tool. For documentation, see
Rem https://kb.opclabs.com/UADemoPublisher_Basics . In some cases, you may have to specify the interface name to be used.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Const UABuiltInType_Boolean = 1
Const UABuiltInType_Int32 = 6
Const UABuiltInType_DateTime = 13

' Define the PubSub connection we will work with.
Dim SubscribeDataSetArguments: Set SubscribeDataSetArguments = CreateObject("OpcLabs.EasyOpc.UA.PubSub.OperationModel.EasyUASubscribeDataSetArguments")
Dim ConnectionDescriptor: Set ConnectionDescriptor = SubscribeDataSetArguments.DataSetSubscriptionDescriptor.ConnectionDescriptor
ConnectionDescriptor.ResourceAddress.ResourceDescriptor.UrlString = "opc.udp://239.0.0.1"
' In some cases you may have to set the interface (network adapter) name that needs to be used, similarly to
' the statement below. Your actual interface name may differ, of course.
' ConnectionDescriptor.ResourceAddress.InterfaceName = "Ethernet"

' Define the filter. Publisher Id (unsigned 16-bits) is 30, and the writer group Id is 101.
' The dataset writer Id (1) must not be specified in the filter, because it does not appear in the message.
SubscribeDataSetArguments.DataSetSubscriptionDescriptor.Filter.PublisherId.SetUInt16Identifier 30
SubscribeDataSetArguments.DataSetSubscriptionDescriptor.Filter.WriterGroupDescriptor.WriterGroupId = 101

' Define the metadata. For UADP, the order of field metadata must correspond to the order of fields in the dataset message.
Dim MetaData: Set MetaData = CreateObject("OpcLabs.EasyOpc.UA.PubSub.Configuration.UADataSetMetaData")
'
Dim Field1: Set Field1 = CreateObject("OpcLabs.EasyOpc.UA.PubSub.Configuration.UAFieldMetaData")
Field1.BuiltInType = UABuiltInType_Boolean
Field1.Name = "BoolToggle"
MetaData.Add(Field1)
'
Dim Field2: Set Field2 = CreateObject("OpcLabs.EasyOpc.UA.PubSub.Configuration.UAFieldMetaData")
Field2.BuiltInType = UABuiltInType_Int32
Field2.Name = "Int32"
MetaData.Add(Field2)
'
Dim Field3: Set Field3 = CreateObject("OpcLabs.EasyOpc.UA.PubSub.Configuration.UAFieldMetaData")
Field3.BuiltInType = UABuiltInType_Int32
Field3.Name = "Int32Fast"
MetaData.Add(Field3)
'
Dim Field4: Set Field4 = CreateObject("OpcLabs.EasyOpc.UA.PubSub.Configuration.UAFieldMetaData")
Field4.BuiltInType = UABuiltInType_DateTime
Field4.Name = "DateTime"
MetaData.Add(Field4)
'
Set SubscribeDataSetArguments.DataSetSubscriptionDescriptor.DataSetMetaData = MetaData

' Define the specific communication parameters for the dataset subscription.
' The dataset offset is needed with messages that do not contain dataset writer Ids and use RawData field
' encoding. An exception to this rule is when the dataset is the only or first in the dataset message payload,
' which is also the case here, but we are specifying the dataset offset anyway, for illustration.
SubscribeDataSetArguments.DataSetSubscriptionDescriptor.CommunicationParameters.UadpDataSetReaderMessageParameters.DataSetOffset = 15

' Instantiate the subscriber object and hook events.
Dim Subscriber: Set Subscriber = CreateObject("OpcLabs.EasyOpc.UA.PubSub.EasyUASubscriber")
WScript.ConnectObject Subscriber, "Subscriber_"

WScript.Echo "Subscribing..."
Subscriber.SubscribeDataSet SubscribeDataSetArguments

WScript.Echo "Processing dataset message events for 20 seconds..."
WScript.Sleep 20*1000

WScript.Echo "Unsubscribing..."
Subscriber.UnsubscribeAllDataSets

WScript.Echo "Waiting for 1 second..."
' Unsubscribe operation is asynchronous, messages may still come for a short while.
WScript.Sleep 1*1000

WScript.Echo "Finished."



Sub Subscriber_DataSetMessage(Sender, e)
    ' Display the dataset.
    If e.Succeeded Then
        ' An event with null DataSetData just indicates a successful connection.
        If Not (e.DataSetData Is Nothing) Then
            WScript.Echo
            WScript.Echo "Dataset data: " & e.DataSetData
            Dim Pair: For Each Pair in e.DataSetData.FieldDataDictionary
                WScript.Echo Pair
            Next
        End If
    Else
        WScript.Echo
        WScript.Echo "*** Failure: " & e.ErrorMessageBrief
    End If
End Sub



' Example output:
'
'Subscribing...
'Processing dataset message events for 20 seconds...
'
'Dataset data: Good; Data; publisher=(UInt16)30, group=101, fields: 4
'[BoolToggle, False {System.Boolean}; Good]
'[Int32, 3072 {System.Int32}; Good]
'[Int32Fast, 894 {System.Int32}; Good]
'[DateTime, 10/1/2019 12:21:14 PM {System.DateTime}; Good]
'
'Dataset data: Good; Data; publisher=(UInt16)30, group=101, fields: 4
'[BoolToggle, False {System.Boolean}; Good]
'[Int32, 3072 {System.Int32}; Good]
'[Int32Fast, 920 {System.Int32}; Good]
'[DateTime, 10/1/2019 12:21:14 PM {System.DateTime}; Good]
'
'Dataset data: Good; Data; publisher=(UInt16)30, group=101, fields: 4
'[BoolToggle, False {System.Boolean}; Good]
'[Int32, 3073 {System.Int32}; Good]
'[Int32Fast, 1003 {System.Int32}; Good]
'[DateTime, 10/1/2019 12:21:15 PM {System.DateTime}; Good]
'
'Dataset data: Good; Data; publisher=(UInt16)30, group=101, fields: 4
'[BoolToggle, False {System.Boolean}; Good]
'[Int32, 3073 {System.Int32}; Good]
'[Int32Fast, 1074 {System.Int32}; Good]
'[DateTime, 10/1/2019 12:21:15 PM {System.DateTime}; Good]
'
'Dataset data: Good; Data; publisher=(UInt16)30, group=101, fields: 4
'[BoolToggle, True {System.Boolean}; Good]
'[Int32, 3074 {System.Int32}; Good]
'[Int32Fast, 1140 {System.Int32}; Good]
'[DateTime, 10/1/2019 12:21:16 PM {System.DateTime}; Good]
'
'...

Rem#endregion Example
