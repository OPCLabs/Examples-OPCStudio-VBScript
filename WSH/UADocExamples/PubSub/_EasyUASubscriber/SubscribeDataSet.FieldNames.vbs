Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to subscribe to dataset messages and specify field names, without having the full metadata.
Rem
Rem In order to produce network messages for this example, run the UADemoPublisher tool. For documentation, see
Rem https://kb.opclabs.com/UADemoPublisher_Basics . In some cases, you may have to specify the interface name to be used.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Const UAPublisherIdType_UInt64 = 4

' Define the PubSub connection we will work with.
Dim SubscribeDataSetArguments: Set SubscribeDataSetArguments = CreateObject("OpcLabs.EasyOpc.UA.PubSub.OperationModel.EasyUASubscribeDataSetArguments")
Dim ConnectionDescriptor: Set ConnectionDescriptor = SubscribeDataSetArguments.DataSetSubscriptionDescriptor.ConnectionDescriptor
ConnectionDescriptor.ResourceAddress.ResourceDescriptor.UrlString = "opc.udp://239.0.0.1"
' In some cases you may have to set the interface (network adapter) name that needs to be used, similarly to
' the statement below. Your actual interface name may differ, of course.
' ConnectionDescriptor.ResourceAddress.InterfaceName = "Ethernet"

' Define the filter. Publisher Id (unsigned 64-bits) is 31, and the dataset writer Id is 1.
SubscribeDataSetArguments.DataSetSubscriptionDescriptor.Filter.PublisherId.SetIdentifier UAPublisherIdType_UInt64, 31
SubscribeDataSetArguments.DataSetSubscriptionDescriptor.Filter.DataSetWriterDescriptor.DataSetWriterId = 1

' Define the metadata. For UADP, the order of field metadata must correspond to the order of fields in the dataset message.
' Since the encoding is not RawData, we do not have to specify the type information for the fields.
Dim MetaData: Set MetaData = CreateObject("OpcLabs.EasyOpc.UA.PubSub.Configuration.UADataSetMetaData")
'
Dim Field1: Set Field1 = CreateObject("OpcLabs.EasyOpc.UA.PubSub.Configuration.UAFieldMetaData")
Field1.Name = "BoolToggle"
MetaData.Add(Field1)
'
Dim Field2: Set Field2 = CreateObject("OpcLabs.EasyOpc.UA.PubSub.Configuration.UAFieldMetaData")
Field2.Name = "Int32"
MetaData.Add(Field2)
'
Dim Field3: Set Field3 = CreateObject("OpcLabs.EasyOpc.UA.PubSub.Configuration.UAFieldMetaData")
Field3.Name = "Int32Fast"
MetaData.Add(Field3)
'
Dim Field4: Set Field4 = CreateObject("OpcLabs.EasyOpc.UA.PubSub.Configuration.UAFieldMetaData")
Field4.Name = "DateTime"
MetaData.Add(Field4)
'
Set SubscribeDataSetArguments.DataSetSubscriptionDescriptor.DataSetMetaData = MetaData

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

WScript.Echo "Finished."



' Example output:
'
'Subscribing...
'Processing dataset message events for 20 seconds...
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[BoolToggle, True {System.Boolean}; Good]
'[Int32, 25 {System.Int32}; Good]
'[Int32Fast, 928 {System.Int32}; Good]
'[DateTime, 10/3/2019 10:43:01 AM {System.DateTime}; Good]
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[Int32, 26 {System.Int32}; Good]
'[Int32Fast, 1007 {System.Int32}; Good]
'[DateTime, 10/3/2019 10:43:02 AM {System.DateTime}; Good]
'[BoolToggle, True {System.Boolean}; Good]
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[Int32Fast, 1113 {System.Int32}; Good]
'[DateTime, 10/3/2019 10:43:02 AM {System.DateTime}; Good]
'[BoolToggle, True {System.Boolean}; Good]
'[Int32, 26 {System.Int32}; Good]
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[BoolToggle, False {System.Boolean}; Good]
'[Int32, 27 {System.Int32}; Good]
'[Int32Fast, 1201 {System.Int32}; Good]
'[DateTime, 10/3/2019 10:43:03 AM {System.DateTime}; Good]
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[Int32Fast, 1260 {System.Int32}; Good]
'[DateTime, 10/3/2019 10:43:03 AM {System.DateTime}; Good]
'[BoolToggle, False {System.Boolean}; Good]
'[Int32, 27 {System.Int32}; Good]
'
'...

Rem#endregion Example
