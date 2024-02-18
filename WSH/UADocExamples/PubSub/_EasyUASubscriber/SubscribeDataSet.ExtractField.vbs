Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to subscribe to dataset messages and extract data of a specific field.
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
' If the field names were contained in the dataset message (such as in JSON), or if we knew the metadata from some other 
' source, this step would not be needed.
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

WScript.Echo "Finished."



Sub Subscriber_DataSetMessage(Sender, e)
    ' Display the dataset.
    If e.Succeeded Then
        ' An event with null DataSetData just indicates a successful connection.
        If Not (e.DataSetData Is Nothing) Then
            ' Extract field data, looking up the field by its name.
            Dim Int32FastValueData: Set Int32FastValueData = e.DataSetData.FieldDataDictionary.Item("Int32Fast")
            WScript.Echo Int32FastValueData
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
'6502 {System.Int32} @2019-10-06T10:02:01.254,647,600,00; Good
'6538 {System.Int32} @2019-10-06T10:02:01.755,010,700,00; Good
'6615 {System.Int32} @2019-10-06T10:02:02.255,780,200,00; Good
'6687 {System.Int32} @2019-10-06T10:02:02.756,495,900,00; Good
'6769 {System.Int32} @2019-10-06T10:02:03.257,320,200,00; Good
'6804 {System.Int32} @2019-10-06T10:02:03.757,667,300,00; Good
'6877 {System.Int32} @2019-10-06T10:02:04.258,405,000,00; Good
'6990 {System.Int32} @2019-10-06T10:02:04.759,532,900,00; Good
'7063 {System.Int32} @2019-10-06T10:02:05.260,257,200,00; Good
'7163 {System.Int32} @2019-10-06T10:02:05.761,261,800,00; Good
'7255 {System.Int32} @2019-10-06T10:02:06.262,176,800,00; Good
'7321 {System.Int32} @2019-10-06T10:02:06.762,839,800,00; Good
'7397 {System.Int32} @2019-10-06T10:02:07.263,598,900,00; Good
'7454 {System.Int32} @2019-10-06T10:02:07.764,168,900,00; Good
'7472 {System.Int32} @2019-10-06T10:02:08.264,350,400,00; Good
'...

Rem#endregion Example
