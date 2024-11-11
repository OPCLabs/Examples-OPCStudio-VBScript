Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to subscribe to dataset messages on an OPC-UA PubSub connection, and then unsubscribe from that
Rem dataset.
Rem
Rem In order to produce network messages for this example, run the UADemoPublisher tool. For documentation, see
Rem https://kb.opclabs.com/UADemoPublisher_Basics . In some cases, you may have to specify the interface name to be used.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

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

' Instantiate the subscriber object and hook events.
Dim Subscriber: Set Subscriber = CreateObject("OpcLabs.EasyOpc.UA.PubSub.EasyUASubscriber")
WScript.ConnectObject Subscriber, "Subscriber_"

WScript.Echo "Subscribing..."
Dim dataSetHandle: dataSetHandle = Subscriber.SubscribeDataSet(SubscribeDataSetArguments)

WScript.Echo "Processing dataset message events for 20 seconds..."
WScript.Sleep 20*1000

WScript.Echo "Unsubscribing..."
Subscriber.UnsubscribeDataSet dataSetHandle

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
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[#0, True {System.Boolean}; Good]
'[#1, 7134 {System.Int32}; Good]
'[#2, 7364 {System.Int32}; Good]
'[#3, 10/1/2019 10:42:16 AM {System.DateTime}; Good]
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[#1, 7135 {System.Int32}; Good]
'[#2, 7429 {System.Int32}; Good]
'[#3, 10/1/2019 10:42:17 AM {System.DateTime}; Good]
'[#0, True {System.Boolean}; Good]
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[#2, 7495 {System.Int32}; Good]
'[#3, 10/1/2019 10:42:17 AM {System.DateTime}; Good]
'[#0, True {System.Boolean}; Good]
'[#1, 7135 {System.Int32}; Good]
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[#1, 7136 {System.Int32}; Good]
'[#2, 7560 {System.Int32}; Good]
'[#3, 10/1/2019 10:42:18 AM {System.DateTime}; Good]
'[#0, True {System.Boolean}; Good]
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[#2, 7626 {System.Int32}; Good]
'[#3, 10/1/2019 10:42:18 AM {System.DateTime}; Good]
'[#0, True {System.Boolean}; Good]
'[#1, 7136 {System.Int32}; Good]
'
'...

Rem#endregion Example
