Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to subscribe to all dataset messages with specific publisher Id, on an OPC-UA PubSub connection
Rem with UDP UADP mapping.
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

' Define the arguments for subscribing to the dataset, where the filter is (unsigned 64-bit) publisher Id 31.
SubscribeDataSetArguments.DataSetSubscriptionDescriptor.Filter.PublisherId.SetIdentifier UAPublisherIdType_UInt64, 31

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
'Dataset data: Good; Event; publisher=(UInt64)31, writer=51, fields: 4
'[#0, True {System.Boolean}; Good]
'[#1, 1237 {System.Int32}; Good]
'[#2, 2514 {System.Int32}; Good]
'[#3, 10/1/2019 9:03:59 AM {System.DateTime}; Good]
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[#0, False {System.Boolean}; Good]
'[#1, 1239 {System.Int32}; Good]
'[#2, 2703 {System.Int32}; Good]
'[#3, 10/1/2019 9:04:01 AM {System.DateTime}; Good]
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=4, fields: 16
'[#0, False {System.Boolean}; Good]
'[#1, 215 {System.Byte}; Good]
'[#2, 1239 {System.Int16}; Good]
'[#3, 1239 {System.Int32}; Good]
'[#4, 1239 {System.Int64}; Good]
'[#5, 87 {System.Int16}; Good]
'[#6, 1239 {System.Int32}; Good]
'[#7, 1239 {System.Int64}; Good]
'[#8, 1239 {System.Decimal}; Good]
'[#9, 1239 {System.Single}; Good]
'[#10, 1239 {System.Double}; Good]
'[#11, Romeo {System.String}; Good]
'[#12, [20] {175, 186, 248, 246, 215, ...} {System.Byte[]}; Good]
'[#13, d4492ca8-35c8-4b98-8edf-6ffa5ca041ca {System.Guid}; Good]
'[#14, 10/1/2019 9:04:01 AM {System.DateTime}; Good]
'[#15, [10] {1239, 1240, 1241, 1242, 1243, ...} {System.Int64[]}; Good]
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[#2, 2722 {System.Int32}; Good]
'[#3, 10/1/2019 9:04:01 AM {System.DateTime}; Good]
'[#0, False {System.Boolean}; Good]
'[#1, 1239 {System.Int32}; Good]
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=3, fields: 100
'[#0, 39 {System.Int64}; Good]
'[#1, 139 {System.Int64}; Good]
'[#2, 239 {System.Int64}; Good]
'[#3, 339 {System.Int64}; Good]
'[#4, 439 {System.Int64}; Good]
'[#5, 539 {System.Int64}; Good]
'[#6, 639 {System.Int64}; Good]
'[#7, 739 {System.Int64}; Good]
'[#8, 839 {System.Int64}; Good]
'[#9, 939 {System.Int64}; Good]
'[#10, 1039 {System.Int64}; Good]
'...

Rem#endregion Example
