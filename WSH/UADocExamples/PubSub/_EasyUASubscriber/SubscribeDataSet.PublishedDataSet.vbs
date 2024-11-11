Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to subscribe to dataset messages, specifying just the published dataset name, and resolving all
Rem the dataset subscription arguments from an OPC-UA PubSub configuration file.
Rem
Rem In order to produce network messages for this example, run the UADemoPublisher tool. For documentation, see
Rem https://kb.opclabs.com/UADemoPublisher_Basics . In some cases, you may have to specify the interface name to be used.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Const UAPubSubResolverKind_PublisherFile = 3

Dim SubscribeDataSetArguments: Set SubscribeDataSetArguments = CreateObject("OpcLabs.EasyOpc.UA.PubSub.OperationModel.EasyUASubscribeDataSetArguments")

' Specify the published dataset name, and let all other subscription arguments be resolved automatically.
SubscribeDataSetArguments.DataSetSubscriptionDescriptor.PublishedDataSetName = "AllTypes-Dynamic"

' Define the PubSub resolver. We want the information be resolved from a PubSub binary configuration file that
' we have. The file itself is included alongside the script.
SubscribeDataSetArguments.DataSetSubscriptionDescriptor.ResolverDescriptor.PublisherFileResourceDescriptor.UrlString = "UADemoPublisher-Default.uabinary"
SubscribeDataSetArguments.DataSetSubscriptionDescriptor.ResolverDescriptor.ResolverKind = UAPubSubResolverKind_PublisherFile

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
'Dataset data: Good; Data; publisher=(UInt64)31, writer=4, fields: 16
'[BoolToggle, False {System.Boolean}; Good]
'[Byte, 137 {System.Byte}; Good]
'[Int16, 10377 {System.Int16}; Good]
'[Int32, 43145 {System.Int32}; Good]
'[Int64, 43145 {System.Int64}; Good]
'[SByte, 9 {System.Int16}; Good]
'[UInt16, 43145 {System.Int32}; Good]
'[UInt32, 43145 {System.Int64}; Good]
'[UInt64, 43145 {System.Decimal}; Good]
'[Float, 43145 {System.Single}; Good]
'[Double, 43145 {System.Double}; Good]
'[String, Lima {System.String}; Good]
'[ByteString, [20] {176, 63, 39, 37, 31, ...} {System.Byte[]}; Good]
'[Guid, 45a99b50-e265-41f2-adea-d0bcedc3ff4b {System.Guid}; Good]
'[DateTime, 10/3/2019 7:15:34 AM {System.DateTime}; Good]
'[UInt32Array, [10] {43145, 43146, 43147, 43148, 43149, ...} {System.Int64[]}; Good]
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=4, fields: 16
'[BoolToggle, True {System.Boolean}; Good]
'[Byte, 138 {System.Byte}; Good]
'[Int16, 10378 {System.Int16}; Good]
'[Int32, 43146 {System.Int32}; Good]
'[Int64, 43146 {System.Int64}; Good]
'[SByte, 10 {System.Int16}; Good]
'[UInt16, 43146 {System.Int32}; Good]
'[UInt32, 43146 {System.Int64}; Good]
'[UInt64, 43146 {System.Decimal}; Good]
'[Float, 43146 {System.Single}; Good]
'[Double, 43146 {System.Double}; Good]
'[String, Mike {System.String}; Good]
'[Guid, a0f06d75-9896-4fa3-9724-b564359da21b {System.Guid}; Good]
'[DateTime, 10/3/2019 7:15:34 AM {System.DateTime}; Good]
'[UInt32Array, [10] {43146, 43147, 43148, 43149, 43150, ...} {System.Int64[]}; Good]
'[ByteString, [20] {176, 63, 39, 37, 31, ...} {System.Byte[]}; Good]
'
'Dataset data: Good; Data; publisher=(UInt64)31, writer=4, fields: 16
'[DateTime, 10/3/2019 7:15:35 AM {System.DateTime}; Good]
'[BoolToggle, True {System.Boolean}; Good]
'[Byte, 138 {System.Byte}; Good]
'[Int16, 10378 {System.Int16}; Good]
'[Int32, 43146 {System.Int32}; Good]
'[Int64, 43146 {System.Int64}; Good]
'[SByte, 10 {System.Int16}; Good]
'[UInt16, 43146 {System.Int32}; Good]
'[UInt32, 43146 {System.Int64}; Good]
'[UInt64, 43146 {System.Decimal}; Good]
'[Float, 43146 {System.Single}; Good]
'[Double, 43146 {System.Double}; Good]
'[String, Mike {System.String}; Good]
'[ByteString, [20] {176, 63, 39, 37, 31, ...} {System.Byte[]}; Good]
'[Guid, a0f06d75-9896-4fa3-9724-b564359da21b {System.Guid}; Good]
'[UInt32Array, [10] {43146, 43147, 43148, 43149, 43150, ...} {System.Int64[]}; Good]
'
'...

Rem#endregion Example
