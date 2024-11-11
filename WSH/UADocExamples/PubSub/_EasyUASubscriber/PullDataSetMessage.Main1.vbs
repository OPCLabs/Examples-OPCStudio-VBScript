Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to subscribe to all dataset messages on an OPC-UA PubSub connection, and pull events, and display
Rem the incoming datasets.
Rem
Rem In order to produce network messages for this example, run the UADemoPublisher tool. For documentation, see
Rem https://kb.opclabs.com/UADemoPublisher_Basics . In some cases, you may have to specify the interface name to be used.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

' Define the PubSub connection we will work with.
Dim SubscribeDataSetArguments: Set SubscribeDataSetArguments = CreateObject("OpcLabs.EasyOpc.UA.PubSub.OperationModel.EasyUASubscribeDataSetArguments")
Dim ConnectionDescriptor: Set ConnectionDescriptor = SubscribeDataSetArguments.DataSetSubscriptionDescriptor.ConnectionDescriptor
ConnectionDescriptor.ResourceAddress.ResourceDescriptor.UrlString = "opc.udp://239.0.0.1"
' In some cases you may have to set the interface (network adapter) name that needs to be used, similarly to
' the statement below. Your actual interface name may differ, of course.
' ConnectionDescriptor.ResourceAddress.InterfaceName = "Ethernet"

' Instantiate the subscriber object.
Dim Subscriber: Set Subscriber = CreateObject("OpcLabs.EasyOpc.UA.PubSub.EasyUASubscriber")
' In order to use event pull, you must set a non-zero queue capacity upfront. 
Subscriber.PullDataSetMessageQueueCapacity = 1000

WScript.Echo "Subscribing..."
Subscriber.SubscribeDataSet SubscribeDataSetArguments

WScript.Echo "Processing dataset message events for 20 seconds..."
Dim endTime: endTime = Now() + 20*(1/24/60/60)
Do
    Dim EventArgs: Set EventArgs = Subscriber.PullDataSetMessage(2*1000)
    If Not (EventArgs Is Nothing) Then
        ' Display the dataset.
        If EventArgs.Succeeded Then
            ' An event with null DataSetData just indicates a successful connection.
            If Not (EventArgs.DataSetData Is Nothing) Then
                WScript.Echo
                WScript.Echo "Dataset data: " & EventArgs.DataSetData
                Dim Pair: For Each Pair in EventArgs.DataSetData.FieldDataDictionary
                    WScript.Echo Pair
                Next
            End If
        Else
            WScript.Echo
            WScript.Echo "*** Failure: " & EventArgs.ErrorMessageBrief
        End If
    End If    
Loop While Now() < endTime

WScript.Echo "Unsubscribing..."
Subscriber.UnsubscribeAllDataSets

WScript.Echo "Waiting for 1 second..."
' Unsubscribe operation is asynchronous, messages may still come for a short while.
WScript.Sleep 1*1000

WScript.Echo "Finished."



' Example output:
'
'Subscribing...
'Processing dataset message events for 20 seconds...
'
''Dataset data: Good; Data; publisher="32", writer=1, class=eae79794-1af7-4f96-8401-4096cd1d8908, fields: 4
'[#0, True {System.Boolean} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#1, 7945 {System.Int32} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#2, 5246 {System.Int32} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#3, 9/30/2019 11:19:14 AM {System.DateTime} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'
'Dataset data: Good; Data; publisher="32", writer=3, class=96976b7b-0db7-46c3-a715-0979884b55ae, fields: 100
'[#0, 45 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#1, 145 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#2, 245 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#3, 345 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#4, 445 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#5, 545 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#6, 645 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#7, 745 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#8, 845 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#9, 945 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'[#10, 1045 {System.Int64} @0001-01-01T00:00:00.000 @@0001-01-01T00:00:00.000; Good]
'...

Rem#endregion Example
