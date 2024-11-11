Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to subscribe to dataset messages and specify a filter, resolving logical parameters to physical
Rem from an OPC-UA PubSub configuration file in binary format. The metadata obtained through the resolution is used to decode
Rem fixed layout messages with RawData field encoding.
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

' Define the PubSub connection we will work with, using its logical name in the PubSub configuration.
SubscribeDataSetArguments.DataSetSubscriptionDescriptor.ConnectionDescriptor.Name = "FixedLayoutConnection"
' In some cases you may have to set the interface (network adapter) name that needs to be used, similarly to
' the statement below. Your actual interface name may differ, of course.
SubscribeDataSetArguments.DataSetSubscriptionDescriptor.ConnectionDescriptor.ResourceAddress.InterfaceName = "Ethernet"

' Define the filter. The writer group and the dataset writer are specified using their logical names in the
' PubSub configuration. The publisher Id in the filter will be taken from the logical PubSub connection.
SubscribeDataSetArguments.DataSetSubscriptionDescriptor.Filter.WriterGroupDescriptor.Name = "FixedLayoutGroup"
SubscribeDataSetArguments.DataSetSubscriptionDescriptor.Filter.DataSetWriterDescriptor.Name = "SimpleWriter"

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

Sub Subscriber_ResolverAccess(Sender, e)
    ' Display resolution information.
    WScript.Echo e
End Sub



' Example output:
'
'Subscribing...
'Processing dataset message events for 20 seconds...
'[PublisherFile: UADemoPublisher-Default.uabinary] (no exception)
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
