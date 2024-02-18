Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to feed the packet capture file into the PubSub subscriber, instead of connecting to the message
Rem oriented middleware (receiving the messages from the network).
Rem
Rem The OpcLabs.Pcap assembly needs to be referenced in your project (or otherwise made available, together with its
Rem dependencies) for the capture files to work. Refer to the documentation for more information.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Const UAPublisherIdType_UInt64 = 4

' Define the PubSub connection we will work with.
Dim SubscribeDataSetArguments: Set SubscribeDataSetArguments = CreateObject("OpcLabs.EasyOpc.UA.PubSub.OperationModel.EasyUASubscribeDataSetArguments")
Dim ConnectionDescriptor: Set ConnectionDescriptor = SubscribeDataSetArguments.DataSetSubscriptionDescriptor.ConnectionDescriptor
ConnectionDescriptor.ResourceAddress.ResourceDescriptor.UrlString = "opc.eth://FF-FF-FF-FF-FF-FF"
' Use packets from the specified Ethernet capture file. The file itself is at the root of the project, and we
' have specified that it has to be copied to the project's output directory.
' Note that .pcap is the default file name extension, and can thus be omitted.
ConnectionDescriptor.UseEthernetCaptureFile "UADemoPublisher-Ethernet.pcap"

' Alternative setup for Ethernet with VLAN tagging:
'ConnectionDescriptor.ResourceAddress.ResourceDescriptor.UrlString = "opc.eth://FF-FF-FF-FF-FF-FF:2"
'ConnectionDescriptor.UseEthernetCaptureFile "UADemoPublisher-EthernetVlan.pcap"

' Alternative setup for UDP over IPv4:
'ConnectionDescriptor.ResourceAddress.ResourceDescriptor.UrlString = "opc.udp://239.0.0.1"
'ConnectionDescriptor.UseUdpCaptureFile "UADemoPublisher-UDP.pcap"

' Alternative setup for UDP over IPv6:
'ConnectionDescriptor.ResourceAddress.ResourceDescriptor.UrlString = "opc.udp://[ff02::1]"
'ConnectionDescriptor.UseUdpCaptureFile "UADemoPublisher-UDP6.pcap"

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
'Dataset data: 2019-10-31T16:04:59.145,266,700,00; Good; Data; publisher=(UInt64)31, writer=1, fields: 4
'[#0, True {System.Boolean} @2019-10-31T16:04:59.145,266,700,00; Good]
'[#1, 0 {System.Int32} @2019-10-31T16:04:59.145,266,700,00; Good]
'[#2, 767 {System.Int32} @2019-10-31T16:04:59.145,266,700,00; Good]
'[#3, 10/31/2019 4:04:59 PM {System.DateTime} @2019-10-31T16:04:59.145,266,700,00; Good]
'
'Dataset data: 2019-10-31T16:04:59.170,047,500,00; Good; Data; publisher=(UInt64)31, writer=3, fields: 100
'[#0, 0 {System.Int64} @2019-10-31T16:04:59.170,047,500,00; Good]
'[#1, 100 {System.Int64} @2019-10-31T16:04:59.170,047,500,00; Good]
'[#2, 200 {System.Int64} @2019-10-31T16:04:59.170,047,500,00; Good]
'[#3, 300 {System.Int64} @2019-10-31T16:04:59.170,047,500,00; Good]
'[#4, 400 {System.Int64} @2019-10-31T16:04:59.170,047,500,00; Good]
'[#5, 500 {System.Int64} @2019-10-31T16:04:59.170,047,500,00; Good]
'[#6, 600 {System.Int64} @2019-10-31T16:04:59.170,047,500,00; Good]
'[#7, 700 {System.Int64} @2019-10-31T16:04:59.170,047,500,00; Good]
'[#8, 800 {System.Int64} @2019-10-31T16:04:59.170,047,500,00; Good]
'[#9, 900 {System.Int64} @2019-10-31T16:04:59.170,047,500,00; Good]
'[#10, 1000 {System.Int64} @2019-10-31T16:04:59.170,047,500,00; Good]
'...

Rem#endregion Example
