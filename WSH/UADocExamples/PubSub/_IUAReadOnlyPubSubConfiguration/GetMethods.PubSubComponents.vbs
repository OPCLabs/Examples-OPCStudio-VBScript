Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example obtains and prints out information about PubSub connections, writer groups, and dataset writers in the
Rem OPC UA PubSub configuration.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

' Instantiate the publish-subscribe client object.
Dim PublishSubscribeClient: Set PublishSubscribeClient = CreateObject("OpcLabs.EasyOpc.UA.PubSub.InformationModel.EasyUAPublishSubscribeClient")

On Error Resume Next
DumpPubSubComponents
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

WScript.Echo "Finished."



Sub DumpPubSubComponents()
    WScript.Echo "Loading the configuration..."
    ' Load the PubSub configuration from a file. The file itself is included alongside the script.
    Dim PubSubConfiguration: Set PubSubConfiguration = PublishSubscribeClient.LoadReadOnlyConfiguration("UADemoPublisher-Default.uabinary")

    ' Alternatively, using the statements below, you can access a live configuration residing in an OPC UA Server
    ' with appropriate information model.
    'Dim EndpointDescriptor: Set EndpointDescriptor = CreateObject("OpcLabs.EasyOpc.UA.UAEndpointDescriptor")
    'EndpointDescriptor.UrlString = "opc.tcp://localhost:48010"
    'Dim PubSubConfiguration: Set PubSubConfiguration = PublishSubscribeClient.AccessReadOnlyConfiguration(EndpointDescriptor)

    ' Get the names of PubSub connections in the configuration.
    Dim ConnectionNames: Set ConnectionNames = PubSubConfiguration.ListConnectionNames
    Dim pubSubConnectionName: For Each pubSubConnectionName In ConnectionNames
        WScript.Echo "PubSub connection: " & pubSubConnectionName

        ' You can use the statement below to obtain parameters of the PubSub connection.
        'Dim PubSubConnectionElement: Set PubSubConnectionElement = PubSubConfiguration.GetConnectionElement(pubSubConnectionName)

        ' Get names of the writer groups on this PubSub connection.
        Dim WriterGroupNames: Set WriterGroupNames = PubSubConfiguration.ListWriterGroupNames(pubSubConnectionName)
        Dim writerGroupName: For Each writerGroupName In WriterGroupNames
            WScript.Echo "  Writer group: " & writerGroupName

            ' You can use the statement below to obtain parameters of the writer group.
            'Dim WriterGroupElement: Set WriterGroupElement = PubSubConfiguration.GetWriterGroupElement(pubSubConnectionName, writerGroupName)

            ' Get names of the dataset writers on this writer group.
            Dim DataSetWriterNames: Set DataSetWriterNames = PubSubConfiguration.ListDataSetWriterNames(pubSubConnectionName, writerGroupName)
            Dim dataSetWriterName: For Each dataSetWriterName In DataSetWriterNames
                WScript.Echo "    Dataset writer: " & dataSetWriterName

                ' You can use the statement below to obtain parameters of the dataset writer.
                'Dim DataSetWriterElement: Set DataSetWriterElement = _
                '    PubSubConfiguration.GetDataSetWriterElement(pubSubConnectionName, writerGroupName, dataSetWriterName)
            Next
        Next
    Next
End Sub



' Example output:
'
'Loading the configuration...
'PubSub connection: FixedLayoutConnection
'  Writer group: FixedLayoutGroup
'    Dataset writer: SimpleWriter
'    Dataset writer: AllTypesWriter
'    Dataset writer: MassTestWriter
'PubSub connection: DynamicLayoutConnection
'  Writer group: DynamicLayoutGroup
'    Dataset writer: SimpleWriter
'    Dataset writer: MassTestWriter
'    Dataset writer: AllTypes-DynamicWriter
'    Dataset writer: EventSimpleWriter
'PubSub connection: FlexibleLayoutConnection
'  Writer group: FlexibleLayoutGroup
'    Dataset writer: SimpleWriter
'    Dataset writer: MassTestWriter
'    Dataset writer: AllTypes-DynamicWriter
'Finished.

Rem#endregion Example
