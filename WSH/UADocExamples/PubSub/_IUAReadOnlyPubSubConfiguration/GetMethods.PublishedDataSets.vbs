Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example obtains and prints out information about all published datasets in the OPC UA PubSub configuration.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

' Instantiate the publish-subscribe client object.
Dim PublishSubscribeClient: Set PublishSubscribeClient = CreateObject("OpcLabs.EasyOpc.UA.PubSub.InformationModel.EasyUAPublishSubscribeClient")

On Error Resume Next
DumpPublishedDataSets
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

WScript.Echo "Finished."



Sub DumpPublishedDataSets()
    WScript.Echo "Loading the configuration..."
    ' Load the PubSub configuration from a file. The file itself is included alongside the script.
    Dim PubSubConfiguration: Set PubSubConfiguration = PublishSubscribeClient.LoadReadOnlyConfiguration("UADemoPublisher-Default.uabinary")

    ' Alternatively, using the statements below, you can access a live configuration residing in an OPC UA Server
    ' with appropriate information model.
    'Dim EndpointDescriptor: Set EndpointDescriptor = CreateObject("OpcLabs.EasyOpc.UA.UAEndpointDescriptor")
    'EndpointDescriptor.UrlString = "opc.tcp://localhost:48010"
    'Dim PubSubConfiguration: Set PubSubConfiguration = PublishSubscribeClient.AccessReadOnlyConfiguration(EndpointDescriptor)

    ' Get the names of PubSub connections in the configuration, regardless of the folder they reside in.
    Dim PublishedDataSetNames: Set PublishedDataSetNames = PubSubConfiguration.ListAllPublishedDataSetNames

    Dim publishedDataSetName: For Each publishedDataSetName In PublishedDataSetNames
        WScript.Echo "Published dataset: " & publishedDataSetName

        ' You can use the statement below to obtain parameters of the published dataset.
        'Dim PublishedDataSetElement: Set PublishedDataSetElement = PubSubConfiguration.GetPublishedDataSetElement(Nothing, publishedDataSetName)
    Next
End Sub



' Example output:
'
'Loading the configuration...
'Published dataset: Simple
'Published dataset: AllTypes
'Published dataset: MassTest
'Published dataset: AllTypes-Dynamic
'Finished.

Rem#endregion Example
