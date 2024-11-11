Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to obtain nodes under a given node of the OPC-UA address space. 
Rem For each node, it displays its browse name and node ID.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Dim EndpointDescriptor: Set EndpointDescriptor = CreateObject("OpcLabs.EasyOpc.UA.UAEndpointDescriptor")
EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"

Dim NodeDescriptor: Set NodeDescriptor = CreateObject("OpcLabs.EasyOpc.UA.UANodeDescriptor")
Dim BrowsePathParser: Set BrowsePathParser = CreateObject("OpcLabs.EasyOpc.UA.Navigation.Parsing.UABrowsePathParser")
BrowsePathParser.DefaultNamespaceUriString = "http://test.org/UA/Data/"
NodeDescriptor.BrowsePath = BrowsePathParser.Parse("[ObjectsFolder]/Data/Static/UserScalar")

Dim BrowseParameters: Set BrowseParameters = CreateObject("OpcLabs.EasyOpc.UA.UABrowseParameters")
BrowseParameters.StandardName = "AllForwardReferences"

' Instantiate the client object
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")

' Perform the operation
On Error Resume Next
Dim NodeElements: Set NodeElements = Client.Browse(EndpointDescriptor, NodeDescriptor, BrowseParameters)
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

' Display results
Dim NodeElement: For Each NodeElement In NodeElements
    WScript.Echo NodeElement.BrowseName & ": " & NodeElement.NodeId
Next

Rem#endregion Example
