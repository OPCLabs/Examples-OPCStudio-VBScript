
Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to obtain objects under the "Server" node in the address space.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Dim endpointDescriptor: endpointDescriptor = _
    "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    '"http://opcua.demo-this.com:51211/UA/SampleServer"  
    '"https://opcua.demo-this.com:51212/UA/SampleServer/"

' Instantiate the client object
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")

' Obtain objects under "Server" node
Dim ServerNodeId: Set ServerNodeId = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
ServerNodeId.StandardName = "Server"
On Error Resume Next
Dim NodeElementCollection: Set NodeElementCollection = Client.BrowseObjects(endpointDescriptor, ServerNodeId.ExpandedText)
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

' Display results
Dim NodeElement: For Each NodeElement In NodeElementCollection
    WScript.Echo 
    WScript.Echo "nodeElement.NodeId: " & NodeElement.NodeId
    WScript.Echo "nodeElement.NodeId.ExpandedText: " & NodeElement.NodeId.ExpandedText
    WScript.Echo "nodeElement.DisplayName: " & NodeElement.DisplayName
Next

' Example output:
'
'nodeElement.NodeId: Server_ServerCapabilities
'nodeElement.NodeId.ExpandedText: nsu=http://opcfoundation.org/UA/ ;i=2268
'nodeElement.DisplayName: ServerCapabilities
'
'nodeElement.NodeId: Server_ServerDiagnostics
'nodeElement.NodeId.ExpandedText: nsu=http://opcfoundation.org/UA/ ;i=2274
'nodeElement.DisplayName: ServerDiagnostics
'
'nodeElement.NodeId: Server_VendorServerInfo
'nodeElement.NodeId.ExpandedText: nsu=http://opcfoundation.org/UA/ ;i=2295
'nodeElement.DisplayName: VendorServerInfo
'
'nodeElement.NodeId: Server_ServerRedundancy
'nodeElement.NodeId.ExpandedText: nsu=http://opcfoundation.org/UA/ ;i=2296
'nodeElement.DisplayName: ServerRedundancy
'
'nodeElement.NodeId: Server_Namespaces
'nodeElement.NodeId.ExpandedText: nsu=http://opcfoundation.org/UA/ ;i=11715
'nodeElement.DisplayName: Namespaces

Rem#endregion Example
