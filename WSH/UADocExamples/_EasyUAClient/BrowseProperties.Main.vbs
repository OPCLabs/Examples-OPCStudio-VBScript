
Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to obtain properties under the "Server" node in the address space.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Dim endpointDescriptor: endpointDescriptor = _
    "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    '"http://opcua.demo-this.com:51211/UA/SampleServer"  
    '"https://opcua.demo-this.com:51212/UA/SampleServer/"

' Instantiate the client object
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")

' Obtain properties under "Server" node
Dim ServerNodeId: Set ServerNodeId = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
ServerNodeId.StandardName = "Server"
On Error Resume Next
Dim NodeElementCollection: Set NodeElementCollection = Client.BrowseProperties(endpointDescriptor, ServerNodeId.ExpandedText)
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
'nodeElement.NodeId: Server_ServerArray
'nodeElement.NodeId.ExpandedText: nsu=http://opcfoundation.org/UA/ ;i=2254
'nodeElement.DisplayName: ServerArray
'
'nodeElement.NodeId: Server_NamespaceArray
'nodeElement.NodeId.ExpandedText: nsu=http://opcfoundation.org/UA/ ;i=2255
'nodeElement.DisplayName: NamespaceArray
'
'nodeElement.NodeId: Server_ServiceLevel
'nodeElement.NodeId.ExpandedText: nsu=http://opcfoundation.org/UA/ ;i=2267
'nodeElement.DisplayName: ServiceLevel
'
'nodeElement.NodeId: Server_Auditing
'nodeElement.NodeId.ExpandedText: nsu=http://opcfoundation.org/UA/ ;i=2994
'nodeElement.DisplayName: Auditing

Rem#endregion Example
