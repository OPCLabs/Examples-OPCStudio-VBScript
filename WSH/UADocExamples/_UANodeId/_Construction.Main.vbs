Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows different ways of constructing OPC UA node IDs.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

' A node ID specifies a namespace (either by an URI or by an index), and an identifier.
' The identifier can be numeric (an integer), string, GUID, or opaque.

Const UANodeClass_All = 255


' A node ID can be specified in string form (so-called expanded text). 
' The code below specifies a namespace URI (nsu=...), and an integer identifier (i=...).
' Assigning an expanded text to a node ID parses the value being assigned and sets all corresponding
' properties accordingly.
Dim NodeId1: Set NodeId1 = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
NodeId1.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10853"
WScript.Echo NodeId1


' Similarly, with a string identifier (s=...).
Dim NodeId2: Set NodeId2 = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
NodeId2.ExpandedText = "nsu=http://test.org/UA/Data/ ;s=someIdentifier"
WScript.Echo NodeId2


' Actually, "s=" can be omitted (not recommended, though)
Dim NodeId3: Set NodeId3 = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
NodeId3.ExpandedText = "nsu=http://test.org/UA/Data/ ;someIdentifier"
WScript.Echo NodeId3
' Notice that the output is normalized - the "s=" is added again.


' Similarly, with a GUID identifier (g=...)
Dim NodeId4: Set NodeId4 = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
NodeId4.ExpandedText = "nsu=http://test.org/UA/Data/ ;g=BAEAF004-1E43-4A06-9EF0-E52010D5CD10"
WScript.Echo NodeId4
' Notice that the output is normalized - uppercase letters in the GUI are converted to lowercase, etc.


' Similarly, with an opaque identifier (b=..., in Base64 encoding).
Dim NodeId5: Set NodeId5 = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
NodeId5.ExpandedText = "nsu=http://test.org/UA/Data/ ;b=AP8="
WScript.Echo NodeId5


' Namespace index can be used instead of namespace URI. The server is allowed to change the namespace 
' indices between sessions (except for namespace 0), and for this reason, you should avoid the use of
' namespace indices, and rather use the namespace URIs whenever possible.
Dim NodeId6: Set NodeId6 = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
NodeId6.ExpandedText = "ns=2;i=10853"
WScript.Echo NodeId6


' Namespace index can be also specified together with namespace URI. This is still safe, but may be 
' a bit quicker to perform, because the client can just verify the namespace URI instead of looking 
' it up.
Dim NodeId7: Set NodeId7 = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
NodeId7.ExpandedText = "nsu=http://test.org/UA/Data/ ;ns=2;i=10853"
WScript.Echo NodeId7


' When neither namespace URI nor namespace index are given, the node ID is assumed to be in namespace
' with index 0 and URI "http://opcfoundation.org/UA/", which is reserved by OPC UA standard. There are 
' many standard nodes that live in this reserved namespace, but no nodes specific to your servers will 
' be in the reserved namespace, and hence the need to specify the namespace with server-specific nodes.
Dim NodeId8: Set NodeId8 = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
NodeId8.ExpandedText = "i=2254"
WScript.Echo NodeId8


' If you attempt to pass in a string that does not conform to the syntax rules, 
' a UANodeIdFormatException is thrown.
Dim NodeId9: Set NodeId9 = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
On Error Resume Next
NodeId9.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=notAnInteger"
If Err.Number = 0 Then
    WScript.Echo NodeId9
Else
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
End If
On Error Goto 0


' There is a parser object that can be used to parse the expanded texts of node IDs. 
Dim NodeIdParser10: Set NodeIdParser10 = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.Parsing.UANodeIdParser")
Dim NodeId10: Set NodeId10 = NodeIdParser10.Parse("nsu=http://test.org/UA/Data/ ;i=10853", False)
WScript.Echo NodeId10


' The parser can be used if you want to parse the expanded text of the node ID but do not want 
' exceptions be thrown.
Dim NodeIdParser11: Set NodeIdParser11 = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.Parsing.UANodeIdParser")
Dim NodeId11
Dim StringParsingError: Set StringParsingError = NodeIdParser11.TryParse("nsu=http://test.org/UA/Data/ ;i=notAnInteger", False, NodeId11)
If StringParsingError Is Nothing Then
    WScript.Echo NodeId11
Else
    WScript.Echo "*** Failure: " & StringParsingError.Message
End If


' You can also use the parser if you have node IDs where you want the default namespace be different 
' from the standard "http://opcfoundation.org/UA/".
Dim NodeIdParser12: Set NodeIdParser12 = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.Parsing.UANodeIdParser")
NodeIdParser12.DefaultNamespaceUriString = "http://test.org/UA/Data/"
Dim NodeId12: Set NodeId12 = NodeIdParser12.Parse("i=10853", False)
WScript.Echo NodeId12


' You can create a "null" node ID. Such node ID does not actually identify any valid node in OPC UA, but 
' is useful as a placeholder or as a starting point for further modifications of its properties.
Dim NodeId14: Set NodeId14 = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
WScript.Echo NodeId14


' Properties of a node ID can be modified individually. The advantage of this approach is that you do 
' not have to care about syntax of the node ID expanded text.
Dim NodeId15: Set NodeId15 = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
NodeId15.NamespaceUriString = "http://test.org/UA/Data/"
NodeId15.Identifier = 10853
WScript.Echo NodeId15


' If you know the type of the identifier upfront, it is safer to use typed properties that correspond 
' to specific types of identifier. Here, with an integer identifier.
Dim NodeId17: Set NodeId17 = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
NodeId17.NamespaceUriString = "http://test.org/UA/Data/"
NodeId17.NumericIdentifier = 10853
WScript.Echo NodeId17


' Similarly, with a string identifier.
Dim NodeId18: Set NodeId18 = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
NodeId18.NamespaceUriString = "http://test.org/UA/Data/"
NodeId18.StringIdentifier = "someIdentifier"
WScript.Echo NodeId18


' If you have GUID in its string form, the node ID object can parse it for you.
Dim NodeId20: Set NodeId20 = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
NodeId20.NamespaceUriString = "http://test.org/UA/Data/"
NodeId20.GuidIdentifierString = "BAEAF004-1E43-4A06-9EF0-E52010D5CD10"
WScript.Echo NodeId20


' And, with an opaque identifier.
Dim NodeId21: Set NodeId21 = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
NodeId21.NamespaceUriString = "http://test.org/UA/Data/"
NodeId21.OpaqueIdentifier = Array(&H00, &HFF)
WScript.Echo NodeId21


' We have built-in a list of all standard nodes specified by OPC UA. You can simply refer to these node IDs in your code.
' You can refer to any standard node using its name (in a string form).
' Note that assigning a non-existing standard name is not allowed, and throws ArgumentException.
Dim NodeId26: Set NodeId26 = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
NodeId26.StandardName = "TypesFolder"
WScript.Echo NodeId26
' When the UANodeId equals to one of the standard nodes, it is output in the shortened form - as the standard name only.


' When you browse for nodes in the OPC UA server, every returned node element contains a node ID that
' you can use further.
Dim Client27: Set Client27 = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")
Dim EndpointDescriptor: Set EndpointDescriptor = CreateObject("OpcLabs.EasyOpc.UA.UAEndpointDescriptor")
EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
' Browse from the Server node.
Dim ServerNodeId: Set ServerNodeId = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
ServerNodeId.StandardName = "Server"
Dim ServerNodeDescriptor: Set ServerNodeDescriptor = CreateObject("OpcLabs.EasyOpc.UA.UANodeDescriptor")
ServerNodeDescriptor.NodeId = ServerNodeId
' Browse all References.
Dim ReferencesNodeId: Set ReferencesNodeId = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
ReferencesNodeId.StandardName = "References"
'
Dim BrowseParameters: Set BrowseParameters = CreateObject("OpcLabs.EasyOpc.UA.UABrowseParameters")
BrowseParameters.NodeClasses = UANodeClass_All  ' this is the default, anyway
BrowseParameters.ReferenceTypeIds.Add ReferencesNodeId
'
On Error Resume Next
Dim NodeElementCollection27: Set NodeElementCollection27 = Client27.Browse( _
    EndpointDescriptor, ServerNodeDescriptor, BrowseParameters)
If Err.Number = 0 Then
    If NodeElementCollection27.Count <> 0 Then
        Dim NodeId27: Set NodeId27 = NodeElementCollection27(0).NodeId
        WScript.Echo NodeId27
    End If
Else
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
End If
On Error Goto 0

Rem#endregion Example
