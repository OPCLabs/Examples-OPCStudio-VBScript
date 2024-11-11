Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem Shows how to obtain object describing the data type of complex data node with OPC UA Complex Data plug-in.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

' Define which server we will work with.
Dim endpointDescriptor: endpointDescriptor = _
    "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    '"http://opcua.demo-this.com:51211/UA/SampleServer"  
    '"https://opcua.demo-this.com:51212/UA/SampleServer/"

' Instantiate the client object.
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")

' Obtain the data type ID.
'
' In many cases, you would be able to obtain the data type ID of a particular node by reading its DataType
' attribute. 
' The sample server, however, shows a more advanced approach in which the data type ID refers to an abstract data type, and 
' the actual values are then sub-types of this base data type. This abstract data type does not have any encodings 
' associated with it and it is therefore not possible to extract its description from the server. We therefore use 
' a hard-coded data type ID for one of the sub-types in this example.
Dim dataTypeId: dataTypeId = "nsu=http://test.org/UA/Data/ ;i=9440"    ' ScalarValueDataType

' Get the IEasyUAClientComplexData service from the client. This is needed for advanced complex data 
' operations.
Dim ComplexData: Set ComplexData = _
    Client.GetServiceByName("OpcLabs.EasyOpc.UA.Plugins.ComplexData.IEasyUAClientComplexData, OpcLabs.EasyOpcUA")

' Resolve the data type ID to the data type object, containing description of the data type.
Dim ModelNodeDescriptor: Set ModelNodeDescriptor = CreateObject("OpcLabs.EasyOpc.UA.InformationModel.UAModelNodeDescriptor")
ModelNodeDescriptor.EndpointDescriptor.UrlString = endpointDescriptor
ModelNodeDescriptor.NodeDescriptor.NodeId.ExpandedText = dataTypeId
Dim EncodingName: Set EncodingName = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UAQualifiedName")
EncodingName.StandardName = "DefaultBinary"
Dim DataTypeResult: Set DataTypeResult = ComplexData.ResolveDataType(ModelNodeDescriptor, EncodingName)
If Not DataTypeResult.Succeeded Then
    WScript.Echo "*** Failure: "  & DataTypeResult.Exception.GetBaseException().Message
    WScript.Quit
End If

' The actual data type is in the Value property.
Dim DataType: Set DataType = DataTypeResult.Value

' Display basic information about what we have obtained.
WScript.Echo DataType

' If we want to see the whole hierarchy of the received data type, we can format it with the "V" (verbose)
' specifier. 
WScript.Echo
WScript.Echo DataType.ToString_2("V", Nothing)

' For processing the internals of the data type, refer to examples for GenericData class.


' Example output (truncated):
'
'ScalarValueDataType = structured
'
'ScalarValueDataType = structured
'  [BooleanValue] Boolean = primitive(System.Boolean)
'  [ByteStringValue] ByteString = primitive(System.Byte[])
'  [ByteValue] Byte = primitive(System.Byte)
'  [DateTimeValue] DateTime = primitive(System.DateTime)
'  [DoubleValue] Double = primitive(System.Double)
'  [EnumerationValue] Int32 = primitive(System.Int32)
'  [ExpandedNodeIdValue] ExpandedNodeId = structured
'    [ByteString] optional ByteStringNodeId = structured
'      [Identifier] ByteString = primitive(System.Byte[])
'      [NamespaceIndex] UInt16 = primitive(System.UInt16)
'    [FourByte] optional FourByteNodeId = structured
'      [Identifier] UInt16 = primitive(System.UInt16)
'      [NamespaceIndex] Byte = primitive(System.Byte)
'    [Guid] optional GuidNodeId = structured
'      [Identifier] Guid = primitive(System.Guid)
'      [NamespaceIndex] UInt16 = primitive(System.UInt16)
'    [NamespaceURI] optional CharArray = primitive(System.String)
'    [NamespaceURISpecified] switch Bit = primitive(System.Boolean)
'    [NodeIdType] switch NodeIdType = enumeration(6)
'      TwoByte = 0
'      FourByte = 1
'      Numeric = 2
'      String = 3
'      Guid = 4
'      ByteString = 5
'    [Numeric] optional NumericNodeId = structured
'      [Identifier] UInt32 = primitive(System.UInt32)
'      [NamespaceIndex] UInt16 = primitive(System.UInt16)
'    [ServerIndex] optional UInt32 = primitive(System.UInt32)
'    [ServerIndexSpecified] switch Bit = primitive(System.Boolean)
'    [String] optional StringNodeId = structured
'      [Identifier] CharArray = primitive(System.String)
'      [NamespaceIndex] UInt16 = primitive(System.UInt16)
'    [TwoByte] optional TwoByteNodeId = structured
'      [Identifier] Byte = primitive(System.Byte)
'  [FloatValue] Float = primitive(System.Single)
'  [GuidValue] Guid = primitive(System.Guid)
'  [Int16Value] Int16 = primitive(System.Int16)
'  [Int32Value] Int32 = primitive(System.Int32)
'  [Int64Value] Int64 = primitive(System.Int64)
'  [Integer] Variant = structured
'    [ArrayDimensions] optional sequence[*] of Int32 = primitive(System.Int32)
'    [ArrayDimensionsSpecified] switch sequence[1] of Bit = primitive(System.Boolean)
'    [ArrayLength] length optional Int32 = primitive(System.Int32)
'    [ArrayLengthSpecified] switch sequence[1] of Bit = primitive(System.Boolean)
'    [Boolean] optional sequence[*] of Boolean = primitive(System.Boolean)
'    [Byte] optional sequence[*] of Byte = primitive(System.Byte)

Rem#endregion Example
