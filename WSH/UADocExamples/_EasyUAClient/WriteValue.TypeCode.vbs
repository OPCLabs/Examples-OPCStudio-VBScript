Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to write a value into a single node, specifying a type code explicitly.
Rem 
Rem Reasons for specifying the type explicitly might be:
Rem - The data type in the server has subtypes, and the client therefore needs to pick the subtype to be written.
Rem - The data type that the reports is incorrect.
Rem - Writing with an explicitly specified type is more efficient.
Rem 
Rem TypeCode is easy to use, but it does not cover all possible types. It is also possible to specify the .NET Type, using
Rem a different overload of the WriteValue method.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Const TypeCode_Int32 = 9

Dim endpointDescriptor: endpointDescriptor = _
    "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    '"http://opcua.demo-this.com:51211/UA/SampleServer"  
    '"https://opcua.demo-this.com:51212/UA/SampleServer/"

' Instantiate the client object
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")

' Prepare the arguments
Dim WriteValueArguments1: Set WriteValueArguments1 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.UAWriteValueArguments")
WriteValueArguments1.EndpointDescriptor.UrlString = endpointDescriptor
WriteValueArguments1.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10221"
WriteValueArguments1.Value = 12345
WriteValueArguments1.ValueTypeCode = TypeCode_Int32

Dim arguments(0)
Set arguments(0) = WriteValueArguments1

' Modify value of a node
Dim results: results = Client.WriteMultipleValues(arguments)
Dim WriteResult: Set WriteResult = results(0)
If Not WriteResult.Succeeded Then
    WScript.Echo "*** Failure: " & WriteResult.Exception.GetBaseException().Message
End If

Rem#endregion Example
