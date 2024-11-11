Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to read the attributes of 4 OPC-UA nodes specified by browse paths at once, and display the 
Rem results.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Dim BrowsePathParser: Set BrowsePathParser = CreateObject("OpcLabs.EasyOpc.UA.Navigation.Parsing.UABrowsePathParser")
BrowsePathParser.DefaultNamespaceUriString = "http://test.org/UA/Data/"

Dim ReadArguments1: Set ReadArguments1 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.UAReadArguments")
ReadArguments1.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
' Note: Add error handling around the following statement if the browse path is not guaranteed to be syntactically valid.
ReadArguments1.NodeDescriptor.BrowsePath = BrowsePathParser.Parse("[ObjectsFolder]/Data/Dynamic/Scalar/FloatValue")

Dim ReadArguments2: Set ReadArguments2 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.UAReadArguments")
ReadArguments2.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
' Note: Add error handling around the following statement if the browse path is not guaranteed to be syntactically valid.
ReadArguments2.NodeDescriptor.BrowsePath = BrowsePathParser.Parse("[ObjectsFolder]/Data/Dynamic/Scalar/SByteValue")

Dim ReadArguments3: Set ReadArguments3 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.UAReadArguments")
ReadArguments3.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
' Note: Add error handling around the following statement if the browse path is not guaranteed to be syntactically valid.
ReadArguments3.NodeDescriptor.BrowsePath = BrowsePathParser.Parse("[ObjectsFolder]/Data/Static/Array/UInt16Value")

Dim ReadArguments4: Set ReadArguments4 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.UAReadArguments")
ReadArguments4.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
' Note: Add error handling around the following statement if the browse path is not guaranteed to be syntactically valid.
ReadArguments4.NodeDescriptor.BrowsePath = BrowsePathParser.Parse("[ObjectsFolder]/Data/Static/UserScalar/Int32Value")

Dim arguments(3)
Set arguments(0) = ReadArguments1
Set arguments(1) = ReadArguments2
Set arguments(2) = ReadArguments3
Set arguments(3) = ReadArguments4

' Instantiate the client object
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")

' Perform the operation
Dim results: results = Client.ReadMultiple(arguments)

' Display results
Dim i: For i = LBound(results) To UBound(results)
    Dim AttributeDataResult: Set AttributeDataResult = results(i)
    If AttributeDataResult.Succeeded Then
        WScript.Echo "results[" & i & "].AttributeData: " & AttributeDataResult.AttributeData
    Else
        WScript.Echo "results[" & i & "] *** Failure: " & AttributeDataResult.ErrorMessageBrief
    End If
Next

' Example output:
'results[0].AttributeData: 4.187603E+21 {System.Single} @2019-11-09T14:05:46.268 @@2019-11-09T14:05:46.268; Good
'results[1].AttributeData: -98 {System.Int16} @2019-11-09T14:05:46.268 @@2019-11-09T14:05:46.268; Good
'results[2].AttributeData: [58] {38240, 11129, 64397, 22845, 30525, ...} {System.Int32[]} @2019-11-09T14:00:07.543 @@2019-11-09T14:05:46.268; Good
'results[3].AttributeData: 1280120396 {System.Int32} @2019-11-09T14:00:07.590 @@2019-11-09T14:05:46.268; Good

Rem#endregion Example
