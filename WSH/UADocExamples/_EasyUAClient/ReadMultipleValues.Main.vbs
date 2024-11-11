Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to read the Value attributes of 3 different nodes at once. Using the same method, it is also possible 
Rem to read multiple attributes of the same node.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

' Instantiate the client object
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")

Dim ReadArguments1: Set ReadArguments1 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.UAReadArguments")
ReadArguments1.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
ReadArguments1.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10845"

Dim ReadArguments2: Set ReadArguments2 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.UAReadArguments")
ReadArguments2.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
ReadArguments2.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10853"

Dim ReadArguments3: Set ReadArguments3 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.UAReadArguments")
ReadArguments3.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
ReadArguments3.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10855"

Dim arguments(2)
Set arguments(0) = ReadArguments1
Set arguments(1) = ReadArguments2
Set arguments(2) = ReadArguments3

' Obtain values. By default, the Value attributes of the nodes will be read.
Dim results: results = Client.ReadMultipleValues(arguments)

' Display results
Dim i: For i = LBound(results) To UBound(results)
    Dim ValueResult: Set ValueResult = results(i)
    If ValueResult.Succeeded Then
        WScript.Echo "Value: " & ValueResult.Value
    Else
        WScript.Echo "*** Failure: " & ValueResult.ErrorMessageBrief
    End If
Next

' Example output:
'
'Value: 8
'Value: -8.06803E+21
'Value: Strawberry Pig Banana Snake Mango Purple Grape Monkey Purple? Blueberry Lemon^            

Rem#endregion Example
