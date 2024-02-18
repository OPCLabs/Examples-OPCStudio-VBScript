Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to read 3 values (without status or timestamps) at once, from the device (data source).
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

' Instantiate the client object
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")

Dim ReadArguments1: Set ReadArguments1 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.UAReadArguments")
ReadArguments1.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
ReadArguments1.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10845"
ReadArguments1.ReadParameters.MaximumAge = 0.0

Dim ReadArguments2: Set ReadArguments2 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.UAReadArguments")
ReadArguments2.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
ReadArguments2.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10853"
ReadArguments2.ReadParameters.MaximumAge = 0.0

Dim ReadArguments3: Set ReadArguments3 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.UAReadArguments")
ReadArguments3.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
ReadArguments3.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10855"
ReadArguments3.ReadParameters.MaximumAge = 0.0

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
