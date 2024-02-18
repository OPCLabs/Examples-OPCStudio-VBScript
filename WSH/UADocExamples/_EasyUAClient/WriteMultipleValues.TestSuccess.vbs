Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to write values into 3 nodes at once, test for success of each write and display the exception 
Rem message in case of failure.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

' Instantiate the client object
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")

Dim WriteValueArguments1: Set WriteValueArguments1 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.UAWriteValueArguments")
WriteValueArguments1.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
WriteValueArguments1.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10221"
WriteValueArguments1.Value = 23456

Dim WriteValueArguments2: Set WriteValueArguments2 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.UAWriteValueArguments")
WriteValueArguments2.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
WriteValueArguments2.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10226"
WriteValueArguments2.Value = "This string cannot be converted to Double"

Dim WriteValueArguments3: Set WriteValueArguments3 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.UAWriteValueArguments")
WriteValueArguments3.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
WriteValueArguments3.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;s=UnknownNode"
WriteValueArguments3.Value = "ABC"

Dim arguments(2)
Set arguments(0) = WriteValueArguments1
Set arguments(1) = WriteValueArguments2
Set arguments(2) = WriteValueArguments3

' Modify values of nodes
Dim results: results = Client.WriteMultipleValues(arguments)

' Display results
Dim i: For i = LBound(results) To UBound(results)
    Dim WriteResult: Set WriteResult = results(i)
    If WriteResult.Succeeded Then
        WScript.Echo "Result " & i & " success"
    Else
        WScript.Echo "Result " & i & ": " & WriteResult.Exception.GetBaseException().Message
    End If
Next

Rem#endregion Example
