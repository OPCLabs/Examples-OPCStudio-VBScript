Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to write data (a value, timestamps and status code) into 3 nodes at once, test for success of each 
Rem write and display the exception message in case of failure.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Const GoodOrSuccess = 0

' Instantiate the client object
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")

Dim StatusCode: Set StatusCode = CreateObject("OpcLabs.EasyOpc.UA.UAStatusCode")
StatusCode.Severity = GoodOrSuccess

Dim WriteArguments1: Set WriteArguments1 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.UAWriteArguments")
WriteArguments1.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
WriteArguments1.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10221"
Dim AttributeData1: Set AttributeData1 = CreateObject("OpcLabs.EasyOpc.UA.UAAttributeData")
AttributeData1.Value = 23456
AttributeData1.StatusCode = StatusCode
AttributeData1.SourceTimestamp = Now
WriteArguments1.AttributeData = AttributeData1

Dim WriteArguments2: Set WriteArguments2 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.UAWriteArguments")
WriteArguments2.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
WriteArguments2.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10226"
Dim AttributeData2: Set AttributeData2 = CreateObject("OpcLabs.EasyOpc.UA.UAAttributeData")
AttributeData2.Value = 2.3456789
AttributeData2.StatusCode = StatusCode
AttributeData2.SourceTimestamp = Now
WriteArguments2.AttributeData = AttributeData2

Dim WriteArguments3: Set WriteArguments3 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.UAWriteArguments")
WriteArguments3.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
WriteArguments3.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10227"
Dim AttributeData3: Set AttributeData3 = CreateObject("OpcLabs.EasyOpc.UA.UAAttributeData")
AttributeData3.Value = "ABC"
AttributeData3.StatusCode = StatusCode
AttributeData3.SourceTimestamp = Now
WriteArguments3.AttributeData = AttributeData3

Dim arguments(2)
Set arguments(0) = WriteArguments1
Set arguments(1) = WriteArguments2
Set arguments(2) = WriteArguments3

' Modify data of nodes' attributes
Dim results: results = Client.WriteMultiple(arguments)

' Display results
Dim i: For i = LBound(results) To UBound(results)
    Dim WriteResult: Set WriteResult = results(i)
    ' The target server may not support this, and in such case failures will occur.
    If WriteResult.Succeeded Then
        WScript.Echo "Result " & i & " success"
    Else
        WScript.Echo "Result " & i & ": " & WriteResult.Exception.GetBaseException().Message
    End If
Next

Rem#endregion Example
