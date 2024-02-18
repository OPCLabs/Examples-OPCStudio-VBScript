Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to read the attributes of 4 OPC-UA nodes at once, and display the results.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Dim ReadArguments1: Set ReadArguments1 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.UAReadArguments")
ReadArguments1.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
ReadArguments1.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10853"

Dim ReadArguments2: Set ReadArguments2 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.UAReadArguments")
ReadArguments2.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
ReadArguments2.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10845"

Dim ReadArguments3: Set ReadArguments3 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.UAReadArguments")
ReadArguments3.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
ReadArguments3.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10304"

Dim ReadArguments4: Set ReadArguments4 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.UAReadArguments")
ReadArguments4.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
ReadArguments4.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10389"

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
        WScript.Echo "results(" & i & ").AttributeData: " & AttributeDataResult.AttributeData
    Else
        WScript.Echo "results(" & i & ") *** Failure: " & AttributeDataResult.ErrorMessageBrief
    End If
Next

Rem#endregion Example
