Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to call multiple methods, and pass arguments to and from them.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Dim inputs1(10)
inputs1(0) = False
inputs1(1) = 1
inputs1(2) = 2
inputs1(3) = 3
inputs1(4) = 4
inputs1(5) = 5
inputs1(6) = 6
inputs1(7) = 7
inputs1(8) = 8
inputs1(9) = 9
inputs1(10) = 10

Dim typeCodes1(10)
typeCodes1(0) = 3    ' TypeCode.Boolean
typeCodes1(1) = 5    ' TypeCode.SByte
typeCodes1(2) = 6    ' TypeCode.Byte
typeCodes1(3) = 7    ' TypeCode.Int16
typeCodes1(4) = 8    ' TypeCode.UInt16
typeCodes1(5) = 9    ' TypeCode.Int32
typeCodes1(6) = 10   ' TypeCode.UInt32
typeCodes1(7) = 11   ' TypeCode.Int64
typeCodes1(8) = 12   ' TypeCode.UInt64
typeCodes1(9) = 13   ' TypeCode.Single
typeCodes1(10) = 14  ' TypeCode.Double

Dim CallArguments1: Set CallArguments1 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.UACallArguments")
CallArguments1.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
CallArguments1.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10755"
CallArguments1.MethodNodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10756"
CallArguments1.InputArguments = inputs1
CallArguments1.InputTypeCodes = typeCodes1

Dim inputs2(11)
inputs2(0) = False
inputs2(1) = 1
inputs2(2) = 2
inputs2(3) = 3
inputs2(4) = 4
inputs2(5) = 5
inputs2(6) = 6
inputs2(7) = 7
inputs2(8) = 8
inputs2(9) = 9
inputs2(10) = 10
inputs2(11) = "eleven"

Dim typeCodes2(11)
typeCodes2(0) = 3    ' TypeCode.Boolean
typeCodes2(1) = 5    ' TypeCode.SByte
typeCodes2(2) = 6    ' TypeCode.Byte
typeCodes2(3) = 7    ' TypeCode.Int16
typeCodes2(4) = 8    ' TypeCode.UInt16
typeCodes2(5) = 9    ' TypeCode.Int32
typeCodes2(6) = 10   ' TypeCode.UInt32
typeCodes2(7) = 11   ' TypeCode.Int64
typeCodes2(8) = 12   ' TypeCode.UInt64
typeCodes2(9) = 13   ' TypeCode.Single
typeCodes2(10) = 14  ' TypeCode.Double
typeCodes2(11) = 18  ' TypeCode.String

Dim CallArguments2: Set CallArguments2 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.UACallArguments")
CallArguments2.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
CallArguments2.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10755"
CallArguments2.MethodNodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;i=10774"
CallArguments2.InputArguments = inputs2
CallArguments2.InputTypeCodes = typeCodes2

Dim arguments(1)
Set arguments(0) = CallArguments1
Set arguments(1) = CallArguments2

' Instantiate the client object
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")

' Perform the operation
Dim results: results = Client.CallMultipleMethods(arguments)

' Display results
Dim i: For i = LBound(results) To UBound(results)
    WScript.Echo
    WScript.Echo "results(" & i & "):"
    Dim Result: Set Result = results(i)

    If Result.Exception Is Nothing Then
        Dim outputs: outputs = Result.ValueArray
        Dim j: For j = LBound(outputs) To UBound(outputs)
            On Error Resume Next
            WScript.Echo Space(4) & "outputs(" & j & "): " & outputs(j)
            If Err <> 0 Then WScript.Echo Space(4) & "*** Error"   ' occurrs with types not recognized by VBScript
            On Error Goto 0
        Next
    Else
        WScript.Echo "*** Error: " & Result.Exception
    End If
Next

Rem#endregion Example
