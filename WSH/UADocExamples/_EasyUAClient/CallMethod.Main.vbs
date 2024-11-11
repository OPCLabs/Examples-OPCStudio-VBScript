Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to call a single method, and pass arguments to and from it.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Dim inputs(10)
inputs(0) = False
inputs(1) = 1
inputs(2) = 2
inputs(3) = 3
inputs(4) = 4
inputs(5) = 5
inputs(6) = 6
inputs(7) = 7
inputs(8) = 8
inputs(9) = 9
inputs(10) = 10

Dim typeCodes(10)
typeCodes(0) = 3    ' TypeCode.Boolean
typeCodes(1) = 5    ' TypeCode.SByte
typeCodes(2) = 6    ' TypeCode.Byte
typeCodes(3) = 7    ' TypeCode.Int16
typeCodes(4) = 8    ' TypeCode.UInt16
typeCodes(5) = 9    ' TypeCode.Int32
typeCodes(6) = 10   ' TypeCode.UInt32
typeCodes(7) = 11   ' TypeCode.Int64
typeCodes(8) = 12   ' TypeCode.UInt64
typeCodes(9) = 13   ' TypeCode.Single
typeCodes(10) = 14  ' TypeCode.Double

' Instantiate the client object
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")

' Perform the operation
On Error Resume Next
Dim outputs: outputs = Client.CallMethod( _
    "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer", _
    "nsu=http://test.org/UA/Data/ ;i=10755", _
    "nsu=http://test.org/UA/Data/ ;i=10756", _
    inputs, _
    typeCodes)
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

' Display results
Dim i: For i = LBound(outputs) To UBound(outputs)
    On Error Resume Next
    WScript.Echo "outputs(" & i & "): " & outputs(i)
    If Err <> 0 Then WScript.Echo "*** Error"   ' occurrs with types not recognized by VBScript
    On Error Goto 0
Next

Rem#endregion Example
