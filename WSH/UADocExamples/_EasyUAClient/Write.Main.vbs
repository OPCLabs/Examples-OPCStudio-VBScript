Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to write data (a value, timestamps and status code) into a single attribute of a node.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Const GoodOrSuccess = 0

' Instantiate the client object
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")

' Modify data of a node's attribute
Dim StatusCode: Set StatusCode = CreateObject("OpcLabs.EasyOpc.UA.UAStatusCode")
StatusCode.Severity = GoodOrSuccess
Dim AttributeData: Set AttributeData = CreateObject("OpcLabs.EasyOpc.UA.UAAttributeData")
AttributeData.Value = 12345
AttributeData.StatusCode = StatusCode
AttributeData.SourceTimestamp = Now

' Perform the operation
On Error Resume Next
Client.Write "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer", "nsu=http://test.org/UA/Data/ ;i=10221", _
             AttributeData
' The target server may not support this, and in such case a failure will occur.
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

Rem#endregion Example
