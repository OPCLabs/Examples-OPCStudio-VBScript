Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to read and display data of an attribute (value, timestamps, and status code).
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

' Instantiate the client object
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")

' Obtain attribute data. By default, the Value attribute of a node will be read.
On Error Resume Next
Dim AttributeData: Set AttributeData = Client.Read("opc.tcp://opcua.demo-this.com:51210/UA/SampleServer", _
                                                   "nsu=http://test.org/UA/Data/ ;i=10853")
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

' Display results
WScript.Echo "Value: " & AttributeData.Value
WScript.Echo "ServerTimestamp: " & AttributeData.ServerTimestamp
WScript.Echo "SourceTimestamp: " & AttributeData.SourceTimestamp
WScript.Echo "StatusCode: " & AttributeData.StatusCode

' Example output:
'
'Value: -2.230064E-31
'ServerTimestamp: 11/6/2011 1:34:30 PM
'SourceTimestamp: 11/6/2011 1:34:30 PM
'StatusCode: Good

Rem#endregion Example
