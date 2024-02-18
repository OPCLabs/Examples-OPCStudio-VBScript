Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to obtain all method nodes under a given node of the OPC-UA address space. 
Rem For each node, it displays its browse name and node ID.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

' Instantiate the client object
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")

' Perform the operation
On Error Resume Next
Dim NodeElements: Set NodeElements = Client.BrowseMethods("opc.tcp://opcua.demo-this.com:51210/UA/SampleServer", "nsu=http://test.org/UA/Data/ ;i=10755")
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

' Display results
Dim NodeElement: For Each NodeElement In NodeElements
    WScript.Echo NodeElement.BrowseName & ": " & NodeElement.NodeId
Next


' Example output:
'ScalarMethod1: nsu=http://test.org/UA/Data/ ;ns=2;i=10756
'ScalarMethod2: nsu=http://test.org/UA/Data/ ;ns=2;i=10759
'ScalarMethod3: nsu=http://test.org/UA/Data/ ;ns=2;i=10762
'ArrayMethod1: nsu=http://test.org/UA/Data/ ;ns=2;i=10765
'ArrayMethod2: nsu=http://test.org/UA/Data/ ;ns=2;i=10768
'ArrayMethod3: nsu=http://test.org/UA/Data/ ;ns=2;i=10771
'UserScalarMethod1: nsu=http://test.org/UA/Data/ ;ns=2;i=10774
'UserScalarMethod2: nsu=http://test.org/UA/Data/ ;ns=2;i=10777
'UserArrayMethod1: nsu=http://test.org/UA/Data/ ;ns=2;i=10780
'UserArrayMethod2: nsu=http://test.org/UA/Data/ ;ns=2;i=10783

Rem#endregion Example
