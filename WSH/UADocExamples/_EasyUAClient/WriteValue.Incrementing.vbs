Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to write an ever-incrementing value to an OPC UA variable.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Const endpointDescriptorUrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
Const nodeIdExpandedText = "nsu=http://test.org/UA/Data/ ;i=10221"
' Example settings with Softing dataFEED OPC Suite: 
'Const endpointDescriptorUrlString = "opc.tcp://localhost:4980/Softing_dataFEED_OPC_Suite_Configuration1"
'Const nodeIdExpandedText = "nsu=Local%20Items ;s=Local Items.EAK_Test1.EAK_Testwert1_I4"

' Instantiate the client object
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")

' 
Dim i: i = 0

Do While True
    WScript.Echo "@" & Time & ": Writing " & i
    On Error Resume Next
    Client.WriteValue endpointDescriptorUrlString, nodeIdExpandedText, i
    If Err.Number <> 0 Then
        WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
        WScript.Quit
    End If
    On Error Goto 0
    i = (i + 1) And &H7FFFFFFF
    WScript.Sleep 2*1000
Loop

Rem#endregion Example
