Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example demonstrates the loggable entries originating in the OPC-UA client engine and the EasyUAClient component.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

' The management object allows access to static behavior - here, the shared LogEntry event.
Dim ClientManagement: Set ClientManagement = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClientManagement")
WScript.ConnectObject ClientManagement, "ClientManagement_"

' Do something - invoke an OPC read, to trigger some loggable entries.
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")
On Error Resume Next
Dim value: value = Client.ReadValue("opc.tcp://opcua.demo-this.com:51210/UA/SampleServer", "nsu=http://test.org/UA/Data/ ;i=10853")
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

WScript.Echo "Processing log entry events for 1 minute..."
WScript.Sleep 60*1000



' Event handler for the LogEntry event. It simply prints out the event.
Sub ClientManagement_LogEntry(Sender, e)
	WScript.Echo e
End Sub

Rem#endregion Example
