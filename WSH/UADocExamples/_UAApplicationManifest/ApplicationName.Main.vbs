Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example demonstrates how to set the application name for the application instance certificate.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

' The management object allows access to static behavior.
WScript.Echo "Obtaining the client management object..."
Dim ClientManagement: Set ClientManagement = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClientManagement")
WScript.ConnectObject ClientManagement, "ClientManagement_"

WScript.Echo "Obtaining the application interface..."
Dim Application: Set Application = CreateObject("OpcLabs.EasyOpc.UA.Application.EasyUAApplication")

' Set the application name, which determines the subject of the client certificate.
' Note that this only works once in each host process.
WScript.Echo "Setting the application name..."
Application.ApplicationParameters.ApplicationManifest.ApplicationName = "QuickOPC - VBScript example application"

WScript.Echo "Creating a client object..."
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")

' Do something - invoke an OPC read, to trigger some loggable entries.
' If you are doing server development: Instantiate and start the server here, instead of invoking the client.
WScript.Echo "Reading a value..."
On Error Resume Next
Dim value: value = Client.ReadValue("opc.tcp://opcua.demo-this.com:51210/UA/SampleServer", "nsu=http://test.org/UA/Data/ ;i=10853")
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

' The certificate will be located or created in a directory similar to:
' C:\Users\All Users\OPC Foundation\CertificateStores\UA Applications\certs\
' and its subject will be as given by the application name.

WScript.Echo "Processing log entry events for 10 seconds..."
WScript.Sleep 10*1000

WScript.Echo "Finished."



' Event handler for the LogEntry event.
' Print the loggable entry containing client certificate parameters.
Sub ClientManagement_LogEntry(Sender, e)
    If e.EventId = 161 Then WScript.Echo e
End Sub

Rem#endregion Example
