Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example demonstrates how to place the application instance certificate in the platform-specific (Windows, Linux, 
Rem ...) certificate store.
Rem Note: COM is only available on Windows.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

WScript.Echo "Obtaining the application interface..."
Dim Application: Set Application = CreateObject("OpcLabs.EasyOpc.UA.Application.EasyUAApplication")

' Set the application certificate store path, which determines the location of the client certificate.
' Note that this only works once in each host process.
WScript.Echo "Setting the application certificate store path..."
Application.ApplicationParameters.ApplicationManifest.InstanceOwnStorePath = "CurrentUser\My"

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

' The certificate will be located or created in the specified platform-specific certificate store.
' On Windows, when viewed by the certmgr.msc tool, it will be under
' Certificates - Current User -> Personal -> Certificates.

WScript.Echo "Finished."
Rem#endregion Example
