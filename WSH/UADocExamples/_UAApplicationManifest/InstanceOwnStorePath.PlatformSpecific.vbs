Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example demonstrates how to place the client certificate in the platform-specific (Windows, Linux, ...) certificate 
Rem store.
Rem Note: COM is only available on Windows.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

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
