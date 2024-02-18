Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how in a console application, the user is asked to allow a server instance certificate with
Rem mismatched domain name.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

' Define which server we will work with.
' Note that extra '.' at the end of the domain name. For the purpose of this example, it allows us to address
' the same domain, but cause a mismatch with what the names that are listed in the server instance certificate.
Dim endpointDescriptor: endpointDescriptor = "opc.tcp://opcua.demo-this.com.:51210/UA/SampleServer"

' Instantiate the client object.
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")
' Enforce the endpoint domain check.
Client.Isolated = True
Client.IsolatedParameters.SessionParameters.CheckEndpointDomain = True

' Obtain attribute data.
' The component automatically triggers the necessary user interaction during the first operation.
On Error Resume Next
Dim AttributeData: Set AttributeData = Client.Read(endpointDescriptor, "nsu=http://test.org/UA/Data/ ;i=10853")
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
'OPC-UA Endpoint Domain Mismatch
'The effective host name in endpoint URL returned by the server does not match any of the domain names in the server certificate.
'Endpoint URL as returned by the server: opc.tcp://opcua.demo-this.com.:51210/UA/SampleServer
'The server certificate is for following domain names or IP addresses: opcua.demo-this.com
'This may be an indication of a spoofing attempt. Do you want to allow the endpoint anyway? [Y/n]: Y
'Value: -1.285897E+14
'ServerTimestamp: 11/28/2019 1:34:23 PM
'SourceTimestamp: 11/28/2019 1:34:23 PM
'StatusCode: Good

Rem#endregion Example
