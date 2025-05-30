Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem Shows how to obtain a new application certificate pack from the certificate manager (GDS), and store it for subsequent 
Rem usage.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

' Define which GDS we will work with.
Dim GdsEndpointDescriptor: Set GdsEndpointDescriptor = CreateObject("OpcLabs.EasyOpc.UA.UAEndpointDescriptor")
GdsEndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:58810/GlobalDiscoveryServer"
GdsEndpointDescriptor.UserIdentity.UserNameTokenInfo.UserName = "appadmin"
GdsEndpointDescriptor.UserIdentity.UserNameTokenInfo.Password = "demo"

' Obtain the application interface.
Dim Application: Set Application = CreateObject("OpcLabs.EasyOpc.UA.Application.EasyUAApplication")

' Display which application we are about to work with.
Dim ApplicationElement: Set ApplicationElement = Application.GetApplicationElement
WScript.Echo "Application URI string: " & Application.GetApplicationElement.ApplicationUriString

Rem Obtain a new application certificate pack from the certificate manager (GDS), and store it for subsequent usage.
On Error Resume Next
Dim CertificateDictionary: Set CertificateDictionary = Application.ObtainNewCertificatePack(GdsEndpointDescriptor)
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

' Display results
Dim Pair: For Each Pair In CertificateDictionary
    WScript.Echo
    WScript.Echo "Certificate type Id: " & Pair.Key
    WScript.Echo "Certificate: " & Pair.Value
Next

Rem#endregion Example
