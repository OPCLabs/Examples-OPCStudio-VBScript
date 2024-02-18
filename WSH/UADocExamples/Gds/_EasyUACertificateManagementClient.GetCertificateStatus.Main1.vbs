Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem Shows how to check if an application needs to update its certificate.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

' Define which GDS we will work with.
Dim GdsEndpointDescriptor: Set GdsEndpointDescriptor = CreateObject("OpcLabs.EasyOpc.UA.UAEndpointDescriptor")
GdsEndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:58810/GlobalDiscoveryServer"
GdsEndpointDescriptor.UserIdentity.UserNameTokenInfo.UserName = "appadmin"
GdsEndpointDescriptor.UserIdentity.UserNameTokenInfo.Password = "demo"

' Register our client application with the GDS, so that we obtain an application ID that we need later.
' Obtain the application interface.
Dim Application: Set Application = CreateObject("OpcLabs.EasyOpc.UA.Application.EasyUAApplication")
On Error Resume Next
Dim ApplicationId: Set ApplicationId = Application.RegisterToGds(GdsEndpointDescriptor)
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0
WScript.Echo "Application ID: " & ApplicationId

' Instantiate the certificate management client object
Dim CertificateManagementClient: Set CertificateManagementClient = _
    CreateObject("OpcLabs.EasyOpc.UA.Gds.EasyUACertificateManagementClient")

' Check if the application needs to update its certificate.
Dim NullNodeId: Set NullNodeId = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
On Error Resume Next
Dim UpdateRequired: UpdateRequired = CertificateManagementClient.GetCertificateStatus( _
    GdsEndpointDescriptor, ApplicationId, NullNodeId, NullNodeId)
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

' Display results
WScript.Echo "Update required: " & UpdateRequired


' Example output:
'Application ID: nsu=http://opcfoundation.org/UA/GDS/applications/ ;ns=2;g=aec94459-f513-4979-8619-8383555fca61
'Update required: False

Rem#endregion Example
