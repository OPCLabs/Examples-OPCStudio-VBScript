Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem Shows how to find all registrations in the GDS.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Const UAApplicationTypes_All = 7

' Define which GDS we will work with.
Dim GdsEndpointDescriptor: Set GdsEndpointDescriptor = CreateObject("OpcLabs.EasyOpc.UA.UAEndpointDescriptor")
GdsEndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:58810/GlobalDiscoveryServer"
GdsEndpointDescriptor.UserIdentity.UserNameTokenInfo.UserName = "appadmin"
GdsEndpointDescriptor.UserIdentity.UserNameTokenInfo.Password = "demo"

' Instantiate the global discovery client object
Dim GlobalDiscoveryClient: Set GlobalDiscoveryClient = CreateObject("OpcLabs.EasyOpc.UA.Gds.EasyUAGlobalDiscoveryClient")

' Find all (client or server) applications registered in the GDS.
Dim startingRecordId: startingRecordId = 0
Dim maximumRecordsToReturn: maximumRecordsToReturn = 0
Dim applicationName: applicationName = ""
Dim applicationUriString: applicationUriString = ""
Dim productUriString: productUriString = ""
Dim serverCapabilities: serverCapabilities = Array()
Dim lastCounterResetTime
Dim nextRecordId
Dim applicationDescriptionArray
On Error Resume Next
GlobalDiscoveryClient.QueryApplications _
    GdsEndpointDescriptor, _
    startingRecordId, _
    maximumRecordsToReturn, _
    applicationName, _
    applicationUriString, _
    UAApplicationTypes_All, _
    productUriString, _
    serverCapabilities, _ 
    lastCounterResetTime, _
    nextRecordId, _
    applicationDescriptionArray
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

' For each application returned by the query, find its registrations in the GDS.
Dim ApplicationDescription
For Each ApplicationDescription In applicationDescriptionArray
    WScript.Echo
    WScript.Echo "Application URI string: " & ApplicationDescription.ApplicationUriString

    On Error Resume Next
    Dim applicationRecordArray: applicationRecordArray = GlobalDiscoveryClient.FindApplications( _
        gdsEndpointDescriptor, _
        ApplicationDescription.ApplicationUriString)
    If Err.Number <> 0 Then
        WScript.Echo "  *** Failure: " & Err.Source & ": " & Err.Description
    Else
        Dim ApplicationRecord
        For Each ApplicationRecord In applicationRecordArray
            ' Display results
            WScript.Echo "  Application ID: " & ApplicationRecord.ApplicationId
        Next
    End If
    On Error Goto 0
Next


' Example output:
'
'Application URI string: urn:sampleserver
'  Application ID: nsu=http://opcfoundation.org/UA/GDS/applications/ ;ns=2;g=09ecaa08-6ec6-462c-a214-1e66a3099107
'
'Application URI string: urn:alarmconditionserver
'  Application ID: nsu=http://opcfoundation.org/UA/GDS/applications/ ;ns=2;g=783e1e9a-8036-43b6-928f-97488c460266
'
'Application URI string: urn:PC:MultiTargetUADocExamples:5.54.1026.1:neutral:null
'  Application ID: nsu=http://opcfoundation.org/UA/GDS/applications/ ;ns=2;g=9e700ea5-55a6-4c3c-ba9f-b91c890dc519
'
'Application URI string: urn:PC:UADocExamples:5.56.0.16:neutral:null
'  Application ID: nsu=http://opcfoundation.org/UA/GDS/applications/ ;ns=2;g=e182e28c-086b-4fc7-82c7-70ca7cda3033
'
'Application URI string: urn:PC:cscript:5.812.10240.16384
'  Application ID: nsu=http://opcfoundation.org/UA/GDS/applications/ ;ns=2;g=aec94459-f513-4979-8619-8383555fca61

Rem#endregion Example
