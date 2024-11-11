Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to obtain information about OPC UA servers available on the network.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

' Instantiate the client object
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")

' Obtain collection of application elements
On Error Resume Next
Dim DiscoveryElementCollection: Set DiscoveryElementCollection = _
    Client.DiscoverNetworkServers(Nothing, "opcua.demo-this.com")
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

' Display results
Dim DiscoveryElement: For Each DiscoveryElement In DiscoveryElementCollection
    WScript.Echo
    WScript.Echo "Server name: " & DiscoveryElement.ServerName
    WScript.Echo "Discovery URI string: " & DiscoveryElement.DiscoveryUriString
    Dim s: s = ""
    Dim serverCapability: For Each serverCapability In DiscoveryElement.ServerCapabilities
        s = s & serverCapability & " "
    Next
    WScript.Echo "Server capabilities: " & s
Next

Rem#endregion Example
