Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to let the user browse for a host (computer) and an endpoint of an OPC-UA server residing on it.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Const DialogResult_OK = 1

Dim HostAndEndpointDialog: Set HostAndEndpointDialog = CreateObject("OpcLabs.EasyOpc.UA.Forms.Browsing.UAHostAndEndpointDialog")
HostAndEndpointDialog.EndpointDescriptor.Host = "opcua.demo-this.com"

Dim dialogResult: dialogResult = HostAndEndpointDialog.ShowDialog
WScript.Echo dialogResult

If dialogResult <> DialogResult_OK Then
    WScript.Quit
End If

' Display results
WScript.Echo "HostElement: " & HostAndEndpointDialog.HostElement
WScript.Echo "DiscoveryElement: " & HostAndEndpointDialog.DiscoveryElement

Rem#endregion Example
