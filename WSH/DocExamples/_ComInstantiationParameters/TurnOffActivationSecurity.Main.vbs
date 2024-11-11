Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This examples shows how to turn off the activation security when the OPC server object is created.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

'

Dim ComManagement: Set ComManagement = CreateObject("OpcLabs.BaseLib.Runtime.InteropServices.ComManagement")
ComManagement.Configuration.InstantiationParameters.TurnOffActivationSecurity = True

Dim ClientManagement: Set ClientManagement = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClientManagement")
ClientManagement.SharedParameters.ClientParameters.ComInstantiationParameters.TurnOffActivationSecurity = True

'

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")
On Error Resume Next
Dim value: value = Client.ReadItemValue("localhost", "OPCLabs.KitServer.2", "Simulation.Random")
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0
WScript.Echo "value: " & value
Rem#endregion Example
