Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to write a value, timestamp and quality into a single item.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")
On Error Resume Next
Client.WriteItem "", "OPCLabs.KitServer.2", "Simulation.Register_I4", 12345, DateSerial(1980, 1, 1), 192
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0
Rem#endregion Example
