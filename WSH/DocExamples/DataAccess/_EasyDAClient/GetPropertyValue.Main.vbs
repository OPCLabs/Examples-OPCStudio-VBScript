Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example shows how to get a value of a single OPC property.
Rem
Rem Note that some properties may not have a useful value initially (e.g. until the item is activated in a group), which also the
Rem case with Timestamp property as implemented by the demo server. This behavior is server-dependent, and normal. You can run 
Rem IEasyDAClient.ReadItemValue.Main.vbs shortly before this example, in order to obtain better property values. Your code may 
Rem also subscribe to the item in order to assure that it remains active.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Const Timestamp = 4

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")

On Error Resume Next
Dim value: value = Client.GetPropertyValue("", "OPCLabs.KitServer.2", "Simulation.Random", Timestamp)
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

WScript.Echo value
Rem#endregion Example
