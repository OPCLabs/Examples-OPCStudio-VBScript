Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to subscribe to changes of a single item and display the value of the item with each change.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")
WScript.ConnectObject Client, "Client_"

Client.SubscribeItem "", "OPCLabs.KitServer.2", "Simulation.Random", 1000

WScript.Echo "Processing item changed events for 1 minute..."
WScript.Sleep 60*1000



Sub Client_ItemChanged(Sender, e)
    If Not (e.Succeeded) Then
        WScript.Echo "*** Failure: " & e.ErrorMessageBrief
        Exit Sub
    End If

    WScript.Echo e.Vtq.ToString
End Sub
Rem#endregion Example
