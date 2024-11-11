Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example shows how change the update rate of an existing subscription.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")
WScript.ConnectObject Client, "Client_"

WScript.Echo "Subscribing..."
Dim handle: handle = Client.SubscribeItem("", "OPCLabs.KitServer.2", "Simulation.Random", 1000)

WScript.Echo "Waiting for 10 seconds..."
WScript.Sleep 10*1000

WScript.Echo "Changing subscription..."
Client.ChangeItemSubscription handle, 100

WScript.Echo "Waiting for 10 seconds..."
WScript.Sleep 10*1000

WScript.Echo "Unsubscribing..."
Client.UnsubscribeAllItems

WScript.Echo "Waiting for 10 seconds..."
WScript.Sleep 10*1000

WScript.DisconnectObject Client
Set Client = Nothing



Sub Client_ItemChanged(Sender, e)
    If Not (e.Succeeded) Then
        WScript.Echo "*** Failure: " & e.ErrorMessageBrief
        Exit Sub
    End If

	WScript.Echo e.Vtq
End Sub
Rem#endregion Example
