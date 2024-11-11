Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to unsubscribe from specific event notifications.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Dim ServerDescriptor: Set ServerDescriptor = CreateObject("OpcLabs.EasyOpc.ServerDescriptor")
ServerDescriptor.ServerClass = "OPCLabs.KitEventServer.2"

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.AlarmsAndEvents.EasyAEClient")
WScript.ConnectObject Client, "Client_"

WScript.Echo "Subscribing..."
Dim SubscriptionParameters: Set SubscriptionParameters = CreateObject("OpcLabs.EasyOpc.AlarmsAndEvents.AESubscriptionParameters")
SubscriptionParameters.NotificationRate = 1000
Dim handle: handle = Client.SubscribeEvents(ServerDescriptor, SubscriptionParameters, True, Nothing)

WScript.Echo "Waiting for 10 seconds..."
WScript.Sleep 10*1000

WScript.Echo "Unsubscribing..."
Client.UnsubscribeEvents handle

WScript.Echo "Waiting for 10 seconds..."
WScript.Sleep 10*1000



Rem Notification event handler
Sub Client_Notification(Sender, e)
    If Not (e.Succeeded) Then
        WScript.Echo "*** Failure: " & e.ErrorMessageBrief
        Exit Sub
    End If

	If Not e.EventData Is Nothing Then WScript.Echo e.EventData.Message
End Sub

Rem#endregion Example
