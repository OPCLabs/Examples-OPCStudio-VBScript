Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to subscribe to events and display the event message with each notification. It also shows how to
Rem unsubscribe afterwards.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Dim ServerDescriptor: Set ServerDescriptor = CreateObject("OpcLabs.EasyOpc.ServerDescriptor")
ServerDescriptor.ServerClass = "OPCLabs.KitEventServer.2"

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.AlarmsAndEvents.EasyAEClient")
WScript.ConnectObject Client, "Client_"

WScript.Echo "Subscribing events..."
Dim SubscriptionParameters: Set SubscriptionParameters = CreateObject("OpcLabs.EasyOpc.AlarmsAndEvents.AESubscriptionParameters")
SubscriptionParameters.NotificationRate = 1000
Dim handle: handle = Client.SubscribeEvents(ServerDescriptor, SubscriptionParameters, True, Nothing)

WScript.Echo "Processing event notifications for 1 minute..."
WScript.Sleep 60*1000

WScript.Echo "Unsubscribing events..."
Client.UnsubscribeEvents handle

WScript.Echo "Finished."



Rem Notification event handler
Sub Client_Notification(Sender, e)
    If Not e.Succeeded Then
        WScript.Echo "*** Failure: " & e.ErrorMessageBrief
        Exit Sub
    End If

	If Not e.EventData Is Nothing Then WScript.Echo e.EventData.Message
End Sub

Rem#endregion Example
