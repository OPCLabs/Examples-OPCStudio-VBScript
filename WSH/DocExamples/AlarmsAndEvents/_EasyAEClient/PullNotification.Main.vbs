Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to subscribe to events and obtain the notification events by pulling them.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Dim ServerDescriptor: Set ServerDescriptor = CreateObject("OpcLabs.EasyOpc.ServerDescriptor")
ServerDescriptor.ServerClass = "OPCLabs.KitEventServer.2"

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.AlarmsAndEvents.EasyAEClient")
' In order to use event pull, you must set a non-zero queue capacity upfront.
Client.PullNotificationQueueCapacity = 1000

WScript.Echo "Subscribing events..."
Dim SubscriptionParameters: Set SubscriptionParameters = CreateObject("OpcLabs.EasyOpc.AlarmsAndEvents.AESubscriptionParameters")
SubscriptionParameters.NotificationRate = 1000
Dim handle: handle = Client.SubscribeEvents(ServerDescriptor, SubscriptionParameters, True, Nothing)

WScript.Echo "Processing event notifications for 1 minute..."
Dim endTime: endTime = Now() + 60*(1/24/60/60)
Do
    Dim EventArgs: Set EventArgs = Client.PullNotification(2*1000)
    If Not (EventArgs Is Nothing) Then
        ' Handle the notification event
        WScript.Echo EventArgs
    End If    
Loop While Now() < endTime

WScript.Echo "Unsubscribing events..."
Client.UnsubscribeEvents handle

WScript.Echo "Finished."

Rem#endregion Example
