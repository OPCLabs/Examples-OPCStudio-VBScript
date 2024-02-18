Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to subscribe to item changes and obtain the events by pulling them.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")
' In order to use event pull, you must set a non-zero queue capacity upfront.
Client.PullItemChangedQueueCapacity = 1000

WScript.Echo "Subscribing item changes..."
Client.SubscribeItem "", "OPCLabs.KitServer.2", "Simulation.Random", 1000

WScript.Echo "Processing item changes for 1 minute..."
Dim endTime: endTime = Now() + 60*(1/24/60/60)
Do
    Dim EventArgs: Set EventArgs = Client.PullItemChanged(2*1000)
    If Not (EventArgs Is Nothing) Then
        ' Handle the notification event
        WScript.Echo EventArgs
    End If    
Loop While Now() < endTime

WScript.Echo "Unsubscribing item changes..."
Client.UnsubscribeAllItems

WScript.Echo "Finished."

Rem#endregion Example
