Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to subscribe to OPC XML-DA item changes and obtain the events by pulling them.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Dim ItemSubscriptionArguments1: Set ItemSubscriptionArguments1 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.EasyDAItemSubscriptionArguments")
ItemSubscriptionArguments1.ServerDescriptor.UrlString = "http://opcxml.demo-this.com/XmlDaSampleServer/Service.asmx"
ItemSubscriptionArguments1.ItemDescriptor.ItemID = "Dynamic/Analog Types/Int"
ItemSubscriptionArguments1.GroupParameters.RequestedUpdateRate = 1000

Dim arguments(0)
Set arguments(0) = ItemSubscriptionArguments1

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")
' In order to use event pull, you must set a non-zero queue capacity upfront.
Client.PullItemChangedQueueCapacity = 1000

WScript.Echo "Subscribing item changes..."
Client.SubscribeMultipleItems arguments

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
