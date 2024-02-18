Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example shows how change the update rate of an existing OPC XML-DA subscription.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Dim ItemSubscriptionArguments1: Set ItemSubscriptionArguments1 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.EasyDAItemSubscriptionArguments")
ItemSubscriptionArguments1.ServerDescriptor.UrlString = "http://opcxml.demo-this.com/XmlDaSampleServer/Service.asmx"
ItemSubscriptionArguments1.ItemDescriptor.ItemID = "Dynamic/Analog Types/Int"
ItemSubscriptionArguments1.GroupParameters.RequestedUpdateRate = 2000

Dim arguments(0)
Set arguments(0) = ItemSubscriptionArguments1

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")
WScript.ConnectObject Client, "Client_"

WScript.Echo "Subscribing..."
Dim handles: handles = Client.SubscribeMultipleItems(arguments)


WScript.Echo "Waiting for 20 seconds..."
WScript.Sleep 20*1000

WScript.Echo "Changing subscription..."
Client.ChangeItemSubscription handles(0), 500

WScript.Echo "Waiting for 10 seconds..."
WScript.Sleep 10*1000

WScript.Echo "Unsubscribing..."
Client.UnsubscribeAllItems

WScript.Echo "Waiting for 10 seconds..."
WScript.Sleep 10*1000

WScript.DisconnectObject Client
Set Client = Nothing


' Item changed event handler
Sub Client_ItemChanged(Sender, e)
    If Not (e.Succeeded) Then
        WScript.Echo "*** Failure: " & e.ErrorMessageBrief
        Exit Sub
    End If

	WScript.Echo e.Vtq
End Sub
Rem#endregion Example
