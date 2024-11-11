Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example shows how change the update rate of multiple existing subscriptions.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Dim ItemSubscriptionArguments1: Set ItemSubscriptionArguments1 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.EasyDAItemSubscriptionArguments")
ItemSubscriptionArguments1.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
ItemSubscriptionArguments1.ItemDescriptor.ItemID = "Simulation.Random"
ItemSubscriptionArguments1.GroupParameters.RequestedUpdateRate = 1000

Dim ItemSubscriptionArguments2: Set ItemSubscriptionArguments2 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.EasyDAItemSubscriptionArguments")
ItemSubscriptionArguments2.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
ItemSubscriptionArguments2.ItemDescriptor.ItemID = "Trends.Ramp (1 min)"
ItemSubscriptionArguments2.GroupParameters.RequestedUpdateRate = 1000

Dim ItemSubscriptionArguments3: Set ItemSubscriptionArguments3 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.EasyDAItemSubscriptionArguments")
ItemSubscriptionArguments3.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
ItemSubscriptionArguments3.ItemDescriptor.ItemID = "Trends.Sine (1 min)"
ItemSubscriptionArguments3.GroupParameters.RequestedUpdateRate = 1000

Dim ItemSubscriptionArguments4: Set ItemSubscriptionArguments4 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.EasyDAItemSubscriptionArguments")
ItemSubscriptionArguments4.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
ItemSubscriptionArguments4.ItemDescriptor.ItemID = "Simulation.Register_I4"
ItemSubscriptionArguments4.GroupParameters.RequestedUpdateRate = 1000

Dim arguments(3)
Set arguments(0) = ItemSubscriptionArguments1
Set arguments(1) = ItemSubscriptionArguments2
Set arguments(2) = ItemSubscriptionArguments3
Set arguments(3) = ItemSubscriptionArguments4

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")
WScript.ConnectObject Client, "Client_"

WScript.Echo "Subscribing..."
Dim handleArray: handleArray = Client.SubscribeMultipleItems(arguments)

WScript.Echo "Waiting for 10 seconds..."
WScript.Sleep 10*1000

WScript.Echo "Changing subscriptions..."

Dim HandleGroupArguments1: Set HandleGroupArguments1 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAHandleGroupArguments")
HandleGroupArguments1.Handle = handleArray(0)
HandleGroupArguments1.GroupParameters.RequestedUpdateRate = 100

Dim HandleGroupArguments2: Set HandleGroupArguments2 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAHandleGroupArguments")
HandleGroupArguments2.Handle = handleArray(1)
HandleGroupArguments2.GroupParameters.RequestedUpdateRate = 100

Dim subscriptionChangeArguments(1)
Set subscriptionChangeArguments(0) = HandleGroupArguments1
Set subscriptionChangeArguments(1) = HandleGroupArguments2

Client.ChangeMultipleItemSubscriptions subscriptionChangeArguments

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
