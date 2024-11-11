Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example shows how to subscribe to changes of a single OPC XML-DA item and display the value of the item with each change.
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
WScript.ConnectObject Client, "Client_"

WScript.Echo "Subscribing..."
Dim handles: handles = Client.SubscribeMultipleItems(arguments)


WScript.Echo "Processing item changed events for 1 minute..."
WScript.Sleep 60*1000

WScript.Echo "Unsubscribing..."
Client.UnsubscribeAllItems

WScript.Echo "Waiting for 5 seconds..."
WScript.Sleep 5*1000

WScript.DisconnectObject Client
Set Client = Nothing


' Item changed event handler
Sub Client_ItemChanged(Sender, e)
    If e.Succeeded Then
	WScript.Echo e.Vtq
    Else
        WScript.Echo "*** Failure: " & e.ErrorMessageBrief
    End If
End Sub
Rem#endregion Example
