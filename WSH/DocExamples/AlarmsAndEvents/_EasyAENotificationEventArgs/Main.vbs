Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example subscribe to events, and displays rich information available with each event notification. 
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

WScript.Echo "Processing event notifications for 1 minute..."
WScript.Sleep 60*1000

Client.UnsubscribeEvents handle



Rem Notification event handler
Sub Client_Notification(Sender, e)
    On Error Resume Next
	WScript.Echo
	WScript.Echo "e.Exception.Message: " & e.Exception.Message
	WScript.Echo "e.Exception.Source: " & e.Exception.Source
	WScript.Echo "e.Exception.ErrorCode: " & e.Exception.ErrorCode
	WScript.Echo "e.Arguments.State: " & e.Arguments.State
	WScript.Echo "e.Arguments.ServerDescriptor.MachineName: " & e.Arguments.ServerDescriptor.MachineName
	WScript.Echo "e.Arguments.ServerDescriptor.ServerClass: " & e.Arguments.ServerDescriptor.ServerClass
	WScript.Echo "e.Arguments.SubscriptionParameters.Active: " & e.Arguments.SubscriptionParameters.Active
	WScript.Echo "e.Arguments.SubscriptionParameters.NotificationRate: " & e.Arguments.SubscriptionParameters.NotificationRate
    Rem IMPROVE: Display Arguments.SubscriptionParameters.Filter details
	WScript.Echo "e.Arguments.SubscriptionParameters.Filter: " & e.Arguments.SubscriptionParameters.Filter  
    Rem IMPROVE: Display Arguments.SubscriptionParameters.ReturnedAttributesByCategory details
	WScript.Echo "e.Arguments.SubscriptionParameters.ReturnedAttributesByCategory: " & e.Arguments.SubscriptionParameters.ReturnedAttributesByCategory
	WScript.Echo "e.Refresh: " & e.Refresh
	WScript.Echo "e.RefreshComplete: " & e.RefreshComplete
	WScript.Echo "e.EnabledChanged: " & e.EnabledChanged 
	WScript.Echo "e.ActiveChanged: " & e.ActiveChanged 
	WScript.Echo "e.AcknowledgedChanged: " & e.AcknowledgedChanged 
	WScript.Echo "e.QualityChanged: " & e.QualityChanged  
	WScript.Echo "e.SeverityChanged: " & e.SeverityChanged 
	WScript.Echo "e.SubconditionChanged: " & e.SubconditionChanged 
	WScript.Echo "e.MessageChanged: " & e.MessageChanged
	WScript.Echo "e.AttributeChanged: " & e.AttributeChanged 
	WScript.Echo "e.EventData.QualifiedSourceName: " & e.EventData.QualifiedSourceName 
	WScript.Echo "e.EventData.Time: " & e.EventData.Time
	WScript.Echo "e.EventData.TimeLocal: " & e.EventData.TimeLocal
	WScript.Echo "e.EventData.Message: " & e.EventData.Message
	WScript.Echo "e.EventData.EventType: " & e.EventData.EventType 
	WScript.Echo "e.EventData.CategoryId: " & e.EventData.CategoryId 
	WScript.Echo "e.EventData.Severity: " & e.EventData.Severity 
    Rem IMPROVE: Display EventData.AttributeValues details
	WScript.Echo "e.EventData.AttributeValues: " & e.EventData.AttributeValues 
	WScript.Echo "e.EventData.ConditionName: " & e.EventData.ConditionName 
	WScript.Echo "e.EventData.SubconditionName: " & e.EventData.SubconditionName 
	WScript.Echo "e.EventData.Enabled: " & e.EventData.Enabled 
	WScript.Echo "e.EventData.Active: " & e.EventData.Active 
	WScript.Echo "e.EventData.Acknowledged: " & e.EventData.Acknowledged 
	WScript.Echo "e.EventData.Quality: " & e.EventData.Quality 
	WScript.Echo "e.EventData.AcknowledgeRequired: " & e.EventData.AcknowledgeRequired 
	WScript.Echo "e.EventData.ActiveTime: " & e.EventData.ActiveTime
	WScript.Echo "e.EventData.ActiveTimeLocal: " & e.EventData.ActiveTimeLocal
	WScript.Echo "e.EventData.Cookie: " & e.EventData.Cookie 
	WScript.Echo "e.EventData.ActorId: " & e.EventData.ActorId 
End Sub

Rem#endregion Example
