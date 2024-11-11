Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example shows how to for a refresh for all active conditions and inactive, unacknowledged conditions.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Dim DAClient: Set DAClient = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")
Dim AEClient: Set AEClient = CreateObject("OpcLabs.EasyOpc.AlarmsAndEvents.EasyAEClient")
WScript.ConnectObject AEClient, "AEClient_"

WScript.Echo "Processing event notifications..."
Dim ServerDescriptor: Set ServerDescriptor = CreateObject("OpcLabs.EasyOpc.ServerDescriptor")
ServerDescriptor.ServerClass = "OPCLabs.KitEventServer.2"
Dim SourceDescriptor1: Set SourceDescriptor1 = CreateObject("OpcLabs.EasyOpc.AlarmsAndEvents.AENodeDescriptor")
SourceDescriptor1.QualifiedName = "Simulation.ConditionState1"
Dim SourceDescriptor2: Set SourceDescriptor2 = CreateObject("OpcLabs.EasyOpc.AlarmsAndEvents.AENodeDescriptor")
SourceDescriptor2.QualifiedName = "Simulation.ConditionState1"
Dim SourceDescriptor3: Set SourceDescriptor3 = CreateObject("OpcLabs.EasyOpc.AlarmsAndEvents.AENodeDescriptor")
SourceDescriptor3.QualifiedName = "Simulation.ConditionState1"
Dim SubscriptionParameters: Set SubscriptionParameters = CreateObject("OpcLabs.EasyOpc.AlarmsAndEvents.AESubscriptionParameters")
SubscriptionParameters.Filter.Sources = Array(SourceDescriptor1, SourceDescriptor2, SourceDescriptor3)
SubscriptionParameters.NotificationRate = 1000
Dim handle: handle = AEClient.SubscribeEvents(ServerDescriptor, SubscriptionParameters, True, Nothing)

Rem The component will perform auto-refresh at this point, give it time to happen
WScript.Echo "Waiting for 10 seconds..."
WScript.Sleep 10*1000

Rem Set some events to active state, which will cause them to appear in refresh
On Error Resume Next
WScript.Echo "Activating conditions and waiting for 10 seconds..."
DAClient.WriteItemValue "", "OPCLabs.KitServer.2", "SimulateEvents.ConditionState1.Activate", True
DAClient.WriteItemValue "", "OPCLabs.KitServer.2", "SimulateEvents.ConditionState2.Activate", True
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0
WScript.Sleep 10*1000

WScript.Echo "Refreshing subscription and waiting for 10 seconds..."
AEClient.RefreshEventSubscription handle
WScript.Sleep 10*1000

AEClient.UnsubscribeEvents handle



Rem Notification event handler
Sub AEClient_Notification(Sender, e)
    If Not (e.Succeeded) Then
        WScript.Echo "*** Failure: " & e.ErrorMessageBrief
        Exit Sub
    End If

    WScript.Echo 
    WScript.Echo "Refresh: " & e.Refresh
    WScript.Echo "RefreshComplete: " & e.RefreshComplete

    If Not (e.EventData Is Nothing) Then
        With e.EventData
    	    WScript.Echo "EventData.QualifiedSourceName: " & .QualifiedSourceName
    	    WScript.Echo "EventData.Message: " & .Message
    	    WScript.Echo "EventData.Active: " & .Active
    	    WScript.Echo "EventData.Acknowledged: " & .Acknowledged
        End With
    End If
End Sub
Rem#endregion Example
