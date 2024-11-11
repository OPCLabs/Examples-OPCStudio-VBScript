Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example shows how to acknowledge an event condition in the OPC server.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Dim DAClient: Set DAClient = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")
Dim AEClient: Set AEClient = CreateObject("OpcLabs.EasyOpc.AlarmsAndEvents.EasyAEClient")

WScript.Echo "Hooking event handler..."
WScript.ConnectObject AEClient, "AEClient_"

Dim ServerDescriptor: Set ServerDescriptor = CreateObject("OpcLabs.EasyOpc.ServerDescriptor")
ServerDescriptor.ServerClass = "OPCLabs.KitEventServer.2"

Dim SourceDescriptor: Set SourceDescriptor = CreateObject("OpcLabs.EasyOpc.AlarmsAndEvents.AENodeDescriptor")
SourceDescriptor.QualifiedName = "Simulation.ConditionState1"

WScript.Echo "Processing event notifications for 1 minute..."
Dim SubscriptionParameters: Set SubscriptionParameters = CreateObject("OpcLabs.EasyOpc.AlarmsAndEvents.AESubscriptionParameters")
SubscriptionParameters.Filter.Sources = Array(SourceDescriptor)
SubscriptionParameters.NotificationRate = 1000
Dim handle: handle = AEClient.SubscribeEvents(ServerDescriptor, SubscriptionParameters, True, Nothing)

WScript.Echo "Give the refresh operation time to complete: Waiting for 5 seconds..."
WScript.Sleep 5*1000

WScript.Echo "Triggering an acknowledgeable event..."
On Error Resume Next
DAClient.WriteItemValue "", "OPCLabs.KitServer.2", "SimulateEvents.ConditionState1.Activate", True
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

Dim done: done = False
Dim endTime: endTime = Now() + 5*(1/24/60/60)
While (Not done) And (Now() < endTime)
    WScript.Sleep 1000
WEnd

WScript.Echo "Give some time to also receive the acknowledgement notification: Waiting for 5 seconds..."
WScript.Sleep 5*1000

WScript.Echo "Unsubscribing events..."
AEClient.UnsubscribeEvents handle

WScript.Echo "Unhooking event handler..."
WScript.DisconnectObject AEClient

WScript.Echo "Finished."



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
    	    WScript.Echo "EventData.AcknowledgeRequired: " & .AcknowledgeRequired

            If .AcknowledgeRequired Then
                WScript.Echo ">>>>> ACKNOWLEDGING THIS EVENT"
                On Error Resume Next
                AEClient.AcknowledgeCondition ServerDescriptor, SourceDescriptor, "Simulated", _
                    .ActiveTime, .Cookie, "aUser", ""
                If Err.Number <> 0 Then
                    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
                    Exit Sub
                End If
                On Error Goto 0
                WScript.Echo ">>>>> EVENT ACKNOWLEDGED"
                done = True
            End If
        End With
    End If
End Sub
Rem#endregion Example
