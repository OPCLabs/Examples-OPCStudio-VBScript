Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example shows how to work with Software Toolbox TOP Server 5 Alarms and Events.
Rem Use simdemo_WithA&E.opf configuration file and write a value above 1000 to Channel1.Device1.Tag1 or Channel1.Device1.Tag2.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Const AEEventTypes_All = 7

'Dim progID: progID = "Kepware.KEPServerEX_AE.V5"
Dim progID: progID = "SWToolbox.TOPServer_AE.V5"

Dim ServerDescriptor: Set ServerDescriptor = CreateObject("OpcLabs.EasyOpc.ServerDescriptor")
ServerDescriptor.ServerClass = progID

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.AlarmsAndEvents.EasyAEClient")

Rem Browse for some areas and sources

On Error Resume Next
Dim AreaElements: Set AreaElements = Client.BrowseAreas("", progID, "")
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

Dim AreaElement: For Each AreaElement In AreaElements
    WScript.Echo "AreaElements(""" & AreaElement.Name & """):"
    With AreaElement
        WScript.Echo Space(4) & ".QualifiedName: " & .QualifiedName
    End With

    On Error Resume Next
    Dim SourceElements: Set SourceElements = Client.BrowseSources("", progID, AreaElement.QualifiedName)
    If Err.Number <> 0 Then
        WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
        WScript.Quit
    End If
    On Error Goto 0

    Dim SourceElement: For Each SourceElement In SourceElements
        WScript.Echo Space(4) & "SourceElement(""" & SourceElement.Name & """):"
        With SourceElement
            WScript.Echo Space(8) & ".QualifiedName: " & .QualifiedName
        End With
    Next
Next

Rem Query for event categories

On Error Resume Next
Dim CategoryElements: Set CategoryElements = Client.QueryEventCategories(ServerDescriptor, AEEventTypes_All)
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

Dim CategoryElement: For Each CategoryElement In CategoryElements
    WScript.Echo "CategoryElements(" & CategoryElement.CategoryId & ").Description: " & CategoryElement.Description
Next

Rem Subscribe to events, wait, and unsubscribe

WScript.ConnectObject Client, "Client_"

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
