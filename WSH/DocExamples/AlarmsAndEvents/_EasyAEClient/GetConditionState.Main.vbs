Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example shows how to obtain current state information for the condition instance corresponding to a Source and 
Rem certain ConditionName.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Dim ServerDescriptor: Set ServerDescriptor = CreateObject("OpcLabs.EasyOpc.ServerDescriptor")
ServerDescriptor.ServerClass = "OPCLabs.KitEventServer.2"

Dim SourceDescriptor: Set SourceDescriptor = CreateObject("OpcLabs.EasyOpc.AlarmsAndEvents.AENodeDescriptor")
SourceDescriptor.QualifiedName = "Simulation.ConditionState1"

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.AlarmsAndEvents.EasyAEClient")
On Error Resume Next
Dim ConditionState: Set ConditionState = Client.GetConditionState(ServerDescriptor, SourceDescriptor, "Simulated", Array())
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

WScript.Echo "ConditionState:"
With ConditionState
    WScript.Echo Space(4) & ".ActiveSubcondition: " & .ActiveSubcondition
    WScript.Echo Space(4) & ".Enabled: " & .Enabled
    WScript.Echo Space(4) & ".Active: " & .Active
    WScript.Echo Space(4) & ".Acknowledged: " & .Acknowledged
    WScript.Echo Space(4) & ".Quality: " & .Quality
    Rem Note that IAEConditionState has many more properties
End With
Rem#endregion Example
