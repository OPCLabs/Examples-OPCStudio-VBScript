Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example shows all information available about categories that particular OPC servers do support.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Sub DumpServerElements(ByVal ServerElements)
Dim ServerElement: For Each ServerElement In ServerElements
    WScript.Echo "Categories of """ & ServerElement.ProgID & """:"
    With ServerElement.ServerCategories
        WScript.Echo Space(4) & ".OpcAlarmsAndEvents10: " & .OpcAlarmsAndEvents10
        WScript.Echo Space(4) & ".OpcDataAccess10: " & .OpcDataAccess10
        WScript.Echo Space(4) & ".OpcDataAccess20: " & .OpcDataAccess20
        WScript.Echo Space(4) & ".OpcDataAccess30: " & .OpcDataAccess30
        WScript.Echo Space(4) & ".ToString(): " & .ToString()
    End With
Next
End Sub



Dim DAClient: Set DAClient = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")
WScript.Echo
WScript.Echo "OPC DATA ACCESS"
On Error Resume Next
Dim DAServerElements: Set DAServerElements = DAClient.BrowseServers("")
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0
DumpServerElements DAServerElements

Dim AEClient: Set AEClient = CreateObject("OpcLabs.EasyOpc.AlarmsAndEvents.EasyAEClient")
WScript.Echo
WScript.Echo "OPC ALARMS AND EVENTS"
On Error Resume Next
Dim AEServerElements: Set AEServerElements = AEClient.BrowseServers("")
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0
DumpServerElements AEServerElements

Rem#endregion Example
