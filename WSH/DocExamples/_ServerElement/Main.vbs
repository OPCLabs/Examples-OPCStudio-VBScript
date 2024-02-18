Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example shows all information available about OPC servers.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")
On Error Resume Next
Dim ServerElements: Set ServerElements = Client.BrowseServers("")
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

Dim ServerElement: For Each ServerElement In ServerElements
    WScript.Echo "Information about server """ & ServerElement & """:"
    With ServerElement
        WScript.Echo Space(4) & ".ServerClass: " & .ServerClass
        WScript.Echo Space(4) & ".ClsidString: " & .ClsidString
        WScript.Echo Space(4) & ".ProgId: " & .ProgId
        WScript.Echo Space(4) & ".Description: " & .Description
        WScript.Echo Space(4) & ".Vendor: " & .Vendor
        WScript.Echo Space(4) & ".ServerCategories.ToString(): " & .ServerCategories.ToString()
        WScript.Echo Space(4) & ".VersionIndependentProgId: " & .VersionIndependentProgId
    End With
Next
Rem#endregion Example
