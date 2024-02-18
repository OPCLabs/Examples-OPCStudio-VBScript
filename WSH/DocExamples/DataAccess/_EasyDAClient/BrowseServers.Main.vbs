Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example shows how to obtain all ProgIDs of all OPC Data Access servers on the local machine.
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
    WScript.Echo "ServerElements(""" & ServerElement.ClsidString & """).ProgId: " & ServerElement.ProgId
Next
Rem#endregion Example
