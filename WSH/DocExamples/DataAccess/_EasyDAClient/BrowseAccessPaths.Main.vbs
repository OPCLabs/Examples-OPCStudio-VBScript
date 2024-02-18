Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example shows how to obtain all access paths available for an item.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Dim ServerDescriptor: Set ServerDescriptor = CreateObject("OpcLabs.EasyOpc.ServerDescriptor")
ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"

Dim NodeDescriptor: Set NodeDescriptor = CreateObject("OpcLabs.EasyOpc.DataAccess.DANodeDescriptor")
NodeDescriptor.ItemID = "Simulation.Random"

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")
On Error Resume Next
Dim accessPaths: accessPaths = Client.BrowseAccessPaths(ServerDescriptor, NodeDescriptor)
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

Dim i: For i = LBound(accessPaths) To UBound(accessPaths)
    WScript.Echo "accessPaths(" & i & "): " & accessPaths(i)
Next
Rem#endregion Example
