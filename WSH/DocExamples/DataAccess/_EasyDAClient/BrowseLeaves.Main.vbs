Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to obtain all leaves under the "Simulation" branch of the address space. For each leaf, it displays 
Rem the ItemID of the node.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")
On Error Resume Next
Dim LeafElements: Set LeafElements = Client.BrowseLeaves("", "OPCLabs.KitServer.2", "Simulation")
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

Dim LeafElement: For Each LeafElement In LeafElements
    WScript.Echo "LeafElements(""" & LeafElement.Name & """).ItemId: " & LeafElement.ItemId
Next
Rem#endregion Example
