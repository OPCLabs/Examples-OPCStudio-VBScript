Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example shows how to obtain all nodes under the "Simulation" branch of the address space. For each node, it displays
Rem whether the node is a branch or a leaf.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Dim ServerDescriptor: Set ServerDescriptor = CreateObject("OpcLabs.EasyOpc.ServerDescriptor")
ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"

Dim NodeDescriptor: Set NodeDescriptor = CreateObject("OpcLabs.EasyOpc.DataAccess.DANodeDescriptor")
NodeDescriptor.ItemID = "Simulation"

Dim BrowseParameters: Set BrowseParameters = CreateObject("OpcLabs.EasyOpc.DataAccess.DABrowseParameters")

Dim Client: Set Client= CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")
On Error Resume Next
Dim NodeElements: Set NodeElements = Client.BrowseNodes(ServerDescriptor, NodeDescriptor, BrowseParameters)
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

Dim NodeElement: For Each NodeElement In NodeElements
    WScript.Echo "NodeElements(""" & NodeElement.Name & """):"
    With NodeElement
        WScript.Echo Space(4) & ".IsBranch: " & .IsBranch
        WScript.Echo Space(4) & ".IsLeaf: " & .IsLeaf
    End With
Next
Rem#endregion Example
