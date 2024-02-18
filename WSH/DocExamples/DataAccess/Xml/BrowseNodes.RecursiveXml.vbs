Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example shows how to recursively browse the nodes in the OPC XML-DA address space.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Dim beginTime: beginTime = Timer
Dim branchCount: branchCount = 0
Dim leafCount: leafCount = 0

Dim ServerDescriptor: Set ServerDescriptor = CreateObject("OpcLabs.EasyOpc.ServerDescriptor")
ServerDescriptor.UrlString = "http://opcxml.demo-this.com/XmlDaSampleServer/Service.asmx"

Dim NodeDescriptor: Set NodeDescriptor = CreateObject("OpcLabs.EasyOpc.DataAccess.DANodeDescriptor")
NodeDescriptor.ItemID = ""

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")
On Error Resume Next
BrowseFromNode Client, ServerDescriptor, NodeDescriptor
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

Dim endTime: endTime = Timer

WScript.Echo ""
WScript.Echo "Browsing has taken (milliseconds): " & (endTime - beginTime) * 1000
WScript.Echo "Branch count: " & branchCount
WScript.Echo "Leaf count: " & leafCount


Sub BrowseFromNode(Client, ServerDescriptor, ParentNodeDescriptor)
    ' Obtain all node elements under ParentNodeDescriptor
    Dim BrowseParameters: Set BrowseParameters = CreateObject("OpcLabs.EasyOpc.DataAccess.DABrowseParameters")
    Dim NodeElementCollection: Set NodeElementCollection = Client.BrowseNodes(serverDescriptor, parentNodeDescriptor, browseParameters)
    ' Remark: that BrowseNodes(...) may also throw OpcException; a production code should contain handling for 
    ' it, here omitted for brevity.

    Dim NodeElement: For Each NodeElement In NodeElementCollection
        WScript.Echo NodeElement
        
        ' If the node is a branch, browse recursively into it.
        If NodeElement.IsBranch Then
            branchCount = branchCount + 1
            BrowseFromNode Client, ServerDescriptor, NodeElement.ToDANodeDescriptor
        Else
            leafCount = leafCount + 1
        End If
    Next

End Sub

Rem#endregion Example
