
Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to browse objects under the "Objects" node and display notifiers.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

' Start browsing from the "Objects" node
Dim ObjectsNodeId: Set ObjectsNodeId = CreateObject("OpcLabs.EasyOpc.UA.AddressSpace.UANodeId")
ObjectsNodeId.StandardName = "Objects"
On Error Resume Next
BrowseFrom ObjectsNodeId
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0



Sub BrowseFrom(NodeId)
    Dim endpointDescriptor
    endpointDescriptor = "opc.tcp://opcua.demo-this.com:62544/Quickstarts/AlarmConditionServer"

    WScript.Echo 
    WScript.Echo 
    WScript.Echo "Parent node: " & NodeId

    ' Instantiate the client object
    Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")

    ' Obtain notifiers
    Dim NotifierNodeElementCollection: Set NotifierNodeElementCollection = Client.BrowseNotifiers( _
        endpointDescriptor, NodeId.ExpandedText)

    ' Display notifiers
    If NotifierNodeElementCollection.Count <> 0 Then
        WScript.Echo 
        WScript.Echo "Notifiers:"
        Dim NotifierNodeElement: For Each NotifierNodeElement In NotifierNodeElementCollection
            WScript.Echo NotifierNodeElement
        Next
    End If
    
    ' Obtain objects
    Dim ObjectNodeElementCollection: Set ObjectNodeElementCollection = Client.BrowseObjects( _
        endpointDescriptor, NodeId.ExpandedText)

    ' Recurse
    Dim ObjectNodeElement: For Each ObjectNodeElement In ObjectNodeElementCollection
        BrowseFrom ObjectNodeElement.NodeId
    Next
End Sub



' Example output (truncated):
'
'
'Parent node: ObjectsFolder
'
'
'Parent node: Server
'
'Notifiers:
'Green -> nsu=http://opcfoundation.org/Quickstarts/AlarmCondition ;ns=2;s=0:/Green (Object)
'Yellow -> nsu=http://opcfoundation.org/Quickstarts/AlarmCondition ;ns=2;s=0:/Yellow (Object)
'
'
'Parent node: Server_ServerCapabilities
'...

Rem#endregion Example
