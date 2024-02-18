Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example

Rem This example shows use the output data (node) from one dialog invocation as input data in a subsequent dialog 
Rem invocation.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Const DialogResult_OK = 1

' Define the starting node for the first dialog.

Dim BrowseNodeDescriptor0: Set BrowseNodeDescriptor0 = CreateObject("OpcLabs.EasyOpc.UA.Forms.Browsing.UABrowseNodeDescriptor")
BrowseNodeDescriptor0.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
'BrowseNodeDescriptor0.NodeDescriptor.NodeId.StandardName = "Objects"

' Set the starting node of the first dialog, show the first dialog, and let the user select a first node.

Dim BrowseDialog1: Set BrowseDialog1 = CreateObject("OpcLabs.EasyOpc.UA.Forms.Browsing.UABrowseDialog")
Set BrowseDialog1.InputsOutputs.CurrentNodeDescriptor = BrowseNodeDescriptor0
Dim dialogResult1: dialogResult1 = BrowseDialog1.ShowDialog
If dialogResult1 <> DialogResult_OK Then
    WScript.Quit
End If
Dim BrowseNodeDescriptor1: Set BrowseNodeDescriptor1 = BrowseDialog1.InputsOutputs.CurrentNodeDescriptor

' Display the first node chosen.

WScript.Echo
WScript.Echo BrowseNodeDescriptor1.NodeDescriptor.NodeId
WScript.Echo BrowseNodeDescriptor1.NodeDescriptor.BrowsePath

' Set the starting node of the second dialog to be the node chosen by the user before, show the second dialog, and let the
' use select a second node.

Dim BrowseDialog2: Set BrowseDialog2 = CreateObject("OpcLabs.EasyOpc.UA.Forms.Browsing.UABrowseDialog")
Set BrowseDialog2.InputsOutputs.CurrentNodeDescriptor = BrowseNodeDescriptor1
Dim dialogResult2: dialogResult2 = BrowseDialog2.ShowDialog
If dialogResult2 <> DialogResult_OK Then
    WScript.Quit
End If
Dim BrowseNodeDescriptor2: Set BrowseNodeDescriptor2 = BrowseDialog2.InputsOutputs.CurrentNodeDescriptor

' Display the second node chosen.

WScript.Echo
WScript.Echo BrowseNodeDescriptor2.NodeDescriptor.NodeId
WScript.Echo BrowseNodeDescriptor2.NodeDescriptor.BrowsePath

Rem#endregion Example
