Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to let the user browse for multiple OPC-UA nodes.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Const DialogResult_OK = 1

Const UAElementType_Host = 1

Dim BrowseDialog: Set BrowseDialog = CreateObject("OpcLabs.EasyOpc.UA.Forms.Browsing.UABrowseDialog")
BrowseDialog.InputsOutputs.CurrentNodeDescriptor.EndpointDescriptor.Host = "opcua.demo-this.com"
BrowseDialog.Mode.AnchorElementType = UAElementType_Host
BrowseDialog.Mode.MultiSelect = True


Dim dialogResult: dialogResult = BrowseDialog.ShowDialog
WScript.Echo dialogResult

If dialogResult <> DialogResult_OK Then
    WScript.Quit
End If

' Display results
Dim SelectionElements: Set SelectionElements = BrowseDialog.Outputs.SelectionElements
Dim i: For i = 0 To SelectionElements.Count - 1
    Dim Element: Set Element = SelectionElements(i)
    WScript.Echo "SelectionElements(" & i & "): " & Element.NodeElement
Next

Rem#endregion Example
