Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example

Rem This example shows how the current node and selected nodes can be persisted between dialog invocations.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Const DialogResult_OK = 1

' The variables that persist the current and selected nodes.

Dim CurrentNodeDescriptor: Set CurrentNodeDescriptor = CreateObject("OpcLabs.EasyOpc.UA.Forms.Browsing.UABrowseNodeDescriptor")
Dim SelectionDescriptors: Set SelectionDescriptors = CreateObject("OpcLabs.EasyOpc.UA.Forms.Browsing.UABrowseNodeDescriptorCollection")

' The initial current node (optional).

CurrentNodeDescriptor.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"

' Repeatedly show the dialog until the user cancels it.

Do
    Dim BrowseDialog: Set BrowseDialog = CreateObject("OpcLabs.EasyOpc.UA.Forms.Browsing.UABrowseDialog")
    BrowseDialog.Mode.MultiSelect = True

    ' Set the dialog inputs from the persistence variables.

    Set BrowseDialog.InputsOutputs.CurrentNodeDescriptor = CurrentNodeDescriptor
    BrowseDialog.InputsOutputs.SelectionDescriptors.Clear
    Dim BrowseNodeDescriptor: For Each BrowseNodeDescriptor In SelectionDescriptors
        BrowseDialog.InputsOutputs.SelectionDescriptors.Add BrowseNodeDescriptor
    Next

    Dim dialogResult1: dialogResult1 = BrowseDialog.ShowDialog
    If dialogResult1 <> DialogResult_OK Then
        Exit Do
    End If

    ' Update the persistence variables with the dialog output.

    Set CurrentNodeDescriptor = BrowseDialog.InputsOutputs.CurrentNodeDescriptor
    selectionDescriptors.Clear
    For Each BrowseNodeDescriptor In BrowseDialog.InputsOutputs.SelectionDescriptors
        SelectionDescriptors.Add BrowseNodeDescriptor
    Next

Loop While True

Rem#endregion Example
