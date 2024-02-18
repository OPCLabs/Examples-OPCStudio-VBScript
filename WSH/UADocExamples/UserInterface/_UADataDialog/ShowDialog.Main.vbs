Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to let the user browse for an OPC-UA data node (a Data Variable or a Property). 
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Const DialogResult_OK = 1

Dim DataDialog: Set DataDialog = CreateObject("OpcLabs.EasyOpc.UA.Forms.Browsing.UADataDialog")
DataDialog.EndpointDescriptor.UrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
DataDialog.UserPickEndpoint = True

Dim dialogResult: dialogResult = DataDialog.ShowDialog
WScript.Echo dialogResult

If dialogResult <> DialogResult_OK Then
    WScript.Quit
End If

' Display results
WScript.Echo "EndpointDescriptor: " & DataDialog.EndpointDescriptor
WScript.Echo "NodeElement: " & DataDialog.NodeElement

Rem#endregion Example
