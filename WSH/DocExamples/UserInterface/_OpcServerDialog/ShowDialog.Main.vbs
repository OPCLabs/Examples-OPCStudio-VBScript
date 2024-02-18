Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to let the user browse for an OPC "Classic" server.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Const DialogResult_OK = 1

Dim ServerDialog: Set ServerDialog = CreateObject("OpcLabs.EasyOpc.Forms.Browsing.OpcServerDialog")
'ServerDialog.Location = ""
Dim dialogResult: dialogResult = ServerDialog.ShowDialog
WScript.Echo dialogResult

If dialogResult <> DialogResult_OK Then
    WScript.Quit
End If

' Display results
WScript.Echo ServerDialog.ServerElement

Rem#endregion Example
