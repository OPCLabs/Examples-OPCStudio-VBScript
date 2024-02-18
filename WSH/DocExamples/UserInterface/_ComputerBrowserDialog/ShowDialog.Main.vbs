Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to let the user browse for computers on the network.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Const DialogResult_OK = 1

Dim Dialog: Set Dialog = CreateObject("OpcLabs.BaseLib.Forms.Browsing.Specialized.ComputerBrowserDialog")
Dim dialogResult: dialogResult = Dialog.ShowDialog
WScript.Echo dialogResult

If dialogResult <> DialogResult_OK Then
    WScript.Quit
End If

' Display results
WScript.Echo Dialog.SelectedName

Rem#endregion Example
