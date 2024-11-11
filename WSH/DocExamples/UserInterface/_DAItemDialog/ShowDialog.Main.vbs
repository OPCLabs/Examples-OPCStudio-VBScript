Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to let the user browse for an OPC Data Access item. 
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Const DialogResult_OK = 1

Dim ItemDialog: Set ItemDialog = CreateObject("OpcLabs.EasyOpc.DataAccess.Forms.Browsing.DAItemDialog")
ItemDialog.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
Dim dialogResult: dialogResult = ItemDialog.ShowDialog
WScript.Echo dialogResult

If dialogResult <> DialogResult_OK Then
    WScript.Quit
End If

' Display results
WScript.Echo "NodeElement: " & ItemDialog.NodeElement

Rem#endregion Example
