Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example shows how to obtain all branches at the root of the address space. For each branch, it displays whether 
Rem it may have child nodes.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")
On Error Resume Next
Dim BranchElements: Set BranchElements = Client.BrowseBranches("", "OPCLabs.KitServer.2", "")
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

Dim BranchElement: For Each BranchElement In BranchElements
    WScript.Echo "BranchElements(""" & BranchElement.Name & """).HasChildren: " & BranchElement.HasChildren
Next
Rem#endregion Example
