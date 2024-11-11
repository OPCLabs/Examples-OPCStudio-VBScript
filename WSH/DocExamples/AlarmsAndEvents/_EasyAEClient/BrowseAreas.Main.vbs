Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example shows how to obtain all areas directly under the root (denoted by empty string for the parent).
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.AlarmsAndEvents.EasyAEClient")
On Error Resume Next
Dim NodeElements: Set NodeElements = Client.BrowseAreas("", "OPCLabs.KitEventServer.2", "")
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

Dim NodeElement: For Each NodeElement In NodeElements
    WScript.Echo "NodeElements(""" & NodeElement.Name & """):"
    With NodeElement
        WScript.Echo Space(4) & ".QualifiedName: " & .QualifiedName
    End With
Next
Rem#endregion Example
