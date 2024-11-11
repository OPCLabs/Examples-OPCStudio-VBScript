Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example shows information available about OPC event category.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Const AEEventTypes_All = 7

Dim ServerDescriptor: Set ServerDescriptor = CreateObject("OpcLabs.EasyOpc.ServerDescriptor")
ServerDescriptor.ServerClass = "OPCLabs.KitEventServer.2"

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.AlarmsAndEvents.EasyAEClient")
On Error Resume Next
Dim CategoryElements: Set CategoryElements = Client.QueryEventCategories(ServerDescriptor, AEEventTypes_All)
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

Dim CategoryElement: For Each CategoryElement In CategoryElements
    WScript.Echo "Information about category " & CategoryElement & ":"
    With CategoryElement
        WScript.Echo Space(4) & ".CategoryId: " & .CategoryId
        WScript.Echo Space(4) & ".Description: " & .Description
        WScript.Echo Space(4) & ".ConditionElements:"
        Dim ConditionElement: For Each ConditionElement In .ConditionElements: WScript.Echo Space(8) & ConditionElement: Next
        WScript.Echo Space(4) & ".AttributeElements:"
        Dim AttributeElement: For Each AttributeElement In .AttributeElements: WScript.Echo Space(8) & AttributeElement: Next
    End With
Next
Rem#endregion Example
