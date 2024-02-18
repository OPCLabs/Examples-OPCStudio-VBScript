Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example shows information available about OPC event attribute.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

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
    WScript.Echo "Category " & CategoryElement & ":"
    With CategoryElement
        Dim AttributeElement: For Each AttributeElement In .AttributeElements
            WScript.Echo Space(4) & "Information about attribute " & AttributeElement & ":"
            With AttributeElement
                WScript.Echo Space(8) & ".AttributeId: " & .AttributeId
                WScript.Echo Space(8) & ".Description: " & .Description
                WScript.Echo Space(8) & ".DataType: " & .DataType
            End With
        Next
    End With
Next
Rem#endregion Example
