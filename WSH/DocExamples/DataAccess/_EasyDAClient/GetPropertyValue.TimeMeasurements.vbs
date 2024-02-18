Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example measures the time needed to get values of all OPC properties of a single OPC item "one by one".
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Dim ServerDescriptor: Set ServerDescriptor = CreateObject("OpcLabs.EasyOpc.ServerDescriptor")
ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"

Dim NodeDescriptor: Set NodeDescriptor = CreateObject("OpcLabs.EasyOpc.DataAccess.DANodeDescriptor")
NodeDescriptor.ItemID = "Simulation.ReadValue_I4"

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")

Dim PropertyElementCollection
On Error Resume Next
Set PropertyElementCollection = Client.BrowseProperties(ServerDescriptor, NodeDescriptor)
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

'EasyDAClient.ReadItemValue "", "OPCLabs.KitServer.2", "Simulation.ReadValue_I4"
Dim startTime: startTime = Timer
Dim PropertyElement: For Each PropertyElement In PropertyElementCollection
    Dim propertyID: Set propertyID = PropertyElement.PropertyID
    On Error Resume Next
    Dim value: value = Client.GetPropertyValue("", "OPCLabs.KitServer.2", "Simulation.ReadValue_I4", propertyID.NumericalValue)
    If Err.Number <> 0 Then
        WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
        WScript.Quit
    End If
    On Error Goto 0
    'WScript.Echo value
Next
WScript.Echo "Time taken (milliseconds): " & (Timer - startTime)*1000

Rem#endregion Example
