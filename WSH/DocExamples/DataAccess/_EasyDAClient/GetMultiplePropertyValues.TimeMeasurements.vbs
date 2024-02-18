Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example measures the time needed to get values of all OPC properties of a single OPC item all at once.
Rem This example shows how to get value of multiple OPC properties.
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

Dim count: count = PropertyElementCollection.Count

Dim arguments(): Redim arguments(count - 1)
Dim i: i = 0
Dim PropertyElement: For Each PropertyElement In PropertyElementCollection
    Dim PropertyArguments: Set PropertyArguments = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAPropertyArguments")
    PropertyArguments.ServerDescriptor = ServerDescriptor
    PropertyArguments.NodeDescriptor = NodeDescriptor
    PropertyArguments.PropertyDescriptor.PropertyID = PropertyElement.PropertyId

    Set arguments(i) = PropertyArguments
    i = i + 1
Next

'EasyDAClient.ReadItemValue "", "OPCLabs.KitServer.2", "Simulation.ReadValue_I4"
Dim startTime: startTime = Timer
Dim results: results = Client.GetMultiplePropertyValues(arguments)
WScript.Echo "Time taken (milliseconds): " & (Timer - startTime)*1000

'For i = LBound(results) To UBound(results)
'    If results(i).Exception Is Nothing Then 
'        WScript.Echo "results(" & i & ").Value: " & results(i).Value
'    Else
'        WScript.Echo "results(" & i & ").Exception.Message: " & results(i).Exception.Message
'    End If
''Next

Rem#endregion Example
