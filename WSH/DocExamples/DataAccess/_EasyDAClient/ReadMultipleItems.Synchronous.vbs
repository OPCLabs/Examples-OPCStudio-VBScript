Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example shows how to read 4 items at once synchronously, and display their values, timestamps and qualities.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Dim ReadItemArguments1: Set ReadItemArguments1 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAReadItemArguments")
ReadItemArguments1.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
ReadItemArguments1.ItemDescriptor.ItemID = "Simulation.Random"

Dim ReadItemArguments2: Set ReadItemArguments2 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAReadItemArguments")
ReadItemArguments2.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
ReadItemArguments2.ItemDescriptor.ItemID = "Trends.Ramp (1 min)"

Dim ReadItemArguments3: Set ReadItemArguments3 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAReadItemArguments")
ReadItemArguments3.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
ReadItemArguments3.ItemDescriptor.ItemID = "Trends.Sine (1 min)"

Dim ReadItemArguments4: Set ReadItemArguments4 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAReadItemArguments")
ReadItemArguments4.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
ReadItemArguments4.ItemDescriptor.ItemID = "Simulation.Register_I4"

Dim arguments(3)
Set arguments(0) = ReadItemArguments1
Set arguments(1) = ReadItemArguments2
Set arguments(2) = ReadItemArguments3
Set arguments(3) = ReadItemArguments4

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")

' Specify that only synchronous method is allowed. By default, both synchronous and asynchronous methods are allowed, and 
' the component picks a suitable method automatically. Disallowing asynchronous method leaves only the synchronous method
' available for selection.
Client.InstanceParameters.Mode.AllowAsynchronousMethod = False

Dim results: results = Client.ReadMultipleItems(arguments)

Dim i: For i = LBound(results) To UBound(results)
    Dim VtqResult: Set VtqResult = results(i)
    If VtqResult.Succeeded Then
        WScript.Echo "results(" & i & ").Vtq.ToString(): " & VtqResult.Vtq.ToString()
    Else
        WScript.Echo "results(" & i & ") *** Failure: " & VtqResult.ErrorMessageBrief
    End If
Next
Rem#endregion Example
