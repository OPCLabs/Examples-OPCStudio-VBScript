Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to write values, timestamps and qualities into 3 items at once.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Dim ItemVtqArguments1: Set ItemVtqArguments1 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAItemVtqArguments")
ItemVtqArguments1.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
ItemVtqArguments1.ItemDescriptor.ItemID = "Simulation.Register_I4"
ItemVtqArguments1.Vtq.Value = 23456
ItemVtqArguments1.Vtq.TimestampLocal = Now()
ItemVtqArguments1.Vtq.Quality.NumericalValue = 192

Dim ItemVtqArguments2: Set ItemVtqArguments2 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAItemVtqArguments")
ItemVtqArguments2.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
ItemVtqArguments2.ItemDescriptor.ItemID = "Simulation.Register_R8"
ItemVtqArguments2.Vtq.Value = 2.34567890
ItemVtqArguments2.Vtq.TimestampLocal = Now()
ItemVtqArguments2.Vtq.Quality.NumericalValue = 192

Dim ItemVtqArguments3: Set ItemVtqArguments3 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAItemVtqArguments")
ItemVtqArguments3.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
ItemVtqArguments3.ItemDescriptor.ItemID = "Simulation.Register_BSTR"
ItemVtqArguments3.Vtq.Value = "ABC"
ItemVtqArguments3.Vtq.TimestampLocal = Now()
ItemVtqArguments3.Vtq.Quality.NumericalValue = 192

Dim arguments(2)
Set arguments(0) = ItemVtqArguments1
Set arguments(1) = ItemVtqArguments2
Set arguments(2) = ItemVtqArguments3

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")
Dim results: results = Client.WriteMultipleItems(arguments)

Dim i: For i = LBound(results) To UBound(results)
    Dim OperationResult: Set OperationResult = results(i)
    If OperationResult.Succeeded Then
        WScript.Echo "Result " & i & ": success"
    Else
        WScript.Echo "Result " & i & ": " & OperationResult.Exception.GetBaseException.Message
    End If
Next
Rem#endregion Example
