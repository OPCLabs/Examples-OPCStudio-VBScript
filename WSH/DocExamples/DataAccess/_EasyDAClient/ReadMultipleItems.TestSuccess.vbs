Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example shows how to read 2 items (first valid, second invalid), test for success of each read and display either 
Rem the DAVtq or the Exception.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Dim ReadItemArguments1: Set ReadItemArguments1 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAReadItemArguments")
ReadItemArguments1.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
ReadItemArguments1.ItemDescriptor.ItemID = "Simulation.Random"

Dim ReadItemArguments2: Set ReadItemArguments2 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAReadItemArguments")
ReadItemArguments2.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
ReadItemArguments2.ItemDescriptor.ItemID = "UnknownItem"

Dim arguments(1)
Set arguments(0) = ReadItemArguments1
Set arguments(1) = ReadItemArguments2

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")
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
