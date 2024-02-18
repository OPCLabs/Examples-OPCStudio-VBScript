Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example shows how to read 4 items from the device, and display their values, timestamps and qualities.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

' Selects the data source for OPC reads (from device, from OPC cache, or dynamically determined).
' The data source (memory, OPC cache or OPC device) selection will be based on the desired value age and current status of 
' data received from the server.
Const DADataSource_ByValueAge = 0
' OPC reads will be fulfilled from the cache in the OPC server.
Const DADataSource_Cache = 1
' OPC reads will be fulfilled from the device by the OPC server.
Const DADataSource_Device = 2

Dim ReadItemArguments1: Set ReadItemArguments1 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAReadItemArguments")
ReadItemArguments1.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
ReadItemArguments1.ItemDescriptor.ItemID = "Simulation.Random"
ReadItemArguments1.ReadParameters.DataSource = DADataSource_Device  ' read will be from device

Dim ReadItemArguments2: Set ReadItemArguments2 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAReadItemArguments")
ReadItemArguments2.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
ReadItemArguments2.ItemDescriptor.ItemID = "Trends.Ramp (1 min)"
ReadItemArguments2.ReadParameters.DataSource = DADataSource_Device  ' read will be from device

Dim ReadItemArguments3: Set ReadItemArguments3 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAReadItemArguments")
ReadItemArguments3.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
ReadItemArguments3.ItemDescriptor.ItemID = "Trends.Sine (1 min)"
ReadItemArguments3.ReadParameters.DataSource = DADataSource_Device  ' read will be from device

Dim ReadItemArguments4: Set ReadItemArguments4 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAReadItemArguments")
ReadItemArguments4.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
ReadItemArguments4.ItemDescriptor.ItemID = "Simulation.Register_I4"
ReadItemArguments4.ReadParameters.DataSource = DADataSource_Device  ' read will be from device

Dim arguments(3)
Set arguments(0) = ReadItemArguments1
Set arguments(1) = ReadItemArguments2
Set arguments(2) = ReadItemArguments3
Set arguments(3) = ReadItemArguments4

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
