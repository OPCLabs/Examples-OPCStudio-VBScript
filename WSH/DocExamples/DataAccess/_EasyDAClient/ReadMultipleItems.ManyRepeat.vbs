Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example repeatedly reads a large number of items.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Const repeatCount = 10
Const numberOfItems = 1000

WScript.Echo "Creating array of arguments..."
Dim arguments(): ReDim arguments(numberOfItems - 1)
Dim i: For i = 0 To numberOfItems - 1
    Dim copy: copy = Int(i / 100) + 1
    Dim phase: phase = i Mod 100
    Dim itemId: itemId = "Simulation.Incrementing.Copy_" & copy & ".Phase_" & phase
    WScript.Echo itemId

    Dim ReadItemArguments: Set ReadItemArguments = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAReadItemArguments")
    ReadItemArguments.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
    ReadItemArguments.ItemDescriptor.ItemID = itemId

    Set arguments(i) = ReadItemArguments
Next

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")

Dim iRepeat: For iRepeat = 1 To repeatCount
    WScript.Echo "Reading items..."
    Dim results: results = Client.ReadMultipleItems(arguments)

    Dim successCount: successCount = 0
    For i = LBound(results) To UBound(results)
        Dim VtqResult: Set VtqResult = results(i)
        If VtqResult.Succeeded Then successCount = successCount + 1
    Next
    WScript.Echo "Success count: " & successCount
Next
Rem#endregion Example
