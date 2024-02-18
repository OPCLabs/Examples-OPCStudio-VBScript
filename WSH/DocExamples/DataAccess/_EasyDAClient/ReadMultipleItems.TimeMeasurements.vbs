Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example measures the time needed to read 2000 items all at once, and in 20 groups by 100 items.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Const numberOfGroups = 100
Const itemsInGroup = 20
Dim totalItems: totalItems = numberOfGroups*itemsInGroup

Rem Make the measurements 10 times; note that first time the times might be longer.
Dim i: For i = 1 To 10

    Rem Pause - we do not want the component to use the values it has in memory
    WScript.Sleep 2*1000

    WScript.Echo
    WScript.Echo "Reading all items at once, and measuring the time..."
    Dim startTime1: startTime1 = Timer
    ReadAllAtOnce
    WScript.Echo "ReadAllAtOnce has taken (milliseconds): " & (Timer - startTime1)*1000

    Rem Pause - we do not want the component to use the values it has in memory
    WScript.Sleep 2*1000

    WScript.Echo
    WScript.Echo "Reading items in groups, and measuring the time..."
    Dim startTime2: startTime2 = Timer
    ReadInGroups
    WScript.Echo "ReadInGroups has taken (milliseconds): " & (Timer - startTime2)*1000
Next



Rem Read all items at once
Sub ReadAllAtOnce
    Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")

    Rem Create an array of item Ids for all items
    Dim arguments(): ReDim arguments(totalItems - 1)
    Dim index: index = 0
    Dim iLoop: For iLoop = 0 To numberOfGroups - 1
        Dim iItem: For iItem = 0 To itemsInGroup - 1
            Dim ReadItemArguments: Set ReadItemArguments = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAReadItemArguments")
            ReadItemArguments.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
            ReadItemArguments.ItemDescriptor.ItemID = "Simulation.Incrementing.Copy_" & (iLoop + 1) & ".Phase_" & (iItem + 1)

            Set arguments(index) = ReadItemArguments
            index = index + 1
        Next
    Next

    Rem Perform the OPC read
    Dim results: results = Client.ReadMultipleItems(arguments)

    Rem Count successful results
    Dim successCount: successCount = 0
    Dim i: For i = LBound(results) To UBound(results)
        Dim VtqResult: Set VtqResult = results(i)
        If VtqResult.Succeeded Then successCount = successCount + 1
    Next

    If successCount <> totalItems Then
        WScript.Echo "Warning: There were some failures, success count is " & successCount
    End If
End Sub



Rem Read items in groups
Sub ReadInGroups
    Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")

    Dim successCount: successCount = 0
    Dim arguments(): ReDim arguments(itemsInGroup - 1)
    Dim iLoop: For iLoop = 0 To numberOfGroups - 1
        'WScript.Echo iloop

        Rem Create an array of item Ids for items in one group
        Dim iItem: For iItem = 0 To itemsInGroup - 1
            Dim ReadItemArguments: Set ReadItemArguments = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAReadItemArguments")
            ReadItemArguments.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
            ReadItemArguments.ItemDescriptor.ItemID = "Simulation.Incrementing.Copy_" & (iLoop + 1) & ".Phase_" & (iItem + 1)
            Set arguments(iItem) = ReadItemArguments
        Next

        Rem Perform the OPC read
        Dim results: results = Client.ReadMultipleItems(arguments)

        Rem Count successful results (totalling to previous value)
        For iItem = LBound(results) To UBound(results)
            Dim VtqResult: Set VtqResult = results(iItem)
            If VtqResult.Succeeded Then successCount = successCount + 1
        Next
    Next

    If successCount <> totalItems Then
        WScript.Echo "Warning: There were some failures, success count is " & successCount
    End If
End Sub

Rem#endregion Example
