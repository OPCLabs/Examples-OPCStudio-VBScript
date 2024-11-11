Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to write values into 3 items at once, test for success of each write and display the exception 
Rem message in case of failure.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Dim ItemValueArguments1: Set ItemValueArguments1 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAItemValueArguments")
ItemValueArguments1.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
ItemValueArguments1.ItemDescriptor.ItemID = "Simulation.Register_I4"
ItemValueArguments1.Value = 23456

Dim ItemValueArguments2: Set ItemValueArguments2 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAItemValueArguments")
ItemValueArguments2.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
ItemValueArguments2.ItemDescriptor.ItemID = "Simulation.Register_R8"
ItemValueArguments2.Value = "This string cannot be converted to VT_R8"

Dim ItemValueArguments3: Set ItemValueArguments3 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAItemValueArguments")
ItemValueArguments3.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
ItemValueArguments3.ItemDescriptor.ItemID = "UnknownItem"
ItemValueArguments3.Value = "ABC"

Dim arguments(2)
Set arguments(0) = ItemValueArguments1
Set arguments(1) = ItemValueArguments2
Set arguments(2) = ItemValueArguments3

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")
Dim results: results = Client.WriteMultipleItemValues(arguments)

Dim i: For i = LBound(results) To UBound(results)
    Dim OperationResult: Set OperationResult = results(i)
    If OperationResult.Succeeded Then
        WScript.Echo "Result " & i & ": success"
    Else
        WScript.Echo "Result " & i & ": " & OperationResult.Exception.GetBaseException.Message
    End If
Next
Rem#endregion Example
