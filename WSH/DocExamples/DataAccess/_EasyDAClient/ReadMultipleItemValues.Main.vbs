Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to read values of 4 items at once, and display them.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

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
Dim results: results = Client.ReadMultipleItemValues(arguments)

Dim i: For i = LBound(results) To UBound(results)
    Dim ValueResult: Set ValueResult = results(i)
    If ValueResult.Succeeded Then
        WScript.Echo "results(" & i & ").Value: " & ValueResult.Value
    Else
        WScript.Echo "results(" & i & ") *** Failure: " & ValueResult.ErrorMessageBrief
    End If
Next
Rem#endregion Example
