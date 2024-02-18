Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to write a value into a single OPC XML-DA item.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Dim ItemValueArguments1: Set ItemValueArguments1 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAItemValueArguments")
ItemValueArguments1.ServerDescriptor.UrlString = "http://opcxml.demo-this.com/XmlDaSampleServer/Service.asmx"
ItemValueArguments1.ItemDescriptor.ItemID = "Static/Analog Types/Int"
ItemValueArguments1.Value = 12345

Dim arguments(0)
Set arguments(0) = ItemValueArguments1

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")
Dim results: results = Client.WriteMultipleItemValues(arguments)

Dim OperationResult: Set OperationResult = results(0)
If OperationResult.Succeeded Then
    WScript.Echo "Result: success"
Else
    WScript.Echo "Result: " & OperationResult.Exception.GetBaseException.Message
End If
Rem#endregion Example
