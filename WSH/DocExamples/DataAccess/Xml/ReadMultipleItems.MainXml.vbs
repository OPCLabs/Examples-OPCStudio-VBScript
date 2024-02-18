Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example shows how to read 4 items from an OPC XML-DA server at once, and display their values, timestamps
Rem and qualities.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Dim ReadItemArguments1: Set ReadItemArguments1 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAReadItemArguments")
ReadItemArguments1.ServerDescriptor.UrlString = "http://opcxml.demo-this.com/XmlDaSampleServer/Service.asmx"
ReadItemArguments1.ItemDescriptor.ItemID = "Dynamic/Analog Types/Double"

Dim ReadItemArguments2: Set ReadItemArguments2 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAReadItemArguments")
ReadItemArguments2.ServerDescriptor.UrlString = "http://opcxml.demo-this.com/XmlDaSampleServer/Service.asmx"
ReadItemArguments2.ItemDescriptor.ItemID = "Dynamic/Analog Types/Double[]"

Dim ReadItemArguments3: Set ReadItemArguments3 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAReadItemArguments")
ReadItemArguments3.ServerDescriptor.UrlString = "http://opcxml.demo-this.com/XmlDaSampleServer/Service.asmx"
ReadItemArguments3.ItemDescriptor.ItemID = "Dynamic/Analog Types/Int"

Dim ReadItemArguments4: Set ReadItemArguments4 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAReadItemArguments")
ReadItemArguments4.ServerDescriptor.UrlString = "http://opcxml.demo-this.com/XmlDaSampleServer/Service.asmx"
ReadItemArguments4.ItemDescriptor.ItemID = "SomeUnknownItem"

Dim arguments(3)
Set arguments(0) = ReadItemArguments1
Set arguments(1) = ReadItemArguments2
Set arguments(2) = ReadItemArguments3
Set arguments(3) = ReadItemArguments4

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")
Dim results: results = Client.ReadMultipleItems(arguments)

Dim i: For i = LBound(results) To UBound(results)
    Dim VtqResult: Set VtqResult = results(i)
    If Not (VtqResult.Exception Is Nothing) Then
        WScript.Echo "results(" & i & ") *** Failure: " & VtqResult.ErrorMessageBrief
    Else
        WScript.Echo "results(" & i & ").Vtq.ToString(): " & VtqResult.Vtq.ToString()
    End If
Next
Rem#endregion Example
