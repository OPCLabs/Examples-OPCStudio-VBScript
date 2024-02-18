Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows different ways of constructing generic data.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

' Create enumeration data with value of 1.
Dim EnumerationData: Set EnumerationData = CreateObject("OpcLabs.BaseLib.DataTypeModel.EnumerationData")
EnumerationData.Value = 1
WScript.Echo EnumerationData

' Create opaque data from an array of 2 bytes, specifying its size as 15 bits.
Dim OpaqueData1: Set OpaqueData1 = CreateObject("OpcLabs.BaseLib.DataTypeModel.OpaqueData")
OpaqueData1.SetByteArrayValue Array(&HAA, &H55), 15
WScript.Echo OpaqueData1

' Create opaque data from a bit array.
Dim OpaqueData2: Set OpaqueData2 = CreateObject("OpcLabs.BaseLib.DataTypeModel.OpaqueData")
Dim BitArray: Set BitArray = OpaqueData2.Value
BitArray.Length = 5
BitArray(0) = False
BitArray(1) = True
BitArray(2) = False
BitArray(3) = True
BitArray(4) = False
WScript.Echo OpaqueData2

' Create primitive data with System.Double value of 180.0.
Dim PrimitiveData1: Set PrimitiveData1 = CreateObject("OpcLabs.BaseLib.DataTypeModel.PrimitiveData")
PrimitiveData1.Value = 180.0
WScript.Echo PrimitiveData1

' Create primitive data with System.String value.
Dim PrimitiveData2: Set PrimitiveData2 = CreateObject("OpcLabs.BaseLib.DataTypeModel.PrimitiveData")
PrimitiveData2.Value = "Temperature is too high!"
WScript.Echo PrimitiveData2

' Create sequence data with two elements, using the Add method.
Dim SequenceData1: Set SequenceData1 = CreateObject("OpcLabs.BaseLib.DataTypeModel.SequenceData")
SequenceData1.Elements.Add OpaqueData1
SequenceData1.Elements.Add OpaqueData2
WScript.Echo SequenceData1

' Create structured data with two members, using the Add method.
Dim StructuredData1: Set StructuredData1 = CreateObject("OpcLabs.BaseLib.DataTypeModel.StructuredData")
StructuredData1.Add "Message", PrimitiveData2
StructuredData1.Add "Status", EnumerationData
WScript.Echo StructuredData1

Rem#endregion Example
