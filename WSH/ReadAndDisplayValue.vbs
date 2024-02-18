Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

' Create EasyOPC-DA component 
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")

' Read item value and display it
' Note: An exception can be thrown from the statement below in case of failure. See other examples for proper error 
' handling  practices!
WScript.Echo Client.ReadItemValue("", "OPCLabs.KitServer.2", "Demo.Single")
