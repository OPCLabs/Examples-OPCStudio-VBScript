Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem For this example to work, you must first register the API simplifier using "RegSvr32 APISimplifier.wsc".
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

' Create the QuickOPC API simplifier object
Dim Simplifier: Set Simplifier = CreateObject("OPCLabs.QuickOPC.APISimplifier")

' Read value and display it
' Note: An exception can be thrown from the statement below in case of failure. See other examples for proper error 
' handling  practices!
WScript.Echo Simplifier.UARead("nsu=http://test.org/UA/Data/ ;i=10853")
