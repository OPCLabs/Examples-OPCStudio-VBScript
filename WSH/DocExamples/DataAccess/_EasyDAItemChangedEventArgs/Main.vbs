Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example subscribes to changes of 2 items separately, and displays rich information available with each item changed
Rem event notification.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")
WScript.ConnectObject Client, "Client_"

Client.SubscribeItem "", "OPCLabs.KitServer.2", "Simulation.Random", 5*1000
Client.SubscribeItem "", "OPCLabs.KitServer.2", "Trends.Ramp (1 min)", 5*1000

WScript.Echo "Processing item changed events for 1 minute..."
WScript.Sleep 60*1000



Sub Client_ItemChanged(Sender, e)
    On Error Resume Next
	WScript.Echo
	WScript.Echo "e.Arguments.State: " & e.Arguments.State
	WScript.Echo "e.Arguments.ServerDescriptor.MachineName: " & e.Arguments.ServerDescriptor.MachineName
	WScript.Echo "e.Arguments.ServerDescriptor.ServerClass: " & e.Arguments.ServerDescriptor.ServerClass
	WScript.Echo "e.Arguments.ItemDescriptor.ItemId: " & e.Arguments.ItemDescriptor.ItemId
	WScript.Echo "e.Arguments.ItemDescriptor.AccessPath: " & e.Arguments.ItemDescriptor.AccessPath
	WScript.Echo "e.Arguments.ItemDescriptor.RequestedDataType: " & e.Arguments.ItemDescriptor.RequestedDataType
	WScript.Echo "e.Arguments.GroupParameters.Locale: " & e.Arguments.GroupParameters.Locale
	WScript.Echo "e.Arguments.GroupParameters.RequestedUpdateRate: " & e.Arguments.GroupParameters.RequestedUpdateRate
	WScript.Echo "e.Arguments.GroupParameters.PercentDeadband: " & e.Arguments.GroupParameters.PercentDeadband
	WScript.Echo "e.Exception.Message: " & e.Exception.Message
	WScript.Echo "e.Exception.Source: " & e.Exception.Source
	WScript.Echo "e.Exception.ErrorCode: " & e.Exception.ErrorCode
	WScript.Echo "e.Vtq.Value: " & e.Vtq.Value
	WScript.Echo "e.Vtq.Timestamp: " & e.Vtq.Timestamp
	WScript.Echo "e.Vtq.TimestampLocal: " & e.Vtq.TimestampLocal
	WScript.Echo "e.Vtq.Quality: " & e.Vtq.Quality
End Sub
Rem#endregion Example
