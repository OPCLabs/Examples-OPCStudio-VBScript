Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example shows how to get value of multiple OPC properties, and handle errors.
Rem
Rem Note that some properties may not have a useful value initially (e.g. until the item is activated in a group), which also the
Rem case with Timestamp property as implemented by the demo server. This behavior is server-dependent, and normal. You can run 
Rem IEasyDAClient.ReadMultipleItemValues.Main.vbs shortly before this example, in order to obtain better property values. Your 
Rem code may also subscribe to the items in order to assure that they remain active.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Const Timestamp = 4
Const AccessRights = 5

' Get the values of Timestamp and AccessRights properties of two items.

Dim PropertyArguments1: Set PropertyArguments1 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAPropertyArguments")
PropertyArguments1.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
PropertyArguments1.NodeDescriptor.ItemID = "Simulation.Random"
PropertyArguments1.PropertyDescriptor.PropertyID.NumericalValue = Timestamp

Dim PropertyArguments2: Set PropertyArguments2 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAPropertyArguments")
PropertyArguments2.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
PropertyArguments2.NodeDescriptor.ItemID = "Simulation.Random"
PropertyArguments2.PropertyDescriptor.PropertyID.NumericalValue = AccessRights

Dim PropertyArguments3: Set PropertyArguments3 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAPropertyArguments")
PropertyArguments3.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
PropertyArguments3.NodeDescriptor.ItemID = "Trends.Ramp (1 min)"
PropertyArguments3.PropertyDescriptor.PropertyID.NumericalValue = Timestamp

Dim PropertyArguments4: Set PropertyArguments4 = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAPropertyArguments")
PropertyArguments4.ServerDescriptor.ServerClass = "OPCLabs.KitServer.2"
PropertyArguments4.NodeDescriptor.ItemID = "Trends.Ramp (1 min)"
PropertyArguments4.PropertyDescriptor.PropertyID.NumericalValue = AccessRights

Dim arguments(3)
Set arguments(0) = PropertyArguments1
Set arguments(1) = PropertyArguments2
Set arguments(2) = PropertyArguments3
Set arguments(3) = PropertyArguments4

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")
Dim results: results = Client.GetMultiplePropertyValues(arguments)

Dim i: For i = LBound(results) To UBound(results)
    If results(i).Exception Is Nothing Then 
        WScript.Echo "results(" & i & ").Value: " & results(i).Value
    Else
        WScript.Echo "results(" & i & ").Exception.Message: " & results(i).Exception.Message
    End If
Next
Rem#endregion Example
