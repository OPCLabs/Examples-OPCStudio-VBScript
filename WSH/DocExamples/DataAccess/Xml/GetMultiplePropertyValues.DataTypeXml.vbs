Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example 
Rem This example shows how to obtain a data type of all OPC XML-DA items under a branch.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Const DABrowseFilter_Leaves = 3
Const DAPropertyIds_DataType = 1

Dim ServerDescriptor: Set ServerDescriptor = CreateObject("OpcLabs.EasyOpc.ServerDescriptor")
ServerDescriptor.UrlString = "http://opcxml.demo-this.com/XmlDaSampleServer/Service.asmx"

Dim NodeDescriptor: Set NodeDescriptor = CreateObject("OpcLabs.EasyOpc.DataAccess.DANodeDescriptor")
NodeDescriptor.ItemID = "Static/Analog Types"

Dim BrowseParameters: Set BrowseParameters = CreateObject("OpcLabs.EasyOpc.DataAccess.DABrowseParameters")
BrowseParameters.BrowseFilter = DABrowseFilter_Leaves

Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.DataAccess.EasyDAClient")

' Browse for all leaves under the "Static/Analog Types" branch
Dim NodeElementCollection: Set NodeElementCollection = client.BrowseNodes(serverDescriptor, nodeDescriptor, browseParameters)

' Create list of node descriptors, one for each leaf obtained
Dim arguments()
Redim arguments(nodeElementCollection.Count)
Dim i: i = 0
Dim NodeElement: For Each NodeElement In NodeElementCollection
    ' filter out hint leafs that do not represent real OPC XML-DA items (rare)
    If Not NodeElement.IsHint Then
	Dim PropertyArguments: Set PropertyArguments = CreateObject("OpcLabs.EasyOpc.DataAccess.OperationModel.DAPropertyArguments")
	PropertyArguments.ServerDescriptor = ServerDescriptor
	PropertyArguments.NodeDescriptor = NodeElement.ToDANodeDescriptor
        PropertyArguments.PropertyDescriptor.PropertyId.InternalValue = DAPropertyIds_DataType
	Set arguments(i) = PropertyArguments
	i = i + 1
    End If    
Next

Dim propertyArgumentArray()
ReDim propertyArgumentArray(i - 1)
Dim j: For j = 0 To i - 1
    Set propertyArgumentArray(j) = arguments(j)
Next

' Get the value of DataType property; it is a 16-bit signed integer
Dim valueResultArray: valueResultArray = client.GetMultiplePropertyValues(propertyArgumentArray)

For j = 0 To i - 1
    Dim NodeDescriptor2: Set NodeDescriptor2 = propertyArgumentArray(j).NodeDescriptor
    ' Check if there has been an error getting the property value
    If Not (valueResultArray(j).Exception Is Nothing) Then
        WScript.Echo NodeDescriptor2.NodeId & " *** Failure: " & valueResultArray(j).Exception.Message
    Else
        ' Display the obtained data type
        Dim VarType: Set VarType = CreateObject("OpcLabs.BaseLib.ComInterop.VarType")
        VarType.InternalValue = valueResultArray(j).Value
        WScript.Echo nodeDescriptor2.NodeId & ": " & VarType
    End If
Next

Rem#endregion Example
