Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to read a range of values from an array.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Dim endpointDescriptor: endpointDescriptor = _
    "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    '"http://opcua.demo-this.com:51211/UA/SampleServer"  
    '"https://opcua.demo-this.com:51212/UA/SampleServer/"

' Instantiate the client object
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")

' Prepare the arguments, indicating that just the elements 2 to 4 should be returned.
Dim IndexRangeList: Set IndexRangeList = CreateObject("OpcLabs.EasyOpc.UA.UAIndexRangeList")
Dim IndexRange: Set IndexRange = CreateObject("OpcLabs.EasyOpc.UA.UAIndexRange")
IndexRange.Minimum = 2
IndexRange.Maximum = 4
IndexRangeList.Add IndexRange
'
Dim ReadArguments1: Set ReadArguments1 = CreateObject("OpcLabs.EasyOpc.UA.OperationModel.UAReadArguments")
ReadArguments1.EndpointDescriptor.UrlString = endpointDescriptor
ReadArguments1.NodeDescriptor.NodeId.ExpandedText = "nsu=http://test.org/UA/Data/ ;ns=2;i=10305"
ReadArguments1.IndexRangeList = IndexRangeList

Dim arguments(0)
Set arguments(0) = ReadArguments1

' Obtain the value.
Dim results: results = Client.ReadMultipleValues(arguments)
Dim ValueResult: Set ValueResult = results(0)
If Not ValueResult.Succeeded Then
    WScript.Echo "*** Failure: " & ValueResult.Exception.GetBaseException().Message
    WScript.Quit
End If
' VBScript can only handle well arrays of VARIANTs; most other COM tool will be able to use simply the .Value property.
Dim arrayValue: arrayValue = ValueResult.RegularizedValue

' Display results
Dim i: For i = 0 To 2
    WScript.Echo "arrayValue[" & i & "]: " & arrayValue(i)
Next


' Example output:
'arrayValue[0]: 180410224
'arrayValue[1]: 1919239969
'arrayValue[2]: 1700185172

Rem#endregion Example
