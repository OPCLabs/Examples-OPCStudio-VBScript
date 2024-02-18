Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem Shows how to configure the OPC UA Complex Data plug-in to use a shared data type model provider.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

' Define which server and node we will work with.
Dim endpointDescriptor: endpointDescriptor = _
    "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"
    '"http://opcua.demo-this.com:51211/UA/SampleServer"  
    '"https://opcua.demo-this.com:51212/UA/SampleServer/"
Dim nodeDescriptor: nodeDescriptor = _
    "nsu=http://test.org/UA/Data/ ;i=10239"  ' [ObjectsFolder]/Data.Static.Scalar.StructureValue


' We will create two instances of EasyUAClient class, and configure each of them to use the shared data type
' model provider.

' Configure the first client object.
Dim Client1: Set Client1 = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")
Dim ComplexDataClientPluginParameters1: Set ComplexDataClientPluginParameters1 = Client1.InstanceParameters.PluginConfigurations.Find( _
    "OpcLabs.EasyOpc.UA.Plugins.ComplexData.UAComplexDataClientPluginParameters")
ComplexDataClientPluginParameters1.IsolatedDataTypeModelProvider = False

' Configure the second client object.
Dim Client2: Set Client2 = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")
Dim ComplexDataClientPluginParameters2: Set ComplexDataClientPluginParameters2 = Client2.InstanceParameters.PluginConfigurations.Find( _
    "OpcLabs.EasyOpc.UA.Plugins.ComplexData.UAComplexDataClientPluginParameters")
ComplexDataClientPluginParameters2.IsolatedDataTypeModelProvider = False

' We will now read the same complex data node using the two client objects.
'
' There is no noticeable difference in the results from the default state in which the client objects are
' set to use per-instance data type model provider. But, with the shared data type model provider, the metadata
' obtained during the read on the first client object and cached inside the data type model provider are reused
' during the read on the second client object, making this and the subsequent operations more efficient.

' Read the complex data node using the first client.
On Error Resume Next
Dim Value1: Set Value1 = Client1.ReadValue(endpointDescriptor, nodeDescriptor)
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0
WScript.Echo Value1

' Read the complex data node using the second client.
On Error Resume Next
Dim Value2: Set Value2 = Client2.ReadValue(endpointDescriptor, nodeDescriptor)
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0
WScript.Echo Value2

Rem#endregion Example
