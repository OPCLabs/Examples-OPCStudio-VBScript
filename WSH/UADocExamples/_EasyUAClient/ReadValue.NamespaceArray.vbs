Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to read value of server's NamespaceArray, and display the namespace URIs in it.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Const endpointDescriptorUrlString = "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer"

' Instantiate the client object
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")

' Perform the operation
On Error Resume Next
Dim value: value = Client.ReadValue(endpointDescriptorUrlString, "i=2255")
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

' Work around the fact that elements of object arrays returned cannot be retrieved in VBScript directly.
Dim ElasticVector: Set ElasticVector = CreateObject("OpcLabs.BaseLib.Collections.ElasticVector")
ElasticVector.Assign(value)

' Display results
Dim i: For i = ElasticVector.LowerBound To ElasticVector.UpperBound
    WScript.Echo i & ": " & ElasticVector(i)
Next

' Example output:
'
'0: http://opcfoundation.org/UA/
'1: urn:DEMO-5:UA Sample Server
'2: http://test.org/UA/Data/
'3: http://test.org/UA/Data//Instance
'4: http://opcfoundation.org/UA/Boiler/
'5: http://opcfoundation.org/UA/Boiler//Instance
'6: http://opcfoundation.org/UA/Diagnostics
'7: http://samples.org/UA/memorybuffer
'8: http://samples.org/UA/memorybuffer/Instance

Rem#endregion Example
