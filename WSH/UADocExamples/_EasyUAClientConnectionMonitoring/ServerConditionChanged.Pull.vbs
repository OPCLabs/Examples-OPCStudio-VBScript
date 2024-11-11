Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to monitor connections to and disconnections from the OPC UA server with event pull mechanism.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

' Instantiate the client object.
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")

' Obtain the client connection monitoring service.
Dim ClientConnectionMonitoring: Set ClientConnectionMonitoring = Client.GetServiceByName("OpcLabs.EasyOpc.UA.Services.IEasyUAClientConnectionMonitoring, OpcLabs.EasyOpcUA")
If ClientConnectionMonitoring Is Nothing Then
    WScript.Echo "The client connection monitoring service is not available."
    WScript.Quit
End If

' In order to use event pull, you must set a non-zero queue capacity upfront.
ClientConnectionMonitoring.PullServerConditionChangedQueueCapacity = 1000

WScript.Echo "Reading (1)"
' The first read will cause a connection to the server.
Dim AttributeData1: Set AttributeData1 = Client.Read("opc.tcp://opcua.demo-this.com:51210/UA/SampleServer", "nsu=http://test.org/UA/Data/ ;i=10853")
WScript.Echo AttributeData1

WScript.Echo "Reading (2)"
' The second read, because it closely follows the first one, will reuse the connection that is already open.
Dim AttributeData2: Set AttributeData2 = Client.Read("opc.tcp://opcua.demo-this.com:51210/UA/SampleServer", "nsu=http://test.org/UA/Data/ ;i=10853")
WScript.Echo AttributeData2

WScript.Echo "Processing server condition changed events for 10 seconds..."
' Because we only started the processing after we have made the reads, there are now events related to opening the 
' connection already in the queue, and they will be consumed first.
' Since the connection is now not used for some time, it will be closed.
Dim endTime: endTime = Now() + 10*(1/24/60/60)
Do
    Dim EventArgs: Set EventArgs = ClientConnectionMonitoring.PullServerConditionChanged(2*1000)
    If Not (EventArgs Is Nothing) Then
        ' Handle the server condition changed event.
        WScript.Echo EventArgs
    End If    
Loop While Now() < endTime

WScript.Echo "Finished."

Rem#endregion Example
