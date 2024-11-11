Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to subscribe to changes of a single monitored item, pull events, and display each change.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

' Instantiate the client object
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")
' In order to use event pull, you must set a non-zero queue capacity upfront.
Client.PullDataChangeNotificationQueueCapacity = 1000

WScript.Echo "Subscribing..."
Client.SubscribeDataChange "opc.tcp://opcua.demo-this.com:51210/UA/SampleServer", "nsu=http://test.org/UA/Data/ ;i=10853", 1000

WScript.Echo "Processing data change events for 1 minute..."
Dim endTime: endTime = Now() + 60*(1/24/60/60)
Do
    Dim EventArgs: Set EventArgs = Client.PullDataChangeNotification(2*1000)
    If Not (EventArgs Is Nothing) Then
        ' Handle the notification event
        WScript.Echo EventArgs
    End If    
Loop While Now() < endTime

Rem#endregion Example
