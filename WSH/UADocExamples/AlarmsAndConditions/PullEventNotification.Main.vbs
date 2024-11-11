Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem This example shows how to subscribe to event notifications, pull events, and display each incoming event.
Rem
Rem Find all latest examples here: https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .
Rem OPC client and subscriber examples in VBScript on GitHub: https://github.com/OPCLabs/Examples-QuickOPC-VBScript .
Rem Missing some example? Ask us for it on our Online Forums, https://www.opclabs.com/forum/index ! You do not have to own
Rem a commercial license in order to use Online Forums, and we reply to every post.

Option Explicit

Const UAObjectIds_Server = "nsu=http://opcfoundation.org/UA/;i=2253"

' Instantiate the client object and hook events
Dim Client: Set Client = CreateObject("OpcLabs.EasyOpc.UA.EasyUAClient")
' In order to use event pull, you must set a non-zero queue capacity upfront.
Client.PullEventNotificationQueueCapacity = 1000

WScript.Echo "Subscribing..."
Client.SubscribeEvent "opc.tcp://opcua.demo-this.com:62544/Quickstarts/AlarmConditionServer", UAObjectIds_Server, 1000

WScript.Echo "Processing event notifications for 30 seconds..."
Dim endTime: endTime = Now() + 30*(1/24/60/60)
Do
    Dim EventArgs: Set EventArgs = Client.PullEventNotification(2*1000)
    If Not (EventArgs Is Nothing) Then
        ' Handle the notification event
        WScript.Echo EventArgs
    End If    
Loop While Now() < endTime



' Example output (truncated):
'Subscribing...
'Processing event notifications for 30 seconds...
'[] Success
'[] Success; Refresh; RefreshInitiated
'[] Success; Refresh; (10 field results) [EastTank] 100! "The dialog was activated" @9/10/2019 8:08:23 PM
'[] Success; Refresh; (10 field results) [EastTank] 500! "The alarm was acknoweledged." @10/14/2019 4:00:13 PM
'[] Success; Refresh; (10 field results) [EastTank] 100! "The alarm was acknoweledged." @11/9/2019 9:56:23 AM
'[] Success; Refresh; (10 field results) [EastTank] 500! "The alarm was acknoweledged." @10/14/2019 4:00:17 PM
'[] Success; Refresh; (10 field results) [EastTank] 100! "The alarm severity has increased." @9/10/2019 8:09:07 PM
'[] Success; Refresh; (10 field results) [EastTank] 100! "The alarm severity has increased." @9/10/2019 8:10:09 PM
'[] Success; Refresh; (10 field results) [NorthMotor] 100! "The dialog was activated" @9/10/2019 8:08:25 PM
'[] Success; Refresh; (10 field results) [NorthMotor] 500! "The alarm was acknoweledged." @10/14/2019 4:00:02 PM
'[] Success; Refresh; (10 field results) [NorthMotor] 500! "The alarm was acknoweledged." @10/14/2019 4:00:16 PM
'[] Success; Refresh; (10 field results) [NorthMotor] 300! "The alarm severity has increased." @11/9/2019 10:29:42 AM
'[] Success; Refresh; (10 field results) [NorthMotor] 100! "The alarm severity has increased." @9/10/2019 8:09:11 PM
'[] Success; Refresh; (10 field results) [NorthMotor] 100! "The alarm severity has increased." @9/10/2019 8:10:19 PM
'[] Success; Refresh; (10 field results) [WestTank] 100! "The dialog was activated" @9/10/2019 8:08:25 PM
'[] Success; Refresh; (10 field results) [WestTank] 300! "The alarm was acknoweledged." @10/14/2019 4:00:12 PM
'[] Success; Refresh; (10 field results) [WestTank] 300! "The alarm severity has increased." @11/9/2019 10:29:42 AM
'[] Success; Refresh; (10 field results) [WestTank] 300! "The alarm was acknoweledged." @10/14/2019 4:00:04 PM
'[] Success; Refresh; (10 field results) [WestTank] 100! "The alarm severity has increased." @9/10/2019 8:08:58 PM
'[] Success; Refresh; (10 field results) [WestTank] 100! "The alarm severity has increased." @9/10/2019 8:09:48 PM
'[] Success; Refresh; (10 field results) [SouthMotor] 100! "The dialog was activated" @9/10/2019 8:08:25 PM
'[] Success; Refresh; (10 field results) [SouthMotor] 300! "The alarm severity has increased." @11/9/2019 10:29:42 AM
'[] Success; Refresh; (10 field results) [SouthMotor] 700! "The alarm was acknoweledged." @10/14/2019 4:00:21 PM
'[] Success; Refresh; (10 field results) [SouthMotor] 500! "The alarm was acknoweledged." @10/14/2019 4:00:03 PM
'[] Success; Refresh; (10 field results) [SouthMotor] 100! "The alarm severity has increased." @9/10/2019 8:09:02 PM
'[] Success; Refresh; (10 field results) [SouthMotor] 100! "The alarm severity has increased." @9/10/2019 8:09:59 PM
'[] Success; Refresh; RefreshComplete
'[] Success; (10 field results) [Internal] 500! "Raising Events" @11/9/2019 10:29:43 AM
'[] Success; (10 field results) [Internal] 500! "Events Raised" @11/9/2019 10:29:43 AM
'[] Success; (10 field results) [Internal] 500! "Raising Events" @11/9/2019 10:29:44 AM
'[] Success; (10 field results) [Internal] 500! "Events Raised" @11/9/2019 10:29:44 AM
'[] Success; (10 field results) [Internal] 500! "Raising Events" @11/9/2019 10:29:45 AM
'[] Success; (10 field results) [Internal] 500! "Events Raised" @11/9/2019 10:29:45 AM
'[] Success; (10 field results) [Internal] 500! "Raising Events" @11/9/2019 10:29:46 AM
'[] Success; (10 field results) [Internal] 500! "Events Raised" @11/9/2019 10:29:46 AM
'[] Success; (10 field results) [Internal] 500! "Raising Events" @11/9/2019 10:29:47 AM
'[] Success; (10 field results) [Internal] 500! "Events Raised" @11/9/2019 10:29:47 AM
'[] Success; (10 field results) [Internal] 500! "Raising Events" @11/9/2019 10:29:48 AM
'[] Success; (10 field results) [Internal] 500! "Events Raised" @11/9/2019 10:29:48 AM
'[] Success; (10 field results) [Internal] 500! "Raising Events" @11/9/2019 10:29:49 AM
'[] Success; (10 field results) [Internal] 500! "Events Raised" @11/9/2019 10:29:49 AM
'[] Success; (10 field results) [Internal] 500! "Raising Events" @11/9/2019 10:29:50 AM
'[] Success; (10 field results) [Internal] 500! "Events Raised" @11/9/2019 10:29:50 AM
'[] Success; (10 field results) [Internal] 500! "Raising Events" @11/9/2019 10:29:51 AM
'[] Success; (10 field results) [Internal] 500! "Events Raised" @11/9/2019 10:29:51 AM
'[] Success; (10 field results) [Internal] 500! "Raising Events" @11/9/2019 10:29:52 AM
'[] Success; (10 field results) [Internal] 500! "Events Raised" @11/9/2019 10:29:52 AM
'[] Success; (10 field results) [Internal] 500! "Raising Events" @11/9/2019 10:29:53 AM
'[] Success; (10 field results) [NorthMotor] 500! "The alarm severity has increased." @11/9/2019 10:29:53 AM
'[] Success; (10 field results) [Internal] 500! "Events Raised" @11/9/2019 10:29:53 AM
'[] Success; (10 field results) [WestTank] 500! "The alarm severity has increased." @11/9/2019 10:29:53 AM
'[] Success; (10 field results) [SouthMotor] 500! "The alarm severity has increased." @11/9/2019 10:29:53 AM
'[] Success; (10 field results) [Internal] 500! "Raising Events" @11/9/2019 10:29:54 AM
'[] Success; (10 field results) [Internal] 500! "Events Raised" @11/9/2019 10:29:54 AM
'[] Success; (10 field results) [Internal] 500! "Raising Events" @11/9/2019 10:29:55 AM
'...

Rem#endregion Example
