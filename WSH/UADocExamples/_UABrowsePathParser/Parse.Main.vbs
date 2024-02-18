Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem Parses an absolute  OPC-UA browse path and displays its starting node and elements.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Dim BrowsePathParser: Set BrowsePathParser = CreateObject("OpcLabs.EasyOpc.UA.Navigation.Parsing.UABrowsePathParser")
On Error Resume Next
Dim BrowsePath: Set BrowsePath = BrowsePathParser.Parse("[ObjectsFolder]/Data/Static/UserScalar")
If Err.Number <> 0 Then
    WScript.Echo "*** Failure: " & Err.Source & ": " & Err.Description
    WScript.Quit
End If
On Error Goto 0

' Display results
WScript.Echo "StartingNodeId: " & BrowsePath.StartingNodeId

WScript.Echo "Elements:"
Dim BrowsePathElement: For Each BrowsePathElement In BrowsePath.Elements
    WScript.Echo BrowsePathElement
Next

Rem#endregion Example
