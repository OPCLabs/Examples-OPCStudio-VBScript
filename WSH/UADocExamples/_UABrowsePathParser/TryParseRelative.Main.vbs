Rem $Header: $
Rem Copyright (c) CODE Consulting and Development, s.r.o., Plzen. All rights reserved.

Rem#region Example
Rem Attempts to parse a relative OPC-UA browse path and displays its elements.
Rem
Rem Find all latest examples here : https://opclabs.doc-that.com/files/onlinedocs/OPCLabs-OpcStudio/Latest/examples.html .

Option Explicit

Dim BrowsePathElements: Set BrowsePathElements = CreateObject("OpcLabs.EasyOpc.UA.Navigation.UABrowsePathElementCollection")

Dim BrowsePathParser: Set BrowsePathParser = CreateObject("OpcLabs.EasyOpc.UA.Navigation.Parsing.UABrowsePathParser")
Dim StringParsingError: Set StringParsingError = BrowsePathParser.TryParseRelative("/Data.Dynamic.Scalar.CycleComplete", BrowsePathElements)

' Display results
If Not (StringParsingError Is Nothing) Then
    WScript.Echo "*** Error: " & StringParsingError
    WScript.Quit
End If

Dim BrowsePathElement: For Each BrowsePathElement In BrowsePathElements
    WScript.Echo BrowsePathElement
Next

Rem#endregion Example
