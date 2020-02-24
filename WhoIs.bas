Attribute VB_Name = "WhoIs"
'this started from https://www.datanumen.com/blogs/create-whois-lookup-tool-via-excel-vba/
'but quickly mutated to using MSXML2.XMLHTTP so I could query RDAP instead of parsing html


Public Sub whoismacro()
    Dim v_lrow As Long
    Application.DisplayStatusBar = True
    Application.ScreenUpdating = True
    v_lrow = Sheets("Pivot Table").Range("A" & Rows.Count).End(xlUp).Row
    Dim r As Long
    
    For r = 2 To v_lrow
        Application.StatusBar = "Macro is running... Now fetching whois info for domain at Row : " & r & " /// Total Rows : " & v_lrow
        Sheets("Pivot Table").Range("C" & r, "E" & r).Value = whoislookup(Sheets("Pivot Table").Range("A" & r).Value)
    Next r
    Application.StatusBar = "Ready"
End Sub

Private Function whoislookup(ByRef theIP As Variant) As Variant

'On Error Resume Next

Dim msxmlhttp As Object, winhttp As Object, JSON As Object, i As Integer
Dim b, e, strUrl As String
Dim cidr, arr(2) As Variant

On Error Resume Next
    
    'start with ARIN, it will redirect to RIPE, AFRINIC, APNIC, LACNIC...but errors out with MSXML2
    strUrl = "https://rdap.arin.net/registry/ip/" & theIP
    
    Set msxmlhttp = CreateObject("MSXML2.XMLHTTP")
    msxmlhttp.Open "GET", strUrl, False
    msxmlhttp.send
    If msxmlhttp.responseText = "" Then
    'So it it is NOT ARIN, switch to WinHttp, and use different name of company
    'Use https://jsonformatter.org/json-parser to help find the right arrays of objects
        Set winhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
        winhttp.Open "GET", strUrl, False
        winhttp.send
        response = "[" & winhttp.responseText & "]"
        Set JSON = ParseJson(response)
        i = 1
        For Each Item In JSON
            b = Item("startAddress")
            e = Item("endAddress")
            arr(0) = Item("name")
            arr(1) = IPv4toCidr.IPv4toCidr(b, e)
            arr(2) = Item("country")
            i = i + 1
        Next
    Else
        response = "[" & msxmlhttp.responseText & "]"
        Set JSON = ParseJson(response)
        i = 1
        For Each Item In JSON
            b = Item("startAddress")
            e = Item("endAddress")
            arr(0) = Item("entities")(1)("vcardArray")(2)(2)(4)
            arr(1) = IPv4toCidr.IPv4toCidr(b, e)
            i = i + 1
        Next
    End If
    
whoislookup = arr

End Function


