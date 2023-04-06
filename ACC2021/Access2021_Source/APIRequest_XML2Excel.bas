Attribute VB_Name = "APIRequest_XML2Excel"
Option Compare Database
Option Explicit

' set up a reference to the following Libraries
' --->Microsoft Excel 16.0 Object Library
' --->Microsoft XML, v6.0
' using the References dialog box

Sub RequestData_XML_toExcel()

Dim URL As String
Dim xlApp As Excel.Application
Dim wkb As Excel.Workbook
Dim ws As Excel.Worksheet
Dim httpReq As Object
Dim Resp As New MSXML2.DOMDocument60
Dim Header As Boolean
Dim n As Integer
Dim i As Integer
Dim c As Integer
Dim startTime As Single
Dim endTime As Single
Dim strExcelFile As String

On Error GoTo ErrorHandler

URL = "https://vpic.nhtsa.dot.gov/api/vehicles/getallmakes?format=xml"

strExcelFile = "C:\VBAAccess2021_ByExample\MakesList.xlsx"

' delete Excel file if it exists
If Dir(strExcelFile) <> "" Then Kill strExcelFile

Set httpReq = CreateObject("MSXML2.ServerXMLHttp")
httpReq.Open "GET", URL, False

httpReq.send
' check if we sucessfully made the request
If httpReq.status <> 200 Then
    Debug.Print httpReq.status & ":" & httpReq.statusText
    Set httpReq = Nothing
    Exit Sub
End If

'Assume we got through to the API
Debug.Print httpReq.responseText

Resp.LoadXML httpReq.responseText

' set up an Excel workbook for our data
Set xlApp = New Excel.Application
xlApp.Visible = True

Set wkb = xlApp.Workbooks.Add
Set ws = wkb.Sheets(1)
ws.Activate

Header = False
i = 1

ws.UsedRange.Clear

Dim NodeList As MSXML2.IXMLDOMNodeList
Dim iNode As MSXML2.IXMLDOMNode

startTime = Timer() ' start the timer

Set NodeList = Resp.getElementsByTagName("AllVehicleMakes")
If NodeList.length <> 0 Then
    For c = 0 To NodeList.length - 1
        For n = 0 To NodeList.item(c).childNodes.length - 1
            Set iNode = NodeList.item(c).childNodes(n)
            
            If Header = False Then
                ws.[A1].Offset(0, n).Value = iNode.baseName
            End If
            ws.[A1].Offset(i, n).Value = iNode.Text
        Next n
    i = i + 1
    ' this will get only the first 100 items from the response
    ' and will shorten the processing time
    ' comment it out to get all records (processing may take over 5 minutes)
    If i = 101 Then Exit For
    Next
       
    'autofit columns
    ws.Columns.AutoFit
    
    'save and close the workbook
    wkb.SaveAs strExcelFile
    wkb.Close
    
    endTime = Timer() ' end the timer
    
    Debug.Print "Processing completed in " & _
        Format(Round(endTime - startTime, 2), "#0.00") & " seconds."
    
    MsgBox "Processing completed in " & _
        Int((endTime - startTime) / 60) & " minutes, " & _
        Round(endTime - startTime, 2) Mod 60 & " seconds.", _
        vbInformation
    
    Debug.Print "Excel worksheet with All Vehicle Makes is completed."
Else
    Debug.Print "There was nothing to process. Please check for errors."
End If

ExitHere:
'cleanup
    Set httpReq = Nothing
    Set ws = Nothing
    Set wkb = Nothing
    xlApp.Quit
    Set xlApp = Nothing
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & ":" & Err.Description
    Resume ExitHere
End Sub

