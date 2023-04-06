Attribute VB_Name = "Module2"
Option Compare Database
Option Explicit


' call this from Immediate Window: CreateReport_Crosstab 1996

Public Function CreateReport_Crosstab(yr As String) As Boolean

Dim db As DAO.Database
Dim qdf As DAO.QueryDef
Dim rst As DAO.Recordset
Dim fld As DAO.Field ' recordset field
Dim rpt As Report
Dim rptCtl As Control
Dim i As Integer
Dim c As Integer
Dim mdl As Module
Dim strQryName As String
Dim lngReturn As Long
Dim strCode As String
Dim strRptName As String
Dim txtNew As Access.TextBox ' textbox control
Dim lblNew As Access.Label ' label control
Dim lngTop As Long ' top value of control position
Dim lngLeft As Long ' left value of controls position

' initialise position variables
lngTop = 0
lngLeft = 0
     
' specify report and query name
strQryName = "Quarterly Orders by Product Crosstab"
strRptName = "myReport Product Crosstab"

Set db = CurrentDb()

'call function to create query
CreateCrossTab_Qry (strQryName)

' get the parameter query
Set qdf = db.QueryDefs(strQryName)

'provide the parameter value
qdf.Parameters(0) = yr

'Open a Recordset based on the parameter query
Set rst = qdf.OpenRecordset()

'Create the Report object
Set rpt = CreateReport

'Add ReportHeader/Footer to the empty report
RunCommand acCmdReportHdrFtr

rpt.Caption = strQryName

'Create report controls
Set rptCtl = CreateReportControl(rpt.Name, acImage, acHeader)
With rptCtl
    .Height = 400
    .Width = 400
    .Top = 0
    .Name = "Auto_Logo0"
    .Picture = CurrentProject.path & "\Images\redstar.gif"
End With

' add a Report Title to the ReportHeader
Set rptCtl = CreateReportControl(rpt.Name, acLabel, acHeader)

With rptCtl
    .Height = 400
    .Width = 2200
    .Top = 0
    .Left = 500
    .Name = "Auto_Header0"
    .Caption = "Custom Demo Report"
    .FontName = "Haettenschweiler"
    .FontSize = 18
    .TextAlign = 2
    .SizeToFit
End With

' Adjust Width of ReportHeader Section
With rpt
    .Section(1).Visible = True
    .Section(1).Height = 500
End With

Set rptCtl = CreateReportControl(rpt.Name, acLabel, acPageHeader)
With rptCtl
    .Height = 280
    .Width = 1200
    .Top = 0
    .Caption = "Order Year:"
End With

Set rptCtl = CreateReportControl(rpt.Name, acTextBox, acPageHeader)
With rptCtl
    .Height = 280
    .Width = 1080
    .Top = 0
    .Left = 1250
    .ControlSource = "=[Forms]![frmCrossCriteria]![cboYear]"
End With

' Create Label Title
 Set lblNew = CreateReportControl(rpt.Name, acLabel, _
 acPageHeader, , strQryName, 0, 0)
 With lblNew
    .Height = 280
    .Width = 4000
    .Top = 0
    .Left = 2500
    .FontBold = True
    .FontSize = 12
    .SizeToFit
End With

Debug.Print rst.Fields.Count

' Create corresponding label and text box controls for each field
For i = 0 To rst.Fields.Count - 1
    ' Create new text box control and size to fit data.
    Set txtNew = CreateReportControl(rpt.Name, acTextBox, _
    acDetail, , rst.Fields(i).Name, lngLeft, 0)
    If i = 0 Then
       txtNew.Width = 3500
    Else
       'txtNew.SizeToFit
    End If

    ' Create new label control and size to fit data.
    Set lblNew = CreateReportControl(rpt.Name, acLabel, _
     acPageHeader, , rst.Fields(i).Name, _
     lngLeft, 400, 1400, txtNew.Height)

   With lblNew
       .FontSize = 12
       .FontName = "Haettenschweiler"
       .TextAlign = 2
   End With

    ' Increment top value for next control
    If i = 0 Then
       lngLeft = lngLeft + txtNew.Width '+ 25
       Debug.Print lngLeft
    Else
       lngLeft = lngLeft + txtNew.Width '+ 25
        Debug.Print lngLeft
    End If
Next
     
With rpt
' Create timestamp on footer
    With CreateReportControl(.Name, acLabel, _
         acPageFooter, , Now(), 0, 0)
    End With

    .Section(acDetail).Height = 420
    .Section("detail").Height = 0
    .RecordSource = strQryName
    'make PageHeaderSection visible
    .Section(3).Visible = True
    .DefaultView = 0 'show report in Print Preview
End With

' return reference to report module
Set mdl = rpt.Module

' add event procedure
lngReturn = mdl.CreateEventProc("Open", "Report")
strCode = strCode & vbCrLf & "Dim strDocName As String"
strCode = strCode & vbCrLf & "strDocName = ""frmCrossCriteria"""
strCode = strCode & vbCrLf & "' Open form."
strCode = strCode & vbCrLf & "DoCmd.OpenForm strDocName"
strCode = strCode & ", , , , , acDialog"
strCode = strCode & vbCrLf & "If IsLoaded(strDocName) "
strCode = strCode & "= False Then Cancel = True"

' Insert text into body of procedure.
mdl.InsertLines lngReturn + 1, strCode

' Insert other event procedures from a text file
mdl.AddFromFile "C:\VBAAccess2021_ByExample\ReportEvents.txt"

' Print the name of report assigned by Access
Debug.Print rpt.Name

'save the report with specified name
DoCmd.Save , strRptName

'open the report in Print Preview
DoCmd.OpenReport strRptName, acViewPreview

Set rpt = Nothing
rst.Close

End Function


