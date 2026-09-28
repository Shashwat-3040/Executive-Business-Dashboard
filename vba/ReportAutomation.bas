Attribute VB_Name = "ReportAutomation"
'================================================================================
' ReportAutomation.bas
' Executive Business Performance 360 Dashboard - NorthStar Retail Group
'
' Automates the recurring reporting workflow: full recalculation, a timestamped
' snapshot of the KPI headlines, and an exportable PDF -- the kind of manual
' collation this workbook is designed to remove from someone's week.
'
' HOW TO IMPORT:
'   1. Open the workbook in Excel (Windows) and enable macros.
'   2. Press Alt+F11 to open the VBA editor.
'   3. File > Import File... and select this ReportAutomation.bas.
'   4. Save the workbook as .xlsm (macro-enabled) to keep the module attached.
'
' HOW TO RUN:
'   Alt+F8 > pick a macro below > Run. Or wire RefreshAllFormulas /
'   GenerateMonthlyReportSnapshot to a button on the Dashboard sheet
'   (Developer > Insert > Button, then Assign Macro).
'================================================================================

Option Explicit

' Runs automatically whenever the workbook is opened, so the dashboard is always
' showing current numbers without anyone remembering to hit F9 first.
Private Sub Workbook_Open()
    RefreshAllFormulas
End Sub

' -------------------------------------------------------------------------------
' RefreshAllFormulas
' Forces a full recalculation of every sheet and every open workbook. Useful
' after Raw_Sales / Dim_Customer / OpEx_Data / Inventory_Summary are refreshed
' from Power Query (Data > Refresh All) so every downstream KPI sheet is current.
' -------------------------------------------------------------------------------
Sub RefreshAllFormulas()
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    Application.StatusBar = "Refreshing all data connections and formulas..."

    ' Refresh any Power Query / external data connections first
    On Error Resume Next
    ThisWorkbook.RefreshAll
    On Error GoTo 0

    ' Then force a full formula recalculation (not just changed cells)
    Application.CalculateFullRebuild

    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Dashboard refreshed as of " & Format(Now, "mmm d, yyyy h:mm AM/PM") & ".", vbInformation, "Refresh Complete"
End Sub

' -------------------------------------------------------------------------------
' GenerateMonthlyReportSnapshot
' Copies the current KPI headlines from Dashboard / EBITDA / Margin_Leakage /
' Churn / Inventory_Turnover into a new, timestamped "Report_" sheet as static
' values -- a point-in-time snapshot for board packs or month-end archives,
' without disturbing the live formulas elsewhere in the workbook.
' -------------------------------------------------------------------------------
Sub GenerateMonthlyReportSnapshot()
    Dim reportName As String
    Dim ws As Worksheet
    Dim dash As Worksheet
    Dim r As Long

    reportName = "Report_" & Format(Now, "yyyy_mm_dd")

    ' Avoid duplicate sheet names if run twice in one day
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(reportName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = reportName
    Set dash = ThisWorkbook.Sheets("Dashboard")

    With ws
        .Range("B2").Value = "Executive KPI Snapshot"
        .Range("B2").Font.Bold = True
        .Range("B2").Font.Size = 14
        .Range("B3").Value = "Generated " & Format(Now, "mmmm d, yyyy h:mm AM/PM")
        .Range("B3").Font.Italic = True

        r = 6
        .Cells(r, 2).Value = "Metric": .Cells(r, 3).Value = "Value"
        .Range(.Cells(r, 2), .Cells(r, 3)).Font.Bold = True
        r = r + 1

        AddSnapshotRow ws, r, "Total Net Revenue (FY24-FY25)", dash.Range("B7").Value, "$#,##0": r = r + 1
        AddSnapshotRow ws, r, "YoY Revenue Growth", dash.Range("D7").Value, "0.0%": r = r + 1
        AddSnapshotRow ws, r, "Company EBITDA Margin", dash.Range("F7").Value, "0.0%": r = r + 1
        AddSnapshotRow ws, r, "Overall Customer Churn Rate", dash.Range("H7").Value, "0.0%": r = r + 1
        AddSnapshotRow ws, r, "Avg. Inventory Turnover (x/yr)", dash.Range("J7").Value, "0.00": r = r + 1
        AddSnapshotRow ws, r, "South Region Margin Leakage vs. Avg", dash.Range("L7").Value, "0.0%": r = r + 1

        .Columns("B:C").AutoFit
    End With

    MsgBox "Snapshot saved to sheet '" & reportName & "'.", vbInformation, "Report Generated"
End Sub

' Helper: writes one label/value row as static values (not formulas) so the
' snapshot stays frozen even after the live sheets recalculate later.
Private Sub AddSnapshotRow(ws As Worksheet, r As Long, label As String, val As Variant, fmt As String)
    ws.Cells(r, 2).Value = label
    ws.Cells(r, 3).Value = val
    ws.Cells(r, 3).NumberFormat = fmt
End Sub

' -------------------------------------------------------------------------------
' ExportDashboardToPDF
' Exports just the Dashboard sheet to a PDF next to the workbook, named with
' today's date -- the artifact that used to get built by hand for distribution.
' -------------------------------------------------------------------------------
Sub ExportDashboardToPDF()
    Dim pdfPath As String
    pdfPath = ThisWorkbook.Path & Application.PathSeparator & _
              "Executive_Dashboard_" & Format(Now, "yyyy_mm_dd") & ".pdf"

    ThisWorkbook.Sheets("Dashboard").ExportAsFixedFormat _
        Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

    MsgBox "Dashboard exported to:" & vbCrLf & pdfPath, vbInformation, "Export Complete"
End Sub
