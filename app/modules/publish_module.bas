Attribute VB_Name = "publish_module"
Sub publish_report()

'-------------------------------------------------------------------------------
answer = MsgBox("Želiš oddati poroèilo?", vbQuestion + vbYesNo + vbDefaultButton2, "")

If answer <> vbYes Then
    Exit Sub
End If

Set main = ThisWorkbook.Worksheets("report")

os_st = main.Range("D4").Value
month_d = main.Range("F4").Value
leto = main.Range("F5").Value
por_ure = main.Range("F8").Value

If os_st = "" Or month_d = "" Or leto = "" Or por_ure = "/" Then
    MsgBox ("Report filled out incorrectly_message")
    Exit Sub
End If

Dim dtToday As Date
dtToday = Date
dtReport = CDate(WorksheetFunction.Eomonth_d(DateSerial(leto, month_d, 1), 0))

If dtReport > dtToday Then
    answer = MsgBox("Future report ok?_message", vbQuestion + vbYesNo + vbDefaultButton2, "")
    
    If answer <> vbYes Then
        Exit Sub
    End If
End If
'------------------------------------------------------------------------------------------
datum1 = month_d & "_" & leto
ime_datoteke = os_st & "_" & datum1

Dim wb As Workbook
Set wb = Workbooks.Add
Set ws = wb.ActiveSheet

main.Cells.Copy ws.Cells
wb.ActiveSheet.Name = "report"
wb.Worksheets("report").Range("K:V").ClearContents

'odstrani vse VBA gumbe
Dim btn As Shape
For Each btn In wb.Worksheets("report").Shapes
    btn.Delete
Next

'odstrani vse povezave
Dim links As Variant
Dim x As Long
links = ActiveWorkbook.LinkSources(Type:=xlLinkTypeExcelLinks)
For x = 1 To UBound(links)
    wb.BreakLink Name:=links(x), Type:=xlLinkTypeExcelLinks
Next x

root_dir = "\\path\"
dir0 = root_dir & os_st
dir1 = dir0 & "\" & leto

If Dir(dir0, vbDirectory) = vbNullString Then
    MkDir dir0
End If

If Dir(dir1, vbDirectory) = vbNullString Then
    MkDir dir1
End If

wb.SaveAs fileName:=dir1 & "\" & ime_datoteke & ".xlsx"
wb.Close

MsgBox ("Success_message.")

'------------------------------------------------------------------------------------------
ThisWorkbook.Close savechanges:=False
'-------------------------------------------------------------------------------

End Sub
