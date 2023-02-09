Attribute VB_Name = "import_module"

Sub import_sub()
'------------------------------------------------------------------------------

os_st = ThisWorkbook.Worksheets("report").Range("D4").Value
If os_st = "" Then
    MsgBox ("Enter ID.")
    Exit Sub
End If

MsgBox ("Select old report file_message.")

Dim directory As String, fileName As String, sheet As Worksheet, total As Integer
Dim fd As Office.FileDialog

Set fd = Application.FileDialog(msoFileDialogFilePicker)

root_dir = "\\dir"
dir1 = root_dir & os_st & "\"

With fd
    .AllowMultiSelect = False
    .Title = "Select file."
    .Filters.Clear
    .InitialFileName = dir1
    
    If .Show = True Then
      datoteka_string = Dir(.SelectedItems(1))
    End If
End With

If datoteka_string = "" Then Exit Sub

Set datoteka = Workbooks.Open(datoteka_string, UpdateLinks:=False)
Set staro_porocilo = datoteka.Worksheets("report")
Set porocilo = ThisWorkbook.Worksheets("report")

porocilo.Range("F13").Value = staro_porocilo.Range("F13").Value
porocilo.Range("H12:H51").Value = staro_porocilo.Range("H12:H51").Value
porocilo.Range("C14:F51").Value = staro_porocilo.Range("C14:F51").Value

datoteka.Close savechanges:=False

'-------------------------------------------------------------------------------
End Sub
