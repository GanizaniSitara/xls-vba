Sub MarkAutoSysJobs()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row

    ' Insert a new column to the right of column E
    ws.Columns("F").Insert Shift:=xlToRight

    ' Set up the regex pattern
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = False
    regex.IgnoreCase = True
    regex.Pattern = "\b[a-z]{3}_[a-z0-9]{1,3}_[a-z0-9_]{5,}\b"

    Dim i As Long
    For i = 1 To lastRow
        Dim cellValue As String
        cellValue = ws.Cells(i, "E").Value

        If regex.test(cellValue) Then
            ws.Cells(i, "F").Value = 1
        Else
            ws.Cells(i, "F").Value = 0
        End If
    Next i

    MsgBox "AutoSys job markers added to new column F."
End Sub
