Attribute VB_Name = "Invoice"
Option Explicit

' The three buttons on the Invoice sheet run the public macros below.

' Calculate totals: fills in each line's amount, then the subtotal, tax and total under the items.
Public Sub CalculateTotals()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Invoice")

    Dim lastRow As Long, r As Long
    lastRow = LastItemRow(ws)
    For r = 5 To lastRow
        ws.Cells(r, "D").Formula = "=B" & r & "*C" & r
    Next r

    ws.Cells(lastRow + 2, "C").Value = "Subtotal"
    ws.Cells(lastRow + 2, "D").Formula = "=SUM(D5:D" & lastRow & ")"
    ws.Cells(lastRow + 3, "C").Value = "Tax"
    ws.Cells(lastRow + 3, "D").Formula = "=ROUND(D" & lastRow + 2 & "*$B$2,2)"
    ws.Cells(lastRow + 4, "C").Value = "Total"
    ws.Cells(lastRow + 4, "D").Formula = "=D" & lastRow + 2 & "+D" & lastRow + 3
End Sub

' Apply discount: asks for a percentage, checks with the user, then lowers every unit price by it.
Public Sub ApplyDiscount()
    Dim answer As String
    answer = InputBox("Discount on every unit price, in percent:", "Apply discount", "10")
    If answer = "" Then Exit Sub
    If Not IsNumeric(answer) Then
        MsgBox "Please enter a number, like 10."
        Exit Sub
    End If
    If MsgBox("Lower every unit price by " & answer & "%?", vbYesNo + vbQuestion, "Apply discount") = vbYes Then
        LowerPrices CDbl(answer)
    End If
End Sub

' Lowers every unit price by pct percent, rounded to cents.
Public Sub LowerPrices(ByVal pct As Double)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets("Invoice")
    For r = 5 To LastItemRow(ws)
        ws.Cells(r, "C").Value = Round(ws.Cells(r, "C").Value * (1 - pct / 100), 2)
    Next r
End Sub

' Clear totals: removes what Calculate totals wrote, leaving the items as they are.
Public Sub ClearTotals()
    Dim ws As Worksheet, lastRow As Long
    Set ws = ThisWorkbook.Worksheets("Invoice")
    lastRow = LastItemRow(ws)
    ws.Range("D5:D" & lastRow).ClearContents
    ws.Range("C" & lastRow + 2 & ":D" & lastRow + 4).ClearContents
End Sub

' The row of the last item: the last filled cell in column A.
Private Function LastItemRow(ByVal ws As Worksheet) As Long
    LastItemRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
End Function
