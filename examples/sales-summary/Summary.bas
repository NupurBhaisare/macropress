Attribute VB_Name = "Summary"
Option Explicit

' The two buttons on the Summary sheet run the public macros below.

' Build summary: orders, units and revenue for each region, largest revenue first, with a total and each share.
Public Sub BuildSummary()
    Dim sales As Worksheet, out As Worksheet
    Set sales = ThisWorkbook.Worksheets("Sales")
    Set out = ThisWorkbook.Worksheets("Summary")

    out.Range("A2:E100").ClearContents
    out.Range("A3:E3").Value = Array("Region", "Orders", "Units", "Revenue", "Share")

    Dim lastRow As Long, r As Long, row As Long, regions As Long
    lastRow = sales.Cells(sales.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lastRow
        row = RegionRow(out, sales.Cells(r, "B").Value, regions)
        If row = 0 Then
            regions = regions + 1
            row = 3 + regions
            out.Cells(row, "A").Value = sales.Cells(r, "B").Value
        End If
        out.Cells(row, "B").Value = out.Cells(row, "B").Value + 1
        out.Cells(row, "C").Value = out.Cells(row, "C").Value + sales.Cells(r, "D").Value
        out.Cells(row, "D").Value = out.Cells(row, "D").Value + sales.Cells(r, "E").Value
    Next r

    Dim totalRow As Long
    totalRow = 4 + regions
    out.Range("A4:D" & totalRow - 1).Sort Key1:=out.Range("D4"), Order1:=xlDescending, Header:=xlNo
    out.Cells(totalRow, "A").Value = "Total"
    out.Cells(totalRow, "B").Formula = "=SUM(B4:B" & totalRow - 1 & ")"
    out.Cells(totalRow, "C").Formula = "=SUM(C4:C" & totalRow - 1 & ")"
    out.Cells(totalRow, "D").Formula = "=SUM(D4:D" & totalRow - 1 & ")"
    For row = 4 To totalRow
        out.Cells(row, "E").Formula = "=D" & row & "/D$" & totalRow
    Next row
    out.Range("A2").Value = "Built " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " from " & (lastRow - 1) & " sales"
End Sub

' Email summary: sends the workbook with Outlook. MacroPress does not model Outlook, so it stops this button and
' says so, rather than report a mail that never went out.
Public Sub EmailSummary()
    Dim outlook As Object, mail As Object
    Set outlook = CreateObject("Outlook.Application")
    Set mail = outlook.CreateItem(0)
    mail.To = "team@example.com"
    mail.Subject = "Sales by region"
    mail.Body = "The summary is in the attached workbook."
    mail.Attachments.Add ThisWorkbook.FullName
    mail.Send
End Sub

' The summary row of a region already listed, or 0.
Private Function RegionRow(ByVal out As Worksheet, ByVal region As String, ByVal regions As Long) As Long
    Dim row As Long
    For row = 4 To 3 + regions
        If out.Cells(row, "A").Value = region Then
            RegionRow = row
            Exit Function
        End If
    Next row
End Function
