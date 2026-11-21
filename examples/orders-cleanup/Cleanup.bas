Attribute VB_Name = "Cleanup"
Option Explicit

' The buttons on the Orders and Clean sheets run the public macros below.

' Clean up: copies the orders to the Clean sheet without blank rows or repeated orders, with tidy customer names,
' sorted by date. The Orders sheet is left as it was.
Public Sub CleanUpOrders()
    Dim orders As Worksheet, clean As Worksheet
    Set orders = ThisWorkbook.Worksheets("Orders")
    Set clean = ThisWorkbook.Worksheets("Clean")
    clean.Range("A2:E1000").ClearContents

    Dim lastRow As Long, r As Long, kept As Long, skipped As Long
    lastRow = orders.Cells(orders.Rows.Count, "A").End(xlUp).Row
    kept = 1
    For r = 2 To lastRow
        If IsEmpty(orders.Cells(r, "A").Value) Then
            skipped = skipped + 1
        ElseIf IsListed(clean, orders.Cells(r, "A").Value, kept) Then
            skipped = skipped + 1
        Else
            kept = kept + 1
            clean.Cells(kept, "A").Value = orders.Cells(r, "A").Value
            clean.Cells(kept, "B").Value = orders.Cells(r, "B").Value
            clean.Cells(kept, "C").Value = TitleCase(Trim(orders.Cells(r, "C").Value))
            clean.Cells(kept, "D").Value = orders.Cells(r, "D").Value
            clean.Cells(kept, "E").Value = orders.Cells(r, "E").Value
        End If
    Next r

    clean.Range("A1:E" & kept).Sort Key1:=clean.Range("B1"), Order1:=xlAscending, _
        Key2:=clean.Range("A1"), Order2:=xlAscending, Header:=xlYes
    clean.Activate
    Application.StatusBar = (kept - 1) & " orders kept, " & skipped & " rows skipped"
End Sub

' Show North only: filters the clean orders to the North region.
Public Sub ShowNorthOnly()
    ThisWorkbook.Worksheets("Clean").Range("A1").CurrentRegion.AutoFilter Field:=4, Criteria1:="North"
End Sub

' Show all orders: shows every clean order again.
Public Sub ShowAllOrders()
    With ThisWorkbook.Worksheets("Clean")
        If .FilterMode Then .ShowAllData
    End With
End Sub

' Whether order number id is already on the sheet, in rows 2 to lastRow.
Private Function IsListed(ByVal ws As Worksheet, ByVal id As Variant, ByVal lastRow As Long) As Boolean
    Dim r As Long
    For r = 2 To lastRow
        If ws.Cells(r, "A").Value = id Then
            IsListed = True
            Exit Function
        End If
    Next r
End Function

' "blue HARBOR supply" becomes "Blue Harbor Supply".
Private Function TitleCase(ByVal text As String) As String
    Dim words() As String, i As Long
    words = Split(LCase(text), " ")
    For i = LBound(words) To UBound(words)
        If Len(words(i)) > 0 Then words(i) = UCase(Left(words(i), 1)) & Mid(words(i), 2)
    Next i
    TitleCase = Join(words, " ")
End Function
