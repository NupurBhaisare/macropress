"""The probes: small VBA macros that MacroPress and Excel both run, each writing what it finds to one row of the
workbook's Out sheet. Excel's answers are recorded in excel.json (see README.md here); the test compares.

A test is T(module, name, *lines). A line is VBA as it stands, or a probe, which writes one answer:
  S: expr   the expression as text         N: expr   as a number        V: expr   its type and value
  E: stmt   run the statement and record Err.Number                     X: expr   as text, Err left as it is
A test runs under On Error Resume Next, so a probe that fails records ERR and the error's number.

SHEETS lists the workbook's sheets as (name, column widths, rows), with {"protection": a <sheetProtection> element,
"state": "hidden"} after it for a sheet the file saves so. A cell is a value or (value, style); the styles are the
number formats in STYLES below.
"""

import base64
import datetime as dt
import hashlib


def q(s):
    return '"' + s.replace('"', '""') + '"'


HELPERS = '''Attribute VB_Name = "H"
Public Sub PutS(ByVal r As Long, ByVal c As Long, ByVal s As String)
    ThisWorkbook.Worksheets("Out").Cells(r, c + 1).Value = "[" & s & "]"
End Sub

Public Sub PutN(ByVal r As Long, ByVal c As Long, ByVal x As Double)
    ThisWorkbook.Worksheets("Out").Cells(r, c + 1).Value = x
End Sub

Public Function Show(ByVal v As Variant) As String
    If IsArray(v) Then
        Show = "Array"
    ElseIf IsError(v) Then
        Show = "Error:" & CStr(v)
    ElseIf IsEmpty(v) Then
        Show = "Empty"
    ElseIf IsNull(v) Then
        Show = "Null"
    ElseIf TypeName(v) = "Date" Then
        Show = "Date:" & CStr(CDbl(v))
    ElseIf TypeName(v) = "String" Then
        Show = "String:" & v
    Else
        Show = TypeName(v) & ":" & CStr(v)
    End If
End Function

Public Sub PutV(ByVal r As Long, ByVal c As Long, ByVal v As Variant)
    PutS r, c, Show(v)
End Sub

Public Function Vis(ByVal ws As Object) As String
    Dim r As Long, s As String
    For r = 2 To 13
        If ws.Rows(r).Hidden Then
            s = s & "0"
        Else
            s = s & "1"
        End If
    Next r
    Vis = s
End Function

Public Function ColJoin(ByVal ws As Object, ByVal col As Long, ByVal n As Long) As String
    Dim r As Long, s As String
    For r = 1 To n
        s = s & Show(ws.Cells(r, col).Value2) & "|"
    Next r
    ColJoin = s
End Function

Public Function RowJoin(ByVal ws As Object, ByVal row As Long, ByVal n As Long) As String
    Dim c As Long, s As String
    For c = 1 To n
        s = s & Show(ws.Cells(row, c).Value2) & "|"
    Next c
    RowJoin = s
End Function

Public Function FJoin(ByVal ws As Object, ByVal col As Long, ByVal n As Long) As String
    Dim r As Long, s As String
    For r = 1 To n
        s = s & ws.Cells(r, col).Formula & "|"
    Next r
    FJoin = s
End Function

Public Function NFJoin(ByVal ws As Object, ByVal col As Long, ByVal n As Long) As String
    Dim r As Long, s As String
    For r = 1 To n
        s = s & ws.Cells(r, col).NumberFormat & "|"
    Next r
    NFJoin = s
End Function
'''

TESTS = []
SHEETS = []  # (name, column widths, rows of cells), besides the Out sheet


def T(module, name, *lines):
    TESTS.append((module, "T_" + name, list(lines)))


def S(cell, sheet="S"):
    return f'ThisWorkbook.Worksheets("{sheet}").Range("{cell}")'


# --------------------------------------------------------------------------------------------------------------- numbers
T("TNum", "cstr1", "S: CStr(0.1 + 0.2)", "S: CStr(1 / 3)", "S: CStr(1E+20)", "S: CStr(1E+15)", "S: CStr(123456789012345#)",
  "S: CStr(0.00001)", "S: CStr(0.0000001)", "S: CStr(-0.5)", "S: CStr(100)", "S: CStr(1.5E+300)")
T("TNum", "cstr2", "S: CStr(2 / 3)", "S: CStr(1234567.891)", "S: CStr(0.1 * 3)", "S: CStr(1E+16)", "S: CStr(CSng(0.1))",
  "S: CStr(True)", "S: CStr(CCur(1.23456))", "S: CStr(CLng(5))", "S: 0.1 + 0.2 & \"\"", "S: CStr(1E+14 + 0.5)")
T("TNum", "str", "S: Str(5)", "S: Str(-5)", "S: Str(0.5)", "S: Str(1.5)", "S: CStr(0.5)", "S: 5 & \"\"", "S: 0.5 & \"\"",
  "S: -0.5 & \"\"", "S: CStr(-0.25)", "S: Str(100)")
T("TNum", "intdiv", "N: 7 \\ 2", "N: -7 \\ 2", "N: 7 Mod 3", "N: -7 Mod 3", "N: 7.5 Mod 2", "N: 6.5 \\ 2", "N: 7.5 \\ 2",
  "N: -2 ^ 2", "N: 2 ^ -1", "N: 10 / 4")
T("TNum", "round1", "N: Round(2.5)", "N: Round(3.5)", "N: Round(-2.5)", "N: Round(1.005, 2)", "N: Round(2.675, 2)",
  "N: Round(1.45, 1)", "N: Round(0.125, 2)", "N: Round(2.5, 0)", "N: Round(1234.5678, 2)", "N: Round(-0.5)")
T("TNum", "round2", "N: Round(1.015, 2)", "N: Round(1.025, 2)", "N: Round(8.345, 2)", "N: Round(0.285, 2)",
  "N: Round(1.115, 2)", "N: Round(5.055, 2)", "N: Round(2.345, 2)", "N: Round(10.075, 2)", "N: Round(0.5)", "N: Round(1.5)")
T("TNum", "cint", "N: CInt(2.5)", "N: CInt(3.5)", "N: CLng(-2.5)", "N: Int(-1.5)", "N: Fix(-1.5)", 'N: CInt("12.5")',
  'N: CLng("1e3")', 'N: CInt(" 7 ")', 'N: CDbl("1,234.5")', 'N: CInt("&H1F")')
T("TNum", "coerce", 'V: "5" + 5', 'V: "5" & 5', 'V: "5" + "5"', "V: True + 1", "V: CInt(True)", 'V: "abc" = "ABC"',
  'V: 1 = "1"', 'V: "10" > "9"', 'V: "10" > 9', 'V: 2 + 3 & "a"')
T("TNum", "variantcmp", "a = 1", 'b = "1"', "V: a = b", 'c = "10"', "d = 9", "V: c > d", "V: Empty = 0", 'V: Empty = ""',
  'V: u & "x"', "V: u + 1", "V: IsNull(Null + 1)", 'V: Null & "a"', "V: IsEmpty(u)", "V: a < b")
T("TNum", "overflow", "E: x = 32767 + 1", "E: x = 30000 * 2", "N: 30000& * 2", "E: x = 2147483647 + 1",
  "N: 2147483647# + 1", "v = 2147483647", "N: v + 1", "E: x = 1 / 0", "E: x = 0 / 0", "E: x = Sqr(-1)", "E: x = Log(0)")
T("TNum", "typedint", "Dim i As Integer", "i = 32767", "E: i = i + 1", "N: i", "i = 2.5", "N: i", "i = 3.5", "N: i",
  'E: i = "abc"', "E: i = 40000", "Dim y As Byte", "E: y = 256", "E: y = -1")
T("TNum", "typedmisc", "Dim b As Boolean", "b = 5", "V: b", "Dim s As String", "s = 1.5", "S: s", "Dim l As Long",
  "l = True", "N: l", "Dim c As Currency", "c = 1.23456", "S: CStr(c)", "Dim d As Date", "d = 45000.5",
  'S: Format(d, "yyyy-mm-dd hh:nn:ss")', "c = 1.00005", "S: CStr(c)", "c = 1.00015", "S: CStr(c)")
T("TNum", "bitwise", "V: 5 And 3", "V: 5 Or 3", "V: Not 5", "V: True And 5", "V: 5 Xor 1", "V: Not True", "V: 2 And 1",
  "V: 1 < 2 < 3", "V: Not 1 = 2", "V: True Or 1 / 1")
T("TNum", "bitif", "If 2 And 1 Then", 'x = "T"', "Else", 'x = "F"', "End If", "S: x", "If 2 Then", 'y = "T"', "Else", 'y = "F"',
  "End If", "S: y", 'If "abc" = "abc" And 3 Then', 'z = "T"', "Else", 'z = "F"', "End If", "S: z")
T("TNum", "prec", "N: 2 + 3 * 4 ^ 2", "N: -2 ^ 2", "N: 10 Mod 3 * 2", "N: 10 \\ 3 * 2", "N: 7 \\ 2 Mod 2", "N: 2 ^ 3 ^ 2",
  "S: 1 & 2 + 3", "N: 5 - 3 - 1", "N: 100 / 10 / 5", "N: -3 Mod 2")
T("TNum", "val", 'N: Val("12abc")', 'N: Val(" 1 2 3")', 'N: Val("1e2")', 'N: Val("&HFF")', 'N: Val(".5")', 'N: Val("1,234")',
  'N: Val("-")', 'N: Val("1.2.3")', 'N: Val("12e")', 'N: Val("&O17")')
T("TNum", "isnum1", 'V: IsNumeric("1,234")', 'V: IsNumeric("1e5")', 'V: IsNumeric("$5")', 'V: IsNumeric("")',
  'V: IsNumeric(" 5 ")', 'V: IsNumeric("5%")', 'V: IsNumeric("&H1F")', 'V: IsNumeric("1 2")', 'V: IsNumeric("(5)")',
  'V: IsNumeric("5-")')
T("TNum", "isnum2", 'V: IsNumeric("1.")', 'V: IsNumeric(".")', 'V: IsNumeric("1d2")', 'V: IsNumeric("+5")',
  'V: IsNumeric("1,2,3")', "V: IsNumeric(True)", "V: IsNumeric(u)", 'V: IsNumeric("12 ")', 'V: IsNumeric("1e")',
  "V: IsNumeric(#1/1/2024#)")
T("TNum", "cdbl", 'E: x = CDbl("$5")', 'N: CDbl("(5)")', 'E: x = CDbl("5%")', 'N: CDbl("1e3")', 'N: CDbl(" 12 ")',
  'N: CDbl("1,23")', 'E: x = CDbl("")', 'E: x = CDbl("abc")', 'N: CDbl("&H10")', 'N: CDbl("5-")')
T("TNum", "types1", "S: TypeName(1)", "S: TypeName(1.5)", "S: TypeName(100000)", "S: TypeName(1&)", 'S: TypeName("a")',
  "S: TypeName(True)", "S: TypeName(Now)", "S: TypeName(1 / 2)", "S: TypeName(2 * 3)", "S: TypeName(4 / 2)")
T("TNum", "types2", "S: TypeName(5 \\ 2)", "S: TypeName(2 ^ 2)", "S: TypeName(CCur(1))", "S: TypeName(u)", "S: TypeName(Null)",
  "S: TypeName(Nothing)", "S: TypeName(1 + 1.5)", "S: TypeName(1& + 1)", "S: TypeName(1 + True)", "S: TypeName(#1/1/2024# + 1)")
T("TNum", "strnum", 'V: "5" * "2"', 'V: "5" - 1', "V: 5 & 5", 'V: "1e2" + 1', 'V: "&H10" + 1', 'V: " 5 " + 1',
  'E: x = "5%" + 1', 'V: "1,000" + 1', 'V: "$5" + 1', 'E: x = "abc" + 1')

# --------------------------------------------------------------------------------------------------------------- strings
T("TStr", "sub1", 'S: Left("abc", 5)', 'S: Right("abc", 0)', 'S: Mid("abcdef", 2, 3)', 'S: Mid("abc", 5)', 'S: Mid("abc", 2)',
  'N: InStr("hello", "l")', 'N: InStr(4, "hello", "l")', 'N: InStr("hello", "")', 'N: InStr("", "a")', 'N: InStrRev("hello", "l")')
T("TStr", "sub2", 'N: InStr(1, "Hello", "h", vbTextCompare)', 'N: InStr("Hello", "h")', 'N: InStrRev("hello", "l", 3)',
  'E: x = Left("abc", -1)', 'E: x = Mid("abc", 0)', 'N: Len("")', 'N: Len(" a ")', 'S: Left("abc", 0)',
  'N: InStr(10, "hello", "l")', 'E: x = InStr(0, "a", "a")')
T("TStr", "repl", 'S: Replace("aaa", "a", "b", 1, 2)', 'S: Replace("abc", "", "x")', 'S: Replace("ABC", "b", "x")',
  'S: Replace("ABC", "b", "x", , , vbTextCompare)', 'S: Replace("abcabc", "b", "x", 3)', 'S: Replace("abc", "abc", "")',
  'S: Replace("aXbxc", "x", "-", 1, -1, vbTextCompare)', 'S: Replace("", "a", "b")', 'S: Replace("abc", "b", "")',
  'S: Replace("a.b", ".", "..")')
T("TStr", "split", 'N: UBound(Split("a,b,c", ","))', 'S: Split("a,,b", ",")(1)', 'N: UBound(Split("", ","))',
  'S: Split("a b  c")(2)', 'S: Split("a,b,c", ",", 2)(1)', 'S: Join(Split("a-b-c", "-"), "+")', 'S: Join(Array(1, 2, 3), "")',
  'S: Join(Array("a"), ",")', 'N: UBound(Split("a,b", ",,"))', 'S: Join(Array(1.5, True, "x"))')
T("TStr", "trim", 'S: Trim("  a  b  ")', 'N: Len(Trim(vbTab & "a"))', 'N: Len(Trim("   "))', 'S: LTrim("  a ") & "|"',
  'S: RTrim(" a  ") & "|"', 'S: UCase(ThisWorkbook.Worksheets("D").Range("A1").Value)', 'S: LCase(ThisWorkbook.Worksheets("D").Range("A2").Value)', 'S: UCase("abc1")',
  'N: Len(ThisWorkbook.Worksheets("D").Range("A1").Value)', 'S: StrReverse("abc")')
T("TStr", "cmp", 'N: StrComp("a", "B")', 'N: StrComp("a", "B", vbTextCompare)', 'V: "a" < "B"', 'V: "apple" < "Apple"',
  'V: "Z" < "a"', 'V: "a" = "a "', 'V: "" < "a"', 'N: StrComp("abc", "abc")', 'N: Asc("A")',
  'S: Chr(65) & String(3, "ab") & Space(2) & "x"')
T("TStr", "like", 'V: "abc" Like "a*"', 'V: "abc" Like "A*"', 'V: "a1" Like "a#"', 'V: "abc" Like "a?c"', 'V: "b" Like "[a-c]"',
  'V: "d" Like "[!a-c]"', 'V: "a*c" Like "a[*]c"', 'V: "" Like "*"', 'V: "ab" Like "a?c"', 'V: "aBc" Like "a[A-Z]c"')
T("TStr", "fmt1", 'S: Format(1234.5, "#,##0.00")', 'S: Format(1234.567, "0.00")', 'S: Format(-1234.5, "#,##0")',
  'S: Format(0.125, "0.00")', 'S: Format(2.5, "0")', 'S: Format(1.005, "0.00")', 'S: Format("abc", "@")', 'S: Format(0.5, "0.0")',
  'S: Format(3.5, "Standard")', 'S: Format(1234.5, "Fixed")')
T("TStr", "fmt2", "d = DateSerial(2024, 1, 5) + TimeSerial(13, 5, 9)", 'S: Format(d, "yyyy-mm-dd")', 'S: Format(d, "dd/mm/yyyy")',
  'S: Format(d, "mm/dd/yyyy")', 'S: Format(d, "hh:mm:ss")', 'S: Format(d, "hh:nn:ss")', 'S: Format(d, "yyyy-mm-dd hh:mm:ss")',
  'S: Format(d, "yyyymmdd")', 'S: Format(d, "hh:mm")', 'S: Format(d, "yyyy-mm-dd hh:nn:ss")')
T("TStr", "fmt3", "S: Format(1234.5)")
T("TStr", "fmt4", "S: Format(True)")
T("TStr", "fmt5", 'S: Format(45000, "yyyy-mm-dd")')
T("TStr", "cstrdate", "d = DateSerial(2024, 1, 5)", "S: CStr(d)", 'S: d & ""', "S: CStr(d + 0.5)", "S: CStr(TimeSerial(13, 5, 9))",
  "S: CStr(CDate(0))", "S: CStr(CDate(1.5))")

# --------------------------------------------------------------------------------------------------------------- dates
T("TDate", "dser", "N: CDbl(DateSerial(2024, 13, 1))", "N: CDbl(DateSerial(2024, 2, 30))", "N: CDbl(DateSerial(2024, 0, 1))",
  "N: CDbl(DateSerial(24, 1, 1))", "N: CDbl(DateSerial(30, 1, 1))", "N: CDbl(DateSerial(2024, 1, 0))",
  "N: CDbl(DateSerial(1900, 1, 1))", "N: CDbl(DateSerial(1899, 12, 30))", "N: CDbl(DateSerial(100, 1, 1))",
  "N: CDbl(TimeSerial(25, 0, 0))")
T("TDate", "dadd1", "d = DateSerial(2024, 1, 31)", 'N: CDbl(DateAdd("m", 1, d))',
  'N: CDbl(DateAdd("yyyy", 1, DateSerial(2024, 2, 29)))', 'N: CDbl(DateAdd("d", -1, d))', 'N: CDbl(DateAdd("ww", 2, d))',
  'N: CDbl(DateAdd("q", 1, d))')
T("TDate", "dadd2", "d = DateSerial(2024, 1, 31)", 'N: CDbl(DateAdd("h", 25, d))', 'N: CDbl(DateAdd("n", 90, d))',
  'N: CDbl(DateAdd("s", 86400, d))')
T("TDate", "dadd3", "d = DateSerial(2024, 1, 31)", 'N: CDbl(DateAdd("w", 1, d))', 'N: CDbl(DateAdd("y", 1, d))')
T("TDate", "ddiff1", "a = DateSerial(2023, 12, 31)", "b = DateSerial(2024, 1, 1)", 'N: DateDiff("d", a, b)', 'N: DateDiff("m", a, b)',
  'N: DateDiff("yyyy", a, b)', 'N: DateDiff("q", a, b)')
T("TDate", "ddiff2", "a = DateSerial(2023, 12, 30)", "b = DateSerial(2024, 1, 7)", 'N: DateDiff("ww", a, b)', 'N: DateDiff("w", a, b)')
T("TDate", "ddiff3", "a = DateSerial(2024, 1, 1) + TimeSerial(10, 59, 59)", "b = DateSerial(2024, 1, 1) + TimeSerial(11, 0, 0)",
  'N: DateDiff("h", a, b)', 'N: DateDiff("n", a, b)', 'N: DateDiff("s", a, b)', 'N: DateDiff("d", b, a)',
  'N: DateDiff("m", DateSerial(2024, 1, 31), DateSerial(2024, 2, 1))')
T("TDate", "parts", "d = DateSerial(2024, 1, 5) + TimeSerial(13, 5, 9)", "N: Weekday(d)", "N: Weekday(d, vbMonday)", "N: Year(d)",
  "N: Month(d)", "N: Day(d)", "N: Hour(d)", "N: Minute(d)", "N: Second(d)", "N: CDbl(Int(d))", "N: Weekday(DateSerial(2024, 1, 7))")
T("TDate", "cdate1", 'N: CDbl(CDate("1/5/2024"))', 'N: CDbl(CDate("13/5/2024"))')
T("TDate", "cdate2", 'N: CDbl(CDate("2024-01-05"))', 'N: CDbl(DateValue("2024-01-05 13:00"))')
T("TDate", "cdate3", 'N: CDbl(CDate("5 Jan 2024"))', 'N: CDbl(CDate("January 5, 2024"))')
T("TDate", "isdate", 'V: IsDate("1/5/2024")', 'V: IsDate("abc")', 'V: IsDate("13/13/2024")', 'V: IsDate("2024-02-30")',
  "V: IsDate(45000)", "V: IsDate(DateSerial(2024, 1, 1))")
T("TDate", "arith", "d = DateSerial(2024, 1, 5)", "S: TypeName(d + 1)", "N: CDbl(d + 1)", "S: TypeName(d - DateSerial(2024, 1, 1))",
  "N: d - DateSerial(2024, 1, 1)", "N: CLng(d)", "S: TypeName(CDate(45000))", "N: CDbl(#1/31/2024#)",
  "V: d > DateSerial(2023, 1, 1)")
T("TDate", "timeval", 'N: CDbl(TimeValue("13:05"))')

# --------------------------------------------------------------------------------------------------------------- cell writes
WRITES = ["123", "1e3", "TRUE", "true", " 12 ", "00123", "1,234", "50%", "$5", "(5)", "1 1/2", "1/2", "=1+1", "'abc", "abc", "",
          "0x1F", "1.5.2", "-", ".5", "5.", "1 000", "12-13", "Jan 5", "3:30", "1,23", "12:00 AM", "1e", "+5", "5-", "2024-01-05",
          "5 Jan 2024", "1-2-3", "5 %", "1.5e-3", "TRUE ", "#N/A", "#DIV/0!", "1:2:3", "0.5%", "1/2/3", "12/25", "25/12", "3 pm",
          "1.2e400", "-5", "1e5", "1E+5", "12.50", "0", "-0", "1.", "$1,234.50", "1 2", "FALSE", "yes", "1/0", "Feb 2024", "2e3 "]
for k, text in enumerate(WRITES, 1):
    cell = f'ThisWorkbook.Worksheets("S").Cells({k}, 1)'
    T("TCell", f"w{k:02d}", f"{cell}.Value = {q(text)}", f"S: TypeName({cell}.Value)", f"V: {cell}.Value2", f"S: {cell}.NumberFormat")
TEXT_WRITES = ['"00123"', '"123"', '"1/2"', '"TRUE"', '"=1+1"', "5", "True", "DateSerial(2024, 1, 5)", "1.5"]
for k, expr in enumerate(TEXT_WRITES, 1):
    cell = f'ThisWorkbook.Worksheets("TX").Cells({k}, 1)'
    T("TCell", f"wt{k}", f"{cell}.Value = {expr}", f"S: TypeName({cell}.Value)", f"V: {cell}.Value2", f"S: {cell}.NumberFormat")
T("TCell", "wtypes1", f"{S('C1')}.Value = True", f"V: {S('C1')}.Value", f"{S('C2')}.Value = DateSerial(2024, 1, 5)",
  f"S: TypeName({S('C2')}.Value)", f"V: {S('C2')}.Value2", f"{S('C3')}.Value = CVErr(xlErrNA)", f"V: {S('C3')}.Value",
  f"{S('C5')}.Value = CCur(1.5)", f"S: TypeName({S('C5')}.Value)", f"{S('C8')}.Value = CLng(5)", f"S: TypeName({S('C8')}.Value)")
T("TCell", "wtypes2", f"{S('C6')}.Value = 5", f"{S('C6')}.Value = Empty", f"V: {S('C6')}.Value", f"E: {S('C7')}.Value = Null",
  f"V: {S('C7')}.Value", f"{S('C10')}.Value = DateSerial(2024, 1, 5) + TimeSerial(13, 5, 9)", f"V: {S('C10')}.Value2",
  f"S: TypeName({S('C10')}.Value)", f"{S('C11')}.Value = 1E+308", f"V: {S('C11')}.Value")
T("TCell", "formula_r", f"{S('E1')}.Value = 5", f"S: TypeName({S('E1')}.Formula) & \":\" & {S('E1')}.Formula",
  f"{S('E2')}.Value = \"abc\"", f"S: {S('E2')}.Formula", f"S: \"<\" & {S('E3')}.Formula & \">\"", f"{S('E4')}.Value = True",
  f"S: {S('E4')}.Formula", f"{S('E5')}.Value = DateSerial(2024, 1, 5)", f"S: {S('E5')}.Formula")
T("TCell", "formula_w", f"{S('E1')}.Value = 5", f"{S('E6')}.Formula = \"=E1+1\"", f"V: {S('E6')}.Value", f"S: {S('E6')}.Formula",
  f"{S('E7')}.Formula = \"abc\"", f"V: {S('E7')}.Value", f"{S('E8')}.Formula = \"5\"", f"S: TypeName({S('E8')}.Value)",
  f"{S('E9')}.Value = \"=E1*2\"", f"V: {S('E9')}.Value", f"S: {S('E9')}.Formula")
T("TCell", "multi1", f"{S('G1:H2')}.Value = 5", f"N: {S('H2')}.Value", f"arr = {S('G1:H2')}.Value", "N: LBound(arr, 1)",
  "N: UBound(arr, 1)", "N: UBound(arr, 2)", f"{S('J1:K2')}.Value = arr", f"N: {S('K2')}.Value", f"V: IsArray({S('G1')}.Value)",
  f"x = {S('G1')}", "V: x")
T("TCell", "multi2", f"{S('G5:I5')}.Value = Array(1, 2, 3)", f"N: {S('I5')}.Value", f"{S('G6:G8')}.Value = Array(1, 2, 3)",
  f"N: {S('G8')}.Value", f"{S('J5:L6')}.Value = Array(1, 2)", f"V: {S('L5')}.Value", f"V: {S('K6')}.Value")
T("TCell", "val2val", f"{S('M1')}.Value = 2", f"{S('M2')}.Formula = \"=M1*10\"", f"{S('M2')}.Value = {S('M2')}.Value",
  f"S: {S('M2')}.Formula", f"{S('M3')}.Formula = \"=M1+1\"", f"{S('M4')}.Formula = \"=M3+1\"",
  f"{S('M3:M4')}.Value = {S('M3:M4')}.Value", f"S: {S('M4')}.Formula", f"V: {S('M4')}.Value")
T("TCell", "addr1", f"S: {S('A1:C3')}.Address", f"S: {S('A1:C3')}.Address(False, False)", f"N: {S('A1:C3')}.Count",
  f"N: {S('A1:C3')}.Rows.Count", f"N: {S('A1:C3')}.Columns.Count", f"S: {S('B2')}.Offset(1, 2).Address",
  f"S: {S('B2')}.Resize(2, 3).Address", 'S: ThisWorkbook.Worksheets("S").Cells(2, 3).Address',
  'S: ThisWorkbook.Worksheets("S").Range(ThisWorkbook.Worksheets("S").Cells(1, 1), ThisWorkbook.Worksheets("S").Cells(2, 2)).Address', 'S: ThisWorkbook.Worksheets("S").Range("A1", "B2").Address')
T("TCell", "addr2", 'S: ThisWorkbook.Worksheets("S").Columns(2).Address', 'S: ThisWorkbook.Worksheets("S").Rows("2:3").Address', f"N: {S('C5')}.Row", f"N: {S('C5')}.Column",
  f"S: {S('A1:C3')}.Cells(2, 2).Address", f"S: {S('A1:C3')}(5).Address", f"S: {S('B2:C3')}.Range(\"A1\").Address",
  f"S: {S('B2')}.EntireRow.Address", f"S: {S('B2:C3')}.Cells(3, 3).Address", f"N: {S('A:A')}.Count")
T("TCell", "addr3", f"S: Intersect({S('A1:C3')}, {S('B2:D4')}).Address", f"V: Intersect({S('A1')}, {S('B2')}) Is Nothing",
  f"S: {S('A1:B2')}.Offset(1).Address", f"S: {S('A1')}.Resize(, 3).Address", f"S: {S('B2:C3')}.Rows(2).Address",
  f"S: {S('B2:C3')}.Columns(1).Address", 'S: ThisWorkbook.Worksheets("S").Rows(2).Cells(3).Address', f"E: x = {S('A1')}.Offset(-1, 0).Address",
  'E: x = ThisWorkbook.Worksheets("S").Cells(0, 1).Value', 'E: x = ThisWorkbook.Worksheets("S").Range("ZZZZ1").Value')
T("TCell", "end1", 'N: ThisWorkbook.Worksheets("E").Cells(ThisWorkbook.Worksheets("E").Rows.Count, 1).End(xlUp).Row', 'N: ThisWorkbook.Worksheets("E").Range("A1").End(xlDown).Row',
  'N: ThisWorkbook.Worksheets("E").Range("A5").End(xlDown).Row', 'N: ThisWorkbook.Worksheets("E").Range("A8").End(xlDown).Row', 'N: ThisWorkbook.Worksheets("E").Range("A1").End(xlToRight).Column',
  'N: ThisWorkbook.Worksheets("E").Cells(1, ThisWorkbook.Worksheets("E").Columns.Count).End(xlToLeft).Column', 'N: ThisWorkbook.Worksheets("E").Range("B20").End(xlUp).Row',
  'N: ThisWorkbook.Worksheets("E").Range("A6").End(xlUp).Row', 'N: ThisWorkbook.Worksheets("E").Range("A6").End(xlDown).Row', 'N: ThisWorkbook.Worksheets("E").Range("C1").End(xlToRight).Column')
T("TCell", "region", 'S: ThisWorkbook.Worksheets("E").Range("G3").CurrentRegion.Address', 'S: ThisWorkbook.Worksheets("E").Range("A1").CurrentRegion.Address',
  'S: ThisWorkbook.Worksheets("E").UsedRange.Address', 'N: ThisWorkbook.Worksheets("E").UsedRange.Rows.Count', 'S: ThisWorkbook.Worksheets("E").Range("K10").CurrentRegion.Address',
  'S: ThisWorkbook.Worksheets("E").Range("A6").CurrentRegion.Address', 'S: ThisWorkbook.Worksheets("E").Range("I6").CurrentRegion.Address')
T("TCell", "clear", f"{S('O1')}.Value = 5", f"{S('O1')}.ClearContents", f"V: {S('O1')}.Value",
  f"{S('O2')}.Value = DateSerial(2024, 1, 5)", f"{S('O2')}.ClearContents", f"{S('O2')}.Value = 45000", f"S: TypeName({S('O2')}.Value)",
  f"{S('O3')}.Value = DateSerial(2024, 1, 5)", f"{S('O3')}.Clear", f"{S('O3')}.Value = 45000", f"S: TypeName({S('O3')}.Value)",
  f"{S('O4')}.Value = \"x\"", f"{S('O4')}.Value = \"\"", f"V: {S('O4')}.Value", f"{S('O6')}.Formula = \"=\"\"\"\"\"",
  f"V: {S('O6')}.Value", f"V: IsEmpty({S('O6')}.Value)", f"N: Len({S('O6')}.Value)")
T("TCell", "foreach", f"{S('Q1')}.Value = 1", f"{S('R1')}.Value = 2", f"{S('Q2')}.Value = 3", f"{S('R2')}.Value = 4",
  f"For Each c In {S('Q1:R2')}", "s = s & c.Value", "Next", "S: s", f"For Each c In {S('Q1:R2')}.Rows", "n = n + 1", "Next", "N: n",
  f"N: {S('Q1:R2')}.Cells.Count", f"For Each c In {S('Q1:Q2')}", "c.Value = c.Value * 10", "Next", f"N: {S('Q2')}.Value")
T("TCell", "copy1", f"{S('T1')}.Value = 1", f"{S('T2')}.Formula = \"=T1+1\"", f"{S('T1:T2')}.Copy {S('V1')}", f"S: {S('V2')}.Formula",
  f"N: {S('V2')}.Value", f"{S('T1:T2')}.Copy ThisWorkbook.Worksheets(\"TX\").Range(\"C1\")", 'S: ThisWorkbook.Worksheets("TX").Range("C2").Formula',
  'N: ThisWorkbook.Worksheets("TX").Range("C2").Value')
T("TCell", "copydate", f"{S('Y1')}.Value = DateSerial(2024, 1, 5)", f"{S('Y1')}.Copy {S('Z1')}", f"S: TypeName({S('Z1')}.Value)",
  f"{S('Z2')}.Value = {S('Y1')}.Value", f"S: TypeName({S('Z2')}.Value)", f"{S('Z3')}.Value = {S('Y1')}.Value2",
  f"S: TypeName({S('Z3')}.Value)")
T("TCell", "copy1toN", f"{S('AP1')}.Value = 7", f"{S('AP1')}.Copy {S('AQ1:AQ3')}", f"N: {S('AQ3')}.Value")
T("TCell", "hidden", 'ThisWorkbook.Worksheets("S").Rows(40).Hidden = True', 'V: ThisWorkbook.Worksheets("S").Rows(40).Hidden', f"V: {S('A40')}.EntireRow.Hidden",
  'ThisWorkbook.Worksheets("S").Rows("41:42").Hidden = True', 'V: ThisWorkbook.Worksheets("S").Rows(42).Hidden', 'ThisWorkbook.Worksheets("S").Rows(40).Hidden = False', 'V: ThisWorkbook.Worksheets("S").Rows(40).Hidden')
T("TCell", "hiddencol", 'E: ThisWorkbook.Worksheets("S").Columns(30).Hidden = True', 'V: ThisWorkbook.Worksheets("S").Columns(30).Hidden')
T("TCell", "text1", f"{S('AB1')}.Value = 1234.5", f"S: {S('AB1')}.Text", f"{S('AB2')}.Value = 0.1 + 0.2", f"S: {S('AB2')}.Text",
  f"{S('AB4')}.Value = True", f"S: {S('AB4')}.Text", f"{S('AB7')}.Value = \"abc\"", f"S: {S('AB7')}.Text")
T("TCell", "text2", 'S: ThisWorkbook.Worksheets("FX").Range("M1").Text')
T("TCell", "text3", 'S: ThisWorkbook.Worksheets("FX").Range("M2").Text')
T("TCell", "text4", 'S: ThisWorkbook.Worksheets("FX").Range("M3").Text')
T("TCell", "text5", f"{S('AB3')}.Value = DateSerial(2024, 1, 5)", f"S: {S('AB3')}.Text")
T("TCell", "text6", f"{S('AB5')}.Value = 123456789012#", f"S: {S('AB5')}.Text")
T("TCell", "valuetypes", 'S: TypeName(ThisWorkbook.Worksheets("FX").Range("M1").Value)', 'S: TypeName(ThisWorkbook.Worksheets("FX").Range("M2").Value)',
  'S: TypeName(ThisWorkbook.Worksheets("FX").Range("M3").Value)', 'S: TypeName(ThisWorkbook.Worksheets("FX").Range("D3").Value)', 'S: TypeName(ThisWorkbook.Worksheets("FX").Range("D4").Value)',
  'V: ThisWorkbook.Worksheets("FX").Range("M2").Value2', 'V: ThisWorkbook.Worksheets("FX").Range("D4").Value = 0', 'V: ThisWorkbook.Worksheets("FX").Range("D4").Value = ""',
  'N: Len(ThisWorkbook.Worksheets("FX").Range("D4").Value)', 'N: ThisWorkbook.Worksheets("FX").Range("D4").Value + 1')

# --------------------------------------------------------------------------------------------------------------- sheets
T("TSheet", "names", "N: ThisWorkbook.Worksheets.Count", "S: ThisWorkbook.Worksheets(1).Name", 'N: ThisWorkbook.Worksheets("Out").Index',
  "For Each w In ThisWorkbook.Worksheets", "n = n + 1", "Next", "N: n", "S: ThisWorkbook.Name",
  'E: x = ThisWorkbook.Worksheets("NoSuch").Name', "E: x = ThisWorkbook.Worksheets(999).Name",
  'S: TypeName(ThisWorkbook.Worksheets("S"))', 'S: TypeName(ThisWorkbook.Worksheets("S").Range("A1"))', 'S: ThisWorkbook.Worksheets("S").Parent.Name')
T("TSheet", "active", 'ThisWorkbook.Worksheets("S").Activate', "S: ActiveSheet.Name", 'ThisWorkbook.Worksheets("S").Range("B3").Select', "S: Selection.Address",
  "S: ActiveCell.Address", "ActiveCell.Offset(1, 0).Select", "S: ActiveCell.Address", "Selection.Value = 7",
  'N: ThisWorkbook.Worksheets("S").Range("B4").Value', 'Range("AD1").Value = 3', 'N: ThisWorkbook.Worksheets("S").Range("AD1").Value', 'ThisWorkbook.Worksheets("TX").Select', "S: ActiveSheet.Name",
  'E: ThisWorkbook.Worksheets("S").Range("AD2:AE3").Select', "S: Selection.Address")
T("TSheet", "visible", 'ThisWorkbook.Worksheets("CP").Visible = False', 'N: CLng(ThisWorkbook.Worksheets("CP").Visible)', 'ThisWorkbook.Worksheets("CP").Visible = xlSheetVeryHidden',
  'N: CLng(ThisWorkbook.Worksheets("CP").Visible)', 'ThisWorkbook.Worksheets("CP").Visible = True', 'N: CLng(ThisWorkbook.Worksheets("CP").Visible)')
T("TSheet", "rename", 'ThisWorkbook.Worksheets("FX").Range("P1").Formula = "=RN!A1+1"', 'ThisWorkbook.Worksheets("RN").Name = "Renamed"', 'S: ThisWorkbook.Worksheets("FX").Range("P1").Formula',
  'E: ThisWorkbook.Worksheets("Renamed").Name = "S"', 'E: ThisWorkbook.Worksheets("Renamed").Name = ""', 'ThisWorkbook.Worksheets("Renamed").Name = "RN"', 'S: ThisWorkbook.Worksheets("FX").Range("P1").Formula')

# --------------------------------------------------------------------------------------------------------------- formulas
FORMULAS = [
    "=-2^2", "=2^-1", "=1+2*3", '="a"&1', "=1/0", "=0.1+0.2=0.3", "=(0.1+0.2)-0.3", "=0.1+0.2-0.3", '=1="1"', '="a"<"B"',
    '="a"="A"', "=TRUE+1", "=10%", "=1E3", '=--"5"', "=A1", '=A1&""', "=LEN(A1)", "=A1=0", '=A1=""', "=ISBLANK(A1)",
    '=SUM("5",TRUE)', "=SUM(D1:D5)", "=AVERAGE(D1:D5)", "=COUNT(D1:D5)", "=COUNTA(D1:D5)", "=MAX(D1:D5)", "=MIN(D1:D5)",
    "=ROUND(2.5,0)", "=ROUND(-2.5,0)", "=ROUND(1.005,2)", "=ROUND(2.675,2)", "=ROUND(1234.5,-2)", "=INT(-1.5)", '=LEFT("abc")',
    '=LEFT("abc",5)', '=MID("abcdef",2,3)', '=RIGHT("abc",2)', "=LEN(123.5)", '=TRIM("  a   b ")', '=IF(1,"y","n")',
    '=IF("a",1,2)', '=IFERROR(1/0,"e")', "=VLOOKUP(2,F1:G3,2,FALSE)", "=VLOOKUP(9,F1:G3,2,FALSE)", '=MATCH("B",H1:H5,0)',
    '=MATCH("b",H1:H5,0)', "=INDEX(F1:G3,2,2)", "=INDEX(F1:G3,5,1)", "=AND(TRUE,1)", "=OR(FALSE,0)", "=NOT(0)",
    '=ISNUMBER("1")', "=ISTEXT(1)", '=CONCAT("a",1,TRUE)', "=ABS(-1.5)", '=1+"a"', '="5"+1', '="5"*"2"', "=1/3", "=2^3^2",
    "=1>2", '="b">"a"', "=TRUE=1", '=""=0', "=XLOOKUP(2,F1:F3,G1:G3)", '=XLOOKUP(9,F1:F3,G1:G3,"none")', "=UPPER(H2)&LOWER(H1)",
    "=D2&D3", "=D3+1", "=D1>D2", "=D2>D3", "=SUM(D1,D2)", "=D1+D2", "=D1+D4", "=MAX(D2:D3)", "=AVERAGE(D2:D4)", "=COUNT(D:D)",
    "=SUM(D:D)", "=ROUND(0.285,2)", "=ROUND(1.115,2)", "=1.1*3=3.3", "=IF(0.1*3=0.3,1,0)", "=1-0.9=0.1", "=LEN(D4)", "=D4&D4",
    "=IF(D4,1,2)", "=NOT(D4)", "=SUM(F1:F3)/COUNT(F1:F3)", "=MID(G1,1,1)&G2", "=ROUND(-1.005,2)", "=INT(5)", "=5/2", "=2*3.5",
    '=IF(1>2,"a")', "=IF(1>2,)", '=LEFT(12345,2)', "=LEFT(G1,-1)", '=MID("abc",0,1)', "=MAX(F1:F3,10)", "=MIN(F1:F3,-1)",
    '=COUNTA(D1:D5,"x")', "=AVERAGE(F1:F3)", "=AND(D1:D5)", "=OR(D4)", "=SUM(D3)", "=D3*2", "=-D3", '=TRIM(G1)', "=ISNUMBER(D4)",
    "=ISTEXT(D2)", "=ISBLANK(D4)", "=IFERROR(VLOOKUP(9,F1:G3,2,FALSE),\"none\")", "=INDEX(F1:G3,0,2)", "=INDEX(G1:G3,2)",
    '=MATCH(2,F1:F3,0)', '=MATCH("z",H1:H5,0)', "=1E+308*10", "=-1E+308*10", "=2^1024", "=0^0", "=(-8)^(1/3)", "=1/3*3=1",
    "=TRUE&FALSE", "=TRUE>FALSE", '="A"<"a"', '="abc"<"abd"', '="10"<"9"', "=10<\"9\"", "=TRUE>\"z\"", "=D4<1", '=D4<"a"',
]
for k, f in enumerate(FORMULAS, 1):
    cell = f'ThisWorkbook.Worksheets("FX").Range("B{k}")'
    T("TFx", f"f{k:03d}", f"{cell}.Formula = {q(f)}", f"V: {cell}.Value")
T("TFx", "dyn1", 'ThisWorkbook.Worksheets("FX").Range("R1").Formula2 = "=UNIQUE(H1:H5)"', 'V: ThisWorkbook.Worksheets("FX").Range("R1").Value', 'V: ThisWorkbook.Worksheets("FX").Range("R3").Value',
  'V: ThisWorkbook.Worksheets("FX").Range("R4").Value')
T("TFx", "dyn2", 'ThisWorkbook.Worksheets("FX").Range("S1").Formula2 = "=SORT(F1:F3,1,-1)"', 'V: ThisWorkbook.Worksheets("FX").Range("S1").Value',
  'V: ThisWorkbook.Worksheets("FX").Range("S3").Value')
T("TFx", "dyn3", 'ThisWorkbook.Worksheets("FX").Range("U1").Formula2 = "=SORT(H1:H5)"', 'S: ColJoin(ThisWorkbook.Worksheets("FX"), 21, 5)')
T("TFx", "calcmode", "Application.Calculation = xlCalculationManual", 'ThisWorkbook.Worksheets("FX").Range("T1").Value = 1',
  'ThisWorkbook.Worksheets("FX").Range("T2").Formula = "=T1+1"', 'N: ThisWorkbook.Worksheets("FX").Range("T2").Value', 'ThisWorkbook.Worksheets("FX").Range("T1").Value = 10',
  'N: ThisWorkbook.Worksheets("FX").Range("T2").Value', "Application.Calculation = xlCalculationAutomatic", 'N: ThisWorkbook.Worksheets("FX").Range("T2").Value')
T("TFx", "wsfn", "Set wf = Application.WorksheetFunction", 'N: wf.Sum(ThisWorkbook.Worksheets("FX").Range("D1:D5"))', 'N: wf.Max(ThisWorkbook.Worksheets("FX").Range("D1:D5"))',
  'N: wf.Min(ThisWorkbook.Worksheets("FX").Range("F1:F3"))', 'N: wf.CountA(ThisWorkbook.Worksheets("FX").Range("D1:D5"))', 'N: wf.Sum(1, "2", True)', 'N: wf.Find("b", "abc")',
  'E: x = wf.Find("z", "abc")', 'N: wf.Search("B", "abc")', 'N: wf.Sum(ThisWorkbook.Worksheets("FX").Range("M1:M3"))', 'N: wf.Max(ThisWorkbook.Worksheets("FX").Range("M2"))')

# --------------------------------------------------------------------------------------------------------------- sorting
for sheet, name, extra in (("SortT1", "sort_text_asc", "Order1:=xlAscending"), ("SortT2", "sort_text_desc", "Order1:=xlDescending"),
                           ("SortT3", "sort_text_case", "Order1:=xlAscending, MatchCase:=True")):
    T("TSort", name, f'ThisWorkbook.Worksheets("{sheet}").Range("A1:A22").Sort Key1:=ThisWorkbook.Worksheets("{sheet}").Range("A1"), {extra}, Header:=xlNo',
      f'S: ColJoin(ThisWorkbook.Worksheets("{sheet}"), 1, 22)')
for sheet, name, order in (("SortM1", "sort_mixed_asc", "xlAscending"), ("SortM2", "sort_mixed_desc", "xlDescending")):
    T("TSort", name, f'ThisWorkbook.Worksheets("{sheet}").Range("A15").Value = CVErr(xlErrNA)',
      f'ThisWorkbook.Worksheets("{sheet}").Range("A1:A15").Sort Key1:=ThisWorkbook.Worksheets("{sheet}").Range("A1"), Order1:={order}, Header:=xlNo',
      f'S: ColJoin(ThisWorkbook.Worksheets("{sheet}"), 1, 15)')
T("TSort", "sort_multi", 'ThisWorkbook.Worksheets("SortK1").Range("A1:C7").Sort Key1:=ThisWorkbook.Worksheets("SortK1").Range("A1"), Order1:=xlAscending, '
  'Key2:=ThisWorkbook.Worksheets("SortK1").Range("B1"), Order2:=xlDescending, Header:=xlYes', 'S: ColJoin(ThisWorkbook.Worksheets("SortK1"), 3, 7)')
T("TSort", "sort_stable", 'ThisWorkbook.Worksheets("SortK2").Range("A1:C7").Sort Key1:=ThisWorkbook.Worksheets("SortK2").Range("A1"), Order1:=xlAscending, Header:=xlYes',
  'S: ColJoin(ThisWorkbook.Worksheets("SortK2"), 3, 7)')
T("TSort", "sort_obj", 'Set ws = ThisWorkbook.Worksheets("SortK3")', "ws.Sort.SortFields.Clear", 'ws.Sort.SortFields.Add Key:=ws.Range("B2:B7"), Order:=xlDescending',
  'ws.Sort.SortFields.Add Key:=ws.Range("A2:A7"), Order:=xlAscending', 'ws.Sort.SetRange ws.Range("A1:C7")', "ws.Sort.Header = xlYes",
  "ws.Sort.Apply", 'S: ColJoin(ws, 3, 7)')

# --------------------------------------------------------------------------------------------------------------- filters
FILTERS = [
    ("f_eq", 'Field:=2, Criteria1:="North"'), ("f_wild", 'Field:=2, Criteria1:="N*"'), ("f_ne", 'Field:=2, Criteria1:="<>North"'),
    ("f_blank", 'Field:=2, Criteria1:="="'), ("f_nonblank", 'Field:=2, Criteria1:="<>"'), ("f_gt", 'Field:=3, Criteria1:=">500"'),
    ("f_ge", 'Field:=3, Criteria1:=">=1250"'), ("f_numeq", 'Field:=3, Criteria1:="1250"'), ("f_numeq2", 'Field:=3, Criteria1:="=1250"'),
    ("f_between", 'Field:=3, Criteria1:=">500", Operator:=xlAnd, Criteria2:="<1000"'),
    ("f_or", 'Field:=2, Criteria1:="North", Operator:=xlOr, Criteria2:="East"'),
    ("f_values", 'Field:=2, Criteria1:=Array("North", "East"), Operator:=xlFilterValues'),
    ("f_contains", 'Field:=2, Criteria1:="*or*"'), ("f_qmark", 'Field:=2, Criteria1:="?orth"'),
    ("f_trailing", 'Field:=2, Criteria1:="West"'), ("f_numtext", 'Field:=3, Criteria1:="830.5"'),
    ("f_lt0", 'Field:=3, Criteria1:="<0"'), ("f_num_ne", 'Field:=3, Criteria1:="<>1250"'),
    ("f_gt_text", 'Field:=2, Criteria1:=">North"'), ("f_eq_formatted", 'Field:=3, Criteria1:="1,250.00"'),
]
for k, (name, args) in enumerate(FILTERS, 1):
    sheet = f"F{k:02d}"
    T("TFilt", name, f'ThisWorkbook.Worksheets("{sheet}").Range("A1:E13").AutoFilter {args}', f'S: Vis(ThisWorkbook.Worksheets("{sheet}"))', f'V: ThisWorkbook.Worksheets("{sheet}").FilterMode',
      f'V: ThisWorkbook.Worksheets("{sheet}").AutoFilterMode')
T("TFilt", "f_two", 'ThisWorkbook.Worksheets("F21").Range("A1:E13").AutoFilter Field:=2, Criteria1:="North"',
  'ThisWorkbook.Worksheets("F21").Range("A1:E13").AutoFilter Field:=3, Criteria1:=">1000"', 'S: Vis(ThisWorkbook.Worksheets("F21"))')
T("TFilt", "f_copy", 'ThisWorkbook.Worksheets("F22").Range("A1:E13").AutoFilter Field:=2, Criteria1:="North"',
  'ThisWorkbook.Worksheets("F22").Range("A1:E13").Copy ThisWorkbook.Worksheets("CP").Range("A1")', 'N: ThisWorkbook.Worksheets("CP").Cells(ThisWorkbook.Worksheets("CP").Rows.Count, 1).End(xlUp).Row',
  'S: ColJoin(ThisWorkbook.Worksheets("CP"), 1, 13)')
T("TFilt", "f_write", 'ThisWorkbook.Worksheets("F23").Range("A1:E13").AutoFilter Field:=2, Criteria1:="North"', 'ThisWorkbook.Worksheets("F23").Range("E2:E13").Value = "w"',
  'ThisWorkbook.Worksheets("F23").Range("C2:C13").ClearContents', 'ThisWorkbook.Worksheets("F23").ShowAllData', 'S: ColJoin(ThisWorkbook.Worksheets("F23"), 5, 13)', 'S: ColJoin(ThisWorkbook.Worksheets("F23"), 3, 13)')
T("TFilt", "f_end", 'ThisWorkbook.Worksheets("F24").Range("A1:E13").AutoFilter Field:=2, Criteria1:="East"',
  'N: ThisWorkbook.Worksheets("F24").Cells(ThisWorkbook.Worksheets("F24").Rows.Count, 1).End(xlUp).Row', 'N: ThisWorkbook.Worksheets("F24").Range("A1").End(xlDown).Row',
  'N: ThisWorkbook.Worksheets("F24").UsedRange.Rows.Count', 'N: ThisWorkbook.Worksheets("F24").Range("A1").CurrentRegion.Rows.Count')
T("TFilt", "f_toggle", 'ThisWorkbook.Worksheets("F25").Range("A1:E13").AutoFilter', 'V: ThisWorkbook.Worksheets("F25").AutoFilterMode', 'V: ThisWorkbook.Worksheets("F25").FilterMode',
  'E: ThisWorkbook.Worksheets("F25").ShowAllData', 'ThisWorkbook.Worksheets("F25").Range("A1:E13").AutoFilter', 'V: ThisWorkbook.Worksheets("F25").AutoFilterMode',
  'ThisWorkbook.Worksheets("F25").Range("A1:E13").AutoFilter Field:=2, Criteria1:="North"', 'ThisWorkbook.Worksheets("F25").AutoFilterMode = False', 'S: Vis(ThisWorkbook.Worksheets("F25"))',
  'V: ThisWorkbook.Worksheets("F25").AutoFilterMode')
T("TFilt", "f_field_err", 'E: ThisWorkbook.Worksheets("F26").Range("A1:E13").AutoFilter Field:=9, Criteria1:="x"', 'V: ThisWorkbook.Worksheets("F26").AutoFilterMode')
T("TFilt", "f_cell", 'ThisWorkbook.Worksheets("F27").Range("A1").AutoFilter Field:=2, Criteria1:="South"', 'S: Vis(ThisWorkbook.Worksheets("F27"))')
T("TFilt", "f_refilter", 'ThisWorkbook.Worksheets("F28").Range("A1:E13").AutoFilter Field:=2, Criteria1:="North"',
  'ThisWorkbook.Worksheets("F28").Range("A1:E13").AutoFilter Field:=2, Criteria1:="East"', 'S: Vis(ThisWorkbook.Worksheets("F28"))',
  'ThisWorkbook.Worksheets("F28").Range("A1:E13").AutoFilter Field:=2', 'S: Vis(ThisWorkbook.Worksheets("F28"))', 'V: ThisWorkbook.Worksheets("F28").FilterMode')
T("TFilt", "f_sort_filtered", 'ThisWorkbook.Worksheets("F29").Range("A1:E13").AutoFilter Field:=2, Criteria1:="North"',
  'ThisWorkbook.Worksheets("F29").Range("A1:E13").Sort Key1:=ThisWorkbook.Worksheets("F29").Range("C1"), Order1:=xlDescending, Header:=xlYes', 'ThisWorkbook.Worksheets("F29").ShowAllData',
  'S: ColJoin(ThisWorkbook.Worksheets("F29"), 1, 13)')

# --------------------------------------------------------------------------------------------------------------- language
LANG_PROCS = '''
Const K1 = 5
Const K2 = K1 * 2
Public Const KS = "a" & "b"

Sub Inc(x)
    x = x + 1
End Sub

Sub Mut(ByVal arr)
    arr(0) = 99
End Sub

Sub MutR(arr)
    arr(0) = 77
End Sub

Function Counter()
    Static n As Long
    n = n + 1
    Counter = n
End Function

Function Opt(Optional a, Optional b = 5)
    If IsMissing(a) Then
        Opt = "missing" & b
    Else
        Opt = "given" & b
    End If
End Function

Function Rec(n)
    If n <= 1 Then
        Rec = 1
    Else
        Rec = n * Rec(n - 1)
    End If
End Function

Function NoRet()
End Function

Function TypedRet() As Long
End Function

Function Early()
    Early = 1
    Exit Function
    Early = 2
End Function

Function Total(ParamArray xs() As Variant)
    Dim i As Long
    For i = LBound(xs) To UBound(xs)
        Total = Total + xs(i)
    Next i
End Function

Function Handler1() As String
    On Error GoTo Fail
    x = 1 / 0
    Handler1 = Handler1 & "after"
    Exit Function
Fail:
    Handler1 = "caught" & Err.Number & "|"
    Resume Next
End Function

Function Handler2() As String
    On Error GoTo Fail
    Boom
    Handler2 = Handler2 & "after"
    Exit Function
Fail:
    Handler2 = "caught" & Err.Number & "|"
    Resume Done
Done:
    Handler2 = Handler2 & "done" & Err.Number
End Function

Sub Boom()
    Dim a(1 To 2)
    a(5) = 1
End Sub

Function Handler3() As String
    On Error Resume Next
    x = 1 / 0
    Handler3 = "n" & Err.Number
    On Error GoTo 0
    Handler3 = Handler3 & "|z" & Err.Number
End Function
'''
T("TLang", "byref", "n = 1", "Inc n", "N: n", "Inc (n)", "N: n", "Call Inc(n)", "N: n", "a = Array(1, 2)", "Inc a(0)", "N: a(0)",
  f"{S('AJ1')}.Value = 5", f"Inc {S('AJ1')}.Value", f"N: {S('AJ1')}.Value")
T("TLang", "arrays1", "a = Array(1, 2, 3)", "b = a", "b(0) = 9", "N: a(0)", "Dim c(1 To 2)", "c(1) = 5", "d = c", "d(1) = 6", "N: c(1)",
  "Mut a", "N: a(0)", "MutR a", "N: a(0)", "Dim e()", "E: x = UBound(e)", "ReDim e(1 To 2)", "e(2) = 4", "ReDim Preserve e(1 To 3)",
  "N: e(2)", "N: UBound(e)")
T("TLang", "arrays2", "Dim m(1 To 2, 1 To 3)", "For i = 1 To 2", "For j = 1 To 3", "m(i, j) = i * 10 + j", "Next j", "Next i",
  "N: UBound(m, 2)", "For Each x In m", 's = s & x & ","', "Next", "S: s", "N: LBound(Array(1, 2))", "N: UBound(Array())",
  "E: x = m(3, 1)", "Erase m", "N: m(1, 1)")
T("TLang", "loops", "n = 3", "For i = 1 To n", "n = n + 1", "c = c + 1", "Next", "N: c", "N: i", "For i = 1 To 0", "z = z + 1", "Next",
  "N: z", "For i = 1 To 5", "i = i + 1", "k = k + 1", "Next", "N: k", "For x = 0 To 1 Step 0.1", "f = f + 1", "Next", "N: f",
  "For i = 10 To 1 Step -3", "g = g & i", "Next", "S: g", "For q9 = 1 To 3", "hh = hh + 1", "Next", "N: hh")
T("TLang", "ifs1", "x = 0", "E: If x <> 0 And 1 / x > 1 Then y = 1", "If False Then p = 1: q = 2", "V: q",
  "If True Then p2 = 1 Else p2 = 2: q2 = 3", "V: q2", "If False Then p3 = 1 Else p3 = 2: q3 = 3", "V: q3", "V: IIf(True, 1, 2)",
  "E: x = IIf(True, 1, 1 / 0)")
T("TLang", "select1", "Select Case True", "Case 1 > 2", 's1 = "a"', "Case 2 > 1", 's1 = "b"', "End Select", "S: s1",
  'Select Case "b"', 'Case "a" To "c"', 's2 = "in"', "Case Else", 's2 = "out"', "End Select", "S: s2", "Select Case 5",
  "Case 1 To 3, 5", 's3 = "hit"', "Case Else", 's3 = "miss"', "End Select", "S: s3", 'Select Case "B"', 'Case "b"', 's4 = "same"',
  "Case Else", 's4 = "else"', "End Select", "S: s4")
T("TLang", "caseis", "x = 7", "Select Case x", "Case Is > 5", 's = "big"', "Case Else", 's = "else"', "End Select", "S: s")
T("TLang", "errs1", "E: Err.Raise 513, , \"boom\"", "E: Err.Raise vbObjectError + 1", "E: Error 13", 'E: x = CInt("abc")',
  "Dim r As Object", "E: x = r.Value", "E: x = Worksheets(\"NoSuch\").Name", "Dim arr(1 To 2)", "E: arr(3) = 1")
T("TLang", "errdesc", "Err.Clear", 'Err.Raise 513, , "boom"', "X: Err.Description", "Err.Clear", "x = 1 / 0", "X: Err.Description")
T("TLang", "handlers", "S: Handler1()", "S: Handler2()", "S: Handler3()")
T("TLang", "withs", f"With {S('AL1')}", ".Value = 1", "With .Offset(1, 0)", ".Value = 2", "End With", ".Offset(2, 0).Value = .Value + 10",
  "End With", f"N: {S('AL2')}.Value", f"N: {S('AL3')}.Value")
T("TLang", "consts", "N: K2", "S: KS", "N: vbYes", "N: xlUp", "N: Len(vbCrLf)", "N: Asc(vbTab)", "N: xlPasteValues",
  "N: xlCalculationManual", "N: &HFF", "N: &H8000", "N: &H8000&", "N: 1E3", "N: .5")
T("TLang", "procs", "N: Counter() + Counter()", "S: Opt()", "S: Opt(1, 2)", "S: Opt(, 7)", "S: Opt(b:=9)", "N: Rec(5)", "V: NoRet()",
  "V: TypedRet()", "N: Early()")
T("TLang", "paramarray", "N: Total(1, 2, 3)")
T("TLang", "midstmt", 's = "abc"', 'Mid(s, 2, 1) = "X"', "S: s")
T("TLang", "dollar", 'S: Left$("abc", 2) & Trim$(" a ") & UCase$("a") & Mid$("abc", 2) & Chr$(65)')
T("TBase", "optbase", "Dim a(3)", "N: LBound(a)", "x = Array(10, 20)", "N: LBound(x)", "N: x(1)", 'N: LBound(Split("a,b", ","))',
  "Dim m(2, 3)", "N: LBound(m, 2)", "N: UBound(Array(1, 2, 3))")
T("TEnum", "enum1", "N: Green")
T("TEnum", "enum2", "N: Colour.Green")

MODULE_HEADERS = {"TLang": LANG_PROCS, "TBase": "Option Base 1\n", "TEnum": "Enum Colour\n    Red = 1\n    Green\nEnd Enum\n"}


def module_sources():
    """{module name: VBA source}, and [(row, module, test name)] for the Out sheet."""
    by_module, index = {}, []
    for row, (module, name, lines) in enumerate(TESTS, 1):
        out, col = [f"Public Sub {name}()", "    On Error Resume Next", "    Application.DisplayAlerts = False"], 0
        for line in lines:
            kind = line[:3] if line[:3] in ("S: ", "N: ", "V: ", "E: ", "X: ") else None
            if kind is None:
                out.append("    " + line)
                continue
            col += 1
            body = line[3:]
            if kind == "X: ":  # as text, without clearing Err first
                out.append(f"    PutS {row}, {col}, {body}")
            elif kind == "E: ":
                out += ["    Err.Clear", "    " + body, f"    PutN {row}, {col}, Err.Number", "    Err.Clear"]
            else:
                put = {"S: ": "PutS", "N: ": "PutN", "V: ": "PutV"}[kind]
                out += ["    Err.Clear", f"    {put} {row}, {col}, {body}", "    If Err.Number <> 0 Then",
                        f'        PutS {row}, {col}, "ERR" & Err.Number', "    End If"]
        out += ["    Application.DisplayAlerts = True", "End Sub", ""]
        by_module.setdefault(module, []).extend(out)
        index.append((row, module, name))
    sources = {"H": HELPERS}
    for module, lines in by_module.items():
        sources[module] = f'Attribute VB_Name = "{module}"\n' + MODULE_HEADERS.get(module, "") + "\n" + "\n".join(lines)
    return sources, index


def d(text):
    return dt.date.fromisoformat(text)


SORT_TEXT = ["banana", "Apple", "cherry", "apple", "Banana", "co-op", "coop", "co op", "CO-OP", "it's", "its", "_x", "10", "9",
             "a1", "A2", "éclair", "eclair", "zebra", " lead", "Zebra", "a-1"]
SORT_MIXED = [3, "b", True, 1, "a", False, None, 2.5, "10", (d("2024-01-05"), "date"), "B", 0, -1, ""]
SORT_KEYS = [["grp", "n", "idx"], ["b", 2, 1], ["a", 1, 2], ["b", 1, 3], ["a", 2, 4], ["c", 1, 5], ["a", 1, 6]]
REGIONS = ["North", "south", "NORTH", "East", "West", None, "North East", "north", "South", "East", "West ", "Northern"]
AMOUNTS = [(1250, "money"), 830.5, 412, (2100, "money"), 96.75, 540, 1250, 0, -5, (1000, "money"), 500, 999.99]


FILTER_ROWS = [["Id", "Region", "Amount", "Date", "Note"]] + [
    [100 + i, REGIONS[i], AMOUNTS[i], (d("2026-03-01") + dt.timedelta(days=i), "date"), f"n{i}"] for i in range(12)
]
SHEETS += [("S", [10], []),
           ("TX", [10], [[(None, "text")] for _ in range(20)]),
           ("D", [12], [["straße"], ["ÀÉÎ"], [" lead"]]),
           ("E", [8], [[1, "b1", "c1", None, "e1"], [2, "b2"], [3, None, None, None, None, None, "g3", "h3"],
                       [4, None, None, None, None, None, "g4", "h4"], [5, None, None, None, None, None, "g5", "h5"],
                       [None, None, None, None, None, None, None, None, "diag"], [7], [8]]),
           ("CP", [10], []),
           ("RN", [10], [[1]]),
           ("FX", [10], [[None, None, None, 5, None, 1, "a", "B", None, None, None, None, (1250, "money")],
                         [None, None, None, "x", None, 2, "b", "a", None, None, None, None, (d("2026-03-02"), "date")],
                         [None, None, None, True, None, 3, "c", "b", None, None, None, None, (0.08, "percent")],
                         [None, None, None, None, None, None, None, "A"], [None, None, None, 10, None, None, None, "c"]])]
SHEETS += [(f"SortT{i}", [12], [[v] for v in SORT_TEXT]) for i in (1, 2, 3)]
SHEETS += [(f"SortM{i}", [12], [[v] for v in SORT_MIXED]) for i in (1, 2)]
SHEETS += [(f"SortK{i}", [8, 8, 8], SORT_KEYS) for i in (1, 2, 3)]
SHEETS += [(f"F{i:02d}", [8, 12, 12, 12, 8], FILTER_ROWS) for i in range(1, 30)]


# ------------------------------------------------------------------------------------------------- more of the same
def W(sheet):
    return f'ThisWorkbook.Worksheets("{sheet}")'


D0 = "DateSerial(2024, 1, 1)"
T("TMore", "ddiff_h", f'N: DateDiff("h", {D0} + TimeSerial(10, 0, 0), {D0} + TimeSerial(11, 0, 0))',
  f'N: DateDiff("h", {D0} + TimeSerial(10, 30, 0), {D0} + TimeSerial(11, 29, 0))',
  f'N: DateDiff("h", {D0} + TimeSerial(10, 30, 0), {D0} + TimeSerial(11, 30, 0))',
  f'N: DateDiff("h", {D0} + TimeSerial(10, 59, 0), {D0} + TimeSerial(11, 1, 0))',
  f'N: DateDiff("h", {D0} + TimeSerial(23, 59, 59), {D0} + 1)',
  f'N: DateDiff("h", {D0} + TimeSerial(10, 0, 0), {D0} + TimeSerial(12, 59, 59))',
  f'N: DateDiff("d", {D0} + TimeSerial(23, 59, 59), {D0} + 1)')
T("TMore", "ddiff_n", f'N: DateDiff("n", {D0} + TimeSerial(10, 0, 59), {D0} + TimeSerial(10, 1, 0))',
  f'N: DateDiff("n", {D0} + TimeSerial(10, 0, 30), {D0} + TimeSerial(10, 1, 29))',
  f'N: DateDiff("n", {D0} + TimeSerial(10, 0, 0), {D0} + TimeSerial(10, 0, 59))',
  f'N: DateDiff("s", {D0} + TimeSerial(10, 0, 0), {D0} + TimeSerial(10, 0, 59))',
  f'N: DateDiff("n", {D0} + TimeSerial(10, 0, 0), {D0} + TimeSerial(11, 0, 0))')
T("TMore", "round3", "N: Round(-1.015, 2)", "N: Round(-8.345, 2)", "N: Round(0.045, 2)", "N: Round(1.0005, 3)", "N: Round(123.455, 2)",
  "N: Round(0.615, 2)", "N: Round(1.335, 2)", "N: Round(4.35, 1)", "N: Round(0.15, 1)", "N: Round(0.35, 1)")
T("TMore", "varcmp2", f'{W("S")}.Range("AS1").Value = 5', 'v = "5"', f'V: {W("S")}.Range("AS1").Value = v',
  f'V: {W("S")}.Range("AS1").Value = "5"', "Dim st As String", 'st = "5"', "V: 5 = st", f'V: {W("S")}.Range("AS1").Value = st',
  "Dim nn As Long", "nn = 5", "V: nn = v", "V: v = 5", f'V: {W("S")}.Range("AS1").Value < v')
T("TMore", "textnum", f'V: {W("TXN")}.Range("A1").Value = 10', f'V: {W("TXN")}.Range("A1").Value = {W("TXN")}.Range("A2").Value',
  f'N: {W("TXN")}.Range("A1").Value + {W("TXN")}.Range("A2").Value', f'V: {W("TXN")}.Range("A1").Value & {W("TXN")}.Range("A2").Value',
  f'V: {W("TXN")}.Range("A1").Value > {W("TXN")}.Range("A2").Value', f'S: TypeName({W("TXN")}.Range("A1").Value)',
  f'x = {W("TXN")}.Range("A1").Value', f'y = {W("TXN")}.Range("A2").Value', "V: x = y", "V: x < y")
T("TMore", "prefmt", f'{W("PF")}.Range("A1").Value = 45000', f'S: TypeName({W("PF")}.Range("A1").Value)',
  f'S: {W("PF")}.Range("A1").NumberFormat', f'{W("PF")}.Range("A2").Value = 0.5', f'S: {W("PF")}.Range("A2").NumberFormat',
  f'{W("PF")}.Range("A3").Value = 5', f'S: {W("PF")}.Range("A3").NumberFormat', f'S: {W("PF")}.Range("A4").NumberFormat',
  f'{W("PF")}.Range("A5").Value = "45000"', f'S: TypeName({W("PF")}.Range("A5").Value)')
T("TMore", "hidecopy", f'{W("HC")}.Rows(3).Hidden = True', f'{W("HC")}.Range("A1:A5").Copy {W("HC")}.Range("C1")',
  f'S: ColJoin({W("HC")}, 3, 5)', f'{W("HC")}.Range("A1:A5").Value = 9', f'{W("HC")}.Rows(3).Hidden = False', f'S: ColJoin({W("HC")}, 1, 5)')
T("TMore", "calcmanual2", "Application.Calculation = xlCalculationManual", f'{W("FX")}.Range("W1").Value = 1',
  f'{W("FX")}.Range("W2").Formula = "=W1+1"', f'{W("FX")}.Range("W1").Value = 10', f'N: {W("FX")}.Range("W2").Value',
  "Application.Calculation = xlCalculationAutomatic", f'N: {W("FX")}.Range("W2").Value')
T("TMore", "strmisc", "S: Str(0.25)", "S: Str(-0.5)", "S: CStr(CDate(0.5))", "S: Format(0.5)", "S: CStr(.5)", 'S: "x" & 0.5',
  "S: Trim(Str(12))", "S: CStr(-0.5)", "S: Str(1E+20)", "S: CStr(CDate(0))")
T("TMore", "sorterr", f'{W("SortE")}.Range("A1").Value = 3', f'{W("SortE")}.Range("A2").Value = CVErr(xlErrNA)',
  f'{W("SortE")}.Range("A3").Value = "a"', f'{W("SortE")}.Range("A4").Value = 1', f'{W("SortE")}.Range("A5").Value = CVErr(xlErrDiv0)',
  f'{W("SortE")}.Range("A6").Value = True',
  f'{W("SortE")}.Range("A1:A6").Sort Key1:={W("SortE")}.Range("A1"), Order1:=xlAscending, Header:=xlNo', f'S: ColJoin({W("SortE")}, 1, 6)')
T("TMore", "timecell", f'{W("S")}.Range("AW1").Value = TimeSerial(13, 5, 0)', f'S: TypeName({W("S")}.Range("AW1").Value)',
  f'S: {W("S")}.Range("AW1").NumberFormat', f'V: {W("S")}.Range("AW1").Value2',
  f'{W("S")}.Range("AW2").Value = DateSerial(2024, 1, 5) + TimeSerial(13, 5, 0)', f'S: {W("S")}.Range("AW2").NumberFormat')
T("TMore", "arrbyref2", "Dim q(1 To 2)", "q(1) = 1", "Inc2 q(1)", "N: q(1)")
T("TMore", "defprop", f'Set r = {W("S")}.Range("AX1")', "r = 5", f'N: {W("S")}.Range("AX1").Value', "x = r", "S: TypeName(x)",
  "r.Value = r.Value + 1", "N: r.Value", "If r = 6 Then", 's = "y"', "End If", "S: s", "S: TypeName(r)", "Set r2 = r", "r2.Value = 9",
  "N: r.Value")
T("TMore", "numfmt_read", f'S: {W("FX")}.Range("M1").NumberFormat', f'S: {W("FX")}.Range("M2").NumberFormat',
  f'S: {W("FX")}.Range("M3").NumberFormat', f'S: {W("FX")}.Range("D1").NumberFormat', f'S: {W("TX")}.Range("A12").NumberFormat')
T("TMore", "intlit", "E: x = 200 * 300", "E: x = 100 * 100", "N: 200 * 300&", "E: x = 32000 + 1000", "N: CLng(200) * 300")
T("TMore", "filt_eq2", f'{W("F30")}.Range("A1:E13").AutoFilter Field:=2, Criteria1:="=West"', f'S: Vis({W("F30")})',
  f'{W("F30")}.Range("A1:E13").AutoFilter Field:=2, Criteria1:="north east"', f'S: Vis({W("F30")})',
  f'{W("F30")}.Range("A1:E13").AutoFilter Field:=2, Criteria1:="North*"', f'S: Vis({W("F30")})')
WRITES2 = ["12%", "Jan 2024", "January", "12,5", "1,234,5", "1e-2", "1,000.5", "(1,000)", "-$5", "1,2", "12.5%", "$ 5", "1/2 ", "TRUE", " TRUE",
           "#REF!", "#VALUE!", "1.2.3", "March 5", "5 March 2024", "10:30", "1,000,000", "0,5", "1 %"]
for k, text in enumerate(WRITES2, 1):
    cell = f'{W("S")}.Cells({k}, 60)'
    T("TMore", f"x{k:02d}", f"{cell}.Value = {q(text)}", f"S: TypeName({cell}.Value)", f"V: {cell}.Value2", f"S: {cell}.NumberFormat")

MODULE_HEADERS["TMore"] = "Sub Inc2(x)\n    x = x + 1\nEnd Sub\n"
SHEETS += [("PF", [10], [[(None, "date")], [(None, "percent")], [(None, "money")], [(None, "text")], [(None, "date")]]),
           ("HC", [8], [[1], [2], [3], [4], [5]]), ("TXN", [8], [["10"], [10]]), ("SortE", [8], []),
           ("F30", [8, 12, 12, 12, 8], FILTER_ROWS)]


# ------------------------------------------------------- filtered ranges, hidden rows, text and numbers compared
def flt(sheet):
    return f'{W(sheet)}.Range("A1:E13").AutoFilter Field:=2, Criteria1:="North"'


T("TModel", "fw_single_hidden", flt("F31"), f'{W("F31")}.Range("E3").Value = "h"', f'{W("F31")}.ShowAllData', f'S: ColJoin({W("F31")}, 5, 13)')
T("TModel", "fw_array", "Dim arr(1 To 12, 1 To 1)", "For i = 1 To 12", 'arr(i, 1) = "a" & i', "Next", flt("F32"),
  f'{W("F32")}.Range("E2:E13").Value = arr', f'{W("F32")}.ShowAllData', f'S: ColJoin({W("F32")}, 5, 13)')
T("TModel", "fw_formula", flt("F33"), f'{W("F33")}.Range("F2:F13").Formula = "=A2*2"', f'{W("F33")}.ShowAllData',
  f'S: ColJoin({W("F33")}, 6, 13)')
T("TModel", "fw_clear", flt("F34"), f'{W("F34")}.Range("E2:E13").Clear', f'{W("F34")}.ShowAllData', f'S: ColJoin({W("F34")}, 5, 13)')
T("TModel", "fw_loop", flt("F35"), f'For Each c In {W("F35")}.Range("E2:E13")', 'c.Value = "L"', "Next", f'{W("F35")}.ShowAllData',
  f'S: ColJoin({W("F35")}, 5, 13)')
T("TModel", "fw_copy_into", flt("F36"), f'{W("HC2")}.Range("A1:A3").Copy {W("F36")}.Range("E2")', f'{W("F36")}.ShowAllData',
  f'S: ColJoin({W("F36")}, 5, 13)')
T("TModel", "fw_value_copy", flt("F37"), f'{W("F37")}.Range("E2:E13").Value = {W("F37")}.Range("A2:A13").Value', f'{W("F37")}.ShowAllData',
  f'S: ColJoin({W("F37")}, 5, 13)')
T("TModel", "fw_multi_col", flt("F38"), f'{W("F38")}.Range("E2:F13").Value = "z"', f'{W("F38")}.ShowAllData', f'S: ColJoin({W("F38")}, 5, 13)',
  f'S: ColJoin({W("F38")}, 6, 13)')
T("TModel", "fw_partial", flt("F39"), f'{W("F39")}.Range("E3:E4").Value = "p"', f'{W("F39")}.Range("E5:E7").Value = "q"',
  f'{W("F39")}.Range("C3").ClearContents', f'{W("F39")}.Range("C5:C6").ClearContents', f'{W("F39")}.ShowAllData',
  f'S: ColJoin({W("F39")}, 5, 13)', f'S: ColJoin({W("F39")}, 3, 13)')
T("TModel", "fw_outside", flt("F40"), f'{W("F40")}.Range("H2:H13").Value = "o"', f'{W("F40")}.Range("H15:H16").Value = "u"',
  f'{W("F40")}.ShowAllData', f'S: ColJoin({W("F40")}, 8, 16)')
T("TModel", "fw_formula_arr", flt("F41"), f'{W("F41")}.Range("F2:F13").Value = "=A2+1"', f'{W("F41")}.ShowAllData',
  f'S: ColJoin({W("F41")}, 6, 13)', f'S: {W("F41")}.Range("F4").Formula')
T("TModel", "fw_read", flt("F42"), f'x = {W("F42")}.Range("A2:A13").Value', "N: UBound(x, 1)", "N: x(2, 1)",
  f'N: Application.WorksheetFunction.Sum({W("F42")}.Range("A2:A13"))', f'N: {W("F42")}.Range("A2:A13").Count',
  f'N: {W("F42")}.Range("A2:A13").Rows.Count')

FS_DATA = ["West", "West ", " West", "West  ", "WEST", "Wes", "West1", None, "west", "Wes t", "EastWest", "W"]
for k, crit in enumerate(['"West"', '"=West"', '"West "', '"W*t"', '"<>West"', '" West"', '"=west*"', '"<>*t"'], 1):
    T("TModel", f"fs{k}", f'{W(f"FS{k}")}.Range("A1:A13").AutoFilter Field:=1, Criteria1:={crit}', f'S: Vis({W(f"FS{k}")})')

T("TModel", "eh1", f'{W("EH1")}.Rows("7:8").Hidden = True', f'N: {W("EH1")}.Cells({W("EH1")}.Rows.Count, 1).End(xlUp).Row',
  f'N: {W("EH1")}.Range("A1").End(xlDown).Row', f'N: {W("EH1")}.Range("A5").End(xlDown).Row')
T("TModel", "eh2", f'{W("EH2")}.Rows("3:4").Hidden = True', f'N: {W("EH2")}.Range("A1").End(xlDown).Row',
  f'N: {W("EH2")}.Range("A2").End(xlDown).Row', f'N: {W("EH2")}.Range("A8").End(xlUp).Row', f'N: {W("EH2")}.Range("A5").End(xlUp).Row')
T("TModel", "eh3", f'{W("EH3")}.Columns("B:C").Hidden = True', f'N: {W("EH3")}.Range("A1").End(xlToRight).Column',
  f'{W("EH3")}.Columns("F").Hidden = True', f'N: {W("EH3")}.Cells(1, {W("EH3")}.Columns.Count).End(xlToLeft).Column',
  f'N: {W("EH3")}.Range("A1").End(xlToRight).Column')

T("TModel", "pow", "E: x = (-8) ^ (1 / 3)", "E: x = 0 ^ 0", "N: 0 ^ 0", "E: x = 0 ^ -1", "E: x = 2 ^ 1024", "N: (-8) ^ 3", "N: (-8) ^ -1",
  "E: x = (-8) ^ 0.5", "N: 4 ^ 0.5", "N: 2 ^ 0.5")
T("TModel", "midstmt2", 's = "abcdef"', 'Mid(s, 2, 1) = "X"', "S: s", 's = "abcdef"', 'Mid(s, 2) = "XY"', "S: s", 's = "abcdef"',
  'Mid(s, 5) = "WXYZ"', "S: s", 's = "abcdef"', 'Mid(s, 2, 3) = "Z"', "S: s", 's = "abcdef"', 'Mid$(s, 1, 1) = "Q"', "S: s",
  's = "abcdef"', 'E: Mid(s, 10) = "x"', 's = "abcdef"', 'Mid(s, 2, 10) = "12"', "S: s")
T("TModel", "switch", 'V: Switch(1 = 2, "a", True, "b")', "V: Switch(False, 1)", 'V: Switch(True, "x", True, "y")', "E: Error 13",
  "V: Choose(2, \"a\", \"b\")", "V: Choose(5, \"a\", \"b\")")
T("TModel", "formula_r2", f'{W("S")}.Range("BA1").Value = CVErr(xlErrNA)', f'S: {W("S")}.Range("BA1").Formula',
  f'{W("S")}.Range("BA2").Value = 0.1 + 0.2', f'S: {W("S")}.Range("BA2").Formula', f'{W("S")}.Range("BA3").Value = 1E+20',
  f'S: {W("S")}.Range("BA3").Formula', f'{W("S")}.Range("BA4").Value = 1 / 3', f'S: {W("S")}.Range("BA4").Formula',
  f'{W("S")}.Range("BA5").Value = False', f'S: {W("S")}.Range("BA5").Formula',
  f'{W("S")}.Range("BA6").Value = DateSerial(2024, 1, 5) + TimeSerial(12, 0, 0)', f'S: {W("S")}.Range("BA6").Formula')
T("TModel", "colstyle", f'{W("CS")}.Range("A1").Value = "00123"', f'S: TypeName({W("CS")}.Range("A1").Value)',
  f'V: {W("CS")}.Range("A1").Value2', f'S: {W("CS")}.Range("A1").NumberFormat', f'{W("CS")}.Range("B1").Value = 45000',
  f'S: TypeName({W("CS")}.Range("B1").Value)', f'{W("CS")}.Range("C1").Value = 5', f'S: {W("CS")}.Range("C1").NumberFormat',
  f'S: {W("CS")}.Range("A5").NumberFormat', f'S: {W("CS")}.Range("D5").NumberFormat', f'S: TypeName({W("CS")}.Range("C1").Value)')
T("TModel", "ccur", "S: CStr(CCur(1.00005))", "S: CStr(CCur(2.00015))", "S: CStr(CCur(0.00005))", "S: CStr(CCur(-1.00005))",
  "S: CStr(CCur(1.23455))", "S: CStr(CCur(1.23465))", "S: CStr(CCur(0.12345))", "S: CStr(CCur(1000000.00005))", "S: CStr(CCur(0.00015))",
  "S: CStr(CCur(123.45675))")
NUMERIC = ["($5)", "$(5)", " ( 5 ) ", "-(5)", "(-5)", "5 -", "- 5", "$-5", "-$5", "5$", "1 1/2", "TRUE", "1e5", "1.5e", "5,", ",5",
           "1,2,3", ".", "&HFF", "&O7", "1D5", "12:30", "1/2", "$", "%5", "5%", "+5", "1,23", "(5)", "5-", " 5 ", "1.2.3", "1e", "0x10",
           "1 000", "$ 5", "-", "+", "5.", ".5"]
for i in range(0, len(NUMERIC), 10):
    chunk = NUMERIC[i : i + 10]
    T("TModel", f"isnumb{i // 10}", *[f"V: IsNumeric({q(x)})" for x in chunk])
    T("TModel", f"cdblb{i // 10}", *[f"N: CDbl({q(x)})" for x in chunk])
T("TModel", "str2", "S: Str(0)", "S: Str(1)", "S: Str(-0.25)", "S: Str(100.5)", "S: Str(0.00001)", "S: Str(1 / 3)", "S: Str(-1.5)",
  "S: Str(0.5) & \"|\"", "S: Str(True)", "S: Str(12345.678)")
DATES = ["5 Jan 2024", "Jan 5, 2024", "5-Jan-2024", "5-Jan-24", "January 2024", "Jan 2024", "2024-01-05", "2024-01-05 13:00", "2024/01/05",
         "13:05", "1:05 PM", "13/5/2024", "5/13/2024", "45000", "20240105", "January 5, 2024 1:05 PM", "5 January 2024", "1 Feb 99",
         "1 Feb 29", "1 Feb 30", "Feb 30 2024", "13:05:09", "24:00", "12 am", "2024-1-5", "05-01-2024 "]
for i in range(0, len(DATES), 2):
    T("TModel", f"cdate{i // 2:02d}", *[f"N: CDbl(CDate({q(x)}))" for x in DATES[i : i + 2]], *[f"V: IsDate({q(x)})" for x in DATES[i : i + 2]])
T("TModel", "manual3", "Application.Calculation = xlCalculationManual", f'{W("FX")}.Range("AF1").Value = 1',
  f'{W("FX")}.Range("AF2").Formula = "=AF1+1"', f'{W("FX")}.Range("AF3").Formula = "=AF2*10"', f'N: {W("FX")}.Range("AF3").Value',
  f'{W("FX")}.Range("AF1").Value = 5', f'N: {W("FX")}.Range("AF2").Value', f'N: {W("FX")}.Range("AF3").Value',
  f'{W("FX")}.Range("AF4").Formula = "=AF2+AF3"', f'N: {W("FX")}.Range("AF4").Value', f'{W("FX")}.Calculate',
  f'N: {W("FX")}.Range("AF3").Value', "Application.Calculation = xlCalculationAutomatic", f'N: {W("FX")}.Range("AF4").Value')
T("TModel", "showall", f'E: {W("F31")}.ShowAllData', f'V: {W("EH1")}.AutoFilterMode', f'E: {W("EH1")}.ShowAllData')
T("TModel", "sortfn", f'{W("FX")}.Range("AC1").Value = 3', f'{W("FX")}.Range("AC2").Value = "b"', f'{W("FX")}.Range("AC3").Value = True',
  f'{W("FX")}.Range("AC4").Value = "a"', f'{W("FX")}.Range("AC5").Value = 1', f'{W("FX")}.Range("AC6").Value = False',
  f'{W("FX")}.Range("AD1").Formula2 = "=SORT(AC1:AC6)"', f'S: ColJoin2({W("FX")}, 30, 6)',
  f'{W("FX")}.Range("AE1").Formula2 = "=SORT(AC1:AC6,1,-1)"', f'S: ColJoin2({W("FX")}, 31, 6)')

VC_SETUP = [f'Set ws = {W("VC")}', "Dim st As String", 'st = "5"', "Dim nn As Long", "nn = 5", "Dim sa As String", 'sa = "abc"', "v5 = 5",
            'vs = "5"', 'va = "abc"']
T("TCompare", "vc1", *VC_SETUP, 'V: ws.Range("A1").Value = ws.Range("A2").Value', 'V: ws.Range("A1").Value = "5"',
  'V: ws.Range("A1").Value = "05"', 'V: ws.Range("A1").Value = "Total"', 'V: ws.Range("A1").Value < "Total"',
  'V: ws.Range("A2").Value = 5', 'V: ws.Range("A3").Value = 5', 'V: ws.Range("A2").Value > 4', 'V: ws.Range("A3").Value > 4',
  'V: ws.Range("A4").Value = ""', 'V: ws.Range("A4").Value = 0')
T("TCompare", "vc2", *VC_SETUP, "V: st = nn", "V: sa = nn", "V: nn = vs", "V: nn = va", "V: st = v5", "V: sa = v5", "V: sa > v5", "V: v5 = vs",
  "V: v5 < va", 'V: 5 = "abc"', 'V: "abc" = 5')
T("TCompare", "vc3", *VC_SETUP, 'V: Trim(ws.Range("A2").Value) = ws.Range("A1").Value', 'V: Trim$(ws.Range("A2").Value) = ws.Range("A1").Value',
  'V: CStr(ws.Range("A1").Value) = ws.Range("A2").Value', 'V: CLng(ws.Range("A2").Value) = ws.Range("A1").Value',
  'V: ws.Range("A1").Value = Val(ws.Range("A2").Value)', 'V: Len(ws.Range("A2").Value) = "1"',
  'V: (ws.Range("A2").Value & "") = ws.Range("A1").Value', 'V: (ws.Range("A1").Value + 0) = ws.Range("A2").Value',
  'V: UCase(ws.Range("A2").Value) = ws.Range("A1").Value', 'V: Left(ws.Range("A2").Value, 1) = ws.Range("A1").Value',
  'V: Left$(ws.Range("A2").Value, 1) = ws.Range("A1").Value')
T("TCompare", "vc4", *VC_SETUP, "V: CStr(5) = 5", "V: Str(5) = 5", "V: RetS() = v5", "V: RetV() = v5", "V: RetN() = vs", "V: RetVn() = vs",
  "Dim sarr(1 To 1) As String", 'sarr(1) = "5"', "V: sarr(1) = v5", 'varr = Array("5")', "V: varr(0) = v5", "V: (st & \"\") = v5",
  "V: (vs & \"\") = v5", "V: (v5 & \"\") = nn")
T("TCompare", "vc5", *VC_SETUP, 'V: ws.Range("A6").Value = "True"', 'V: ws.Range("A6").Value = -1', 'vt = "True"',
  'V: ws.Range("A6").Value = vt', 'V: ws.Range("A4").Value = vs', 'V: ws.Range("A4").Value < vs', 'V: ws.Range("A4").Value = v5',
  'V: ws.Range("A4").Value < v5', 'V: ws.Range("A5").Value = ws.Range("A2").Value', 'V: ws.Range("A5").Value = 5',
  'V: ws.Range("A1").Value = ws.Range("A1").Value', 'V: ws.Range("A2") = ws.Range("A1")', 'V: ws.Cells(1, 1) = ws.Cells(2, 1)',
  'V: ws.Cells(1, 1) = "5"')
T("TCompare", "vc_sel", *VC_SETUP, 'Select Case ws.Range("A1").Value', 'Case "5"', 'r1 = "hit"', "Case Else", 'r1 = "miss"', "End Select",
  "S: r1", "Select Case v5", "Case vs", 'r2 = "hit"', "Case Else", 'r2 = "miss"', "End Select", "S: r2", "Select Case nn", "Case vs",
  'r3 = "hit"', "Case Else", 'r3 = "miss"', "End Select", "S: r3", "Select Case v5", 'Case "4" To "6"', 'r4 = "hit"', "Case Else",
  'r4 = "miss"', "End Select", "S: r4", "Select Case vs", "Case 5", 'r5 = "hit"', "Case Else", 'r5 = "miss"', "End Select", "S: r5",
  "Select Case v5", "Case Is > \"4\"", 'r6 = "hit"', "Case Else", 'r6 = "miss"', "End Select", "S: r6")
T("TEnum2", "enum3", "N: ShRed", "N: ShGreen", "N: ShPurple", "N: ShNeg", "N: ShCalc", "N: ShAfter", "N: BxSmall", "N: BxMedium",
  "N: Shade.ShBlue", "Dim c As Shade", "c = ShGreen", "N: c", "S: TypeName(c)", "S: TypeName(ShGreen)")

WRITES3 = ["1,234.5", "1,2345", "12,345,678", ",5", "1,,000", "1,000,00", "$1,23", "5$", "$-5", "($5)", "(5%)", "5E3", "1.5%", ".5%", "%5",
           "0 1/2", "1 3/8", "2 10/16", "-1 1/2", "False", "tRuE", "#NAME?", "#NUM!", "#NULL!", "#n/a", "N/A", "e5", "--5", "+-5", "5+",
           "$", "%", "()", "(5", "5)", " (5) ", "( 5 )", "- 5", "5 -", "1 e3", "1,000", "1,000.", "1.000,5", "001", "1e+3", "1E-3", "1.e3",
           ".e3", "1e3.5", "$1e3", "1e3%", "(1e3)", "1 / 2", "Mar", "May", "may 5", "5 may", "12 am", "12:30 pm", "1:30", "25:00", "1:60",
           "13:00:00", "2024/01/05", "01.05.2024", "1-Jan", "Jan-24", "Q1 2024", "1st", "10%5", "1,234 ", " 1,234", "1 234", "+1,234",
           "+$5", "$5.555", "3e", "1_000", "5 %%", "$5%", "1.5e3", "12e", "1e1000", "-1e400", "1/2/2024 10:00", "5-", "(5.5)", "$(5)",
           "-(5)", "1,234e2", "TRUE FALSE", "=", "=5", "'5", "''5", "5'", "a=1", "+", "++5", "+5%"]
for k, text in enumerate(WRITES3, 1):
    cell = f'{W("S")}.Cells({k}, 70)'
    T("TWrites", f"y{k:03d}", f"{cell}.Value = {q(text)}", f"S: TypeName({cell}.Value)", f"V: {cell}.Value2", f"S: {cell}.NumberFormat")

F3_SETUP = [f'{W("FX")}.Range("Z1").Value = 0.3', f'{W("FX")}.Range("AA1").Value = "hit"', f'{W("FX")}.Range("Z2").Formula = "=0.1+0.2"',
            f'{W("FX")}.Range("Z3").Value = 0.3']
FORMULAS3 = [
    "=1.1-1=0.1", "=(1.1-1)=0.1", "=4.35*100=435", "=4.1-4=0.1", "=0.1+0.2>0.3", "=0.1+0.2<=0.3", "=0.1+0.2>=0.3", "=0.1+0.2<0.3",
    "=0.1+0.2<>0.3", "=1+1E-15=1", "=1+1E-14=1", "=1+2^-44=1", "=1+2^-45=1", "=1+2^-46=1", "=1+2^-47=1", "=1+2^-48=1", "=1+2^-49=1",
    "=1+2^-50=1", "=1+2^-52=1", "=1000000+2^-28=1000000", "=1000000+2^-30=1000000", "=1E+15+1=1E+15", "=1E+14+1=1E+14",
    '=IF(0.1+0.2=0.3,"eq","ne")', "=(1+2^-44)-1", "=(1+2^-46)-1", "=(1+2^-47)-1", "=(1+2^-48)-1", "=(1+2^-49)-1", "=(1+2^-50)-1",
    "=(1+2^-52)-1", "=(0.1+0.2-0.3)", "=(0.1+0.2-0.3)*1", "=0.1+0.2-0.3+0", "=-0.3+0.1+0.2", "=1-0.9-0.1", "=0.5-0.4-0.1",
    "=1*(0.5-0.4-0.1)", "=0.3-(0.1+0.2)", "=SUM(0.1,0.2,-0.3)", "=SUM(0.1,0.2)-0.3", "=1.1-1-0.1", "=43.1-43.2+1", "=100.1-100",
    "=MATCH(0.1+0.2,Z1:Z1,0)", "=VLOOKUP(0.1+0.2,Z1:AA1,2,FALSE)", "=IF(Z2=Z3,1,0)", "=Z2-Z3", "=(Z2-Z3)", "=Z2=Z3", "=ABS(Z2-Z3)",
    "=MAX(Z2,Z3)=Z3", "=INT(0.1+0.2+0.7)", "=INT(4.35*100)", "=4.35*100", "=IF(4.35*100<435,1,0)", "=ROUND(4.35*100,0)", "=Z2>Z3",
    "=XLOOKUP(0.1+0.2,Z1:Z1,AA1:AA1)", "=(-8)^(2/3)", "=(-32)^(1/5)", "=(-8)^0.5", "=(-27)^(1/3)", "=(-8)^-1", "=(-8)^2", "=(-1)^(1/3)",
    "=0^-1", "=2^1023*2", "=10^308*10", '=SUM("a",1)', '=COUNT("5",TRUE,"a")', '=COUNTA("",1)', '=MAX("5",3)', "=MIN(TRUE,3)",
    '=AVERAGE("5",TRUE)', "=SUM(D2)", '=AND("TRUE",1)', '=AND("x",1)', "=OR(D2:D2)", "=AND(D2,TRUE)", "=AND(D4)", "=NOT(D2)",
    "=AND(1,0)", "=OR(D1:D5)", "=AND(D4,TRUE)", "=OR(D3:D4)", "=0.1+0.2-0.3=0", "=(0.1+0.2-0.3)=0", "=1-0.9=0.1", "=10-9.9", "=10-9.9=0.1",
    "=0.1*3", "=0.1*3-0.3", "=1/3+1/3+1/3-1", "=2^0.5*2^0.5-2", "=2^0.5*2^0.5=2",
]
for k, f in enumerate(FORMULAS3, 1):
    cell = f'{W("FX")}.Range("Y{k}")'
    T("TFloat", f"g{k:03d}", *F3_SETUP, f"{cell}.Formula = {q(f)}", f"V: {cell}.Value")

HELPERS += '''
Public Function ColJoin2(ByVal ws As Object, ByVal col As Long, ByVal n As Long) As String
    Dim r As Long, s As String
    For r = 1 To n
        s = s & Show(ws.Cells(r, col).Value) & "|"
    Next r
    ColJoin2 = s
End Function
'''
MODULE_HEADERS["TCompare"] = ('Function RetS() As String\n    RetS = "5"\nEnd Function\nFunction RetV()\n    RetV = "5"\nEnd Function\n'
                          "Function RetN() As Long\n    RetN = 5\nEnd Function\nFunction RetVn()\n    RetVn = 5\nEnd Function\n")
MODULE_HEADERS["TEnum2"] = ("Enum Shade\n    ShRed = 1\n    ShGreen\n    ShBlue = 10\n    ShPurple\n    ShNeg = -5\n    ShCalc = ShBlue * 2\n"
                          "    ShAfter\nEnd Enum\nPrivate Enum BoxSize\n    BxSmall\n    BxMedium\nEnd Enum\n")
SHEETS += [(f"F{i}", [8, 12, 12, 12, 8], FILTER_ROWS) for i in range(31, 43)]
SHEETS += [(f"FS{i}", [12], [["K"]] + [[v] for v in FS_DATA]) for i in range(1, 9)]
SHEETS += [(f"EH{i}", [8], [[1, 2, 3, 4, 5, 6]] + [[n] for n in range(2, 9)]) for i in (1, 2, 3)]
SHEETS += [("VC", [8], [[5], ["5"], ["abc"], [None], ["05"], [True]]),
           ("CS", [(10, "text"), (10, "date"), (10, "money"), 10], [])]


# ------------------------------------------------------------------------------------ what a number format makes of a cell
# The number formats cells carry in the workbook, by style name: one of Excel's own by its number, or a format code.
# (The builder has "bold", "money" (#,##0.00), "date" (yyyy-mm-dd) and "percent" (0%) already.)
STYLES = [("text", 49), ("t20", 20), ("t19", 19), ("t45", 45), ("t46", 46), ("d16", 16), ("d17", 17), ("d22", 22), ("d14", 14),
                ("yyyy", "yyyy"), ("mmmm", "mmmm"), ("dddd", "dddd"), ("hlit", 'h "hours"'), ("mlit", '"m"0'), ("m1", "m"),
                ("mmss0", "mm:ss.0"), ("c5", 5), ("c7", 7), ("c8", 8), ("c42", 42), ("c44", 44), ("n41", 41), ("n43", 43),
                ("euro", "[$€-407] #,##0.00"), ("euroafter", "#,##0.00 [$€-407]"), ("usdlit", '#,##0.00 "USD"'),
                ("dollarq", '"$"#,##0.00'), ("pound", "£#,##0.00"), ("usdcode", "[$USD] #,##0.00"), ("n2", 2), ("n4", 4),
                ("dollarbare", "$0.00"), ("sci", 11), ("frac", 12), ("yen", "[$¥-411]#,##0"), ("n0000", "0.0000"), ("p2", 10),
          ("d15", 15), ("t18", 18), ("n37", 37), ("f13", 13), ("t21", 21), ("n39", 39), ("s48", 48),
          ("unlocked", 0, {"unlocked": True})]
B4_DATES = ["t20", "t19", "t45", "t46", "d16", "d17", "d22", "d14", "yyyy", "mmmm", "dddd", "hlit", "mlit", "m1", "mmss0", "sci", "frac"]
B4_MONEY = ["c5", "c7", "c8", "c42", "c44", "n41", "n43", "euro", "euroafter", "usdlit", "dollarq", "pound", "usdcode", "n2", "n4",
            "dollarbare", "yen", "money", "percent"]
B4 = W("B4")
for k in range(1, len(B4_DATES) + 1):
    T("TFormat", f"dfmt{k:02d}", f"S: TypeName({B4}.Cells({k}, 1).Value)", f"S: {B4}.Cells({k}, 1).NumberFormat")
for k in range(1, len(B4_MONEY) + 1):
    T("TFormat", f"cfmt{k:02d}", f"S: TypeName({B4}.Cells({k}, 2).Value)", f"V: {B4}.Cells({k}, 2).Value", f"V: {B4}.Cells({k}, 2).Value2",
      f"S: {B4}.Cells({k}, 2).NumberFormat")
CUR = ["Dim c As Currency, d2 As Double, i As Integer, l As Long, sg As Single", "c = 10.5", "d2 = 1.5", "i = 2", "l = 3", "sg = 1.5"]
T("TFormat", "cur1", *CUR, "S: TypeName(c + i)", "S: TypeName(c + l)", "S: TypeName(c + d2)", "S: TypeName(c + sg)", "S: TypeName(c * i)",
  "S: TypeName(c * d2)", "S: TypeName(c * c)", "S: TypeName(c / i)", "S: TypeName(c / c)", "S: TypeName(c \\ i)", "S: TypeName(c Mod i)",
  "S: TypeName(c ^ 2)", "S: TypeName(-c)", "S: TypeName(c - c)")
T("TFormat", "cur2", *CUR, 'S: TypeName(c & "")', "S: TypeName(Abs(c))", "S: TypeName(Int(c))", "S: TypeName(Fix(c))", "S: TypeName(Round(c, 1))",
  "S: TypeName(Sgn(c))", "S: TypeName(c + True)", 'S: TypeName(c + "1")', "S: TypeName(c + Empty)", "S: TypeName(c + CDate(1))",
  "S: TypeName(c * l)", "S: TypeName(c - d2)", "S: TypeName(i * c)", "S: TypeName(d2 * c)")
T("TFormat", "cur3", *CUR, "V: c * 1.00005", "V: CCur(1.2346) * CCur(1.2346)", "V: CCur(0.0001) * CCur(0.5)", "V: CCur(10) / CCur(3)",
  "V: CCur(1.5) + 1.00005", "V: CCur(2.5) * 2", "V: CCur(1.23455)", "V: CCur(1.23445)", "V: c = 10.5", "V: CCur(0.1) + CCur(0.2) = CCur(0.3)",
  "V: CCur(0.0003) * CCur(0.5)", "V: CCur(1.1111) * 3", "V: CCur(10) \\ 3", "V: CCur(10.5) Mod 3")
T("TFormat", "cur4", *CUR, "E: x = CCur(922337203685477) + CCur(1)", "E: x = CCur(922337203685477) * 2", "S: TypeName(c * 2.5)",
  "V: CCur(3.3333) * 3", "V: CCur(0.0001) / 2", "S: CStr(CCur(1234.5))", "S: CStr(CCur(0.1))", "S: Format(CCur(1234.5), \"0.00\")",
  "V: CCur(2.5) > 2.4999", "V: -CCur(1.5)", "V: Abs(CCur(-1.25))", "V: Int(CCur(-1.5))", "V: Round(CCur(2.345), 2)", "V: CCur(1.005) * 100")
T("TFormat", "cur5", f"{B4}.Range(\"D1\").Value = CCur(2.5)", f"S: {B4}.Range(\"D1\").NumberFormat", f"S: TypeName({B4}.Range(\"D1\").Value)",
  f"V: {B4}.Range(\"D1\").Value2", f"{B4}.Range(\"D2\").Value = CCur(1234.5678)", f"V: {B4}.Range(\"D2\").Value2",
  f"S: {B4}.Range(\"D2\").NumberFormat", "Dim cc As Currency", "cc = 7.25", f"{B4}.Range(\"D3\").Value = cc",
  f"S: {B4}.Range(\"D3\").NumberFormat", f"{B4}.Range(\"D4\").Value2 = cc", f"S: {B4}.Range(\"D4\").NumberFormat",
  f"{B4}.Range(\"D5\").Formula = cc", f"S: {B4}.Range(\"D5\").NumberFormat")
T("TFormat", "cur6", f"v = {B4}.Cells(2, 2).Value", "S: TypeName(v)", "S: TypeName(v * 3)", "V: v * 3", "V: v / 3", "V: v * 1.1", "V: v + 0.00004",
  f"arr = {B4}.Range(\"B2:B3\").Value", "S: TypeName(arr(1, 1))", f"V: {B4}.Cells(2, 2).Value + {B4}.Cells(2, 2).Value",
  f"V: {B4}.Cells(2, 2).Value2 + {B4}.Cells(2, 2).Value2", "Dim dd As Double", f"dd = {B4}.Cells(2, 2).Value", "V: dd",
  f"V: Application.WorksheetFunction.Sum({B4}.Range(\"B2:B3\"))", f"V: {B4}.Cells(2, 2).Value = 1.2346", f"V: {B4}.Cells(2, 2).Value = 1.23456")
TEXTS4 = [1234.56789, 123456789, 12345678901, 0.000012345, 1234567.891, -1234567.5, 0.1234567891234, 1e15, 1.5e-10, 99999999999,
          123456.789012, 12345678, 1234567.8, 0.1, 1 / 3, 100000000000, 1e-5, 123456789.5, -12345678901, 0.00001234]
for k, v in enumerate(TEXTS4, 1):
    T("TFormat", f"txt{k:02d}", f"{B4}.Cells({k}, 6).Value = {v!r}", f"S: {B4}.Cells({k}, 6).Text")
T("TFormat", "txtnarrow", f"{B4}.Range(\"E1\").Value = 123456", f"S: {B4}.Range(\"E1\").Text", f"{B4}.Range(\"E2\").Value = 12.3456",
  f"S: {B4}.Range(\"E2\").Text", f"{B4}.Range(\"E3\").Value = \"abcdefghij\"", f"S: {B4}.Range(\"E3\").Text",
  f"{B4}.Range(\"E4\").Value = DateSerial(2024, 1, 5)", f"S: {B4}.Range(\"E4\").Text", f"{B4}.Range(\"E5\").Value = True",
  f"S: {B4}.Range(\"E5\").Text", f"{B4}.Range(\"E6\").Value = 1234", f"S: {B4}.Range(\"E6\").Text")
T("TFormat", "valfill", f"{B4}.Range(\"H1\").Value = 1", f"{B4}.Range(\"H2\").Value = 2", f"{B4}.Range(\"H3\").Value = 3",
  f"{B4}.Range(\"I1:I3\").Value = \"=H1*2\"", f"S: {B4}.Range(\"I2\").Formula", f"S: {B4}.Range(\"I3\").Formula", f"V: {B4}.Range(\"I3\").Value")
T("TFormat", "clearfmt", f"{B4}.Range(\"B1\").Clear", f"S: {B4}.Range(\"B1\").NumberFormat", f"V: {B4}.Range(\"B1\").Value",
  f"{B4}.Range(\"B2\").ClearContents", f"S: {B4}.Range(\"B2\").NumberFormat", f"{B4}.Range(\"B2\").Value = 5",
  f"S: TypeName({B4}.Range(\"B2\").Value)", f"{B4}.Range(\"B1\").Value = 5", f"S: TypeName({B4}.Range(\"B1\").Value)",
  f"{W('CS')}.Range(\"A2\").Clear", f"S: {W('CS')}.Range(\"A2\").NumberFormat", f"S: {W('CS')}.Range(\"A3\").NumberFormat",
  f"{W('CS')}.Range(\"A2\").Value = \"007\"", f"V: {W('CS')}.Range(\"A2\").Value", f"{W('CS')}.Columns(2).Clear",
  f"S: {W('CS')}.Range(\"B5\").NumberFormat")
T("TFormat", "pasteblank", f"{B4}.Range(\"K1\").Copy {B4}.Range(\"A8\")", f"S: {B4}.Range(\"A8\").NumberFormat", f"V: {B4}.Range(\"A8\").Value",
  f"{B4}.Range(\"A8\").Value = 45296", f"S: TypeName({B4}.Range(\"A8\").Value)", f"{W('CS')}.Range(\"D1\").Copy {W('CS')}.Range(\"A4\")",
  f"S: {W('CS')}.Range(\"A4\").NumberFormat", f"{W('CS')}.Rows(6).Copy {W('B4')}.Rows(30)", f"S: {B4}.Range(\"A30\").NumberFormat",
  f"S: {B4}.Range(\"D30\").NumberFormat")
T("TFormat", "wsfn2", "Set wf = Application.WorksheetFunction", 'N: wf.Max("5", 3)', 'N: wf.Count("5", True, "a")', 'E: x = wf.Sum("a", 1)',
  'N: wf.Average("5", True)', 'arr = Array(1, "2", True)', "N: wf.Sum(arr)", "N: wf.Count(arr)", 'N: wf.Min("5", 3)', "N: wf.Sum(True, True)",
  'N: wf.Sum("5%", 1)', 'E: x = wf.Max("a", 1)', "N: wf.CountA(arr)", 'N: wf.Sum(" 5 ", 1)')
T("TFormat", "fvalues", f'{W("FS9")}.Range("A1:A13").AutoFilter Field:=1, Criteria1:=Array("West", "Wes"), Operator:=xlFilterValues',
  f'S: Vis({W("FS9")})', f'{W("FS9")}.Range("A1:A13").AutoFilter Field:=1, Criteria1:=Array("west ", "W"), Operator:=xlFilterValues',
  f'S: Vis({W("FS9")})', f'{W("FS9")}.Range("A1:A13").AutoFilter Field:=1, Criteria1:=">West"', f'S: Vis({W("FS9")})',
  f'{W("FS9")}.Range("A1:A13").AutoFilter Field:=1, Criteria1:="West", Operator:=xlOr, Criteria2:="W "', f'S: Vis({W("FS9")})',
  f'{W("FS9")}.Range("A1:A13").AutoFilter Field:=1, Criteria1:="="', f'S: Vis({W("FS9")})',
  f'{W("FS9")}.Range("A1:A13").AutoFilter Field:=1, Criteria1:="<>"', f'S: Vis({W("FS9")})')
T("TFormat", "manual4", "Application.Calculation = xlCalculationAutomatic", f'{W("FX")}.Range("AR1").Value = 1',
  f'{W("FX")}.Range("AR2").Formula = "=AR1+1"', "Application.Calculation = xlCalculationManual", f'{W("FX")}.Range("AR1").Value = 5',
  f'N: {W("FX")}.Range("AR2").Value', f'{W("FX")}.Range("AR3").Formula = "=AR2*10"', f'N: {W("FX")}.Range("AR3").Value',
  f'{W("FX")}.Range("AR2").Formula = "=AR1+2"', f'N: {W("FX")}.Range("AR2").Value', f'N: {W("FX")}.Range("AR3").Value',
  f'{W("FX")}.Range("AR4").Value = "=AR1*2"', f'N: {W("FX")}.Range("AR4").Value', f'{W("FX")}.Range("AR1").Value = 7',
  f'N: {W("FX")}.Range("AR4").Value', f'{W("FX")}.Range("AR5").Formula = "=SUM(AR1:AR4)"', f'N: {W("FX")}.Range("AR5").Value',
  "Application.Calculate", f'N: {W("FX")}.Range("AR5").Value', f'{W("FX")}.Range("AR1").Value = 9', f'{W("FX")}.Range("AR2").Calculate',
  f'N: {W("FX")}.Range("AR2").Value', f'N: {W("FX")}.Range("AR3").Value', "Application.Calculation = xlCalculationAutomatic",
  f'N: {W("FX")}.Range("AR3").Value')

F4_SETUP = F3_SETUP + [f'{W("FX")}.Range("AC1").Value = 0.1', f'{W("FX")}.Range("AC2").Value = 0.2', f'{W("FX")}.Range("AC3").Value = -0.3']
FORMULAS4 = [
    "=SUM(0.1,0.2,-0.3)*1", "=(SUM(0.1,0.2,-0.3))", "=SUM(0.1,0.2,-0.3)+0", "=SUM(-0.3,0.1,0.2)", "=SUM(0.1,0.2,-0.3,0)", "=SUM(AC1:AC3)",
    "=SUM(AC1:AC3)*1", "=ABS(SUM(AC1:AC3))", "=SUM(1,2^-50,-1)", "=SUM(1,2^-49,-1)", "=AVERAGE(0.1,0.2,-0.3)", "=SUM(0.1,0.2,-0.3)=0",
    "=-(0.1+0.2-0.3)", "=IF(TRUE,0.1+0.2-0.3)", '=0.1+0.2-0.3&""', "=(1.9+2^-49)-1.9", "=(1000+2^-40)-1000", "=(1000+2^-41)-1000",
    "=(1000+2^-42)-1000", "=(0.001+2^-59)-0.001", "=(0.001+2^-60)-0.001", "=(0.001+2^-61)-0.001", "=1+2^-50-1", "=2^-50+1-1",
    "=1-(1-2^-50)", "=1-(1-2^-49)", "=(1.5+2^-49)-1.5", "=(1.1+2^-49)-1.1", "=(4.35*100-435)", "=4.35*100-435", "=TRUNC(4.35*100)",
    "=ROUNDDOWN(4.35*100,0)", "=FLOOR(4.35*100,1)", "=INT(0.29*100)", "=INT((0.1+0.7)*10)", "=ROUNDUP(1.1*3,1)", "=INT(1-2^-50)",
    "=INT(1-2^-53)", "=MOD(4.35*100,1)", "=INT(-4.35*100)", "=2.9999999999999996", "=INT(2.9999999999999996)", '=AND("x")',
    '=AND("FALSE",1)', '=OR("x",0)', '=AND("1")', '=OR("TRUE")', "=AND(D2)", "=AND(D1)", "=OR(D2,D4)", "=AND(TRUE,D2:D2)", '=NOT("x")',
    '=NOT("TRUE")', '=IF("x",1,2)', '=IF("TRUE",1,2)', "=AND(D3,D2)", "=OR(D4,FALSE)", '=AND(0,"x")', "=IF(D2,1,2)", '=MAX("a",1)',
    '=AVERAGE("a",1)', '=MIN("5",3)', '=SUM("5%",1)', '=SUM("$5",1)', '=SUM(" 5 ",1)', '=COUNT("")', '=SUM("",1)', '=SUM({1,2,"3"})',
    "=SUM(TRUE,TRUE)", '=COUNT({1,"2",TRUE})', '=MAX({1,"9"})', '=SUM(D1:D5,"3")', '=PRODUCT("2",3)', '=COUNT(D1:D5,"7",TRUE)',
    '=MAX(TRUE,"x")', '="5%"+1', '="$5"+1', '="1,000"+1', '=" 5 "+1', '="1e3"+1', '="(5)"+1', '="5-"+1', '="1 1/2"+1', '="TRUE"+1',
    '="inf"+1', '="nan"+1', '="0x10"+1', '="1_000"+1', '=VALUE("5%")', '=-"5"', '=""+1', "=1E-20=0", "=0.1+0.2=0.30000000000000004",
    "=1.000000000000001=1", "=123456789012345678", "=(-8)^(-1/3)", "=(-8)^(1/4)", "=(-8)^(3/5)", "=(-8)^0.333333333333333",
    "=(-8)^(2/6)", "=(-1)^0.2", "=(-32)^0.2", "=POWER(-8,1/3)", "=(-8)^(1/3)*1", "=0^0.5", "=(-8)^3", "=(-8)^(1/7)", "=(-8)^(1/9)",
    "=(-8)^(1/15)", "=(-8)^(1/1001)", "=(-0.001)^(1/3)", "=(-8)^-3", '=1&2=12', '=1=1=TRUE', "=0.1+0.2-0.3<1E-16", "=INT(0.1+0.2-0.3)",
    "=ROUND(0.1+0.2-0.3,20)", "=SUM(0.1+0.2,-0.3)", "=SUM(0.1+0.2-0.3)", "=MAX(0.1+0.2-0.3,0)", "=0.3-0.1-0.2", "=0.1+0.2-0.3-0",
    "=SUM(Z2,-Z3)", "=SUM(Z2,-0.3)", "=Z2-Z3+0", "=0+Z2-Z3", "=10.1-10-0.1", "=100.1-100-0.1", "=1000.1-1000-0.1", "=1E+10+0.1-1E+10-0.1",
    "=1/3-0.333333333333333", "=1-1/3-2/3", "=5.1-5.05-0.05", "=SUM(1.1,-1,-0.1)", "=SUM(5.1,-5.05,-0.05)", "=SUM(1E+10,0.1,-1E+10,-0.1)",
]
for k, f in enumerate(FORMULAS4, 1):
    cell = f'{W("FX")}.Range("AH{k}")'
    T("TFormula2", f"h{k:03d}", *F4_SETUP, f"{cell}.Formula = {q(f)}", f"V: {cell}.Value")
II = [("AJ2", "=F1:F3"), ("AJ9", "=F1:F3"), ("AK3", "=INDEX(F1:G3,0,2)"), ("AL2", "=F1:F3+1"), ("AL9", "=F1:F3+1"), ("AM2", "=SUM(F1:F3*2)"),
      ("AM9", "=SUM(F1:F3*2)"), ("AN2", "={1,2,3}"), ("AN3", "=LEN(G1:G3)"), ("AO2", "=F:F"), ("AO3", '=IF(F1:F3>1,"y","n")'),
      ("AP2", "=SUMPRODUCT(F1:F3,F1:F3)"), ("AP3", "=MAX(F1:F3*2)"), ("AP9", "=MAX(F1:F3*2)"), ("AJ3", "=INDEX(F1:F3,0)"),
      ("AK2", "=F1:G1"), ("AK9", "=F1:G3"), ("AL3", "=-F1:F3"), ("AN9", "=SUM(IF(F1:F3>1,F1:F3,0))"), ("AO9", "=F1:F3=2"),
      ("AP1", '=F1:F3&"x"'), ("AJ1", "=COUNTIF(F1:F3,F1:F3)"), ("AK1", "=VLOOKUP(F1:F3,F1:G3,2,FALSE)"), ("AL1", "=ROUND(F1:F3,0)")]
for k, (cell, f) in enumerate(II, 1):
    T("TFormula2", f"ii{k:02d}", f'{W("FX")}.Range("{cell}").Formula = {q(f)}', f'V: {W("FX")}.Range("{cell}").Value',
      f'S: {W("FX")}.Range("{cell}").Formula')
T("TFormula2", "ii_f2", f'{W("FX")}.Range("AT2").Formula2 = "=F1:F3+1"', f'V: {W("FX")}.Range("AT2").Value', f'V: {W("FX")}.Range("AT4").Value',
  f'S: {W("FX")}.Range("AT2").Formula', f'S: {W("FX")}.Range("AT2").Formula2')

_B4_VALUES = {"t20": 0.545138888888889, "t19": 0.545138888888889, "t45": 0.02, "t46": 1.5, "hlit": 0.5, "mlit": 5,
              "mmss0": 0.02, "sci": 12345.678, "frac": 1.5, "d22": 45296.5}
_B4_ROWS = []
for k in range(max(len(B4_DATES), len(B4_MONEY))):
    a = B4_DATES[k] if k < len(B4_DATES) else None
    b = B4_MONEY[k] if k < len(B4_MONEY) else None
    _B4_ROWS.append([(_B4_VALUES.get(a, 45296), a) if a else None, (1.23456, b) if b else None])
_B4_ROWS[7] = [(d("2024-01-05"), "d14"), _B4_ROWS[7][1]]
SHEETS += [("B4", [10, 14, 10, 14, 4, 9.14], _B4_ROWS), ("FS9", [12], [["K"]] + [[v] for v in FS_DATA])]


# ---------------------------------------------------------------------------------- Currency, pasted formats, typed digits
B5 = W("B5")
CUR5 = [0.125, 0.135, 2.345, -0.125, 1.005, 0.0049, 1234.5, 0.995, 2.675, 0.005, 0.015, -2.345, 1e-4, 12345678.905]
for k, v in enumerate(CUR5, 1):
    T("TMoney", f"cw{k:02d}", f"{B5}.Cells({k}, 1).Value = CCur({v!r})", f"V: {B5}.Cells({k}, 1).Value2", f"S: {B5}.Cells({k}, 1).NumberFormat")
# column B of B5 is formatted ahead: 0.0000, currency with cents, text, date, percent, General with a value
for k in range(1, 7):
    T("TMoney", f"cf{k:02d}", f"{B5}.Cells({k}, 2).Value = CCur(1.23456)", f"V: {B5}.Cells({k}, 2).Value2", f"S: {B5}.Cells({k}, 2).NumberFormat",
      f"S: TypeName({B5}.Cells({k}, 2).Value)")
T("TMoney", "curcopy", f"{B5}.Range(\"D1\").Value = {B5}.Range(\"C1\").Value", f"V: {B5}.Range(\"D1\").Value2", f"S: {B5}.Range(\"D1\").NumberFormat",
  f"arr = {B5}.Range(\"C1:C2\").Value", f"{B5}.Range(\"E1:E2\").Value = arr", f"V: {B5}.Range(\"E1\").Value2", f"V: {B5}.Range(\"E2\").Value2",
  f"S: {B5}.Range(\"E1\").NumberFormat", f"{B5}.Range(\"F1\").Value2 = {B5}.Range(\"C1\").Value", f"V: {B5}.Range(\"F1\").Value2",
  f"S: {B5}.Range(\"F1\").NumberFormat", f"{B5}.Range(\"G1\").Value = {B5}.Range(\"C1\").Value2", f"V: {B5}.Range(\"G1\").Value2",
  f"S: {B5}.Range(\"G1\").NumberFormat", f"{B5}.Range(\"H1\").Value = {B5}.Range(\"C1\").Value * 2", f"V: {B5}.Range(\"H1\").Value2",
  f"{B5}.Range(\"H2\").Value = {B5}.Range(\"C1\").Value * 1.5", f"V: {B5}.Range(\"H2\").Value2", f"S: {B5}.Range(\"H2\").NumberFormat",
  f"{B5}.Range(\"H3\").Value = {B5}.Range(\"C1\").Value / 3", f"V: {B5}.Range(\"H3\").Value2")
T("TMoney", "v2date", f"{B5}.Range(\"I1\").Value2 = DateSerial(2024, 1, 5)", f"S: {B5}.Range(\"I1\").NumberFormat",
  f"S: TypeName({B5}.Range(\"I1\").Value)", f"{B5}.Range(\"I2\").Value2 = \"=1+1\"", f"V: {B5}.Range(\"I2\").Value",
  f"{B5}.Range(\"I3\").Value2 = \"5%\"", f"V: {B5}.Range(\"I3\").Value2", f"S: {B5}.Range(\"I3\").NumberFormat",
  f"{B5}.Range(\"I4\").Formula = DateSerial(2024, 1, 5)", f"S: {B5}.Range(\"I4\").NumberFormat", f"V: {B5}.Range(\"I4\").Value2",
  f"{B5}.Range(\"I5\").Formula = 2.5", f"V: {B5}.Range(\"I5\").Value2", f"{B5}.Range(\"I6\").Formula = True", f"V: {B5}.Range(\"I6\").Value2")
CS2 = W("CS2")
T("TMoney", "paste5", f"{CS2}.Range(\"C4\").Copy {CS2}.Range(\"D4\")", f"S: {CS2}.Range(\"D4\").NumberFormat",
  f"{CS2}.Rows(7).Copy {B5}.Rows(31)", f"S: {B5}.Range(\"A31\").NumberFormat", f"S: {B5}.Range(\"B31\").NumberFormat",
  f"S: {B5}.Range(\"C31\").NumberFormat", f"S: {B5}.Range(\"D31\").NumberFormat", f"{CS2}.Range(\"D2\").Value = 5",
  f"{CS2}.Range(\"D2\").Copy {CS2}.Range(\"A9\")", f"S: {CS2}.Range(\"A9\").NumberFormat", f"V: {CS2}.Range(\"A9\").Value",
  f"{CS2}.Range(\"A10\").Value = \"0012\"", f"{CS2}.Range(\"A10\").Copy {CS2}.Range(\"D10\")", f"S: {CS2}.Range(\"D10\").NumberFormat",
  f"V: {CS2}.Range(\"D10\").Value", f"{CS2}.Range(\"A10:B10\").Copy {CS2}.Range(\"D12\")", f"S: {CS2}.Range(\"E12\").NumberFormat",
  f"{CS2}.Range(\"B14\").Value = 45296", f"S: TypeName({CS2}.Range(\"B14\").Value)", f"{CS2}.Range(\"B14\").Copy {CS2}.Range(\"D14\")",
  f"S: TypeName({CS2}.Range(\"D14\").Value)", f"S: {CS2}.Range(\"D14\").NumberFormat")
T("TMoney", "typed15", f"{B5}.Range(\"K1\").Value = \"2.9999999999999996\"", f"V: {B5}.Range(\"K1\").Value2",
  f"{B5}.Range(\"K2\").Value = \"1234567890123456789\"", f"V: {B5}.Range(\"K2\").Value2", f"{B5}.Range(\"K3\").Value = \"0.1234567890123456789\"",
  f"V: {B5}.Range(\"K3\").Value2", f"{B5}.Range(\"K4\").Value = \"0.30000000000000004\"", f"V: {B5}.Range(\"K4\").Value2 = 0.3",
  f"{B5}.Range(\"K5\").Value = 0.30000000000000004", f"V: {B5}.Range(\"K5\").Value2 = 0.3", f"{B5}.Range(\"K6\").Value = \"9999999999999999\"",
  f"V: {B5}.Range(\"K6\").Value2", f"{B5}.Range(\"K7\").Value = \"1.00000000000000049\"", f"V: {B5}.Range(\"K7\").Value2 = 1",
  "V: Val(\"1234567890123456789\")", "V: CDbl(\"1234567890123456789\")", "V: CDbl(\"2.9999999999999996\") = 3")
T("TMoney", "shown", f"{B5}.Range(\"L1\").Value = True", f"S: {B5}.Range(\"L1\").Text", f"{B5}.Range(\"L2\").Value = CVErr(xlErrNA)",
  f"S: {B5}.Range(\"L2\").Text", f"{B5}.Range(\"L3\").Value = 12345678", f"S: {B5}.Range(\"L3\").Text", f"{B5}.Range(\"L4\").Value = -1234567",
  f"S: {B5}.Range(\"L4\").Text", f"{B5}.Range(\"L5\").Value = 1234.567", f"S: {B5}.Range(\"L5\").Text", f"{B5}.Range(\"L6\").Value = -123.4567",
  f"S: {B5}.Range(\"L6\").Text", f"{B5}.Range(\"L7\").Value = 0.5", f"S: {B5}.Range(\"L7\").Text", f"{B5}.Range(\"L8\").Value = 123456789",
  f"S: {B5}.Range(\"L8\").Text", f"{B5}.Range(\"L9\").Value = 1E+20", f"S: {B5}.Range(\"L9\").Text", f"{B5}.Range(\"M1\").Value = 123456789",
  f"S: {B5}.Range(\"M1\").Text", f"{B5}.Range(\"M2\").Value = 1234.56789012", f"S: {B5}.Range(\"M2\").Text",
  f"{B5}.Range(\"M3\").Value = 123456789012", f"S: {B5}.Range(\"M3\").Text", f"{B5}.Range(\"M4\").Value = 1234567890123456",
  f"S: {B5}.Range(\"M4\").Text", f"{B5}.Range(\"M5\").Value = -12345.678", f"S: {B5}.Range(\"M5\").Text")
FORMULAS5 = ["=COUNT(1,,2)", "=AVERAGE(1,,2)", "=MAX(-1,)", "=SUM(1,,2)", "=COUNTA(1,,2)", "=AND(TRUE,)", "=SUM(IF(TRUE,D2,0),1)",
             "=SUM(INDEX(D1:D5,2),1)", "=COUNT(INDEX(D1:D5,3))", '=SUM(IF(TRUE,"5",0),1)', "=AND(IF(TRUE,D2,0),TRUE)", "=ROUND(2.5-2^-51,0)",
             "=ROUND(2.5-2^-50,0)", "=ROUND(1.005-2^-52,2)", "=INT(3-2^-51)", "=MIN(,5)", "=OR(,FALSE)", "=COUNT(D3)", "=COUNT(TRUE)",
             "=SUM(D3,1)", "=SUM((D2),1)", '=COUNTA(D4,"")', "=AVERAGE(D2:D4)", "=MAX(D2:D2)", "=MIN(D4)", "=SUM(XLOOKUP(1,F1:F3,G1:G3),1)",
             "=INDEX(F1:G3,2,2)&1", "=SUM(1E+308,1E+308)", "=ABS(-1E+308*2)", "=1E+308+1E+308", "=5/0+1", '=IF(TRUE,"a")&1',
             "=AVERAGE(0.1,0.2,-0.3,0)", "=SUM(0.1,0.2,-0.3)/1", "=MAX(0.3,0.1+0.2)=0.3", "=MIN(0.3,0.1+0.2)", "=0.1+0.2+0.3-0.6",
             "=COUNT(0.1+0.2=0.3)", "=(0.1+0.2=0.3)+0", "=IF(0.1+0.2>0.3,1,0)", "=MATCH(0.3,Z2:Z2,0)", '=XLOOKUP(0.3,Z2:Z2,AA1:AA1,"none")',
             "=VLOOKUP(0.3,Z2:AA2,1,FALSE)", "=COUNT(Z2:Z3)", "=INT(Z2*10)", "=ROUND(Z2,17)", "=Z2*3=0.9", "=1-0.9-0.1=0", "=(1-0.9-0.1)=0"]
for k, f in enumerate(FORMULAS5, 1):
    cell = f'{W("FX")}.Range("AV{k}")'
    T("TFormula3", f"j{k:03d}", *F4_SETUP, f"{cell}.Formula = {q(f)}", f"V: {cell}.Value")
II5 = [("G9", "=INDEX(F1:G3,2)"), ("G10", "=F1:G1"), ("F11", "=INDEX(F1:G3,3)"), ("AW2", "=INDEX(F1:G3,0,1)"), ("AW9", "=INDEX(F1:G3,0,1)"),
       ("AX2", "=INDEX(F1:F3,2)"), ("AX3", "=INDEX(F1:G1,2)"), ("AX4", "=INDEX(F1:G3,2,0)"), ("AX5", "=INDEX(F1:G3,0,0)"),
       ("AX6", "=INDEX(F1:G3,4,1)"), ("AX7", "=INDEX(F1:G3,1,3)"), ("AX8", "=INDEX(F1:F3,1,1)"), ("AX9", "=INDEX(F1:G3,1)")]
for k, (cell, f) in enumerate(II5, 1):
    T("TFormula3", f"ix{k:02d}", f'{W("FX")}.Range("{cell}").Formula = {q(f)}', f'V: {W("FX")}.Range("{cell}").Value')

SHEETS += [("B5", [10, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 9.14, 20],
            [[None, (None, "n0000"), (1.23456, "c8")], [None, (None, "c8"), (2.5, "c8")], [None, (None, "text")],
             [None, (None, "date")], [None, (None, "percent")], [None, 7]]),
           ("CS2", [(10, "text"), (10, "date"), (10, "money"), 10], [])]


# ---------------------------------------------------------------------------- a write into a cell that has a number format
# What a write leaves in a cell that already has a number format: each entry (a row) into each format (a column).
B6_FORMATS = ["percent", "p2", "n2", "money", "date", "d14", "n0000", "c8", "text", "sci", "frac", "t20"]
B6_ENTRIES = ['"5"', '"5%"', '"$5"', '"1,234"', '"0.5"', "5", "DateSerial(2024, 1, 5)", "CCur(1.239)", "True", '"1 1/2"', '"1e3"',
              "TimeSerial(13, 5, 0)", '"=1+1"', "0.25"]
B6 = W("B6")
for i, entry_ in enumerate(B6_ENTRIES, 1):
    for part in range(2):
        lines = []
        for j in range(part * 6 + 1, part * 6 + 7):
            lines += [f"{B6}.Cells({i}, {j}).Value = {entry_}", f"V: {B6}.Cells({i}, {j}).Value2", f"S: {B6}.Cells({i}, {j}).NumberFormat"]
        T("TKept", f"m{i:02d}{'ab'[part]}", *lines)
T("TKept", "bigformula", f'E: {B6}.Range("N1").Formula = "=1E+308*10"', f'E: {B6}.Range("N2").Formula = "=1E+307*10"',
  f'V: {B6}.Range("N2").Value', f'E: {B6}.Range("N3").Value = "=9.99999999999999E+307"', f'V: {B6}.Range("N3").Value',
  f'E: {B6}.Range("N4").Formula = "=SUM("', f'E: {B6}.Range("N5").Formula = "=1+"', f'E: {B6}.Range("N6").Formula = "=NOSUCHFN(1)"',
  f'V: {B6}.Range("N6").Value', f'E: {B6}.Range("N7").Formula = "=A1:"', f'{B6}.Range("N8").Value = "1E+308"', f'V: {B6}.Range("N8").Value',
  f'{B6}.Range("N9").Value = "9.9E+307"', f'V: {B6}.Range("N9").Value')

SHEETS.append(("B6", [12] * 14, [[(None, style) for style in B6_FORMATS] for _ in B6_ENTRIES]))


B7_FORMATS = ["d22", "d15", "t18", "n37", "n41", "c44", "t45", "t46", "f13", "c7", "d14", "percent", "t21", "n39", "s48"]
B7_ENTRIES = ["DateSerial(2024, 1, 5) + TimeSerial(13, 5, 0)", "DateSerial(2024, 1, 5)", "TimeSerial(13, 5, 0)", '"5%"', '"$5"',
              '"1 1/2"', '"1e3"', "CCur(1.239)"]
B7 = W("B7")
for i, entry_ in enumerate(B7_ENTRIES, 1):
    T("TKept2", f"n{i:02d}", *[x for j in range(1, len(B7_FORMATS) + 1)
                             for x in (f"{B7}.Cells({i}, {j}).Value = {entry_}", f"S: {B7}.Cells({i}, {j}).NumberFormat")])

SHEETS.append(("B7", [12] * 15, [[(None, style) for style in B7_FORMATS] for _ in B7_ENTRIES]))


# ------------------------------------------------------------------------------------------ one of each, left to check
FX, TXS = W("FX"), W("TX")
T("TChecks", "emptystring", f'{S("BX1")}.Value = "x"', f'{S("BX1")}.Value = ""', f'V: IsEmpty({S("BX1")}.Value)',
  f'N: Len({S("BX1")}.Formula)', f'V: {S("BX1")}.Value')
T("TChecks", "textformula", f'{TXS}.Range("A15").Formula = "=1+1"', f'V: {TXS}.Range("A15").Value', f'S: {TXS}.Range("A15").NumberFormat',
  f'{TXS}.Range("A16").Formula2 = "=1+1"', f'V: {TXS}.Range("A16").Value', f'{TXS}.Range("A17").Value = 5', f'V: {TXS}.Range("A17").Value',
  f'S: {TXS}.Range("A17").Text')
T("TChecks", "spillblank", f'{FX}.Range("BA1").Value = 3', f'{FX}.Range("BA3").Value = 1', f'{FX}.Range("BB1").Formula2 = "=SORT(BA1:BA3)"',
  f'V: {FX}.Range("BB1").Value', f'V: {FX}.Range("BB2").Value', f'V: {FX}.Range("BB3").Value',
  f'{FX}.Range("BC1").Formula2 = "=UNIQUE(BA1:BA3)"', f'V: {FX}.Range("BC1").Value', f'V: {FX}.Range("BC2").Value',
  f'V: {FX}.Range("BC3").Value')
T("TChecks", "wfsum", "Set wf = Application.WorksheetFunction", "V: wf.Sum(0.1, 0.2, -0.3)", "V: wf.Sum(0.1, 0.2, -0.3, 0)",
  'V: wf.Max("5", "7")', "V: wf.Min(True, 3)", "V: wf.Sum(5.1, -5.05, -0.05)")
T("TChecks", "index0", f'{FX}.Range("BD2").Formula = "=INDEX(F1:G3,0)"', f'V: {FX}.Range("BD2").Value',
  f'{FX}.Range("BE2").Formula = "=SUM(INDEX(F1:G3,0,1))"', f'V: {FX}.Range("BE2").Value',
  f'{FX}.Range("BF2").Formula = "=SUM(INDEX(F1:G3,2,0))"', f'V: {FX}.Range("BF2").Value',
  f'{FX}.Range("BG2").Formula = "=COUNTA(INDEX(F1:G3,0,0))"', f'V: {FX}.Range("BG2").Value',
  f'{FX}.Range("BH2").Formula = "=INDEX(F1:G3,1,-1)"', f'V: {FX}.Range("BH2").Value')
T("TChecks", "af_header", f'{W("FH1")}.Range("A1:E1").AutoFilter Field:=2, Criteria1:="North"', f'S: Vis({W("FH1")})',
  f'{W("FH2")}.Rows(1).AutoFilter Field:=2, Criteria1:="North"', f'S: Vis({W("FH2")})',
  f'{W("FH5")}.Range("B1").AutoFilter Field:=2, Criteria1:="North"', f'S: Vis({W("FH5")})')
T("TChecks", "af_extend", f'{W("FH3")}.Range("A1:E13").AutoFilter Field:=2, Criteria1:="North"', f'{W("FH3")}.Range("A14").Value = 200',
  f'{W("FH3")}.Range("B14").Value = "North"', f'{W("FH3")}.Range("A1:E13").AutoFilter Field:=2, Criteria1:="South"',
  f'V: {W("FH3")}.Rows(14).Hidden', f'S: Vis({W("FH3")})', f'{W("FH3")}.Range("A1:E13").AutoFilter Field:=2, Criteria1:="North"',
  f'V: {W("FH3")}.Rows(14).Hidden')
T("TChecks", "af_unhide", f'{W("FH4")}.Rows(3).Hidden = True', f'{W("FH4")}.Rows(5).Hidden = True',
  f'{W("FH4")}.Range("A1:E13").AutoFilter Field:=2, Criteria1:="south"', f'S: Vis({W("FH4")})', f'{W("FH4")}.ShowAllData',
  f'S: Vis({W("FH4")})', f'{W("FH4")}.AutoFilterMode = False', f'S: Vis({W("FH4")})')
T("TChecks", "filter_mixed", f'{W("FM")}.Range("A1:A8").AutoFilter Field:=1, Criteria1:=">500"', f'S: Vis({W("FM")})',
  f'{W("FM")}.Range("A1:A8").AutoFilter Field:=1, Criteria1:="<600"', f'S: Vis({W("FM")})',
  f'{W("FM")}.Range("A1:A8").AutoFilter Field:=1, Criteria1:="<>600"', f'S: Vis({W("FM")})',
  f'{W("FM")}.Range("A1:A8").AutoFilter Field:=1, Criteria1:=">a"', f'S: Vis({W("FM")})',
  f'{W("FM")}.Range("A1:A8").AutoFilter Field:=1, Criteria1:="700"', f'S: Vis({W("FM")})',
  f'{W("FM")}.Range("A1:A8").AutoFilter Field:=1, Criteria1:="TRUE"', f'S: Vis({W("FM")})')
T("TChecks", "af_new", f'{W("FH6")}.Range("A14").Value = 200', f'{W("FH6")}.Range("B14").Value = "North"',
  f'{W("FH6")}.Range("A1:E13").AutoFilter Field:=2, Criteria1:="South"', f'V: {W("FH6")}.Rows(14).Hidden', f'S: Vis({W("FH6")})',
  f'{W("FH7")}.Range("A15").Value = 200', f'{W("FH7")}.Range("B15").Value = "North"',
  f'{W("FH7")}.Range("A1:E13").AutoFilter Field:=2, Criteria1:="North"',
  f'{W("FH7")}.Range("A1:E13").AutoFilter Field:=2, Criteria1:="South"', f'V: {W("FH7")}.Rows(15).Hidden',
  f'{W("FH7")}.Range("C14").Value = 1', f'{W("FH7")}.Range("A1:E13").AutoFilter Field:=2, Criteria1:="South"',
  f'V: {W("FH7")}.Rows(14).Hidden', f'V: {W("FH7")}.Rows(15).Hidden')
# the last test, as it leaves other sheets selected for a moment
T("TChecks", "group", f'{W("Out")}.Activate', 'Sheets(Array("FH1", "FH2")).Select', "S: ActiveSheet.Name",
  'Sheets(Array("FH2", "FH1")).Select', "S: ActiveSheet.Name", 'Sheets("FH1").Select', 'Sheets("FH3").Select False',
  "S: ActiveSheet.Name", 'Sheets("FH3").Select', 'Sheets(Array("FH1", "FH2")).Select False', "S: ActiveSheet.Name",
  f'{W("Out")}.Select')
SHEETS += [(f"FH{i}", [8, 12, 12, 12, 8], FILTER_ROWS) for i in range(1, 8)]
SHEETS.append(("HC2", [8], [[1], [2], [3]]))
SHEETS.append(("FM", [10], [["K"], [5], ["abc"], [600], ["700"], [True], [None], [400]]))


# ------------------------------------------------------------------------------ group 1: the first lines of a recorded macro
def ws_(sheet):
    return f'Set ws = ThisWorkbook.Worksheets("{sheet}")'


# --- setting NumberFormat
NF, NF2, NF3, NF4, NF5, NF6 = (W(f"NF{i}") for i in ("", 2, 3, 4, 5, 6))
T("TNumFmt", "nf_basic", f'{NF}.Range("A1").Value = 1234.5', f'{NF}.Range("A1").NumberFormat = "0.00"', f'S: {NF}.Range("A1").NumberFormat',
  f'S: {NF}.Range("A1").Text', f'S: TypeName({NF}.Range("A1").Value)', f'{NF}.Range("A2").Value = 1234.5',
  f'{NF}.Range("A2").NumberFormat = "#,##0.00"', f'S: {NF}.Range("A2").Text', f'{NF}.Range("A3").Value = 45296',
  f'{NF}.Range("A3").NumberFormat = "yyyy-mm-dd"', f'S: TypeName({NF}.Range("A3").Value)', f'V: {NF}.Range("A3").Value2',
  f'{NF}.Range("A3").NumberFormat = "0.00"', f'S: TypeName({NF}.Range("A3").Value)', f'{NF}.Range("A4").NumberFormat = "0.0"',
  f'{NF}.Range("A4").Value = 2.25', f'S: {NF}.Range("A4").Text', f'{NF}.Range("A4").Value = "5%"', f'S: {NF}.Range("A4").NumberFormat',
  f'V: {NF}.Range("A4").Value')
T("TNumFmt", "nf_text", f'{NF}.Range("B1").Value = 5', f'{NF}.Range("B1").NumberFormat = "@"', f'S: TypeName({NF}.Range("B1").Value)',
  f'S: {NF}.Range("B1").Text', f'{NF}.Range("B1").Value = {NF}.Range("B1").Value', f'S: TypeName({NF}.Range("B1").Value)',
  f'{NF}.Range("B2").Formula = "=1+1"', f'{NF}.Range("B2").NumberFormat = "@"', f'S: {NF}.Range("B2").Formula', f'V: {NF}.Range("B2").Value',
  f'{NF}.Range("B2").Formula = {NF}.Range("B2").Formula', f'V: {NF}.Range("B2").Value', f'{NF}.Range("B4").NumberFormat = "@"',
  f'{NF}.Range("B4").Value = "007"', f'{NF}.Range("B4").NumberFormat = "General"', f'V: {NF}.Range("B4").Value',
  f'S: {NF}.Range("B4").NumberFormat', f'{NF}.Range("B5").NumberFormat = "@"', f'{NF}.Range("B5").Value = 12', f'V: {NF}.Range("B5").Value')
NF_CODES = ["0.00", "#,##0", "#,##0.00", "0%", "0.00%", "General", "@", "0", "0.0", "#,##0.000", "yyyy-mm-dd", "dd/mm/yyyy",
            "mm/dd/yyyy", "dd-mmm-yy", "mmm-yy", "d-mmm-yyyy", "mmmm yyyy", "hh:mm", "h:mm:ss", "[h]:mm", "dd/mm/yyyy hh:mm",
            "#,##0_);(#,##0)", "#,##0.00_);[Red](#,##0.00)", "0.00;[Red]-0.00", '0.00;-0.00;"-"', "#,##0;-#,##0;-", "$#,##0.00",
            "\\$#,##0.00", '"$"#,##0.00', "[$$-409]#,##0.00", "[$€-407] #,##0.00", "£#,##0.00",
            '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)', "#,##0,", '#,##0,,"M"', "0.00E+00", "# ?/?", "0.00 kg", '0.00" kg"',
            "0.0x", "general", "GENERAL", "MM/DD/YYYY", "DD-MMM-YY", "[red]0.00", "[RED]0.00", "0.00e+00", "h:mm AM/PM", "h:mm am/pm",
            "[$-409]mmmm d, yyyy;@", "[$-F800]dddd, mmmm dd, yyyy", "m/d/yyyy", "m/d/yy", "d/m/yyyy", "yyyy", "mmm", "0.00.00", "[Red",
            "0;0;0;0;0", '"abc', "", " ", "0.00 ", "Text", "abc", "#", "?", "0 0", "00000", "000-00-0000", '@" text"', "0.00;;", ";;;",
            "[Blue]0;[Red]-0", "[>1000]#,##0;0", '[<0]"neg";"pos"', '#,##0.00 "USD"', "#,##0.00 USD", "0.0 %", "0.0%;-0.0%",
            "#,##0.0_);(#,##0.0)", "#,##0.00;(#,##0.00)", "0.00x", "Standard"]
for k, code in enumerate(NF_CODES, 1):
    cell = f"{W('NFC')}.Cells({(k - 1) % 40 + 1}, {(k - 1) // 40 + 1})"
    T("TNumFmt", f"code{k:02d}", f"E: {cell}.NumberFormat = {q(code)}", f"S: {cell}.NumberFormat")
T("TNumFmt", "nf_mixed", f'{NF}.Range("C1").NumberFormat = "0.00"', f'{NF}.Range("C2").NumberFormat = "0%"', f'V: {NF}.Range("C1:C2").NumberFormat',
  f'{NF}.Range("C3:C4").NumberFormat = "0.0"', f'V: {NF}.Range("C3:C4").NumberFormat', f'V: {NF}.Range("C5:C6").NumberFormat',
  f'V: {NF}.Range("C4:C5").NumberFormat', f'V: {NF}.Range("C3:D4").NumberFormat', f'{NF}.Range("C7").Value = 1',
  f'{NF}.Range("C7:C8").NumberFormat = "0.000"', f'S: {NF}.Range("C8").NumberFormat', f'S: {NF}.Range("C7").Text')
T("TNumFmt", "nf_rowscols", f'{NF2}.Rows(3).NumberFormat = "0%"', f'{NF2}.Columns("C").NumberFormat = "0.00"', f'S: {NF2}.Range("C1").NumberFormat',
  f'S: {NF2}.Range("C3").NumberFormat', f'S: {NF2}.Range("A3").NumberFormat', f'S: {NF2}.Range("C1048576").NumberFormat',
  f'V: {NF2}.Columns("C").NumberFormat', f'V: {NF2}.Rows(3).NumberFormat', f'{NF2}.Rows(5).NumberFormat = "0.000"',
  f'S: {NF2}.Range("C5").NumberFormat', f'V: {NF2}.Columns("C").NumberFormat', f'{NF2}.Range("C7").Value = 1.5', f'S: {NF2}.Range("C7").Text',
  f'{NF2}.Range("C7").ClearContents', f'S: {NF2}.Range("C7").NumberFormat', f'{NF2}.Range("C7").Clear', f'S: {NF2}.Range("C7").NumberFormat',
  f'{NF2}.Columns("C").Clear', f'S: {NF2}.Range("C9").NumberFormat', f'S: {NF2}.Range("C3").NumberFormat', f'S: {NF2}.Range("C5").NumberFormat')
T("TNumFmt", "nf_used", f'S: {NF3}.UsedRange.Address', f'{NF3}.Columns("D").NumberFormat = "0.00"', f'S: {NF3}.UsedRange.Address',
  f'{NF3}.Rows(10).NumberFormat = "0%"', f'S: {NF3}.UsedRange.Address', f'{NF3}.Range("F20").NumberFormat = "0.0"', f'S: {NF3}.UsedRange.Address',
  f'{NF3}.Range("B2").Value = 1', f'S: {NF3}.UsedRange.Address', f'S: {NF3}.Range("D10").NumberFormat', f'S: {NF3}.Cells(10, 4).End(xlUp).Address')
T("TNumFmt", "nf_cells", f'{NF4}.Range("B2").Value = 5', f'{NF4}.Cells.NumberFormat = "@"', f'S: {NF4}.Range("A1").NumberFormat',
  f'S: {NF4}.Range("XFD1048576").NumberFormat', f'V: {NF4}.Range("B2").Value', f'{NF4}.Range("Z100").Value = "007"', f'V: {NF4}.Range("Z100").Value',
  f'{NF4}.Range("Z101").Value = 5', f'V: {NF4}.Range("Z101").Value', f'V: {NF4}.Cells.NumberFormat', f'{NF4}.Cells.NumberFormat = "General"',
  f'S: {NF4}.Range("Z100").NumberFormat', f'V: {NF4}.Range("Z100").Value', f'V: {NF4}.Cells.NumberFormat', f'S: {NF4}.UsedRange.Address')
T("TNumFmt", "nf_tall", f'{NF5}.Range("D2:D1048576").NumberFormat = "0.0"', f'S: {NF5}.Range("D1").NumberFormat', f'S: {NF5}.Range("D2").NumberFormat',
  f'S: {NF5}.Range("D999999").NumberFormat', f'V: {NF5}.Columns("D").NumberFormat',
  f'{NF5}.Range({NF5}.Cells(3, 6), {NF5}.Cells(3, 16384)).NumberFormat = "0%"', f'S: {NF5}.Range("E3").NumberFormat',
  f'S: {NF5}.Range("XFD3").NumberFormat', f'S: {NF5}.Range("F3").NumberFormat', f'V: {NF5}.Rows(3).NumberFormat', f'S: {NF5}.Range("D3").NumberFormat',
  f'{NF5}.Range("A5:C1048576").NumberFormat = "0.000"', f'S: {NF5}.Range("B4").NumberFormat', f'S: {NF5}.Range("B6").NumberFormat',
  f'S: {NF5}.UsedRange.Address')
T("TNumFmt", "nf_filtered", flt("F43"), f'{W("F43")}.Range("E2:E13").NumberFormat = "0.0"', f'{W("F43")}.Columns("F").NumberFormat = "0.00"',
  f'{W("F43")}.Range("G1:G13").Value = 1', f'{W("F43")}.ShowAllData', f'S: NFJoin({W("F43")}, 5, 13)', f'S: NFJoin({W("F43")}, 6, 13)',
  f'S: ColJoin({W("F43")}, 7, 13)')
T("TNumFmt", "nf_date", f'{NF6}.Range("A1").Value = 45296.5', f'{NF6}.Range("A1").NumberFormat = "dd/mm/yyyy hh:mm"', f'S: TypeName({NF6}.Range("A1").Value)',
  f'{NF6}.Range("A1").NumberFormat = "hh:mm"', f'S: TypeName({NF6}.Range("A1").Value)', f'{NF6}.Range("A1").NumberFormat = "mmm"',
  f'S: TypeName({NF6}.Range("A1").Value)', f'{NF6}.Range("A1").NumberFormat = "General"', f'S: TypeName({NF6}.Range("A1").Value)',
  f'{NF6}.Range("A1").NumberFormat = "[$-409]mmmm d, yyyy;@"', f'S: TypeName({NF6}.Range("A1").Value)', f'{NF6}.Range("A2").Formula = "=A1+1"',
  f'{NF6}.Range("A2").NumberFormat = "yyyy-mm-dd"', f'S: TypeName({NF6}.Range("A2").Value)', f'{NF6}.Range("A3").NumberFormat = "dd/mm/yyyy"',
  f'{NF6}.Range("A3").Value = 5', f'S: {NF6}.Range("A3").NumberFormat', f'S: TypeName({NF6}.Range("A3").Value)',
  f'{NF6}.Range("A4").NumberFormat = "0.00"', f'{NF6}.Range("A4").Value = DateSerial(2024, 1, 5)', f'S: {NF6}.Range("A4").NumberFormat',
  f'S: TypeName({NF6}.Range("A4").Value)')
T("TNumFmt", "nf_args", f'E: {NF6}.Range("C1").NumberFormat = 0', f'S: {NF6}.Range("C1").NumberFormat', f'E: {NF6}.Range("C2").NumberFormat = Null',
  f'S: {NF6}.Range("C2").NumberFormat', f'E: {NF6}.Range("C3").NumberFormat = Empty', f'S: {NF6}.Range("C3").NumberFormat',
  f'E: {NF6}.Range("C4").NumberFormat = True', f'S: {NF6}.Range("C4").NumberFormat', f'E: {NF6}.Range("C5").NumberFormat = 0.5',
  f'S: {NF6}.Range("C5").NumberFormat')
T("TNumFmt", "nf_array", f'E: {NF6}.Range("C6:C7").NumberFormat = Array("0.00", "0%")', f'S: {NF6}.Range("C6").NumberFormat')


# --- sheets hidden and shown
T("TSheets", "vis_values", ws_("V1"), "ws.Visible = False", "V: ws.Visible", "ws.Visible = xlSheetVeryHidden", "V: ws.Visible", "ws.Visible = True",
  "V: ws.Visible", "ws.Visible = 0", "V: ws.Visible", "E: ws.Visible = 1", "V: ws.Visible", "E: ws.Visible = 5", "V: ws.Visible",
  'E: ws.Visible = "False"', "V: ws.Visible", "E: ws.Visible = -1", "V: ws.Visible", "ws.Visible = xlSheetVisible")
T("TSheets", "vis_use", ws_("V2"), "ws.Visible = False", 'ws.Range("A1").Value = 7', 'N: ws.Range("A1").Value', "E: ws.Activate", "E: ws.Select",
  'E: ws.Range("A1").Select', "S: ActiveSheet.Name", "V: ws.Index", "S: Sheets(ws.Index).Name", "ws.Visible = True")
T("TSheets", "vis_active", 'Sheets("V4").Activate', 'Sheets("V4").Visible = False', "S: ActiveSheet.Name", 'Sheets("V4").Visible = True',
  'Sheets("V5").Visible = False', 'Sheets("V4").Activate', 'Sheets("V4").Visible = False', "S: ActiveSheet.Name", 'Sheets("V4").Visible = True',
  'Sheets("V5").Visible = True', 'Sheets("V6").Activate', 'Sheets("V6").Visible = False', "S: ActiveSheet.Name", 'Sheets("V6").Visible = True',
  f'{W("Out")}.Activate', "S: ActiveSheet.Name")
T("TSheets", "vis_group", 'Sheets(Array("V3", "V4")).Select', 'Sheets("V3").Visible = False', "S: ActiveSheet.Name", 'Sheets("V3").Visible = True',
  f'{W("Out")}.Select', "S: ActiveSheet.Name")


# --- protection
def legacy_hash(password):
    """The password hash of a sheet protected the old way: 16 bits."""
    h = 0
    for i, ch in enumerate(password, 1):
        v = ord(ch) << i
        h ^= (v & 0x7FFF) | (v >> 15)
    return f"{h ^ len(password) ^ 0xCE4B:04X}"


def sha512_hash(password, salt, spin):
    h = hashlib.sha512(salt + password.encode("utf-16-le")).digest()
    for i in range(spin):
        h = hashlib.sha512(h + i.to_bytes(4, "little")).digest()
    return base64.b64encode(h).decode()


_SALT = b"macropress-probe"
PROTECTED = {
    "PRL": f'<sheetProtection password="{legacy_hash("secret")}" sheet="1" objects="1" scenarios="1"/>',
    "PRS": f'<sheetProtection algorithmName="SHA-512" hashValue="{sha512_hash("secret", _SALT, 100000)}" '
           f'saltValue="{base64.b64encode(_SALT).decode()}" spinCount="100000" sheet="1" objects="1" scenarios="1"/>',
    "PRO": '<sheetProtection sheet="1" objects="1" scenarios="1" formatCells="0"/>',
    "PRO2": '<sheetProtection sheet="1" objects="1" scenarios="1" formatCells="0"/>',
}
PR_ROWS = [[1, (2, "unlocked"), 3, 4], [5], [6], [7], [8]]
T("TProtect", "pr_basic", ws_("PR1"), 'ws.Protect "pw"', "V: ws.ProtectContents", 'E: ws.Range("A1").Value = 5', 'E: ws.Range("B1").Value = 5',
  'V: ws.Range("B1").Value', 'V: ws.Range("A1").Value', 'E: ws.Range("A1:B1").Value = 7', 'V: ws.Range("B1").Value', 'E: ws.Range("A1").ClearContents',
  'E: ws.Range("B1").ClearContents', 'E: ws.Range("A1").NumberFormat = "0.00"', 'E: ws.Range("B1").NumberFormat = "0.00"', 'E: ws.Unprotect "wrong"',
  'E: ws.Unprotect "pw"', "V: ws.ProtectContents")
T("TProtect", "pr_ops", ws_("PR2"), 'ws.Protect "pw"', 'E: ws.Range("A1").Locked = False', 'E: ws.Range("A1").Formula = "=1+1"',
  'E: ws.Range("C1").Copy ws.Range("A2")', 'E: ws.Range("C1").Copy ws.Range("B1")', 'V: ws.Range("B1").Value', "E: ws.Rows(3).Hidden = True",
  'E: ws.Range("A2:A5").Sort Key1:=ws.Range("A2"), Order1:=xlDescending, Header:=xlNo', "S: ColJoin(ws, 1, 5)", 'E: ws.Range("A1:D5").AutoFilter',
  'E: ws.Name = "PR2x"', "S: ws.Name", 'ws.Name = "PR2"', "E: ws.Calculate", 'V: ws.Range("A1").Locked', 'V: ws.Range("B1").Locked',
  'V: ws.Range("A1:B1").Locked', 'ws.Unprotect "pw"')
T("TProtect", "pr_uionly", ws_("PR3"), 'ws.Protect "pw", UserInterfaceOnly:=True', 'E: ws.Range("A1").Value = 5', 'E: ws.Range("A1").NumberFormat = "0.00"',
  "E: ws.Rows(3).Hidden = True", "V: ws.Rows(3).Hidden", 'E: ws.Range("A2:A5").Sort Key1:=ws.Range("A2"), Order1:=xlDescending, Header:=xlNo',
  "S: ColJoin(ws, 1, 5)", 'E: ws.Range("A1").Locked = False', "V: ws.ProtectContents", "V: ws.ProtectionMode",
  'E: ws.Range("A1:D5").AutoFilter Field:=1, Criteria1:="5"', "V: ws.AutoFilterMode", "ws.AutoFilterMode = False", "ws.Rows(3).Hidden = False",
  'ws.Unprotect "pw"', "V: ws.ProtectionMode", "V: ws.ProtectContents")
T("TProtect", "pr_allow", ws_("PR4"), 'ws.Protect "pw", AllowFormattingCells:=True', 'E: ws.Range("A1").NumberFormat = "0.00"',
  'S: ws.Range("A1").NumberFormat', 'E: ws.Range("A1").Value = 9', "V: ws.Protection.AllowFormattingCells", "E: ws.Rows(3).Hidden = True",
  'ws.Unprotect "pw"', 'ws.Protect "pw", AllowFormattingRows:=True, AllowFormattingColumns:=True', "E: ws.Rows(3).Hidden = True",
  "E: ws.Columns(5).Hidden = True", "V: ws.Protection.AllowFormattingCells", 'E: ws.Range("A1").NumberFormat = "0%"', 'ws.Unprotect "pw"',
  "ws.Rows(3).Hidden = False", "ws.Columns(5).Hidden = False")
T("TProtect", "pr_filter", ws_("PR5"), 'ws.Range("A1:E13").AutoFilter', 'ws.Protect "pw", AllowFiltering:=True',
  'E: ws.Range("A1:E13").AutoFilter Field:=2, Criteria1:="North"', "S: Vis(ws)", "E: ws.ShowAllData", "S: Vis(ws)", 'E: ws.Range("A1:E13").AutoFilter',
  "V: ws.AutoFilterMode", 'ws.Unprotect "pw"', 'ws.Protect "pw"', 'E: ws.Range("A1:E13").AutoFilter Field:=2, Criteria1:="North"', "S: Vis(ws)",
  'ws.Unprotect "pw"', "ws.AutoFilterMode = False")
T("TProtect", "pr_twice", ws_("PR6"), 'ws.Protect "pw"', 'E: ws.Protect "other"', 'E: ws.Unprotect "pw"', "V: ws.ProtectContents",
  'E: ws.Unprotect "other"', "V: ws.ProtectContents", "ws.Protect", "E: ws.Unprotect", "V: ws.ProtectContents", 'ws.Protect ""',
  'E: ws.Unprotect ""', "V: ws.ProtectContents", 'ws.Protect "pw"', 'E: ws.Protect "pw", AllowFormattingCells:=True',
  "V: ws.Protection.AllowFormattingCells", 'E: ws.Range("A1").NumberFormat = "0.00"', 'E: ws.Unprotect "pw"', "V: ws.ProtectContents",
  'E: ws.Unprotect "pw"')
for sheet in ("PRL", "PRS"):
    T("TProtect", f"pr_file_{sheet.lower()}", ws_(sheet), "V: ws.ProtectContents", 'E: ws.Range("A1").Value = 1', 'E: ws.Range("B1").Value = 1',
      'V: ws.Range("B1").Locked', 'E: ws.Unprotect "wrong"', "V: ws.ProtectContents", 'E: ws.Unprotect "secret"', "V: ws.ProtectContents",
      'E: ws.Range("A1").Value = 1', 'E: ws.Protect "secret"', "V: ws.ProtectContents", 'E: ws.Unprotect "secret"')
T("TProtect", "pr_file_options", ws_("PRO"), "V: ws.ProtectContents", "V: ws.Protection.AllowFormattingCells", 'E: ws.Range("A1").NumberFormat = "0.00"',
  'E: ws.Range("A1").Value = 1', "V: ws.ProtectDrawingObjects", "E: ws.Unprotect", "V: ws.ProtectContents", 'E: ws.Range("A1").Value = 1')
T("TProtect", "pr_locked", ws_("PR7"), 'V: ws.Range("A1").Locked', 'V: ws.Range("B1").Locked', 'V: ws.Range("A1:B1").Locked',
  'ws.Range("C1").Locked = False', 'V: ws.Range("C1").Locked', 'ws.Columns("F").Locked = False', 'V: ws.Range("F7").Locked',
  'V: ws.Columns("F").Locked', 'ws.Range("C1").Clear', 'V: ws.Range("C1").Locked', 'ws.Range("F7").Value = 1', 'ws.Range("F7").ClearContents',
  'V: ws.Range("F7").Locked', 'ws.Range("B1").Copy ws.Range("G1")', 'V: ws.Range("G1").Locked', 'ws.Range("B1").Copy',
  'ws.Range("G2").PasteSpecial xlPasteValues', 'V: ws.Range("G2").Locked', "Application.CutCopyMode = False", 'ws.Range("A1").Copy ws.Range("F8")',
  'V: ws.Range("F8").Locked', 'ws.Protect "pw"', 'E: ws.Range("F7").Value = 3', 'E: ws.Range("G1").Value = 3', 'E: ws.Range("F9").Value = 3',
  'ws.Unprotect "pw"')
T("TProtect", "pr_book", ws_("PB"), 'ThisWorkbook.Protect "pw", True', "V: ThisWorkbook.ProtectStructure", 'E: ws.Name = "PBx"', "E: ws.Visible = False",
  "V: ws.Visible", 'E: ws.Range("A1").Value = 1', 'E: ws.Protect "x"', "V: ws.ProtectContents", 'ws.Unprotect "x"', 'E: ThisWorkbook.Unprotect "wrong"',
  "V: ThisWorkbook.ProtectStructure", 'E: ThisWorkbook.Unprotect "pw"', "V: ThisWorkbook.ProtectStructure", 'E: ws.Name = "PBx"', "S: ws.Name",
  'ws.Name = "PB"', "ws.Visible = True")


# --- FormulaR1C1
RC, RCE = W("RC"), W("RCE")
T("TR1C1", "rc_read", f'{RC}.Range("D5").Formula = "=A1+B$2+$A1+$C$3"', f'S: {RC}.Range("D5").FormulaR1C1',
  f'{RC}.Range("D6").Formula = "=SUM(A1:B2)+SUM(A:A)+SUM($A:$B)+SUM(1:1)+SUM($2:$3)"', f'S: {RC}.Range("D6").FormulaR1C1',
  f'{RC}.Range("D7").Formula = "=RCData!A1+\'My R1C1\'!$B$2"', f'S: {RC}.Range("D7").FormulaR1C1',
  f'{RC}.Range("D8").Formula = "=ROUND(D5,2)&""RC[1]"""', f'S: {RC}.Range("D8").FormulaR1C1', f'{RC}.Range("D9").Value = 5',
  f'S: {RC}.Range("D9").FormulaR1C1', f'S: {RC}.Range("D10").FormulaR1C1', f'{RC}.Range("D11").Formula = "=SUM(D5:D5)"',
  f'S: {RC}.Range("D11").FormulaR1C1', f'{RC}.Range("D12").Formula = "=D11"', f'S: {RC}.Range("D12").FormulaR1C1', f'{RC}.Range("D13").Value = "abc"',
  f'S: {RC}.Range("D13").FormulaR1C1', f'{RC}.Range("D14").Value = True', f'S: {RC}.Range("D14").FormulaR1C1', f'S: {RC}.Range("D5").Formula2R1C1')
T("TR1C1", "rc_write", f'{RC}.Range("F5").FormulaR1C1 = "=R[-1]C[-1]+R1C1+R[-4]C1+R2C[-2]"', f'S: {RC}.Range("F5").Formula',
  f'{RC}.Range("F6").FormulaR1C1 = "=SUM(R1C1:R[-1]C[-1])"', f'S: {RC}.Range("F6").Formula', f'{RC}.Range("F7").FormulaR1C1 = "=SUM(C[-5])+SUM(R[-6])"',
  f'S: {RC}.Range("F7").Formula', f'{RC}.Range("F8").FormulaR1C1 = "=SUM(C1:C2)+SUM(R1:R2)"', f'S: {RC}.Range("F8").Formula',
  f'{RC}.Range("F9").FormulaR1C1 = "=RCData!RC[-5]+\'My R1C1\'!R1C1"', f'S: {RC}.Range("F9").Formula',
  f'{RC}.Range("F10").FormulaR1C1 = "=ROUND(R[-5]C,2)&""R1C1"""', f'S: {RC}.Range("F10").Formula', f'{RC}.Range("F11").FormulaR1C1 = "=r[-1]c+1"',
  f'S: {RC}.Range("F11").Formula', f'{RC}.Range("F12").FormulaR1C1 = "=COUNT(R[-11]C[-5]:R[-1]C[-5])"', f'S: {RC}.Range("F12").Formula',
  f'V: {RC}.Range("F12").Value', f'V: {RC}.Range("F5").Value', f'S: {RC}.Range("F12").FormulaR1C1')
T("TR1C1", "rc_fill", f'{RC}.Range("H2:H4").FormulaR1C1 = "=RC[-7]*2"', f'S: {RC}.Range("H4").Formula', f'V: {RC}.Range("H4").Value',
  f'{RC}.Range("I2:J3").FormulaR1C1 = "=R1C[-7]+R[-1]C1"', f'S: {RC}.Range("J3").Formula', f'S: {RC}.Range("I2").Formula',
  f'{RC}.Range("K2:K3").Formula2R1C1 = "=R[-1]C[-10]"', f'S: {RC}.Range("K3").Formula', f'V: {RC}.Range("K3").Value')
T("TR1C1", "rc_wrap", f'E: {RCE}.Range("A1").FormulaR1C1 = "=R[-1]C"', f'S: {RCE}.Range("A1").Formula', f'E: {RCE}.Range("A2").FormulaR1C1 = "=RC[-1]"',
  f'S: {RCE}.Range("A2").Formula', f'E: {RCE}.Range("B3").FormulaR1C1 = "=R[1048575]C"', f'S: {RCE}.Range("B3").Formula',
  f'E: {RCE}.Range("B4").FormulaR1C1 = "=R[-5]C[-3]"', f'S: {RCE}.Range("B4").Formula', f'{RCE}.Range("C5").Formula = "=C1048576"',
  f'S: {RCE}.Range("C5").FormulaR1C1', f'{RCE}.Range("C6").Formula = "=XFD6"', f'S: {RCE}.Range("C6").FormulaR1C1')
T("TR1C1", "rc_bad", f'E: {RCE}.Range("D1").FormulaR1C1 = "=R0C1"', f'S: {RCE}.Range("D1").Formula', f'E: {RCE}.Range("D2").FormulaR1C1 = "=R1048577C1"',
  f'S: {RCE}.Range("D2").Formula', f'E: {RCE}.Range("D4").Formula = "=R1C1"',
  f'S: {RCE}.Range("D4").Formula', f'V: {RCE}.Range("D4").Value', f'E: {RCE}.Range("D5").Formula = "=R[-1]C"', f'S: {RCE}.Range("D5").Formula',
  f'E: {RCE}.Range("D6").FormulaR1C1 = "=R1C16385"', f'S: {RCE}.Range("D6").Formula', f'E: {RCE}.Range("D7").FormulaR1C1 = "=SUM(R1C1:R2)"',
  f'S: {RCE}.Range("D7").Formula')
T("TR1C1", "rc_a1", f'E: {RCE}.Range("D3").FormulaR1C1 = "=A1+1"', f'S: {RCE}.Range("D3").Formula')
T("TR1C1", "rc_text", f'{TXS}.Range("A18").FormulaR1C1 = "=R1C1"', f'V: {TXS}.Range("A18").Value', f'S: {TXS}.Range("A18").FormulaR1C1',
  f'{RC}.Range("L1").FormulaR1C1 = 5', f'V: {RC}.Range("L1").Value', f'{RC}.Range("L2").FormulaR1C1 = "abc"', f'V: {RC}.Range("L2").Value',
  f'{RC}.Range("L3").FormulaR1C1 = "5%"', f'V: {RC}.Range("L3").Value', f'S: {RC}.Range("L3").NumberFormat', f'{RC}.Range("L4").FormulaR1C1 = ""',
  f'V: {RC}.Range("L4").Value')
T("TR1C1", "rc_words", f'{RC}.Range("N5").FormulaR1C1 = "=ROUND(RC[-13],0)+COUNT(R1C1:R2C2)"', f'S: {RC}.Range("N5").Formula', f'V: {RC}.Range("N5").Value',
  f'{RC}.Range("N6").Formula = "=ROUND(A6,0)+COUNTA(A1:B2)"', f'S: {RC}.Range("N6").FormulaR1C1', f'{RC}.Range("N7").Formula = "=TRUE"',
  f'S: {RC}.Range("N7").FormulaR1C1', f'{RC}.Range("N8").FormulaR1C1 = "=IF(TRUE,R1C1,R2C2)"', f'S: {RC}.Range("N8").Formula',
  f'{RC}.Range("N9").FormulaR1C1 = "=RCData!R1C1:R2C2"', f'S: {RC}.Range("N9").Formula')


# --- filling
FL1, FL2, FL3, PL, PL2 = W("FL1"), W("FL2"), W("FL3"), W("PL"), W("PL2")
T("TFill", "fill_down", f'{FL1}.Range("A1").Value = 5', f'{FL1}.Range("B1").Formula = "=A1*2"', f'{FL1}.Range("C1").Value = "x"',
  f'{FL1}.Range("A1").NumberFormat = "0.00"', f'{FL1}.Range("A3").NumberFormat = "0%"', f'{FL1}.Range("C3").Value = "old"', f'{FL1}.Range("A1:C4").FillDown',
  f'S: ColJoin({FL1}, 1, 5)', f'S: FJoin({FL1}, 2, 5)', f'V: {FL1}.Range("B4").Value', f'S: ColJoin({FL1}, 3, 5)', f'S: NFJoin({FL1}, 1, 5)',
  f'S: {FL1}.Range("A4").Text')
T("TFill", "fill_one", f'{FL2}.Range("A1").Value = 1', f'{FL2}.Range("A2").Value = 2', f'{FL2}.Range("A2").FillDown', f'V: {FL2}.Range("A2").Value',
  f'{FL2}.Range("B2").Formula = "=A2+1"', f'{FL2}.Range("B3").FillDown', f'S: {FL2}.Range("B3").Formula', f'E: {FL2}.Range("C1").FillDown',
  f'V: {FL2}.Range("C1").Value', f'{FL2}.Range("D1").Value = 4', f'E: {FL2}.Range("D1").FillDown', f'V: {FL2}.Range("D1").Value',
  f'{FL2}.Range("E3").Value = 6', f'{FL2}.Range("F3").FillRight', f'V: {FL2}.Range("F3").Value', f'{FL2}.Range("A5").Value = 8',
  f'E: {FL2}.Range("A5").FillRight', f'V: {FL2}.Range("A5").Value')
T("TFill", "fill_dirs", f'{FL3}.Range("A1").Value = 3', f'{FL3}.Range("A2").Formula = "=A1+1"', f'{FL3}.Range("A1:D2").FillRight', f'S: {FL3}.Range("D2").Formula',
  f'V: {FL3}.Range("D1").Value', f'{FL3}.Range("F5").Formula = "=F6*2"', f'{FL3}.Range("F6").Value = 4', f'{FL3}.Range("F2:F5").FillUp', f'S: FJoin({FL3}, 6, 6)',
  f'V: {FL3}.Range("F2").Value', f'{FL3}.Range("K8").Value = 9', f'{FL3}.Range("H8:K8").FillLeft', f'S: RowJoin({FL3}, 8, 12)')
T("TFill", "fill_filtered", flt("F44"), f'{W("F44")}.Range("E2:E13").FillDown', f'{W("F44")}.ShowAllData', f'S: ColJoin({W("F44")}, 5, 13)')
T("TFill", "paste_tile", f'{PL}.Range("A1").Value = 1', f'{PL}.Range("A1").Copy {PL}.Range("C1:D3")', f'S: ColJoin({PL}, 3, 4)', f'S: ColJoin({PL}, 4, 4)',
  f'{PL}.Range("A2").Formula = "=A1+1"', f'{PL}.Range("A2").Copy {PL}.Range("F2:F4")', f'S: FJoin({PL}, 6, 5)', f'{PL}.Range("H1").Value = "a"',
  f'{PL}.Range("H2").Value = "b"', f'E: {PL}.Range("H1:H2").Copy {PL}.Range("J1:J4")', f'S: ColJoin({PL}, 10, 5)',
  f'E: {PL}.Range("H1:H2").Copy {PL}.Range("K1:K3")', f'S: ColJoin({PL}, 11, 4)', f'E: {PL}.Range("H1:H2").Copy {PL}.Range("L1:N1")',
  f'S: ColJoin({PL}, 12, 3) & ColJoin({PL}, 13, 3)', f'E: {PL}.Range("H1:H2").Copy {PL}.Range("P1:Q4")', f'S: ColJoin({PL}, 17, 5)')
T("TFill", "paste_tile2", f'{PL2}.Range("A1").Value = 1', f'{PL2}.Range("B1").Value = 2', f'{PL2}.Range("A2").Value = 3', f'{PL2}.Range("B2").Value = 4',
  f'E: {PL2}.Range("A1:B2").Copy {PL2}.Range("D1:G3")', f'S: RowJoin({PL2}, 1, 8)', f'S: RowJoin({PL2}, 3, 8)', f'{PL2}.Range("A1:B2").Copy',
  f'E: {PL2}.Range("D6:G9").PasteSpecial xlPasteValues', f'S: RowJoin({PL2}, 9, 8)', f'E: {PL2}.Range("A12:C12").PasteSpecial xlPasteValues',
  f'S: RowJoin({PL2}, 12, 4)', f'S: RowJoin({PL2}, 13, 4)', "Application.CutCopyMode = False", f'{PL2}.Range("A1").Copy', f'{PL2}.Range("Z1").Value = 5',
  "V: Application.CutCopyMode", f'E: {PL2}.Range("A20").PasteSpecial xlPasteValues', f'V: {PL2}.Range("A20").Value', "Application.CutCopyMode = False")
AF1, AF2, AF3 = W("AF1"), W("AF2"), W("AF3")
T("TFill", "af_copy", f'{AF1}.Range("A1").Value = 5', f'{AF1}.Range("A1").NumberFormat = "0.00"', f'{AF1}.Range("A1").AutoFill Destination:={AF1}.Range("A1:A4")',
  f'S: ColJoin({AF1}, 1, 5)', f'S: NFJoin({AF1}, 1, 4)', f'{AF1}.Range("B1").Formula = "=A1*10"', f'{AF1}.Range("B1").AutoFill {AF1}.Range("B1:B4")',
  f'S: FJoin({AF1}, 2, 5)', f'{AF1}.Range("C1").Value = "x"', f'{AF1}.Range("C1").AutoFill {AF1}.Range("C1:C3")', f'S: ColJoin({AF1}, 3, 4)',
  f'{AF1}.Range("D1").Value = True', f'{AF1}.Range("D1").AutoFill {AF1}.Range("D1:D3")', f'S: ColJoin({AF1}, 4, 3)',
  f'{AF1}.Range("E1").Value = CVErr(xlErrNA)', f'{AF1}.Range("E1").AutoFill {AF1}.Range("E1:E3")', f'S: ColJoin({AF1}, 5, 3)',
  f'{AF1}.Range("G3").Value = 2', f'{AF1}.Range("G1").AutoFill {AF1}.Range("G1:G3")', f'S: ColJoin({AF1}, 7, 3)', f'V: {AF1}.Range("A1").AutoFill({AF1}.Range("A1:A2"))')
T("TFill", "af_more", f'{AF2}.Range("A1").Value = 1', f'{AF2}.Range("A2").Value = 10', f'{AF2}.Range("B1").Formula = "=A1+1"', f'{AF2}.Range("B2").Formula = "=A2*2"',
  f'{AF2}.Range("B1:B2").AutoFill {AF2}.Range("B1:B6")', f'S: FJoin({AF2}, 2, 6)', f'{AF2}.Range("D6").Formula = "=D7+1"', f'{AF2}.Range("D7").Value = 1',
  f'{AF2}.Range("D6").AutoFill {AF2}.Range("D3:D6")', f'S: FJoin({AF2}, 4, 7)', f'{AF2}.Range("H2").Formula = "=I2*2"', f'{AF2}.Range("H2").AutoFill {AF2}.Range("F2:H2")',
  f'S: {AF2}.Range("F2").Formula', f'{AF2}.Range("A10").Formula = "=B10"', f'{AF2}.Range("B10").Value = 3', f'{AF2}.Range("A10:B10").AutoFill {AF2}.Range("A10:B12")',
  f'S: FJoin({AF2}, 1, 12)', f'S: ColJoin({AF2}, 2, 12)', f'E: {AF2}.Range("A1").AutoFill {AF2}.Range("C1:C4")',
  f'E: {AF2}.Range("A1").AutoFill {AF2}.Range("A1")', f'V: {AF2}.Range("A1").Value', f'E: {AF2}.Range("A1:B1").AutoFill {AF2}.Range("A1:A5")')
T("TFill", "af_types", f'{AF3}.Range("A1").Value = 1', f'{AF3}.Range("A2").Value = 2', f'{AF3}.Range("A1:A2").AutoFill {AF3}.Range("A1:A6"), xlFillCopy',
  f'S: ColJoin({AF3}, 1, 6)', f'{AF3}.Range("B1").Value = 5', f'{AF3}.Range("B1").NumberFormat = "0.0"', f'{AF3}.Range("B3").Value = 9',
  f'{AF3}.Range("B1").AutoFill {AF3}.Range("B1:B3"), xlFillFormats', f'S: ColJoin({AF3}, 2, 3)', f'S: NFJoin({AF3}, 2, 3)', f'{AF3}.Range("C1").Formula = "=A1*3"',
  f'{AF3}.Range("C1").NumberFormat = "0.0"', f'{AF3}.Range("C1").AutoFill {AF3}.Range("C1:C3"), xlFillValues', f'S: FJoin({AF3}, 3, 3)',
  f'S: NFJoin({AF3}, 3, 3)', f'{AF3}.Range("D1").Value = "Item"', f'{AF3}.Range("D1").AutoFill {AF3}.Range("D1:D3"), xlFillCopy', f'S: ColJoin({AF3}, 4, 3)',
  f'{AF3}.Range("E1").Value = "Item 1"', f'{AF3}.Range("E1").AutoFill {AF3}.Range("E1:E3"), xlFillCopy', f'S: ColJoin({AF3}, 5, 3)')
SERIES = [("numbers", ["1", "2"], "A1:A2", "A1:A5", 5, None), ("trend", ["1", "2", "4"], "A1:A3", "A1:A5", 5, None),
          ("date", ["DateSerial(2024, 1, 30)"], "A1", "A1:A3", 3, None), ("item", ['"Item 1"'], "A1", "A1:A3", 3, None),
          ("month", ['"Jan"'], "A1", "A1:A3", 3, None), ("weekday", ['"Mon"'], "A1", "A1:A3", 3, None),
          ("quarter", ['"Q1"'], "A1", "A1:A5", 5, None), ("fillseries", ["1"], "A1", "A1:A3", 3, "xlFillSeries"),
          ("days", ["DateSerial(2024, 1, 30)"], "A1", "A1:A3", 3, "xlFillDays"), ("text2", ['"a"', '"b"'], "A1:A2", "A1:A5", 5, None)]
for k, (name, values, source, dest, n, kind) in enumerate(SERIES, 1):
    sheet = W(f"AS{k}")
    T("TFill", f"af_s_{name}", *[f"{sheet}.Cells({i}, 1).Value = {v}" for i, v in enumerate(values, 1)],
      f'{sheet}.Range("{source}").AutoFill {sheet}.Range("{dest}")' + (f", {kind}" if kind else ""), f"S: ColJoin({sheet}, 1, {n})")

# --- cut
CT = {i: W(f"CT{i}") for i in range(1, 9)}
T("TCut", "cut_refs", f'{CT[1]}.Range("A1").Value = 1', f'{CT[1]}.Range("A2").Value = 2', f'{CT[1]}.Range("A1").NumberFormat = "0.00"',
  f'{CT[1]}.Range("B1").Formula = "=A1*10"', f'{CT[1]}.Range("B2").Formula = "=SUM(A1:A2)"', f'{CT[1]}.Range("B3").Formula = "=SUM(A1:A5)"',
  f'{CT[1]}.Range("B4").Formula = "=C1"', f'{CT[1]}.Range("B5").Formula = "=SUM(C1:C5)"', f'{CT[1]}.Range("B6").Formula = "=SUM(C1:C2)"',
  f'{CT[1]}.Range("B7").Formula = "=$A$1+A$2"', f'{CT[1]}.Range("C1").Value = 9', f'{CT[1]}.Range("A1:A2").Cut Destination:={CT[1]}.Range("C1")',
  f'S: FJoin({CT[1]}, 2, 7)', f'S: ColJoin({CT[1]}, 1, 2)', f'S: ColJoin({CT[1]}, 3, 2)', f'S: NFJoin({CT[1]}, 1, 2)', f'S: NFJoin({CT[1]}, 3, 2)',
  f'V: {CT[1]}.Range("B1").Value', f'V: {CT[1]}.Range("B2").Value', "V: Application.CutCopyMode")
T("TCut", "cut_inside", f'{CT[2]}.Range("A1").Value = 5', f'{CT[2]}.Range("A2").Formula = "=A1+1"', f'{CT[2]}.Range("A3").Formula = "=B1"',
  f'{CT[2]}.Range("B1").Value = 7', f'{CT[2]}.Range("A4").Formula = "=SUM(A1:A2)"', f'{CT[2]}.Range("A1:A4").Cut {CT[2]}.Range("D4")',
  f'S: FJoin({CT[2]}, 4, 8)', f'S: ColJoin({CT[2]}, 4, 8)', f'S: FJoin({CT[2]}, 1, 4)')
T("TCut", "cut_overlap", f'{CT[3]}.Range("A1").Value = 1', f'{CT[3]}.Range("A2").Value = 2', f'{CT[3]}.Range("A3").Value = 3',
  f'{CT[3]}.Range("B1").Formula = "=SUM(A1:A3)"', f'{CT[3]}.Range("B2").Formula = "=A3"', f'{CT[3]}.Range("B3").Formula = "=A4"',
  f'{CT[3]}.Range("A1:A3").Cut {CT[3]}.Range("A2")', f'S: ColJoin({CT[3]}, 1, 4)', f'S: FJoin({CT[3]}, 2, 3)')
T("TCut", "cut_sheets", f'{CT[4]}.Range("A1").Value = 3', f'{CT[4]}.Range("B1").Formula = "=A1*2"', f'{CT[5]}.Range("A1").Formula = "=CT4!A1+1"',
  f'{CT[4]}.Range("A1").Cut {CT[5]}.Range("C3")', f'S: {CT[4]}.Range("B1").Formula', f'S: FJoin({CT[5]}, 1, 1)', f'V: {CT[5]}.Range("A1").Value',
  f'V: {CT[4]}.Range("B1").Value', f'V: {CT[5]}.Range("C3").Value')
T("TCut", "cut_sheets_range", f'{W("CT11")}.Range("A1").Value = 3', f'{W("CT12")}.Range("A2").Formula = "=SUM(CT11!A1:A2)"',
  f'{W("CT11")}.Range("A1").Cut {W("CT12")}.Range("C3")', f'S: {W("CT12")}.Range("A2").Formula')
T("TCut", "cut_paste", f'{CT[6]}.Range("A1").Value = 4', f'{CT[6]}.Range("A2").Formula = "=A1*2"', f'{CT[6]}.Activate', f'{CT[6]}.Range("A1").Cut',
  "V: Application.CutCopyMode", f'{CT[6]}.Range("D1").Select', "ActiveSheet.Paste", "V: Application.CutCopyMode", f'S: ColJoin({CT[6]}, 4, 1)',
  f'S: {CT[6]}.Range("A2").Formula', f'{CT[6]}.Range("A2").Copy', f'{CT[6]}.Range("F2").Select', "ActiveSheet.Paste", f'S: {CT[6]}.Range("F2").Formula',
  "V: Application.CutCopyMode", f"ActiveSheet.Paste Destination:={CT[6]}.Range(\"G5\")", f'S: {CT[6]}.Range("G5").Formula',
  "Application.CutCopyMode = False", f'E: {CT[6]}.Range("A2").Cut', f'E: {CT[6]}.Range("H1").PasteSpecial xlPasteValues', f'V: {CT[6]}.Range("H1").Value',
  "Application.CutCopyMode = False", f'S: {CT[6]}.Range("A2").Formula', f'{W("Out")}.Activate')
T("TCut", "cut_sizes", f'{CT[7]}.Range("A1").Value = 1', f'{CT[7]}.Range("A2").Value = 2', f'E: {CT[7]}.Range("A1:A2").Cut {CT[7]}.Range("C1:C5")',
  f'S: ColJoin({CT[7]}, 3, 5)', f'S: ColJoin({CT[7]}, 1, 2)', f'E: {CT[7]}.Range("C1:C2").Cut {CT[7]}.Range("E1:F1")', f'S: ColJoin({CT[7]}, 5, 2)',
  f'S: ColJoin({CT[7]}, 3, 2)')
T("TCut", "cut_filtered", flt("F45"), f'E: {W("F45")}.Range("E2:E13").Cut {W("F45")}.Range("G2")', f'{W("F45")}.ShowAllData',
  f'S: ColJoin({W("F45")}, 7, 13)', f'S: ColJoin({W("F45")}, 5, 13)')
T("TCut", "cut_rows", f'{CT[8]}.Range("A1").Value = 1', f'{CT[8]}.Range("B1").Value = 2', f'{CT[8]}.Range("A2").Formula = "=A1+B1"',
  f'{CT[8]}.Rows(1).Cut {CT[8]}.Rows(5)', f'S: RowJoin({CT[8]}, 5, 3)', f'S: {CT[8]}.Range("A2").Formula', f'V: {CT[8]}.Range("A2").Value',
  f'S: RowJoin({CT[8]}, 1, 3)')

# --- what the first answers left open
NF7 = W("NF7")
T("TNumFmt", "nf_money", *[x for k, code in enumerate(["$#,##0.00", "\\$#,##0.00", '"$"#,##0.00', "[$$-409]#,##0.00", "£#,##0.00", "#,##0.00"], 1)
                           for x in (f"{NF7}.Cells({k}, 1).Value = 1.5", f"{NF7}.Cells({k}, 1).NumberFormat = {q(code)}",
                                     f"S: TypeName({NF7}.Cells({k}, 1).Value)")])
T("TProtect", "pr_partial", ws_("PR8"), 'ws.Range("A1").Locked = False', 'ws.Range("B1").Locked = True', 'ws.Protect "pw"',
  'E: ws.Range("A1:B1").Value = 7', 'V: ws.Range("A1").Value', 'E: ws.Range("A2:B2").Value = Array(8, 9)', 'V: ws.Range("A2").Value',
  'ws.Unprotect "pw"', 'ws.Range("A3").Locked = False', 'ws.Range("A3").Value = 3', 'ws.Range("B3").Value = 4', 'ws.Protect "pw"',
  'E: ws.Range("A3:B3").ClearContents', 'V: ws.Range("A3").Value', 'E: ws.Range("A4:B4").Formula = "=1+1"', 'V: ws.Range("A4").Value',
  'ws.Range("D1").Value = 6', 'E: ws.Range("D1").Copy ws.Range("A5:B5")', 'V: ws.Range("A5").Value', 'ws.Unprotect "pw"')
T("TProtect", "pr_paste_unlocked", ws_("PR9"), 'ws.Range("C1").NumberFormat = "0.00"', 'ws.Protect "pw"', 'E: ws.Range("C1").Copy ws.Range("B1")',
  'V: ws.Range("B1").Locked', 'S: ws.Range("B1").NumberFormat', 'E: ws.Range("B1").ClearContents', 'E: ws.Range("C1:D1").Copy ws.Range("B1")',
  'ws.Unprotect "pw"', 'ws.Protect "pw"', 'E: ws.Protect "other", AllowFormattingCells:=True', "V: ws.Protection.AllowFormattingCells",
  'E: ws.Unprotect "other"', "V: ws.ProtectContents", 'E: ws.Unprotect "pw"', "V: ws.ProtectContents", 'ThisWorkbook.Protect "pw"',
  "V: ThisWorkbook.ProtectStructure", "V: ThisWorkbook.ProtectWindows", 'ThisWorkbook.Unprotect "pw"', "V: ThisWorkbook.ProtectStructure")
T("TProtect", "pr_fill", ws_("PR10"), 'ws.Range("B2").Locked = False', 'ws.Range("B3").Locked = False', 'ws.Protect "pw"',
  'E: ws.Range("B1:B3").FillDown', 'V: ws.Range("B3").Value', 'E: ws.Range("A1:A3").FillDown', 'V: ws.Range("A3").Value',
  'E: ws.Range("B1").AutoFill ws.Range("B1:B3")', 'E: ws.Range("B2").Cut ws.Range("B3")', 'V: ws.Range("B3").Value', 'ws.Unprotect "pw"')
CT9, CT10 = W("CT9"), W("CT10")
T("TCut", "cut_partial", *[f'{CT9}.Range("A{i}").Value = {i}' for i in range(1, 6)], f'{CT9}.Range("B1").Formula = "=SUM(A1:A5)"',
  f'{CT9}.Range("B2").Formula = "=SUM(A1:A5)"', f'{CT9}.Range("B3").Formula = "=SUM(A2:A4)"', f'{CT9}.Range("B4").Formula = "=SUM(A1:B1)"',
  f'{CT9}.Range("B5").Formula = "=SUM(A4:A5)"', f'{CT9}.Range("A4:A5").Cut {CT9}.Range("D1")', f'S: FJoin({CT9}, 2, 5)',
  f'{CT9}.Range("A2:A3").Cut {CT9}.Range("D5")', f'S: FJoin({CT9}, 2, 5)')
T("TCut", "cut_edge", f'{CT10}.Range("A1").Value = 1', f'{CT10}.Range("A2").Value = 2',
  f'{CT10}.Range("A3").Value = 3', f'{CT10}.Range("B1").Formula = "=SUM(A1:A3)"', f'{CT10}.Range("B2").Formula = "=SUM(A1:A2)"',
  f'{CT10}.Range("A3").Cut {CT10}.Range("A5")', f'S: FJoin({CT10}, 2, 2)', f'{CT10}.Range("A1").Cut {CT10}.Range("C1")', f'S: FJoin({CT10}, 2, 2)')
T("TFill", "fill_back", f'{W("FL5")}.Range("B5").Value = 3', f'{W("FL5")}.Range("B4").FillUp', f'V: {W("FL5")}.Range("B4").Value',
  f'{W("FL5")}.Range("D2").Value = 4', f'{W("FL5")}.Range("C2").FillLeft', f'V: {W("FL5")}.Range("C2").Value', f'{W("FL5")}.Range("B1048576").Value = 2',
  f'E: {W("FL5")}.Range("B1048576").FillUp', f'{W("FL5")}.Range("XFD3").Value = 2', f'E: {W("FL5")}.Range("XFD3").FillLeft',
  f'V: {W("FL5")}.Range("XFD3").Value')
T("TFill", "fill_hidden", f'{W("FL4")}.Range("A1").Value = 5', f'{W("FL4")}.Rows(3).Hidden = True', f'{W("FL4")}.Range("A1:A4").FillDown',
  f'S: ColJoin({W("FL4")}, 1, 4)', f'{W("FL4")}.Range("B1").Value = 7', f'{W("FL4")}.Range("B1").AutoFill {W("FL4")}.Range("B1:B4")',
  f'S: ColJoin({W("FL4")}, 2, 4)', f'{W("FL4")}.Range("C1").Value = 1', f'{W("FL4")}.Range("C1").Copy {W("FL4")}.Range("C2:C4")',
  f'S: ColJoin({W("FL4")}, 3, 4)', f'{W("FL4")}.Rows(3).Hidden = False')
T("TFill", "af_filtered", flt("F46"), f'{W("F46")}.Range("E2").AutoFill {W("F46")}.Range("E2:E13")', f'{W("F46")}.ShowAllData',
  f'S: ColJoin({W("F46")}, 5, 13)', f'{W("F46")}.Range("A1:E13").AutoFilter Field:=2, Criteria1:="North"',
  f'{W("F46")}.Range("C2").Copy {W("F46")}.Range("F2:F13")', f'{W("F46")}.ShowAllData', f'S: ColJoin({W("F46")}, 6, 13)')
T("TFill", "af_square", f'{W("AF5")}.Range("A1").Value = 1', f'{W("AF5")}.Range("A2").Value = 10',
  f'E: {W("AF5")}.Range("A1:A2").AutoFill {W("AF5")}.Range("A1:C4")', f'S: ColJoin({W("AF5")}, 1, 4)', f'S: ColJoin({W("AF5")}, 2, 4)')
T("TFill", "af_kinds", f'{W("AF4")}.Range("E1").Value = 4', f'{W("AF4")}.Range("E1").AutoFill {W("AF4")}.Range("E1:E3"), xlFillSeries', f'S: ColJoin({W("AF4")}, 5, 3)',
  f'{W("AF4")}.Range("F1").Value = 2', f'{W("AF4")}.Range("F2").Value = 5', f'{W("AF4")}.Range("F1:F2").AutoFill {W("AF4")}.Range("F1:F4"), xlFillValues',
  f'S: ColJoin({W("AF4")}, 6, 4)', f'{W("AF4")}.Range("G4").Value = 3', f'{W("AF4")}.Range("G5").Value = 5',
  f'{W("AF4")}.Range("G4:G5").AutoFill {W("AF4")}.Range("G1:G5")', f'S: ColJoin({W("AF4")}, 7, 5)', f'{W("AF4")}.Range("H1").Value = 2.5',
  f'{W("AF4")}.Range("H2").Formula = "=H1"', f'E: {W("AF4")}.Range("H1:H2").AutoFill {W("AF4")}.Range("H1:H4")', f'S: FJoin({W("AF4")}, 8, 4)')

T("TProtect", "pr_again", ws_("PR11"), 'ws.Protect "pw", AllowFormattingCells:=True', 'ws.Protect "pw"', "V: ws.Protection.AllowFormattingCells",
  'ws.Protect "pw", UserInterfaceOnly:=True', "V: ws.ProtectionMode", 'ws.Protect "pw"', "V: ws.ProtectionMode", 'ws.Protect UserInterfaceOnly:=True',
  'E: ws.Unprotect "x"', "V: ws.ProtectContents", 'E: ws.Unprotect "pw"', "V: ws.ProtectContents", "ws.Protect", 'ws.Protect "pw"',
  'E: ws.Unprotect "x"', "V: ws.ProtectContents", 'E: ws.Unprotect "pw"', 'ws.Protect "pw"', 'ws.Protect "other"', 'E: ws.Unprotect "other"',
  "V: ws.ProtectContents", 'E: ws.Unprotect "pw"', "V: ws.ProtectContents")
T("TProtect", "pr_file_again", ws_("PRO2"), "ws.Protect UserInterfaceOnly:=True", "V: ws.Protection.AllowFormattingCells", "V: ws.ProtectionMode",
  'E: ws.Range("A1").Value = 3', "E: ws.Unprotect", "V: ws.ProtectContents")
T("TProtect", "pr_more_ops", ws_("PR12"), 'ws.Range("B1:B3").Locked = False', 'ws.Range("B1").NumberFormat = "0.00"', 'ws.Range("B2").Value = 9',
  'ws.Protect "pw"', 'E: ws.Range("B1:B2").FillDown', 'S: ws.Range("B2").NumberFormat',
  'V: ws.Range("B2").Value', 'E: ws.Range("B1:B3").Sort Key1:=ws.Range("B1"), Order1:=xlDescending, Header:=xlNo', 'E: ws.Visible = False',
  "V: ws.Visible", "ws.Visible = True", 'E: ws.Range("C1").Locked = False', 'ws.Unprotect "pw"', 'ws.Protect "pw", AllowSorting:=True, AllowFormattingCells:=True',
  'E: ws.Range("B1:B3").Sort Key1:=ws.Range("B1"), Order1:=xlDescending, Header:=xlNo', 'S: ColJoin(ws, 2, 3)', 'E: ws.Range("C1").Locked = False',
  'V: ws.Range("C1").Locked', 'ws.Unprotect "pw"')
T("TProtect", "pr_clear", ws_("PR14"), 'ws.Range("B1:B3").Locked = False', 'ws.Range("B1:B2").NumberFormat = "0.00"', 'ws.Range("B1:B2").Value = 9',
  'ws.Protect "pw"', 'E: ws.Range("B1").Clear', 'S: ws.Range("B1").NumberFormat', 'V: ws.Range("B1").Locked', 'V: ws.Range("B1").Value',
  'E: ws.Range("B2").ClearContents', 'S: ws.Range("B2").NumberFormat', 'E: ws.Range("B3").Clear', 'V: ws.Range("B3").Locked', 'ws.Unprotect "pw"')
T("TProtect", "pr_filtered", ws_("PR13"), 'ws.Range("A1:E13").AutoFilter Field:=2, Criteria1:="North"', 'ws.Protect "pw", AllowFiltering:=True',
  "E: ws.ShowAllData", "S: Vis(ws)", "E: ws.AutoFilterMode = False", "V: ws.AutoFilterMode", 'ws.Unprotect "pw"', "ws.AutoFilterMode = False")

SHEETS += [(name, [10], []) for name in ("NF", "NF2", "NF3", "NF4", "NF5", "NF6", "NFC", "NF7")]
SHEETS += [(f"F{i}", [8, 12, 12, 12, 8], FILTER_ROWS) for i in (43, 44, 45, 46)]
SHEETS += [(f"PR{i}", [8] * 5, PR_ROWS) for i in (1, 2, 3, 6, 7, 8, 9, 10, 11, 12, 14)] + [(f"PR{i}", [8, 12, 12, 12, 8], FILTER_ROWS) for i in (4, 5, 13)]
SHEETS += [(name, [8] * 5, PR_ROWS, {"protection": xml}) for name, xml in PROTECTED.items()] + [("PB", [8], [[1]])]
SHEETS += [("RC", [8] * 5, [[1, 2, 3, 4, 5], [6, 7, 8, 9, 10], [11, 12, 13, 14, 15], [16, 17, 18, 19, 20]]), ("RCData", [8], [[100], [200]]),
           ("My R1C1", [8], [[1000], [2000]]), ("RCE", [8], [])]
SHEETS += [(name, [8] * 12, []) for name in ("FL1", "FL2", "FL3", "FL4", "FL5", "PL", "PL2", "AF1", "AF2", "AF3", "AF4", "AF5")]
SHEETS += [(f"AS{k}", [8], []) for k in range(1, len(SERIES) + 1)] + [(f"CT{i}", [8] * 8, []) for i in range(1, 13)]


# ------------------------------------------------------------------------------- group 2.1: arrays in formulas
# AR and AR2 hold the same data: A 1-5, B text, C 10-50, D mixed, E three numbers whose sum all but cancels. Formulas
# go to columns H on, a test to its own column, and the row a formula is in is the row it takes of A1:A5.
AR_ROWS = [[1, "x", 10, 1, 0.1], [2, "yy", 20, "2", 0.2], [3, "x", 30, True, -0.3], [4, "zzz", 40], [5, "x", 50, 4]]


def FA(sheet, cell, formula, dynamic=False):
    """Write a formula (Formula, or Formula2 with dynamic) and read its value."""
    prop = "Formula2" if dynamic else "Formula"
    return [f'{W(sheet)}.Range("{cell}").{prop} = {q(formula)}', f'V: {W(sheet)}.Range("{cell}").Value']


def TA(name, sheet, column, rows_formulas, dynamic=False, module="TArray"):
    T(module, name, *[x for row, f in rows_formulas for x in FA(sheet, f"{column}{row}", f, dynamic)])


TA("ar_const", "AR", "H", [(1, "={1,2;3,4}"), (2, "=SUM({1,2;3,4})"), (3, "=SUM({1,2,3}*2)"), (4, "=SUM({1,2,3}+{10;20})"),
                          (5, "=SUM({1,2,3}+{10,20})"), (6, '={-1,"a",TRUE,#N/A}'), (7, "=INDEX({1,2;3,4},2,1)"),
                          (8, "=MATCH(2,{1,2,3},0)"), (9, '=COUNTA({1,"",TRUE})'), (10, "=SUM({1,#N/A})"), (11, "=COUNT({1,#N/A})"),
                          (12, "=ROWS({1,2;3,4})"), (13, '=AVERAGE({1,2,"x"})'), (14, '=LEN({"ab","c"})'), (15, '=SUM(LEN({"ab","c"}))')])
T("TArray", "ar_const_bad", *[f'E: {W("AR")}.Range("I{k}").Formula = {q(f)}' for k, f in
                              enumerate(["={A1,2}", "={1+1,2}", "={}", "={1,2;3}", "={1,,2}", "={1;2,3;4}", '={"a";"b"}', "={1,-2.5E+3}"], 1)],
  f'V: {W("AR")}.Range("I7").Value', f'V: {W("AR")}.Range("I8").Value', f'S: {W("AR")}.Range("I8").Formula')
TA("ar_ops", "AR", "J", [(3, "=A1:A5^2"), (2, "=A1:A5%"), (4, "=A1:A5>2"), (1, "=(A1:A5)*2"), (5, "=A2:A4*C2:C4"), (6, "=A2:A4*C2:C4"),
                        (7, "=A1:A5*2"), (8, "=A1:A5*{1,2}"), (9, "=B1:B5&C1:C5"), (10, "=-A1:A5"), (11, "=AR2!A1:A5+1"), (12, '=B1:B5="x"')])
TA("ar_ops2", "AR2", "C", [(9, "=A1:C1*2"), (10, "=SUM(A1:A5*{1,2})"), (11, "=A3:A5*{1;2;3}")])  # AR's column C is read whole
TA("ar_ops3", "AR", "K", [(3, "=SUM(A1:A5*{1,2})"), (2, "=A1:A5*{1;2;3;4;5}"), (4, "=SUM(A1:A5*C1:C5)"), (1, "=E1:E3*10"), (5, "=AR2!C1:C5")])
TA("ar_fns", "AR", "L", [(2, "=LEN(B1:B5)"), (3, "=ROUND(E1:E3*10,0)"), (4, "=UPPER(B1:B5)"), (5, "=ISNUMBER(D1:D5)"), (1, "=ABS(E1:E3)"),
                        (6, '=IF(A1:A5>2,"big","small")'), (7, "=NOT(A1:A5>2)"), (8, "=ISBLANK(D1:D5)")])
TA("ar_aggr", "AR", "M", [(4, "=AND(A1:A5>2)"), (2, "=COUNT(A1:A5*1)"), (3, "=AVERAGE(A1:A5*2)"), (5, '=CONCAT(B1:B5&"!")'),
                         (1, "=OR(A1:A5>4)"), (6, "=SUM(IF(A1:A5>2,C1:C5,0))"), (4, "=SUM(IF(A1:A5>2,C1:C5,0))"), (7, "=MAX(A1:A5)"),
                         (8, '=CONCATENATE(B1:B5,"!")'), (3, '=CONCATENATE(B1:B5,"!")')])
TA("ar_sumproduct", "AR", "N", [(1, "=SUMPRODUCT(A1:A5,C1:C5)"), (2, '=SUMPRODUCT((B1:B5="x")*C1:C5)'), (3, "=SUMPRODUCT(--(A1:A5>2))"),
                               (4, "=SUMPRODUCT(A1:A5)"), (5, "=SUMPRODUCT(A1:A5,{1;1;1;1;1})"), (6, "=SUMPRODUCT(A1:A5,{1,1,1,1,1})"),
                               (7, "=SUMPRODUCT(A1:A5,B1:B5)"), (8, "=SUMPRODUCT(A1:A5*B1:B5)"), (9, "=SUMPRODUCT(D1:D5)"),
                               (10, "=SUMPRODUCT(E1:E3)"), (11, "=SUMPRODUCT({1,#N/A})"), (12, "=SUMPRODUCT(A1:A5,2)"), (13, "=SUMPRODUCT(5)"),
                               (14, "=SUMPRODUCT((A:A>2)*1)"), (15, "=SUMPRODUCT((A1:A5>2)+(C1:C5>40))")])
TA("ar_sumproduct2", "AR", "O", [(1, "=SUMPRODUCT(--(B:B=\"x\"),C:C)"), (2, '=SUMPRODUCT((A1:A5>1)*(A1:A5<5)*C1:C5)'), (3, "=SUMPRODUCT(A1:A5/A1:A5)"),
                                (4, "=SUMPRODUCT(1/A1:A5)"), (5, "=SUMPRODUCT({0.1,0.2,-0.3},{1,1,1})"), (6, "=SUMPRODUCT(--({0.1}+{0.2}={0.3}))"),
                                (7, "=SUMPRODUCT(A1:A5*(D1:D5=1))"), (8, "=SUMPRODUCT(A1:A2,C1:C2,A1:A2)"), (9, "=SUMPRODUCT(TRUE)"),
                                (10, '=SUMPRODUCT("3")'), (11, "=SUMPRODUCT(A1:A5,A1:A4)")])
TA("ar_row", "AR", "P", [(7, "=ROW()"), (8, "=COLUMN()"), (9, "=ROW(C5)"), (10, "=COLUMN(C5)"), (11, "=ROW(A2:A4)"), (3, "=ROW(A2:A4)"),
                        (12, "=COLUMN(A1:C1)"), (13, "=ROWS(A1:C5)"), (14, "=COLUMNS(A1:C5)"), (15, "=ROW(INDEX(A1:A5,3))"),
                        (16, "=SUMPRODUCT((ROW(A1:A5)>2)*A1:A5)"), (17, "=ROW(A:A)"), (18, "=ROWS(A:A)"), (19, "=ROW(AR2!F2)"), (20, "=COLUMNS(1:1)")])
TA("ar_refs", "AR", "Q", [(1, "=SUM(IF(TRUE,B1,0),1)"), (2, "=SUM(IF(FALSE,0,A1:A3))"), (3, "=SUM(INDEX(A1:C5,0,3))"), (4, "=SUM(INDEX(A1:C5,2,0))"),
                         (5, "=COUNT(INDEX(D1:D5,3))"), (6, "=SUM(XLOOKUP(2,A1:A5,B1:B5),1)"), (7, "=SUM(IFERROR(B1,0),1)"),
                         (8, "=SUM(IFERROR(A1:A3,0))"), (9, "=IF(TRUE,A1:A5)"), (3, "=IF(TRUE,A1:A5)"), (10, '=SUM(IF(TRUE,"5"),1)'),
                         (11, "=SUM(INDEX({1,\"2\"},2),1)")])
T("TArray", "ar_formula2", *[x for k, f in enumerate(["=A1:A5*2", "=SUM(A1:A5*2)", "=SUM(A1:A5)", "=INDEX(A1:A5,0)", "=LEN(B1:B5)",
                                                       "=SUMPRODUCT(A1:A5,C1:C5)", "={1,2,3}", "=IF(A1:A5>2,1,0)", "=ROW(A1:A3)",
                                                       "=A1+B1", "=A2:A4", "=MAX(A1:A5)", "=AR2!A1:A5"], 1)
                             for x in (f'{W("AR")}.Range("R{k + 1}").Formula = {q(f)}', f'S: {W("AR")}.Range("R{k + 1}").Formula2')])
T("TArray", "ar_at", f'E: {W("AR")}.Range("S3").Formula2 = "=@A1:A5"', f'V: {W("AR")}.Range("S3").Value', f'S: {W("AR")}.Range("S3").Formula',
  f'E: {W("AR")}.Range("S2").Formula2 = "=SUM(@A1:A5*2)"', f'V: {W("AR")}.Range("S2").Value', f'S: {W("AR")}.Range("S2").Formula',
  f'E: {W("AR")}.Range("S4").Formula = "=@A1:A5*2"', f'V: {W("AR")}.Range("S4").Value', f'S: {W("AR")}.Range("S4").Formula2',
  f'E: {W("AR")}.Range("S5").Formula2 = "=@A1:A5+@C1:C5"', f'V: {W("AR")}.Range("S5").Value', f'S: {W("AR")}.Range("S5").Formula')
T("TArray", "ar_formula2b", *[x for k, f in enumerate(["=INDEX(A1:A5,2)", "=IF(TRUE,A1:A5)", "=VLOOKUP(A1:A5,A1:C5,2,FALSE)", "=A1:A5",
                                                        "=SUM(IF(A1:A5>2,C1:C5))", "=ROUND(A1:A5,0)+1", "=-A1:A5", "=(A1:A5)",
                                                        '=A1:A5&"x"', "=SUM({1,2}*A1:A5)", "=ROW()", "=ROWS(A1:A5)",
                                                        "=XLOOKUP(2,A1:A5,B1:B5)", "=INDEX(A1:C5,2,3)", "=IF(A1>0,A2:A3,A4)"], 1)
                              for x in (f'{W("AR")}.Range("U{k}").Formula = {q(f)}', f'S: {W("AR")}.Range("U{k}").Formula2')])
T("TArray", "ar_text", *[x for k, f in enumerate(["=AR2!A1+1", "=sum(a1:a2)", "=ar2!a1", "= A1 + 1", "=SUM( A1 , 2 )", "='AR2'!B2",
                                                  "=Ar2!A1:A2", "=sumproduct(a1:a2,c1:c2)"], 1)
                          for x in (f'{W("AR")}.Range("V{k}").Formula = {q(f)}', f'S: {W("AR")}.Range("V{k}").Formula')])
# the same in dynamic-array formulas, Formula2: each test spills in its own block of AR2
DY = [("ar_dyn1", "H", 1, "=A1:A5*2", ["H1", "H5"]), ("ar_dyn2", "I", 1, "=SUM(A1:A5*2)", ["I1"]), ("ar_dyn3", "J", 1, "=LEN(B1:B5)", ["J2", "J4"]),
      ("ar_dyn4", "K", 1, "=IF(A1:A5>2,C1:C5,0)", ["K1", "K4"]), ("ar_dyn5", "L", 1, "=SUM(IF(A1:A5>2,C1:C5,0))", ["L1"]),
      ("ar_dyn6", "M", 1, "=AND(A1:A5>2)", ["M1"]), ("ar_dyn7", "N", 1, '=CONCAT(B1:B5&"!")', ["N1"]), ("ar_dyn8", "O", 1, "=A1:A3*{1,2}", ["O1", "P3"]),
      ("ar_dyn9", "R", 1, "={1,2,3}+{10;20}", ["R1", "T2"]), ("ar_dyn10", "V", 1, "={1,2,3}+{10,20}", ["V1", "X1"]),
      ("ar_dyn11", "Z", 1, "=ROW(A1:A3)", ["Z1", "Z3"]), ("ar_dyn12", "AA", 1, "=COUNT(A1:A5*1)", ["AA1"]), ("ar_dyn13", "AB", 7, "=ROW()", ["AB7"]),
      ("ar_dyn14", "AC", 1, "=A1:A5>2", ["AC1", "AC5"]), ("ar_dyn15", "AD", 1, "=-A1:A3", ["AD3"]),
      ("ar_dyn16", "AE", 1, "={0.1;0.4}+{0.2;0.5}-{0.3;0.9}", ["AE1", "AE2"]), ("ar_dyn17", "AF", 1, "=1/{1,0,2}", ["AF1", "AG1", "AH1"]),
      ("ar_dyn18", "AI", 1, "=SUM(IFERROR(1/{1,0,2},0))", ["AI1"]), ("ar_dyn19", "AJ", 1, "=COUNT(1/{1,0,2})", ["AJ1"]),
      ("ar_dyn20", "AK", 1, "=SUM(1/{1,0,2})", ["AK1"]), ("ar_dyn21", "AL", 1, "=E1:E3+0", ["AL3"]), ("ar_dyn22", "AM", 1, "=A1:A5*C1:C3", ["AM3", "AM4"]),
      ("ar_dyn23", "AN", 1, "=INDEX(A1:C5,0,3)", ["AN1", "AN5"]), ("ar_dyn24", "AO", 1, '=IF(D1:D5="",0,1)', ["AO3", "AO4"]),
      ("ar_dyn25", "AP", 1, "=ISNUMBER(D1:D5)", ["AP1", "AP2"]), ("ar_dyn26", "AQ", 1, "=E1:E3*10", ["AQ1"])]
for name, col, row, formula, reads in DY:
    T("TArray", name, f'{W("AR2")}.Range("{col}{row}").Formula2 = {q(formula)}', *[f'V: {W("AR2")}.Range("{cell}").Value' for cell in reads])
T("TArray", "ar_legacy_more", *FA("AR", "T1", "=SUM(F1:F3)"), *FA("AR", "T2", "=SUMPRODUCT((B1:B5=\"x\")*(A1:A5>2))"),
  *FA("AR", "T3", "=SUM((A1:A5>2)*1)"), *FA("AR", "T4", "=SUM(--(A1:A5>2))"), *FA("AR", "T5", "=MAX(E1:E3*10)"), *FA("AR", "T6", "=INDEX(C1:C5,A1:A5)"))
SHEETS += [("AR", [8] * 30, AR_ROWS), ("AR2", [8] * 50, AR_ROWS)]

# ------------------------------------------------------------------- group 2.2: conditional sums and counts
# CI's column A holds values of every kind a criterion meets, and B the powers of 2: a SUMIF of B says which rows of A
# matched. C alternates x and y, D mixes numbers with text and TRUE, E and F hold a sum that all but cancels, G three
# criteria. A test writes its formulas to a column of its own (H on); no test writes to A to G.
CI_A = [5, "5", "05", "5.0", " 5", (10, "n2"), "apple", "Apple pie", True, None, "a*b", "x?", "1234567890123456",
        "1234567890123457", 2.5, -3, (d("2024-01-05"), "d14"), "TRUE", "1,000", 0, " apple", "APPLE", "5 ", 1234567890123456]
CI_D = [1, "2", True, 4, None, 8, "abc", 16, 32, 64]
CI_ROWS = [[a, 2**k, "xy"[k % 2], CI_D[k] if k < len(CI_D) else None, *([(0.1, "k"), (0.2, "k"), (-0.3, "k")][k] if k < 3 else (None, None)),
            [5, "apple", ">4"][k] if k < 3 else None] for k, a in enumerate(CI_A)]
CI = W("CI")


def CS(column, criteria, sum_range="B1:B24"):
    """SUMIF over A1:A24 of each criterion, row by row of a column: which rows of A each matched."""
    return list(enumerate([f"=SUMIF(A1:A24,{c},{sum_range})" for c in criteria], 1))


TA("ci_num", "CI", "H", CS("H", ["5", '"5"', '"=5"', '"05"', '"5.0"', '" 5"', '">4"', '"<5"', '"<>5"', '"> 4"', "0", '"1000"',
                                 '"1234567890123457"', '"10.00"', '">=10"']), module="TCond")
TA("ci_text", "CI", "I", CS("I", ['"apple"', '"apple*"', '"*apple"', '"a*b"', '"a~*b"', '"x~?"', '"?"', '"*"', '"<>apple"', '"<>a*"',
                                  '">a"', '"<b"', '"5*"', '"=apple"', '">=Apple pie"']), module="TCond")
TA("ci_blank", "CI", "J", CS("J", ['""', '"="', '"<>"', "A10", '">"', '"<"', "TRUE", '"TRUE"', '"<>TRUE"', '"true"', "A2", "A17",
                                   '">45000"', "FALSE", '"<=0"']), module="TCond")
TA("ci_crit", "CI", "R", CS("R", ['"=5.0"', '"500%"', '"(3)"', '"1e1"', '"=1,000"', '" >4"', '">=5 "', '"==5"', '"=<5"', '"2.50"',
                                  '"= apple"', '"<>5*"', '">"&"4"']) + [(14, f'=COUNTIF(A1:A24,"{"a" * 255}")'),
                                                                        (15, f'=COUNTIF(A1:A24,"{"a" * 256}")')], module="TCond")
TA("ci_count", "CI", "K", list(enumerate(['=COUNTIF(A1:A24,"*")', '=COUNTIF(A1:A24,"<>")', '=COUNTIF(A1:A24,"")', '=COUNTIF(A:A,"")',
                                          '=COUNTIF(A:A,"<>")', "=COUNTIF(A:A,5)", '=COUNTIFS(A1:A24,">0",C1:C24,"x")',
                                          '=COUNTIFS(A1:A24,"<>",C1:C24,"y",B1:B24,">100")', '=COUNTIF(A1:A24,"<>*")',
                                          '=COUNTIF(B:B,">1000")', '=COUNTIFS(A1:A24,">0",C1:C23,"x")', '=COUNTIF(A1:B24,">100")',
                                          '=COUNTIF(C:C,"x")', '=COUNTIFS(C:C,"x",A:A,"<>")', '=COUNTIF(A1:A24,"<>"&"5")'], 1)), module="TCond")
TA("ci_sum", "CI", "L", list(enumerate(["=SUMIF(A1:A24,5)", '=SUMIF(A1:A5,"<>",B1)', '=SUMIF(A1:A5,"<>",B3:B4)', '=SUMIF(A1:A5,"<>",B2:C9)',
                                        '=SUMIF(A1:B2,"<>",B1)', '=SUMIF(C1:C24,"x",D1:D24)', '=SUMIFS(B1:B24,A1:A24,">0",C1:C24,"x")',
                                        '=SUMIFS(B1:B24,A1:A23,">0")', '=SUMIFS(B1:B24,A1:A24,"<>")', '=SUMIF(F1:F3,"k",E1:E3)',
                                        '=SUMIFS(E1:E3,F1:F3,"k")', '=SUMIF(A:A,">0",B:B)', '=SUMIFS(B:B,C:C,"y",A:A,"")',
                                        '=SUMIF(AR!B1:B5,"x",AR!C1:C5)', '=SUMIF(A1:A5,"<>",B23)'], 1)), module="TCond")
TA("ci_avg", "CI", "M", list(enumerate(['=AVERAGEIF(A1:A24,">0")', '=AVERAGEIF(A1:A24,"nothing")', '=AVERAGEIF(C1:C24,"x",B1:B24)',
                                        '=AVERAGEIF(C1:C24,"x",D1:D24)', '=AVERAGEIFS(B1:B24,A1:A24,"<>",C1:C24,"y")',
                                        '=AVERAGEIF(A1:A5,"<>",B1)', '=AVERAGEIF(F1:F3,"k",E1:E3)', '=MAXIFS(B1:B24,A1:A24,"*")',
                                        '=MINIFS(B1:B24,A1:A24,"*")', '=MAXIFS(B1:B24,A1:A24,"nothing")', '=MINIFS(D1:D24,C1:C24,"x")',
                                        '=MAXIFS(B1:B24,A1:A23,"*")', '=MAXIFS(A1:A24,C1:C24,"x")', '=AVERAGEIFS(B1:B24,A1:A24,"nothing")',
                                        '=AVERAGEIF(C1:C24,"y",A1:A24)'], 1)), module="TCond")
# errors, "" and a number off by its last bit, in a test's own columns: O to look in, P the powers of 2, Q to add up
T("TCond", "ci_err", f'{CI}.Range("O1").Value = 1', f'{CI}.Range("O2").Formula = {q("=" + chr(34) * 2)}', f'{CI}.Range("O4").Value = "x"',
  f"{CI}.Range(\"O5\").Value = CVErr(xlErrNA)", f'{CI}.Range("O6").Formula = "=1/0"', f'{CI}.Range("O7").Value = 0.1 + 0.2',
  f'{CI}.Range("O8").Value = "\'#N/A"', "For k = 1 To 8", f"    {CI}.Cells(k, 16).Value = 2 ^ (k - 1)", "Next k",
  f'{CI}.Range("Q1").Value = CVErr(xlErrDiv0)', f'{CI}.Range("Q2").Value = "7"', f'{CI}.Range("Q4").Value = 5', f'{CI}.Range("Q7").Value = 3',
  *[x for row, f in enumerate(['=SUMIF(O1:O8,"",P1:P8)', '=SUMIF(O1:O8,"=",P1:P8)', '=SUMIF(O1:O8,"<>",P1:P8)', '=SUMIF(O1:O8,"*",P1:P8)',
                               '=SUMIF(O1:O8,"#N/A",P1:P8)', '=SUMIF(O1:O8,"#DIV/0!",P1:P8)', '=SUMIF(O1:O8,"<>x",P1:P8)',
                               "=SUMIF(O1:O8,0.3,P1:P8)", '=SUMIF(O1:O8,"<>1",P1:P8)', "=SUMIF(O1:O8,O5,P1:P8)", '=SUMIF(O1:O8,"x",Q1:Q8)',
                               '=SUMIF(O1:O8,"<>x",Q1:Q8)', '=AVERAGEIF(O1:O8,"x",Q1:Q8)', "=SUMIF(O1:O8,1/0,P1:P8)",
                               '=MAXIFS(Q1:Q8,O1:O8,"<>x")'], 1) for x in FA("CI", f"N{row}", f)])
# a criterion given as a range takes the formula's own row of it; an array is gone through value by value
TA("ci_arr", "CI", "S", [(2, "=SUMIF(A1:A24,G1:G3,B1:B24)"), (4, "=SUMIF(A1:A24,G1:G3,B1:B24)"), (5, '=SUM(COUNTIF(A1:A24,{5,"apple"}))'),
                         (6, "=SUMPRODUCT(COUNTIF(A1:A24,G1:G3))"), (1, "=SUM(COUNTIF(A1:A24,G1:G3))"), (7, "=COUNTIF(INDEX(A1:C24,0,1),5)"),
                         (8, '=SUMIF(A1:A24,"apple",INDEX(B1:B24,1))'), (9, "=COUNTIF(IF(TRUE,A1:A24),5)"), (10, '=SUMIF(A1:A24,"apple",B:B)'),
                         (11, '=COUNTIFS(A1:A24,5,A1:A24,"5")'), (12, '=SUMIF(A1:A24,{5,"apple"},B1:B24)'), (13, "=COUNTIF(A1:A24,G1:G3)")], module="TCond")
T("TCond", "ci_view", *[x for k, f in enumerate(["=SUMIF(A1:A24,G1:G3,B1:B24)", "=COUNTIF(A1:A24,G1:G3)", '=SUMIFS(B1:B24,A1:A24,G1:G3,C1:C24,"x")',
                                                  "=COUNTIF(A1:A24,G1)", "=SUM(COUNTIF(A1:A24,G1:G3))", "=SUMPRODUCT(COUNTIF(A1:A24,G1:G3))",
                                                  "=COUNTIF(A:A,G:G)", "=AVERAGEIF(A1:A24,G1:G3)", "=MAXIFS(B1:B24,A1:A24,G1:G3)",
                                                  '=sumif(a1:a24,"x",b1:b24)', '=COUNTIFS(A1:A24,">0",C1:C24,G1:G3)', "=MINIFS(B1:B24,A1:A24,G1:G3)"], 1)
                        for x in (f'{CI}.Range("T{k}").Formula = {q(f)}', f'S: {CI}.Range("T{k}").Formula2')])
# what a SUMIF reads past the sum range written in it, kept current as cells change
T("TCond", "ci_stale", f'{CI}.Range("U1").Formula = "=SUMIF(A1:A5,""<>"",V1)"', f'V: {CI}.Range("U1").Value', f'{CI}.Range("V3").Value = 7',
  f'V: {CI}.Range("U1").Value', f'{CI}.Range("V5").Value = 1', f'V: {CI}.Range("U1").Value', f'{CI}.Range("V6").Value = 100',
  f'V: {CI}.Range("U1").Value', f'{CI}.Range("U2").Formula = "=COUNTIF(V:V,"">0"")"', f'V: {CI}.Range("U2").Value', f'{CI}.Range("V9").Value = 2',
  f'V: {CI}.Range("U2").Value', f'{CI}.Range("U3").Formula = "=AVERAGEIF(A1:A5,""<>"",V1)"', f'V: {CI}.Range("U3").Value',
  f'{CI}.Range("V1").Value = 4', f'V: {CI}.Range("U3").Value', f'{CI}.Range("U4").Formula = "=SUMIFS(V1:V9,V1:V9,"">1"")"',
  f'V: {CI}.Range("U4").Value', f'{CI}.Range("V2").Value = 9', f'V: {CI}.Range("U4").Value', f'V: {CI}.Range("U1").Value')
T("TCond", "ci_dyn", *[x for cell, f, reads in [("W1", "=COUNTIF(A1:A24,G1:G3)", ["W1", "W3"]), ("X1", '=SUMIFS(B1:B24,C1:C24,{"x","y"})', ["X1", "Y1"]),
                                                ("Z1", "=SUMIF(A1:A24,G1:G3,B1:B24)", ["Z1", "Z3"]), ("AA1", "=SUM(COUNTIF(A1:A24,G1:G3))", ["AA1"]),
                                                ("AB1", "=MAXIFS(B1:B24,A1:A24,G1:G3)", ["AB1", "AB3"]),
                                                ("AC1", '=COUNTIF(A1:A24,{5;"apple"})', ["AC1", "AC2"])]
                       for x in [f'{CI}.Range("{cell}").Formula2 = {q(f)}', *[f'V: {CI}.Range("{r}").Value' for r in reads]]])
# a range argument that is not a reference, and arguments left out
T("TCond", "ci_entry", *[x for k, f in enumerate(['=SUMIF({1,2,3},">1")', '=COUNTIF(A1:A5*1,">1")', '=COUNTIF(IF(TRUE,{1,2,3}),">1")',
                                                   '=SUMIF(A1:A24,"x",{1,2,3})', "=SUMIFS(B1:B24,A1:A24)", '=COUNTIFS(A1:A24,">0",C1:C24)',
                                                   "=COUNTIF(A1:A24)"], 1)
                         for x in (f'E: {CI}.Range("AD{k}").Formula = {q(f)}', f'V: {CI}.Range("AD{k}").Value')])
WF = "WorksheetFunction"
CR = {k: f'{CI}.Range("{k}")' for k in ["A1:A24", "B1:B24", "C1:C24", "D1:D24", "A1:A23", "C1:C23", "A1:A5", "B1", "A2", "A:A", "G1:G3", "E1:E3", "F1:F3"]}
T("TCond", "ci_vba1", f'V: {WF}.CountIf({CR["A1:A24"]}, 5)', f'V: {WF}.CountIf({CR["A1:A24"]}, "5")',
  f'V: {WF}.SumIf({CR["A1:A24"]}, ">4", {CR["B1:B24"]})', f'V: {WF}.SumIfs({CR["B1:B24"]}, {CR["A1:A24"]}, ">0", {CR["C1:C24"]}, "x")',
  f'V: {WF}.CountIfs({CR["A1:A24"]}, "<>", {CR["C1:C24"]}, "y")', f'V: {WF}.AverageIf({CR["C1:C24"]}, "x", {CR["B1:B24"]})',
  f'E: x = {WF}.AverageIf({CR["A1:A24"]}, "nothing")', f'V: {WF}.MaxIfs({CR["B1:B24"]}, {CR["A1:A24"]}, "*")',
  f'V: {WF}.MinIfs({CR["B1:B24"]}, {CR["A1:A24"]}, "*")', f'V: {WF}.SumIf({CR["A1:A24"]}, DateSerial(2024, 1, 5), {CR["B1:B24"]})',
  f'V: {WF}.SumIf({CR["A1:A24"]}, True, {CR["B1:B24"]})', f"E: x = {WF}.CountIf(Array(1, 2, 3), 2)",
  f'V: {WF}.SumIf({CR["A1:A5"]}, "<>", {CR["B1"]})', f'E: x = {WF}.SumIfs({CR["B1:B24"]}, {CR["A1:A23"]}, ">0")',
  f'V: {WF}.SumIf({CR["A1:A24"]}, {CR["A2"]}, {CR["B1:B24"]})')
T("TCond", "ci_vba2", f'V: {WF}.CountIf({CR["A:A"]}, "")', f'V: {WF}.SumIf({CR["A1:A24"]}, Empty, {CR["B1:B24"]})',
  f'V: {WF}.SumIf({CR["A1:A24"]}, "apple*", {CR["B1:B24"]})', f'V: {WF}.AverageIfs({CR["B1:B24"]}, {CR["A1:A24"]}, "<>", {CR["C1:C24"]}, "y")',
  f'V: {WF}.SumIf({CR["A1:A24"]}, "", {CR["B1:B24"]})', f'V: {WF}.SumIf({CR["C1:C24"]}, "x", {CR["D1:D24"]})',
  f'V: {WF}.CountIf({CR["A1:A24"]}, {CR["G1:G3"]})', f'V: {WF}.CountIf({CR["A1:A24"]}, "<>" & 5)', f'V: {WF}.SumIf({CR["F1:F3"]}, "k", {CR["E1:E3"]})',
  f'E: x = {WF}.CountIfs({CR["A1:A24"]}, ">0", {CR["C1:C23"]}, "x")', f'V: Application.{WF}.CountIf({CR["A1:A24"]}, "apple")',
  f'V: {WF}.SumIf({CR["A1:A24"]}, 1234567890123457#, {CR["B1:B24"]})', f'V: {WF}.SumIf({CR["A1:A24"]}, CVErr(xlErrNA), {CR["B1:B24"]})',
  f'V: {WF}.SumIf({CR["A1:A24"]}, "<>", {CR["B1:B24"]})', f'V: {WF}.MaxIfs({CR["B1:B24"]}, {CR["A1:A24"]}, "nothing")')
TA("ci_more", "CI", "AE", CS("AE", ['">="', '"<="', '">a*z"', "1", '">FALSE"', '"<TRUE"', '"<>#N/A"', '"=TRUE"', '">=TRUE"', '"=*"', '"~5"',
                                    '"*5*"', '"<>*5*"', '"FALSE"']) + [(15, f'=COUNTIF(A1:A24,"={"a" * 255}")')], module="TCond")
# booleans and errors, in a test's own columns: AG to look in, AH the powers of 2
T("TCond", "ci_more2", f'{CI}.Range("AG1").Value = "abc"', f'{CI}.Range("AG2").Value = True', f'{CI}.Range("AG3").Value = 1',
  f'{CI}.Range("AG4").Value = "\'TRUE"', f'{CI}.Range("AG5").Value = False', f'{CI}.Range("AG6").Value = "\'#N/A"',
  f'{CI}.Range("AG7").Value = CVErr(xlErrNA)', f'{CI}.Range("AG8").Value = "\'-5"', f'{CI}.Range("AG9").Value = 5', "For k = 1 To 9",
  f"    {CI}.Cells(k, 34).Value = 2 ^ (k - 1)", "Next k",
  *[x for row, f in enumerate(["TRUE", "1", '">FALSE"', '"<>#N/A"', "5", '"<=TRUE"', '"<>TRUE"', "FALSE", '"<2"', '"=#N/A"', '"<>1"',
                               '"0"', '">#N/A"'], 1) for x in FA("CI", f"AF{row}", f"=SUMIF(AG1:AG9,{f},AH1:AH9)")])
# text that reads as a date or an amount of money, which regional settings read: in a cell looked in, in a criterion
T("TCond", "ci_regional", f'{CI}.Range("AJ1").Value = "\'1/2"', f'{CI}.Range("AJ2").Value = "\'$5"', f'{CI}.Range("AJ3").Value = 5',
  *FA("CI", "AK1", "=COUNTIF(AJ1:AJ3,5)"), *FA("CI", "AK2", '=COUNTIF(A1:A24,">1/5/2024")'))
# the range arguments Excel refuses as a formula is entered, and arguments left empty
T("TCond", "ci_entry2", *[x for k, f in enumerate(['=COUNTIF(5,">1")', '=COUNTIF("a","a")', '=COUNTIF(SUM(A1:A3),">1")', '=COUNTIF(-A1:A5,">1")',
                                                    "=COUNTIF(,5)", '=SUMIF(A1:A5,">1",)', "=COUNTIF(A1:A5,)"], 1)
                          for x in (f'E: {CI}.Range("AI{k}").Formula = {q(f)}', f'V: {CI}.Range("AI{k}").Value')])
SHEETS += [("CI", [8] * 30, CI_ROWS)]
SHEETS += [(f"V{i}", [8], []) for i in range(1, 7)]  # the last sheets of the workbook: hiding the last is tried

def sheets(index):
    """The workbook's sheets for the builder: Out, with each test's name in column A of its row, then the rest."""
    out = [("Out", [14] + [22] * 15, [[name] for _row, _module, name in index])]
    return [(name, columns, rows, [], *options) for name, columns, rows, *options in out + SHEETS]
