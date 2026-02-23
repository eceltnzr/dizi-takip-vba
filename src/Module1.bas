Attribute VB_Name = "Module1"
Option Explicit

Public gTargetCell As Range

Public Sub EnsureCalendarSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("_TAKVIM")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "_TAKVIM"
        ws.Visible = xlSheetVeryHidden

        ws.Range("B2").Value = "?"
        ws.Range("D2").Value = "AY"
        ws.Range("H2").Value = "?"
        ws.Range("B2:H2").HorizontalAlignment = xlCenter
        ws.Range("B2:H2").Font.Bold = True

        ws.Range("J1").Value = Year(Date)
        ws.Range("K1").Value = Month(Date)
        ws.Range("J1:K1").EntireColumn.Hidden = True

        ws.Range("B3:H3").Value = Array("Pzt", "Sal", "Çar", "Per", "Cum", "Cmt", "Paz")
        ws.Range("B3:H3").Font.Bold = True
        ws.Range("B3:H9").ColumnWidth = 5
        ws.Range("B3:H9").RowHeight = 18
        ws.Range("B3:H9").HorizontalAlignment = xlCenter
        ws.Range("B3:H9").VerticalAlignment = xlCenter
    End If
End Sub

Public Sub ShowCalendar(ByVal Target As Range)
    EnsureCalendarSheet
    Set gTargetCell = Target

    With ThisWorkbook.Worksheets("_TAKVIM")
        .Visible = xlSheetVisible
        .Activate
        DrawCalendar CLng(.Range("J1").Value), CLng(.Range("K1").Value)
    End With
End Sub

Public Sub DrawCalendar(ByVal y As Long, ByVal m As Long)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("_TAKVIM")
    Dim firstDay As Date
    Dim dow As Long, d As Long, r As Long, c As Long, daysInMonth As Long

    ws.Range("D2").Value = Format(DateSerial(y, m, 1), "mmmm yyyy")
    ws.Range("J1").Value = y
    ws.Range("K1").Value = m

    ws.Range("B4:H9").ClearContents

    firstDay = DateSerial(y, m, 1)
    daysInMonth = Day(DateSerial(y, m + 1, 0))
    dow = Weekday(firstDay, vbMonday)

    r = 4
    c = 2 + (dow - 1)

    For d = 1 To daysInMonth
        ws.Cells(r, c).Value = DateSerial(y, m, d)
        ws.Cells(r, c).NumberFormat = "d"
        c = c + 1
        If c > 8 Then
            c = 2
            r = r + 1
        End If
    Next d
End Sub

Public Sub SezonlariAlttanKontrolEt(ByVal ws As Worksheet)

    Dim r As Long
    Dim season As String
    Dim found As Range

    For r = 5 To 19
        season = UCase$(Trim$(CStr(ws.Cells(r, 2).Value))) ' üstte B sütunu

        If season <> "" Then
            Set found = ws.Range("B23:B9999").Find(What:=season, LookIn:=xlValues, LookAt:=xlPart)

            If found Is Nothing Then
                ws.Rows(r).Hidden = True
            Else
                ws.Rows(r).Hidden = False
            End If
        End If
    Next r

End Sub
