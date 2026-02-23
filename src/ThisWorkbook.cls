Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Private Sub Workbook_Open()
    EnsureCalendarSheet
    SezonlariAlttanKontrolEt ActiveSheet
End Sub
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo SafeExit
    If Sh.Name = "Dashboard" Or Sh.Name = "_TAKVIM" Then Exit Sub

    ' Çoklu yapıştırmalarda da çalışsın diye bunu kaldırıyoruz:
    'If Target.CountLarge > 1 Then Exit Sub

    Application.EnableEvents = False

    ' 1) Durum X olunca tarih yaz (Durum = F(6), Tarih = G(7))
    If Target.Row >= 12 And Target.Column = 6 Then
        If UCase$(Trim$(CStr(Target.Value))) = "X" Then
            If Sh.Cells(Target.Row, 7).Value = "" Then
                Sh.Cells(Target.Row, 7).Value = Date
                Sh.Cells(Target.Row, 7).NumberFormat = "yyyy-mm-dd"
            End If
        End If
    End If

If Not Intersect(Target, Sh.Range("B23:B9999")) Is Nothing Then
    TumSayfalardaSezonlariGuncelle
End If

SafeExit:
    Application.EnableEvents = True
End Sub

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    If Sh.Name = "Dashboard" Or Sh.Name = "_TAKVIM" Then Exit Sub
    If Target.Row < 12 Then Exit Sub

    ' Durum hücresi: E sütunu (6) -> çift tıkla X / - değiştir
    If Target.Column = 6 Then
        Cancel = True
        Application.EnableEvents = False

        If Trim$(CStr(Target.Value)) = "X" Then
            Target.Value = "-"          ' izlenmedi
        Else
            Target.Value = "X"          ' izlendi
        End If

        Application.EnableEvents = True
        Exit Sub
    End If

    ' Tarih hücresi: F sütunu (7) -> çift tıkla takvim aç
    If Target.Column = 7 Then
        Cancel = True
        ShowCalendar Target
        Exit Sub
    End If
End Sub

Public Sub TumSayfalardaSezonlariGuncelle()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Dashboard" And ws.Name <> "_TAKVIM" Then
            SezonlariAlttanKontrolEt ws
        End If
    Next ws
End Sub
