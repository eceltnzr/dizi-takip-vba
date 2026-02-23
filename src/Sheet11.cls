Attribute VB_Name = "Sheet11"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Prev/Next ay
    If Not Intersect(Target, Me.Range("B2")) Is Nothing Then
        ChangeMonth -1
        Exit Sub
    End If
    If Not Intersect(Target, Me.Range("H2")) Is Nothing Then
        ChangeMonth 1
        Exit Sub
    End If

    ' Tarih seçimi alanı
    If Intersect(Target, Me.Range("B4:H9")) Is Nothing Then Exit Sub
    If Not IsDate(Target.Value) Then Exit Sub

    If Not gTargetCell Is Nothing Then
        gTargetCell.Value = CDate(Target.Value)
        gTargetCell.NumberFormat = "yyyy-mm-dd"
        Me.Visible = xlSheetVeryHidden
        gTargetCell.Parent.Activate
        gTargetCell.Select
    End If
End Sub

Private Sub ChangeMonth(ByVal delta As Long)
    Dim y As Long, m As Long
    y = CLng(Me.Range("J1").Value)
    m = CLng(Me.Range("K1").Value) + delta
    If m = 0 Then m = 12: y = y - 1
    If m = 13 Then m = 1: y = y + 1
    DrawCalendar y, m
End Sub
