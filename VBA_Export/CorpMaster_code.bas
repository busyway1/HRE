VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CorpMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Not Intersect(Target, ListObjects("Corp").DataBodyRange) Is Nothing Then
        Cancel = True
        Dim cellarray(1 To 9) As Variant
        Set cellarray(1) = ListObjects("Corp").DataBodyRange.Cells(Target.row - ListObjects("Corp").DataBodyRange.row + 1, 1)
        Set cellarray(2) = ListObjects("Corp").DataBodyRange.Cells(Target.row - ListObjects("Corp").DataBodyRange.row + 1, 2)
        Set cellarray(3) = ListObjects("Corp").DataBodyRange.Cells(Target.row - ListObjects("Corp").DataBodyRange.row + 1, 3)
        Set cellarray(4) = ListObjects("Corp").DataBodyRange.Cells(Target.row - ListObjects("Corp").DataBodyRange.row + 1, 4)
        Set cellarray(5) = ListObjects("Corp").DataBodyRange.Cells(Target.row - ListObjects("Corp").DataBodyRange.row + 1, 5)
        Set cellarray(6) = ListObjects("Corp").DataBodyRange.Cells(Target.row - ListObjects("Corp").DataBodyRange.row + 1, 6)
        Set cellarray(7) = ListObjects("Corp").DataBodyRange.Cells(Target.row - ListObjects("Corp").DataBodyRange.row + 1, 7)
        Set cellarray(8) = ListObjects("Corp").DataBodyRange.Cells(Target.row - ListObjects("Corp").DataBodyRange.row + 1, 8)
        Set cellarray(9) = ListObjects("Corp").DataBodyRange.Cells(Target.row - ListObjects("Corp").DataBodyRange.row + 1, 12)
        ShowAlterForm cellarray
    End If
End Sub
Private Sub ShowAlterForm(cellarray() As Variant)
    Dim frm As New frmCorp_Alter
    Dim rowIndex As Long
    rowIndex = cellarray(1).row - ListObjects("Corp").HeaderRowRange.row
    frm.LoadData cellarray(), rowIndex
    frm.Show
End Sub
