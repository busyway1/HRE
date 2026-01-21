VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BSPL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Not Intersect(Target, ListObjects("PTB").DataBodyRange) Is Nothing Then
        If Target.Interior.Color = vbYellow Then
            Cancel = True
            Dim cellarray(1 To 3) As Variant
            Set cellarray(1) = ListObjects("PTB").DataBodyRange.Cells(Target.row - ListObjects("PTB").DataBodyRange.row + 1, 1)
            Set cellarray(2) = ListObjects("PTB").DataBodyRange.Cells(Target.row - ListObjects("PTB").DataBodyRange.row + 1, 2)
            Set cellarray(3) = ListObjects("PTB").DataBodyRange.Cells(Target.row - ListObjects("PTB").DataBodyRange.row + 1, 3)
            ShowEditForm cellarray
        End If
    End If
End Sub

Private Sub ShowEditForm(cellarray() As Variant)
    Dim frm As New frmCoA_Update
    Dim rowIndex As Long
    rowIndex = cellarray(1).row - ListObjects("PTB").HeaderRowRange.row
    frm.LoadData cellarray(), rowIndex
    frm.Show
End Sub

