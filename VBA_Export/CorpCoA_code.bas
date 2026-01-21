VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CorpCoA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Not Intersect(Target, ListObjects("Raw_CoA").DataBodyRange) Is Nothing Then
        If Target.Interior.Color <> vbYellow Then
            Cancel = True
            Dim cellarray(1 To 6) As Variant
            Set cellarray(1) = ListObjects("Raw_CoA").DataBodyRange.Cells(Target.row - ListObjects("Raw_CoA").DataBodyRange.row + 1, 1)
            Set cellarray(2) = ListObjects("Raw_CoA").DataBodyRange.Cells(Target.row - ListObjects("Raw_CoA").DataBodyRange.row + 1, 2)
            Set cellarray(3) = ListObjects("Raw_CoA").DataBodyRange.Cells(Target.row - ListObjects("Raw_CoA").DataBodyRange.row + 1, 3)
            Set cellarray(4) = ListObjects("Raw_CoA").DataBodyRange.Cells(Target.row - ListObjects("Raw_CoA").DataBodyRange.row + 1, 4)
            Set cellarray(5) = ListObjects("Raw_CoA").DataBodyRange.Cells(Target.row - ListObjects("Raw_CoA").DataBodyRange.row + 1, 5)
            Set cellarray(6) = ListObjects("Raw_CoA").DataBodyRange.Cells(Target.row - ListObjects("Raw_CoA").DataBodyRange.row + 1, 6)
            ShowAlterForm cellarray
        End If
    End If
End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    If Not Intersect(Target, ListObjects("Raw_CoA").DataBodyRange) Is Nothing Then
        If Target.Interior.Color <> vbYellow Then
            Cancel = True
            Dim cellarray(1 To 6) As Variant
            Set cellarray(1) = ListObjects("Raw_CoA").DataBodyRange.Cells(Target.row - ListObjects("Raw_CoA").DataBodyRange.row + 1, 1)
            Set cellarray(2) = ListObjects("Raw_CoA").DataBodyRange.Cells(Target.row - ListObjects("Raw_CoA").DataBodyRange.row + 1, 2)
            Set cellarray(3) = ListObjects("Raw_CoA").DataBodyRange.Cells(Target.row - ListObjects("Raw_CoA").DataBodyRange.row + 1, 3)
            Set cellarray(4) = ListObjects("Raw_CoA").DataBodyRange.Cells(Target.row - ListObjects("Raw_CoA").DataBodyRange.row + 1, 4)
            Set cellarray(5) = ListObjects("Raw_CoA").DataBodyRange.Cells(Target.row - ListObjects("Raw_CoA").DataBodyRange.row + 1, 5)
            Set cellarray(6) = ListObjects("Raw_CoA").DataBodyRange.Cells(Target.row - ListObjects("Raw_CoA").DataBodyRange.row + 1, 6)
            ShowDeleteForm cellarray
        End If
    End If
End Sub

Private Sub ShowAlterForm(cellarray() As Variant)
    Dim frm As New frmCoA_Alter
    Dim rowIndex As Long
    rowIndex = cellarray(1).row - ListObjects("Raw_CoA").HeaderRowRange.row
    frm.LoadData cellarray(), rowIndex
    frm.Show
End Sub
Private Sub ShowDeleteForm(cellarray() As Variant)
    Dim frm As New frmCoA_Delete
    Dim rowIndex As Long
    rowIndex = cellarray(1).row - ListObjects("Raw_CoA").HeaderRowRange.row
    frm.LoadData cellarray(), rowIndex
    frm.Show
End Sub
