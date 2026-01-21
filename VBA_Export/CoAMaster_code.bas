VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoAMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Not Intersect(Target, ListObjects("Master").DataBodyRange) Is Nothing Then
        If Target.Interior.Color <> vbYellow Then
            Cancel = True
            
            Dim col As Long
            Dim i As Long
            col = ListObjects("Master").ListColumns.count
            
            Dim cellarray() As Variant
            ReDim cellarray(1 To 2)
            
            Set cellarray(1) = ListObjects("Master").DataBodyRange.Cells(Target.row - ListObjects("Master").DataBodyRange.row + 1, 1)
            Set cellarray(2) = ListObjects("Master").DataBodyRange.Cells(Target.row - ListObjects("Master").DataBodyRange.row + 1, 2)
            
            
            Dim colIndex As Long
            colIndex = Target.Column - ListObjects("Master").Range.Column + 1
            
             If ListObjects("Master").HeaderRowRange(colIndex).Value = "TB Account" Or _
               ListObjects("Master").HeaderRowRange(colIndex).Value = "Account Name" Or _
               ListObjects("Master").HeaderRowRange(colIndex).Value = "�ݾ�" Then
            
                MsgBox "�ش� ���� ������ �� �����ϴ�!", vbCritical, "����"
                Exit Sub
            End If
        
            ShowMasterForm_Alter cellarray, colIndex
        End If
        
        
    ElseIf Not Intersect(Target, ListObjects("Master").HeaderRowRange) Is Nothing Then
        If Target.Interior.Color = RGB(217, 217, 217) Then
            Cancel = True
            
            Dim selectcol As Long
            selectcol = Target.Column - ListObjects("Master").Range.Column + 1
            
            
            If ListObjects("Master").HeaderRowRange(selectcol).Value = "TB Account" Or _
               ListObjects("Master").HeaderRowRange(selectcol).Value = "Account Name" Or _
               ListObjects("Master").HeaderRowRange(selectcol).Value = "BSPL" Or _
               ListObjects("Master").HeaderRowRange(selectcol).Value = "��з�" Or _
               ListObjects("Master").HeaderRowRange(selectcol).Value = "�ߺз�" Or _
               ListObjects("Master").HeaderRowRange(selectcol).Value = "�Һз�" Or _
               ListObjects("Master").HeaderRowRange(selectcol).Value = "���ð���" Or _
               ListObjects("Master").HeaderRowRange(selectcol).Value = "�׷�� ������" Or _
               ListObjects("Master").HeaderRowRange(selectcol).Value = "��ȣ" Or _
               ListObjects("Master").HeaderRowRange(selectcol).Value = "�ݾ�" Then

                MsgBox "�ش� �� ���̿� ���� �߰��� �� �����ϴ�!", vbCritical, "����"
                Exit Sub
            End If
            
        End If
    End If
End Sub
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    If Not Intersect(Target, ListObjects("Master").HeaderRowRange) Is Nothing Then
        If Target.Interior.Color = RGB(217, 217, 217) Then
            Cancel = True
            
            Dim selectcol As Long
            selectcol = Target.Column - ListObjects("Master").Range.Column + 1
            
            If ListObjects("Master").HeaderRowRange(selectcol).Value = "TB Account" Or _
               ListObjects("Master").HeaderRowRange(selectcol).Value = "Account Name" Or _
               ListObjects("Master").HeaderRowRange(selectcol).Value = "BSPL" Or _
               ListObjects("Master").HeaderRowRange(selectcol).Value = "��з�" Or _
               ListObjects("Master").HeaderRowRange(selectcol).Value = "�ߺз�" Or _
               ListObjects("Master").HeaderRowRange(selectcol).Value = "�Һз�" Or _
               ListObjects("Master").HeaderRowRange(selectcol).Value = "���ð���" Or _
               ListObjects("Master").HeaderRowRange(selectcol).Value = "�׷�� ������" Or _
               ListObjects("Master").HeaderRowRange(selectcol).Value = "��ȣ" Or _
               ListObjects("Master").HeaderRowRange(selectcol).Value = "�ݾ�" Or _
               ListObjects("Master").HeaderRowRange(selectcol).Value = "Util" Then
            
                MsgBox "�ش� ���� ������ �� �����ϴ�!", vbCritical, "����"
                Exit Sub
            End If
            
        End If
    End If
End Sub

Private Sub ShowMasterForm_Alter(cellarray() As Variant, colIndex As Long)
    Dim frm As New frmMaster_Alter
    Dim rowIndex As Long
    rowIndex = cellarray(1).row - ListObjects("Master").HeaderRowRange.row
    frm.LoadData cellarray(), rowIndex, colIndex
    frm.Show
End Sub
