VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AddCoA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim cell As Range
    On Error Resume Next
    Call SpeedUp
    For Each cell In Target.Cells
        If Not cell.Validation Is Nothing Then
            If Not cell.Validation.Value Then
                cell.ClearContents
            End If
        End If
    Next cell
    Call SpeedDown
    Set cell = Nothing
End Sub

