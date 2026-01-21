VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPeople 
   ClientHeight    =   5085
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   6216
   OleObjectBlob   =   "frmPeople.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmPeople"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Unload Me
End Sub


Private Sub UserForm_Initialize()
    Dim tbl As ListObject
    Dim rng As Range
    On Error Resume Next
    Me.Caption = AppName & " " & AppType
    Set tbl = HideSheet.ListObjects("People_Work")
    Set rng = tbl.DataBodyRange
    
    ' 리스트박스 초기화
    lbPeople.Clear
    
    If HideSheet.Range("U2").Value = "" Then
        lblName.Caption = "없음"
    Else
        lblName.Caption = HideSheet.Range("U2").Value
    End If
    
    ' 테이블의 데이터를 리스트박스에 추가
    With lbPeople
        .ColumnCount = tbl.ListColumns.count
        .ColumnHeads = False
        .List = tbl.DataBodyRange.Value
    End With
    Set tbl = Nothing
End Sub
Private Sub btnAdd_Click()
    frmAddPerson.Show
    RefreshListBox
End Sub
Private Sub btnDelete_Click()
    Dim tbl As ListObject
    Set tbl = HideSheet.ListObjects("People_Work")
    If lbPeople.ListIndex = -1 Then
        Msg "삭제할 담당자를 선택해주세요.", vbExclamation
        Exit Sub
    End If
    If tbl.ListRows.count = 1 Then
        Msg "담당자는 최소 1명이 있어야 합니다.", vbExclamation
        Exit Sub
    End If
    If Msg("선택한 담당자를 삭제하시겠습니까?", vbQuestion + vbYesNo) = vbYes Then
        tbl.ListRows(lbPeople.ListIndex + 1).Delete
        RefreshListBox
    End If
    Set tbl = Nothing
End Sub
Public Sub RefreshListBox()
    Dim tbl As ListObject
    Set tbl = HideSheet.ListObjects("People_Work")
    lbPeople.List = tbl.DataBodyRange.Value
    Set tbl = Nothing
End Sub
Private Sub btnEdit_Click()
    If lbPeople.ListIndex = -1 Then
        Msg "수정할 담당자를 선택해주세요.", vbExclamation
        Exit Sub
    End If
    ' 선택된 데이터를 수정 폼으로 전달
    With frmEditPerson
        .txtName.text = lbPeople.List(lbPeople.ListIndex, 0)
        .txtClass.text = lbPeople.List(lbPeople.ListIndex, 1)
        .txtEtc.text = lbPeople.List(lbPeople.ListIndex, 2)
        .SelectedIndex = lbPeople.ListIndex
        .Show
    End With
    RefreshListBox
End Sub
Private Sub lbPeople_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, Me.lbPeople
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    UnhookListBoxScroll
End Sub

