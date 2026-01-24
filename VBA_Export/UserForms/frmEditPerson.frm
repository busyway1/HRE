VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEditPerson 
   Caption         =   "UserForm1"
   ClientHeight    =   4155
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5820
   OleObjectBlob   =   "frmEditPerson.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmEditPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SelectedIndex As Long
Private Sub UserForm_Initialize()
    Me.Caption = AppName & " " & AppType
End Sub
Private Sub btnConfirm_Click()
    Dim tbl As ListObject
    Set tbl = HideSheet.ListObjects("People_Work")
    
    ' 입력값 검증
    If txtName.text = "" Or txtClass.text = "" Then
        Msg "성명과 직급은 필수로 기재해야합니다.", vbExclamation
        Exit Sub
    End If
    
    ' 선택된 행의 데이터 수정
    With tbl.ListRows(SelectedIndex + 1)
        .Range(1) = txtName.text
        .Range(2) = txtClass.text
        .Range(3) = txtEtc.text
    End With
    
    Set tbl = Nothing
    Unload Me
End Sub
Private Sub btnCancel_Click()
    Unload Me
End Sub
