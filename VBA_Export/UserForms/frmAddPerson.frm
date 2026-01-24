VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddPerson 
   Caption         =   "UserForm1"
   ClientHeight    =   4140
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5808
   OleObjectBlob   =   "frmAddPerson.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmAddPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub UserForm_Initialize()
    Me.Caption = AppName & " " & AppType
    ' TextBox 초기화
    txtName.text = ""
    txtClass.text = ""
    txtEtc.text = ""
End Sub
Private Sub btnConfirm_Click()
   Dim tbl As ListObject
   Set tbl = HideSheet.ListObjects("People_Work")

   ' 입력값 검증
   If Me.txtName.text = "" Or Me.txtClass.text = "" Then
       Msg "성명과 직급은 필수로 기재해야합니다.", vbExclamation
       Exit Sub
   End If

   ' 동일인 검증
   Dim row As ListRow
   For Each row In tbl.ListRows
       If row.Range(1).Value = Me.txtName.text And row.Range(2).Value = Me.txtClass.text Then
           Msg "이미 등록된 담당자입니다.", vbExclamation
           Exit Sub
       End If
   Next row

   ' 새 담당자 추가
    With tbl.ListRows.Add
        .Range(1) = Me.txtName.text
        .Range(2) = Me.txtClass.text
        .Range(3) = Me.txtEtc.text
    End With

   Set tbl = Nothing: Set row = Nothing
   Unload Me
End Sub
Private Sub btnCancel_Click()
    Unload Me
End Sub
