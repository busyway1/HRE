VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDate 
   Caption         =   "결산연월 설정"
   ClientHeight    =   3345
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5712
   OleObjectBlob   =   "frmDate.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private isNormalClose As Boolean
Private Sub UserForm_Initialize()
    Dim i As Integer
    Dim currentYear As Integer
    Dim tbl As ListObject
    
    On Error Resume Next
    Me.Caption = AppName & " " & AppType
    If Check.Cells(12, 4).Value <> "Complete" Or Check.Cells(13, 4).Value <> "Complete" Then
        GoEnd "이전 단계를 완료해주세요!"
    End If
    With Check.Cells(14, 4)
        .Value = "In Progress"
        .Interior.Color = RGB(255, 235, 156)
        .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
        .Offset(0, 2).Value = GetUserInfo()
    End With
    
    
    Set tbl = HideSheet.ListObjects("결산연월")
    
    currentYear = Year(Date)
    For i = currentYear + 1 To currentYear - 5 Step -1  ' 내림차순으로 변경
        Me.Year_Combo.AddItem i
    Next i
    
    For i = 1 To 12
        Me.Month_Combo.AddItem Format(i, "00")
    Next i
    
    If Not tbl.DataBodyRange Is Nothing Then
        Me.Year_Combo.Value = tbl.DataBodyRange.Cells(1, 1).Value
        Me.Month_Combo.Value = Format(tbl.DataBodyRange.Cells(1, 2).Value, "00")
    Else
        Me.Year_Combo.Value = currentYear
        Me.Month_Combo.Value = Format(month(Date), "00")
    End If
    
    Set tbl = Nothing
End Sub
Private Sub Confirm_Cmd_Click()
    Dim tbl As ListObject
    
    'On Error Resume Next
    
    Set tbl = HideSheet.ListObjects("결산연월")
    
    ' 테이블이 비어있으면 새 행 추가, 아니면 첫 행 업데이트
    If tbl.DataBodyRange Is Nothing Then
        tbl.ListRows.Add
    End If
    tbl.DataBodyRange.Cells(1, 1).Value = Me.Year_Combo.Value
    tbl.DataBodyRange.Cells(1, 2).Value = Me.Month_Combo.Value
    
    
    ' 파워쿼리 새로 고침(우선 너무 오래걸리니까 없앰)
    'Call QueryRefresh
    
    With Check.Cells(14, 4)
        .Value = "Complete"
        .Interior.Color = RGB(198, 239, 206)
        .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
        .Offset(0, 2).Value = GetUserInfo()
        .Offset(0, 3).Value = HideSheet.ListObjects("결산연월").DataBodyRange.Cells(1, 1).Value & "-" & Format(HideSheet.ListObjects("결산연월").DataBodyRange.Cells(1, 2).Value, "00")
    End With
    isNormalClose = True
    Msg "결산연월이 설정되었습니다.", vbInformation
    Unload Me
    Set tbl = Nothing
End Sub
Private Sub Cancel_Cmd_Click()
    Unload Me
    With Check.Cells(14, 4)
        .Value = "Not Started"
        .Interior.Color = RGB(255, 199, 206)
    End With
    GoEnd
End Sub
Private Sub Year_Combo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, Me.Year_Combo
End Sub
Private Sub Month_Combo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, Me.Month_Combo
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    UnhookListBoxScroll
    If Not isNormalClose Then
        With Check.Cells(14, 4)
            .Value = "Not Started"
            .Interior.Color = RGB(255, 199, 206)
        End With
    End If
    GoEnd
End Sub
