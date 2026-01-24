VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSPO 
   Caption         =   "SPO 메인 주소 설정"
   ClientHeight    =   3405
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   8628.001
   OleObjectBlob   =   "frmSPO.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmSPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private isNormalClose As Boolean
Private Sub CommandButton1_Click()
    On Error Resume Next
    HideSheet.Range("E2").Value = Me.TextBox1.Value
    isNormalClose = True
    Unload Me
    With Check.Cells(12, 4)
        .Value = "Complete"
        .Interior.Color = RGB(198, 239, 206)
        .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
        .Offset(0, 2).Value = GetUserInfo()
    End With
    MsgBox "SPO 메인 주소가 기록되었습니다.", vbInformation, "주소 기록"
End Sub
Private Sub CommandButton2_Click()
    On Error Resume Next
    Unload Me
    With Check.Cells(12, 4)
        .Value = "Not Started"
        .Interior.Color = RGB(255, 199, 206)
        .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
        .Offset(0, 2).Value = GetUserInfo()
    End With
    GoEnd
End Sub
Private Sub UserForm_Initialize()
    On Error Resume Next
    Me.Caption = AppName & " " & AppType
    With Check.Cells(12, 4)
        .Value = "In Progress"
        .Interior.Color = RGB(255, 235, 156)
        .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
        .Offset(0, 2).Value = GetUserInfo()
    End With
    Me.TextBox1.text = HideSheet.Range("E2").Value
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    If Not isNormalClose Then
       With Check.Cells(12, 4)
           .Value = "Not Started"
           .Interior.Color = RGB(255, 199, 206)
           .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
           .Offset(0, 2).Value = GetUserInfo()
       End With
       GoEnd
    End If
End Sub
