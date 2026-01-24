VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmScope 
   Caption         =   "Scope 설정"
   ClientHeight    =   6030
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5472
   OleObjectBlob   =   "frmScope.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private isNormalClose As Boolean

Private Sub UserForm_Initialize()
   On Error Resume Next
   Dim i As Integer
   Dim currentYear As Integer
   Dim tbl As ListObject
   Dim tblWho As ListObject
   Dim tblFS As ListObject
   Me.Caption = AppName & " " & AppType
   Me.cmbPeople.Style = fmStyleDropDownList
   If Check.Cells(12, 4).Value <> "Complete" Or Check.Cells(13, 4).Value <> "Complete" Or Check.Cells(14, 4).Value <> "Complete" Then
       GoEnd "이전 단계를 완료해주세요!"
   End If
   With Check.Cells(16, 4)
       .Value = "In Progress"
       .Interior.Color = RGB(255, 235, 156)
       .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
       .Offset(0, 2).Value = GetUserInfo()
       .Offset(0, 3).ClearContents
   End With
   
   Set tbl = CorpMaster.ListObjects("Mode")
   Set tblFS = CorpMaster.ListObjects("FS")
   Set tblWho = CorpMaster.ListObjects("Who")
   
   If tbl.DataBodyRange.Cells(1, 1).Value = "Full-Scope" Then
       Me.full_Option = True
       LoadPeopleList False
   Else
       Me.part_Option = True
       LoadPeopleList True
   End If
   
   If tblFS.DataBodyRange.Cells(1, 1).Value = "별도" Then
       Me.Sep_Option = True
   Else
       Me.Con_Option = True
   End If
   
   Me.cmbPeople.Value = tblWho.DataBodyRange.Cells(1, 1).Value
   Me.lblValue = Range("SmallBusiness").Value
   
   Set tbl = Nothing: Set tblFS = Nothing: Set tblWho = Nothing
End Sub

Private Sub LoadPeopleList(ByVal isPartOption As Boolean)
   Me.cmbPeople.Clear
   Me.cmbPeople.AddItem "총괄"
   
   If Not isPartOption Then
       Dim cell As Range
       For Each cell In Range("People_Work[담당자 성명]")
           Me.cmbPeople.AddItem cell.Value
       Next cell
   End If
End Sub

Private Sub part_Option_Click()
   If Me.lblValue = "Yes" Then
       If Me.Sep_Option Then
           LoadPeopleList True
           Me.cmbPeople.Value = "총괄"
       Else
           Me.part_Option = False
           Me.full_Option = True
           LoadPeopleList False
       End If
   ElseIf Me.lblValue = "No" Then
       If Me.Sep_Option Then
           LoadPeopleList True
           Me.cmbPeople.Value = "총괄"
       Else
           Me.part_Option = False
           Me.full_Option = True
           LoadPeopleList False
       End If
   End If
End Sub

Private Sub full_Option_Click()
   If Me.lblValue = "Yes" Then
       If Me.Con_Option Then
           LoadPeopleList False
       Else
           Me.full_Option = False
           Me.part_Option = True
           LoadPeopleList True
           Me.cmbPeople.Value = "총괄"
       End If
   ElseIf Me.lblValue = "No" Then
       If Me.Con_Option Then
           LoadPeopleList False
       Else
           Me.full_Option = False
           Me.part_Option = True
           LoadPeopleList True
           Me.cmbPeople.Value = "총괄"
       End If
   End If
End Sub

Private Sub Sep_Option_Click()
   If Me.lblValue = "Yes" Or Me.lblValue = "No" Then
       Me.part_Option = True
       Me.full_Option = False
       LoadPeopleList True
       Me.cmbPeople.Value = "총괄"
   End If
End Sub

Private Sub Con_Option_Click()
   If Me.lblValue = "Yes" Or Me.lblValue = "No" Then
       Me.full_Option = True
       Me.part_Option = False
       LoadPeopleList False
   End If
End Sub

Private Sub Confirm_Cmd_Click()
   On Error Resume Next
   Dim tbl As ListObject
   Dim tblWho As ListObject
   Dim tblFS As ListObject
   Set tbl = CorpMaster.ListObjects("Mode")
   Set tblWho = CorpMaster.ListObjects("Who")
   Set tblFS = CorpMaster.ListObjects("FS")
   Call SpeedUp
   
   If Me.full_Option = True Then
       tbl.DataBodyRange.Cells(1, 1).Value = "Full-Scope"
   ElseIf Me.part_Option = True Then
       tbl.DataBodyRange.Cells(1, 1).Value = "지분법"
   End If
   
   If Me.Sep_Option = True Then
       tblFS.DataBodyRange.Cells(1, 1).Value = "별도"
   ElseIf Me.Con_Option = True Then
       tblFS.DataBodyRange.Cells(1, 1).Value = "연결"
   End If
   
   If Me.cmbPeople.Value = "" Then
       Msg "담당자를 선택해주세요.", vbExclamation
       Call SpeedDown
       isNormalClose = False
       Exit Sub
   End If
   
   tblWho.DataBodyRange.Cells(1, 1).Value = Me.cmbPeople.Value
   HideSheet.Range("U2").Value = Me.cmbPeople.Value
   
   Me.Hide
   
   Call OpenProgress("SPO에 접근 중...")
   Dim qt As QueryTable
   Dim qt_AD As QueryTable
   Set qt = HideSheet.ListObjects("Link").QueryTable
   Set qt_AD = HideSheet.ListObjects("Link_취득_처분").QueryTable
   Call CalculateProgress(0.5, "Scope 변경 반영 중...")
   qt.Refresh BackgroundQuery:=False
   qt_AD.Refresh BackgroundQuery:=False
   Application.CalculateUntilAsyncQueriesDone
   Call CalculateProgress(1, "작업 완료")
   Application.Calculation = xlCalculationAutomatic
   With Check.Cells(16, 4)
       .Value = "Complete"
       .Interior.Color = RGB(198, 239, 206)
       .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
       .Offset(0, 2).Value = GetUserInfo()
       .Offset(0, 3).Value = CorpMaster.ListObjects("FS").DataBodyRange.Cells(1, 1).Value & ", " & CorpMaster.ListObjects("Mode").DataBodyRange.Cells(1, 1).Value & ", " & CorpMaster.ListObjects("Who").DataBodyRange.Cells(1, 1).Value & ", " & Application.WorksheetFunction.CountIf(CorpMaster.ListObjects("Corp").ListColumns(11).DataBodyRange, "O") & "개사"
   End With
   
   isNormalClose = True
   Msg "Scope가 설정되었습니다.", vbInformation
   Call SpeedDown
   Set tbl = Nothing: Set tblWho = Nothing: Set qt = Nothing: Set qt_AD = Nothing: Set tblFS = Nothing
End Sub

Private Sub Cancel_Cmd_Click()
   On Error Resume Next
   With Check.Cells(16, 4)
       .Value = "Not Started"
       .Interior.Color = RGB(255, 199, 206)
       .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
       .Offset(0, 2).Value = GetUserInfo()
       .Offset(0, 3).ClearContents
   End With
   GoEnd
   Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
   On Error Resume Next
   UnhookListBoxScroll
   If Not isNormalClose Then
      With Check.Cells(16, 4)
          .Value = "Not Started"
          .Interior.Color = RGB(255, 199, 206)
          .Offset(0, 1).Value = Format(Now(), "yyyy-mm-dd hh:mm")
          .Offset(0, 2).Value = GetUserInfo()
          .Offset(0, 3).ClearContents
      End With
      GoEnd
   End If
End Sub

Private Sub cmbPeople_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   HookListBoxScroll Me, Me.cmbPeople
End Sub
